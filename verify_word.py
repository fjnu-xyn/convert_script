"""
验证生成的 Word 文档是否与 Excel 源文件内容一致
"""

import pandas as pd
from docx import Document
from pathlib import Path
from logger import get_logger

logger = get_logger("verify_word")


def read_excel_robust(excel_path):
    """
    自动查找正确的Sheet和表头
    (与 converter 保持一致的逻辑)
    """
    try:
        xl = pd.ExcelFile(excel_path)
    except Exception as e:
        print(f"无法打开Excel文件: {e}")
        logger.error(f"无法打开Excel文件: {e}")
        return None

    # 1. 查找包含数据的Sheet
    target_sheet = None
    for sheet in xl.sheet_names:
        if '拆分表' in sheet or '功能点' in sheet:
            target_sheet = sheet
            break
    
    if not target_sheet:
        target_sheet = xl.sheet_names[0]

    # 2. 查找表头行
    df_preview = pd.read_excel(excel_path, sheet_name=target_sheet, header=None, nrows=10)
    
    header_row_idx = -1
    for idx, row in df_preview.iterrows():
        row_str = row.astype(str).values
        if any('一级模块' in s for s in row_str) and any('二级模块' in s for s in row_str):
            header_row_idx = idx
            break
    
    if header_row_idx == -1:
        header_row_idx = 0

    # 3. 读取完整数据
    df = pd.read_excel(excel_path, sheet_name=target_sheet, header=header_row_idx)
    return df


def extract_excel_processes(excel_path):
    """从 Excel 提取功能过程列表，并构建 三级模块 -> 功能过程 映射"""
    df = read_excel_robust(excel_path)
    if df is None:
        return [], [], {}

    col_map = {}
    required_cols = {
        'CustomerReq': ['客户需求'],
        'Level1': ['一级模块'],
        'Level2': ['二级模块'],
        'Level3': ['三级模块'],
        'Process': ['功能过程', '功能名称'],
        'Description': ['子过程描述', '功能描述']
    }

    for col_idx, col_name in enumerate(df.columns):
        col_str = str(col_name).strip()
        for key, keywords in required_cols.items():
            if key not in col_map and any(kw in col_str for kw in keywords):
                col_map[key] = col_name

    # 回退索引
    if 'CustomerReq' not in col_map and len(df.columns) > 0:
        col_map['CustomerReq'] = df.columns[0]
    if 'Level1' not in col_map and len(df.columns) > 1:
        col_map['Level1'] = df.columns[1]
    if 'Level2' not in col_map and len(df.columns) > 2:
        col_map['Level2'] = df.columns[2]
    if 'Level3' not in col_map and len(df.columns) > 3:
        col_map['Level3'] = df.columns[3]
    if 'Process' not in col_map and len(df.columns) > 6:
        col_map['Process'] = df.columns[6]
    if 'Description' not in col_map and len(df.columns) > 7:
        col_map['Description'] = df.columns[7]

    if 'Process' not in col_map:
        print("无法在Excel中找到'功能过程'列")
        logger.error("无法在Excel中找到'功能过程'列")
        return [], [], {}

    processes = []
    subprocess_data = []
    level3_map = {}

    customer_req_col = col_map.get('CustomerReq')
    process_col = col_map['Process']
    desc_col = col_map.get('Description')
    l1_col = col_map.get('Level1')
    l2_col = col_map.get('Level2')
    l3_col = col_map.get('Level3')

    invalid_keywords = ['呈现', '查询', '保存', '输入', '校验', '输出']
    df.loc[df[process_col].isin(invalid_keywords), process_col] = None

    cols_to_fill = []
    if customer_req_col: cols_to_fill.append(customer_req_col)
    if l1_col: cols_to_fill.append(l1_col)
    if l2_col: cols_to_fill.append(l2_col)
    if l3_col: cols_to_fill.append(l3_col)
    cols_to_fill.append(process_col)
    df[cols_to_fill] = df[cols_to_fill].ffill()

    if customer_req_col and l1_col and l2_col and l3_col:
        grouped_indices = []
        for _, group in df.groupby([customer_req_col, l1_col, l2_col, l3_col], sort=False):
            grouped_indices.extend(group.index.tolist())
        df = df.loc[grouped_indices]

    last_added_process = None
    for _, row in df.iterrows():
        process_name = row[process_col]
        subprocess_desc = row[desc_col] if desc_col else ""
        level3_value = row[l3_col] if l3_col else "未定义三级模块"

        if pd.isna(process_name) and pd.isna(subprocess_desc):
            continue

        process_name_str = str(process_name).strip() if not pd.isna(process_name) else ""
        subprocess_desc_str = str(subprocess_desc).strip() if not pd.isna(subprocess_desc) else ""
        level3_str = str(level3_value).strip() if not pd.isna(level3_value) else "未定义三级模块"

        subprocess_data.append({
            'process': process_name_str,
            'description': subprocess_desc_str,
            'level3': level3_str,
            'is_keyword': False
        })

        if process_name_str and process_name_str not in invalid_keywords:
            if last_added_process != process_name_str:
                processes.append(process_name_str)
                last_added_process = process_name_str
                level3_map.setdefault(level3_str, []).append(process_name_str)
            else:
                # 同一功能过程的重复行，不再追加到总列表，但仍确保映射已存在
                if level3_str not in level3_map:
                    level3_map[level3_str] = [process_name_str]
    return processes, subprocess_data, level3_map


def extract_word_content(word_path):
    """从 Word 提取结构化内容。
    需求：Heading 5 作为三级模块，输出时严格保持文档出现顺序，
    同名模块不合并（每次出现视为一个独立模块实例）。
    功能过程格式：段落前缀 编号.名称。
    返回: summary_line, processes_list, level3_modules_list
        processes_list: 所有功能过程对象列表
        level3_modules_list: 按出现顺序的模块实例列表，每项 {'module': 原始标题文本, 'processes': [功能过程名,...]}
    """
    doc = Document(word_path)

    summary_line = ""
    processes = []
    current_process = None
    current_level1 = None
    current_level2 = None
    current_level3 = None  # 对应 Heading 5
    in_function_desc = False  # 是否在"功能描述"标题区域内
    # 按出现顺序记录的三级模块实例列表
    level3_modules = []
    current_level3_process_bucket = None  # 指向 level3_modules 当前模块的 processes 列表

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name
        if style_name.startswith('Heading'):
            if style_name == 'Heading 1':
                current_level1 = text
                current_level2 = None
                current_level3 = None
                current_level3_process_bucket = None
                in_function_desc = False
            elif style_name == 'Heading 2':
                current_level2 = text
                current_level3 = None
                current_level3_process_bucket = None
                in_function_desc = False
            elif style_name == 'Heading 5':
                current_level3 = text
                # 新的三级模块实例
                level3_modules.append({'module': current_level3, 'processes': []})
                current_level3_process_bucket = level3_modules[-1]['processes']
                in_function_desc = False
            elif style_name == 'Heading 6':
                # Heading 6 包含"功能描述"和"关键时序图/业务逻辑图"
                if '功能描述' in text:
                    in_function_desc = True
                else:
                    in_function_desc = False
            continue

        # 跳过整体功能列表行
        if text.startswith("整体功能列表") or text.startswith("　整体功能列表"):
            continue
        
        # 跳过"无。"行
        if text == "无。":
            continue

        # 只在"功能描述"区域内处理功能过程
        if in_function_desc:
            # 检查是否是子过程描述的开头（输入-、查询-等）
            is_subprocess_line = any(text.startswith(prefix) for prefix in ['输入-', '查询-', '呈现-', '校验-', '输出-', '处理-', '保存-'])
            
            # 识别带编号的功能过程：格式为 "数字.功能名称"
            is_numbered_process = False
            process_name = text
            if not is_subprocess_line and text and text[0].isdigit() and '.' in text[:4]:
                parts = text.split('.', 1)
                if len(parts) == 2 and parts[0].isdigit():
                    is_numbered_process = True
                    process_name = parts[1].strip()
            
            # 如果是带编号的功能过程行
            if is_numbered_process:
                # 结束当前功能过程
                if current_process:
                    processes.append(current_process)
                # 开始新的功能过程
                current_process = {
                    'name': process_name,
                    'details': [],
                    'level3': current_level3 or '未定义三级模块'
                }
                # 推入当前模块实例的 process 列表
                if current_level3_process_bucket is not None:
                    current_level3_process_bucket.append(process_name)
            elif not is_subprocess_line:
                # 普通段落但不是编号格式，可能是数据问题，暂时忽略或作为子过程
                if current_process is not None:
                    current_process['details'].append(text)
            else:
                # 是子过程描述行，添加到当前功能过程的详情中
                if current_process is not None:
                    current_process['details'].append(text)

    if current_process:
        processes.append(current_process)

    return summary_line, processes, level3_modules


def build_detailed_stats(excel_path, word_path):
    """构建详细的模块统计数据
    返回格式：包含一级、二级、三级模块名称和数量，以及功能过程名称、数量和子过程数量
    """
    # 从 Excel 读取数据
    df = read_excel_robust(excel_path)
    if df is None:
        return []
    
    # 列映射
    col_map = {}
    required_cols = {
        'CustomerReq': ['客户需求'],
        'Level1': ['一级模块'],
        'Level2': ['二级模块'],
        'Level3': ['三级模块'],
        'Process': ['功能过程', '功能名称'],
        'Description': ['子过程描述', '功能描述']
    }
    
    for col_idx, col_name in enumerate(df.columns):
        col_str = str(col_name).strip()
        for key, keywords in required_cols.items():
            if key not in col_map and any(kw in col_str for kw in keywords):
                col_map[key] = col_name
    
    # 回退索引
    if 'CustomerReq' not in col_map and len(df.columns) > 0:
        col_map['CustomerReq'] = df.columns[0]
    if 'Level1' not in col_map and len(df.columns) > 1:
        col_map['Level1'] = df.columns[1]
    if 'Level2' not in col_map and len(df.columns) > 2:
        col_map['Level2'] = df.columns[2]
    if 'Level3' not in col_map and len(df.columns) > 3:
        col_map['Level3'] = df.columns[3]
    if 'Process' not in col_map and len(df.columns) > 6:
        col_map['Process'] = df.columns[6]
    if 'Description' not in col_map and len(df.columns) > 7:
        col_map['Description'] = df.columns[7]
    
    if 'Process' not in col_map:
        return []
    
    process_col = col_map['Process']
    desc_col = col_map.get('Description')
    l1_col = col_map.get('Level1')
    l2_col = col_map.get('Level2')
    l3_col = col_map.get('Level3')
    
    # 清理无效关键字
    invalid_keywords = ['呈现', '查询', '保存', '输入', '校验', '输出']
    df.loc[df[process_col].isin(invalid_keywords), process_col] = None
    
    # 填充合并单元格
    cols_to_fill = []
    if l1_col: cols_to_fill.append(l1_col)
    if l2_col: cols_to_fill.append(l2_col)
    if l3_col: cols_to_fill.append(l3_col)
    cols_to_fill.append(process_col)
    df[cols_to_fill] = df[cols_to_fill].ffill()
    
    # 构建层级统计
    stats = []
    
    # 按一级、二级、三级、功能过程分组
    if l1_col and l2_col and l3_col:
        for (l1, l2, l3), group in df.groupby([l1_col, l2_col, l3_col], sort=False):
            if pd.isna(l1) or pd.isna(l2) or pd.isna(l3):
                continue
            
            # 统计该三级模块下的功能过程
            processes = group[process_col].dropna().unique()
            
            # 对每个功能过程，统计子过程数量
            for process in processes:
                process_rows = group[group[process_col] == process]
                # 子过程数 = 该功能过程的行数（每行一个子过程描述）
                subprocess_count = len(process_rows)
                
                # 获取子过程描述列表
                subprocesses = []
                if desc_col:
                    subprocesses = process_rows[desc_col].dropna().astype(str).tolist()
                
                subprocess_details = "\n".join([f"{i+1}. {s}" for i, s in enumerate(subprocesses)])
                
                stats.append({
                    '一级模块名称': str(l1).strip(),
                    '二级模块名称': str(l2).strip(),
                    '三级模块名称': str(l3).strip(),
                    '功能过程名称': str(process).strip(),
                    '子过程数量': subprocess_count,
                    '子过程详情': subprocess_details
                })
    
    return stats


def verify_consistency(excel_path, word_path):
    """验证 Excel 和 Word 的一致性，并返回详细统计数据
    """
    
    print("=" * 80)
    print("Word 文档内容验证")
    print("=" * 80)
    print()
    logger.info("=" * 80)
    logger.info("Word 文档内容验证")
    logger.info("=" * 80)
    
    # 提取 Excel 数据
    excel_processes, excel_details, _ = extract_excel_processes(excel_path)
    _, word_processes, word_level3_modules = extract_word_content(word_path)
    
    # 验证功能过程数量
    print(f"✓ Excel 功能过程数: {len(excel_processes)}")
    print(f"✓ Word 功能过程数: {len(word_processes)}")
    logger.info(f"✓ Excel 功能过程数: {len(excel_processes)}")
    logger.info(f"✓ Word 功能过程数: {len(word_processes)}")
    
    if len(excel_processes) == len(word_processes):
        print(f"✓ 功能过程数量一致")
        logger.info(f"✓ 功能过程数量一致")
    else:
        print(f"✗ 功能过程数量不一致!")
        logger.warning(f"✗ 功能过程数量不一致!")
        # return False # 继续对比以显示差异
    
    print()
    print("=" * 80)
    print("功能过程对比")
    print("=" * 80)
    logger.info("")
    logger.info("=" * 80)
    logger.info("功能过程对比")
    logger.info("=" * 80)
    
    all_match = True
    # 使用 zip_longest 防止长度不一致时漏掉
    from itertools import zip_longest
    
    for i, (excel_p, word_p) in enumerate(zip_longest(excel_processes, word_processes), 1):
        word_p_name = word_p['name'] if word_p else "MISSING"
        excel_p_name = excel_p if excel_p else "MISSING"
        
        match = excel_p_name == word_p_name
        symbol = "✓" if match else "✗"
        
        if not match:
            msg = f"{symbol} {i}. 不匹配!"
            print(msg)
            logger.info(msg)
            print(f"   Excel: {excel_p_name}")
            logger.info(f"   Excel: {excel_p_name}")
            print(f"   Word:  {word_p_name}")
            logger.info(f"   Word:  {word_p_name}")
            all_match = False
        else:
            if i <= 5 or (len(excel_processes) > 10 and i > len(excel_processes) - 5):
                msg = f"{symbol} {i}. {excel_p_name}"
                print(msg)
                logger.info(msg)
            elif i == 6:
                msg = f"   ... (中间 {len(excel_processes) - 10} 个过程)"
                print(msg)
                logger.info(msg)
    
    # 生成详细模块统计数据
    detailed_stats = build_detailed_stats(excel_path, word_path)
    
    print()
    print("=" * 80)
    if all_match:
        print("✓ 验证通过！Word 文档与 Excel 源文件完全一致")
        logger.info("✓ 验证通过！Word 文档与 Excel 源文件完全一致")
    else:
        print("✗ 验证失败！存在内容不一致")
        logger.warning("✗ 验证失败！存在内容不一致")
    print("=" * 80)
    logger.info("=" * 80)
    
    return all_match, detailed_stats


if __name__ == "__main__":
    logger.info("此脚本仅供 Web 服务内部调用，不支持直接命令行运行")
    logger.info("请通过 Streamlit 应用界面使用校对功能")
