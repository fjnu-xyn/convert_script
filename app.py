import streamlit as st
import os
import shutil
from pathlib import Path
import sys
import io
from contextlib import redirect_stdout
import time
import atexit
import threading
import pandas as pd  # 新增：用于导出统计到 Excel

# 导入转换脚本
# 确保当前目录在 sys.path 中
current_dir = Path(__file__).parent.resolve()
sys.path.append(str(current_dir))

import excel_to_word_converter
import verify_word
import styles
from cleanup_loop import run_loop

# 后台静默清理线程
@st.cache_resource(show_spinner=False)
def start_cleanup_daemon():
    """启动后台清理守护线程（全服务器单例）"""
    # 创建单个线程（进程退出时自动终止）
    daemon_thread = threading.Thread(target=run_loop, daemon=True)
    daemon_thread.start()
    return daemon_thread

def cleanup_files(*file_paths):
    """清理指定的文件"""
    for file_path in file_paths:
        try:
            if file_path and Path(file_path).exists():
                Path(file_path).unlink()
        except Exception as e:
            pass  # 静默失败，不会影响体验

def save_uploaded_file(uploaded_file, target_folder):
    try:
        # 生成带时间戳的唯一文件名
        timestamp = int(time.time() * 1000)
        file_stem = Path(uploaded_file.name).stem
        file_suffix = Path(uploaded_file.name).suffix
        unique_filename = f"{file_stem}_{timestamp}{file_suffix}"
        
        target_path = target_folder / unique_filename
        with open(target_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return target_path
    except Exception as e:
        st.error(f"保存文件失败: {e}")
        return None

def main():
    st.set_page_config(page_title="Excel 转 Word 工具", page_icon="📄", layout="wide")
    styles.load_css()
    
    # 启动后台清理守护线程
    start_cleanup_daemon()
    
    # 使用 session_state 存储模块统计数据和文件路径
    if 'module_stats' not in st.session_state:
        st.session_state.module_stats = []
    if 'current_files' not in st.session_state:
        st.session_state.current_files = {'excel': None, 'word': None}
    # 清理行为：上传新文件或移除上传时立即清理
    
    # 路径配置
    base_dir = Path(__file__).parent.resolve()
    input_dir = base_dir / 'excel_input'
    output_dir = base_dir / 'word_output'

    # 确保目录存在
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    # 左侧边栏：使用说明
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        1. 将 Excel 文件拖入右侧上传区域或直接上传文件。
        2. 点击 **开始转换** 按钮生成 Word 文档。
        3. 转换完成后，需点击 **执行内容校对** 检查一致性。
        4. 点击**下载 Word 文档**按钮下载转换后的 Word 文件。
        5. 校对后右侧会出现模块统计信息，可导出为 Excel 文件（注：**详细数据**在excel中的第二个sheet）。
        """)
    
    # 主内容区和右侧边栏布局
    main_col, stats_col = st.columns([3, 1])
    
    with main_col:
        st.title("COSMIC工具：Excel->Word ")
        st.markdown("---")
        
        # 文件上传区域
        uploaded_file = st.file_uploader("拖拽或选择 Excel 文件", type=['xlsx', 'xls'])

        if uploaded_file is not None:
            # 如果是新文件，清理旧文件
            current_upload_name = uploaded_file.name
            if 'last_upload_name' not in st.session_state or st.session_state.last_upload_name != current_upload_name:
                cleanup_files(st.session_state.current_files.get('excel'), st.session_state.current_files.get('word'))
                st.session_state.current_files = {'excel': None, 'word': None}
                # 新文件上传前清理旧文件，确保不会残留
                st.session_state.last_upload_name = current_upload_name
            
            # 保存文件（如果还没保存）
            if st.session_state.current_files['excel'] is None:
                saved_path = save_uploaded_file(uploaded_file, input_dir)
                if saved_path:
                    st.session_state.current_files['excel'] = str(saved_path)
            else:
                saved_path = Path(st.session_state.current_files['excel'])
            
            if saved_path:
                st.success(f"文件已上传: `{uploaded_file.name}`")
                
                word_filename = saved_path.stem + ".docx"
                word_path = output_dir / word_filename
                
                # 如果Word文件存在但不在记录中，更新记录
                if word_path.exists() and st.session_state.current_files['word'] is None:
                    st.session_state.current_files['word'] = str(word_path)
                
                # 文件下载区
                if word_path.exists():
                    st.markdown("### 📥 下载")
                    with open(word_path, "rb") as file:
                        download_clicked = st.download_button(
                            label="⬇ 下载 Word 文档",
                            data=file,
                            file_name=uploaded_file.name.replace(Path(uploaded_file.name).suffix, '.docx'),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    
                    # 下载后不会删除文件，允许重复下载
                    if download_clicked:
                        st.success("✅ 文件已下载，可重复下载。移除上传或上传新文件将自动清理。")
                    
                    st.markdown("---")
                
                # 操作按钮行
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    convert_clicked = st.button(" 开始转换", type="primary", use_container_width=True)
                with btn_col2:
                    verify_clicked = st.button(" 执行内容校对", use_container_width=True)
                
                # 转换处理
                if convert_clicked:
                    st.markdown("### ⏳ 处理日志")
                    
                    f = io.StringIO()
                    with redirect_stdout(f):
                        try:
                            excel_to_word_converter.excel_to_word(saved_path, word_path, perform_verify=False, open_output=False)
                        except Exception as e:
                            print(f"发生错误: {e}")
                    
                    log_output = f.getvalue()
                    st.code(log_output, language="text")
                    
                    if word_path.exists():
                        st.session_state.current_files['word'] = str(word_path)
                        st.success("✅ 转换成功！")
                        st.toast("转换完成")
                    else:
                        st.error("❌ 转换失败，未生成 Word 文件。")

                # 校对处理
                if verify_clicked:
                    if not word_path.exists():
                        st.warning("⚠️ 请先执行转换，生成 Word 文档后再进行校对。")
                    else:
                        st.markdown("### 📋 校对报告")
                        st.info("📌 校对说明：系统将对比服务器上的 Excel 源文件与生成的 Word 文档内容是否一致。")
                        f_verify = io.StringIO()
                        result = False
                        module_stats = []
                        with redirect_stdout(f_verify):
                            try:
                                result, module_stats = verify_word.verify_consistency(saved_path, word_path)
                            except Exception as e:
                                print(f"校对过程出错: {e}")
                        
                        # 保存到 session_state
                        st.session_state.module_stats = module_stats if module_stats else []
                        
                        verify_log = f_verify.getvalue()
                        
                        if result:
                            st.success("✅ 验证通过！Word 文档与 Excel 源文件内容一致。")
                        else:
                            st.error("❌ 验证失败！发现内容不一致，请查看下方详情。")
                            
                        with st.expander("查看详细校对日志", expanded=False):
                            st.code(verify_log, language="text")
        else:
            # 上传区被清空（用户主动移除文件）：清理当前会话文件
            if st.session_state.current_files.get('excel'):
                cleanup_files(st.session_state.current_files.get('excel'), st.session_state.current_files.get('word'))
                st.session_state.current_files = {'excel': None, 'word': None}
            # 清空统计数据
            st.session_state.module_stats = []
    
    # 右侧边栏：显示模块统计
    with stats_col:
        st.markdown('<div class="stat-container"><div class="stat-header">模块功能统计</div>', unsafe_allow_html=True)
        
        if st.session_state.module_stats:
            # 导出按钮放在顶部
            export_df = pd.DataFrame(st.session_state.module_stats)
            
            # 确保列顺序，使导出更直观
            cols_order = ['一级模块名称', '二级模块名称', '三级模块名称', '功能过程名称', '子过程数量', '子过程详情']
            # 仅保留存在的列
            final_cols = [c for c in cols_order if c in export_df.columns]
            export_df = export_df[final_cols]
            
            # 计算汇总统计
            total_l1 = export_df['一级模块名称'].nunique() if '一级模块名称' in export_df else 0
            total_l2 = export_df['二级模块名称'].nunique() if '二级模块名称' in export_df else 0
            total_l3 = export_df['三级模块名称'].nunique() if '三级模块名称' in export_df else 0
            total_processes = len(export_df)
            total_subprocesses = export_df['子过程数量'].sum() if '子过程数量' in export_df else 0
            
            # 生成导出Excel
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # 汇总表
                summary_data = [
                    {'统计项': '一级模块数量', '数值': total_l1},
                    {'统计项': '二级模块数量', '数值': total_l2},
                    {'统计项': '三级模块数量', '数值': total_l3},
                    {'统计项': '功能过程总数', '数值': total_processes},
                    {'统计项': '子过程总数', '数值': total_subprocesses}
                ]
                pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name='汇总统计')
                
                # 详细数据表 (聚合到三级模块)
                if not export_df.empty:
                    # 按三级模块聚合，只保留模块名称和子过程总数
                    agg_cols = ['一级模块名称', '二级模块名称', '三级模块名称']
                    # 确保这些列存在
                    agg_cols = [c for c in agg_cols if c in export_df.columns]
                    
                    if agg_cols:
                        # sort=False 保持原始出现顺序
                        detailed_df = export_df.groupby(agg_cols, as_index=False, sort=False)['子过程数量'].sum()
                    else:
                        detailed_df = export_df
                else:
                    detailed_df = pd.DataFrame()

                detailed_df.to_excel(writer, index=False, sheet_name='详细数据')
            excel_buffer.seek(0)
            
            st.download_button(
                label="⬇ 导出具体数据统计",
                data=excel_buffer,
                file_name="module_stats.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # 显示汇总信息 (使用 Grid 布局)
            st.markdown(f"""
            <div class="summary-grid">
                <div class="summary-card">
                    <div class="summary-val">{total_l1}</div>
                    <div class="summary-label">一级模块</div>
                </div>
                <div class="summary-card">
                    <div class="summary-val">{total_l2}</div>
                    <div class="summary-label">二级模块</div>
                </div>
                <div class="summary-card">
                    <div class="summary-val">{total_l3}</div>
                    <div class="summary-label">三级模块</div>
                </div>
                <div class="summary-card">
                    <div class="summary-val">{total_processes}</div>
                    <div class="summary-label">功能过程</div>
                </div>
                <div class="summary-card" style="grid-column: span 2;">
                    <div class="summary-val">{total_subprocesses}</div>
                    <div class="summary-label">子过程总数</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 显示模块统计数据
            st.markdown('</div>', unsafe_allow_html=True) # Close container
        else:
            st.info("执行校对后将显示统计信息")
            st.markdown('</div>', unsafe_allow_html=True) # Close container


if __name__ == "__main__":
    main()
