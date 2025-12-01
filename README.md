# Excel to Word Converter

基于 Streamlit 的 Excel 转 Word 文档工具，支持自动解析 COSMIC 格式的 Excel 表格并生成标准化 Word 文档。

## 快速开始

### Linux 服务器部署

#### 方式一：前台运行

```bash
# 1. 赋予执行权限
chmod +x run_linux.sh

# 2. 启动应用
./run_linux.sh
```

应用会在 `http://<服务器IP>:8501` 启动。

#### 方式二：后台服务

```bash
# 安装为 systemd 服务（需 root 权限）
sudo ./run_linux.sh install

# 服务管理命令
sudo systemctl status converter    # 查看状态
sudo systemctl stop converter       # 停止服务
sudo systemctl restart converter    # 重启服务
sudo systemctl disable converter    # 禁用开机自启
```

**服务特性**：
- 开机自动启动
- 异常自动重启（5秒后）

## 依赖环境

- **Python**: 3.8+
- **核心库**:
  - `streamlit` - Web 界面框架
  - `pandas` - Excel 数据处理
  - `python-docx` - Word 文档生成
  - `openpyxl` - Excel 文件读取

## 配置说明

### 日志配置（目前未更新到linux服务器上）

- **默认等级**：`INFO`
- **日志文件**：`logs/app.log`（最大 2MB，保留 3 个备份）
- **调整等级**：设置环境变量 `LOG_LEVEL`
  - Windows: `set LOG_LEVEL=DEBUG`
  - Linux: `export LOG_LEVEL=DEBUG` 或 `LOG_LEVEL=DEBUG ./run_linux.sh`
  - Systemd 服务：编辑 `/etc/systemd/system/converter.service`，修改 `Environment=LOG_LEVEL=...`

### 文件清理配置

编辑 `cleanup_loop.py` 调整：
- `RETENTION_HOURS`: 文件保留时长（默认 1 小时）
- `INTERVAL_SECONDS`: 清理间隔（默认 30 分钟）

## 使用流程

1. **上传 Excel**：拖拽或选择 Excel 文件（需包含模块拆分数据）
2. **开始转换**：点击"开始转换"按钮生成 Word 文档
3. **执行校对**：点击"执行内容校对"验证一致性
4. **下载文档**：点击"下载 Word 文档"获取生成的文件
5. **查看统计**：右侧面板将显示模块统计，详细数据可导出为 Excel

