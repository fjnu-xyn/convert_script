#!/bin/bash

# 获取脚本所在目录，确保在任何地方执行都能找到文件
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR"

# --- 1. 环境检查与准备 ---

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 python3。请先安装 Python 3。"
    exit 1
fi

# 准备虚拟环境
VENV_DIR="venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "正在创建 Python 虚拟环境..."
    python3 -m venv "$VENV_DIR"
    if [ $? -ne 0 ]; then
        echo "错误: 创建虚拟环境失败。请尝试安装 python3-venv 包。"
        exit 1
    fi
fi

# 激活环境
source "$VENV_DIR/bin/activate"

# 安装依赖
echo "正在检查依赖..."
pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install -r requirements_converter.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# --- 2. 功能分支: 安装服务 或 启动应用 ---

if [ "$1" = "install" ]; then
    # === 安装开机自启服务 ===
    if [ "$EUID" -ne 0 ]; then 
        echo "错误: 配置系统服务需要 root 权限。"
        echo "请使用 sudo 运行: sudo ./run_linux.sh install"
        exit 1
    fi

    SERVICE_NAME="converter"
    SERVICE_FILE="/etc/systemd/system/${SERVICE_NAME}.service"
    REAL_USER=${SUDO_USER:-$USER} # 获取 sudo 前的用户

    echo "正在配置 Systemd 服务..."
    echo "  - 服务文件: $SERVICE_FILE"
    echo "  - 运行用户: $REAL_USER"
    echo "  - 工作目录: $SCRIPT_DIR"
    
    cat > "$SERVICE_FILE" <<EOF
[Unit]
Description=Excel to Word Converter Web App
After=network.target

[Service]
Type=simple
User=$REAL_USER
WorkingDirectory=$SCRIPT_DIR
ExecStart=$SCRIPT_DIR/$VENV_DIR/bin/streamlit run app.py --server.port 8501 --server.address 0.0.0.0
Environment=LOG_LEVEL=${LOG_LEVEL:-INFO}
Environment=PYTHONUNBUFFERED=1
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

    systemctl daemon-reload
    systemctl enable $SERVICE_NAME
    systemctl restart $SERVICE_NAME
    
    echo "--------------------------------------------------"
    echo "✅ 开机自启服务已安装并启动！"
    echo "服务名称: $SERVICE_NAME"
    echo "查看状态: sudo systemctl status $SERVICE_NAME"
    echo "停止服务: sudo systemctl stop $SERVICE_NAME"
    echo "--------------------------------------------------"

else
    # === 正常启动应用 ===
    echo "--------------------------------------------------"
    echo "应用正在启动..."
    
    SERVER_IP=$(hostname -I 2>/dev/null | awk '{print $1}')
    [ -z "$SERVER_IP" ] && SERVER_IP="<服务器IP>"
    
    echo "请在浏览器访问: http://$SERVER_IP:8501"
    echo "提示: 运行 'sudo ./run_linux.sh install' 可配置开机自启"
    echo "--------------------------------------------------"
    
    export LOG_LEVEL=${LOG_LEVEL:-INFO}
    export PYTHONUNBUFFERED=1
    streamlit run app.py --server.port 8501 --server.address 0.0.0.0
fi
