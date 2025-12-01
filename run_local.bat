@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM 进入脚本所在目录
cd /d "%~dp0"

echo ========================================
echo 启动 Excel转Word (本地版，使用虚拟环境)
echo ========================================
echo.

REM 优先使用本地虚拟环境 .venv
set VENV_DIR=.venv
set PY_EXE=%VENV_DIR%\Scripts\python.exe
set PIP_EXE=%VENV_DIR%\Scripts\pip.exe
set STREAMLIT_EXE=%VENV_DIR%\Scripts\streamlit.exe

if not exist "%VENV_DIR%" (
    echo [信息] 正在创建 Python 虚拟环境到 %VENV_DIR% ...
    python -m venv "%VENV_DIR%"
    if %errorlevel% neq 0 (
        echo [错误] 创建虚拟环境失败，请确认已安装 Python。
        pause
        exit /b 1
    )
)

echo [信息] 升级 pip 并安装依赖...
"%PY_EXE%" -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
"%PY_EXE%" -m pip install -r requirements_converter.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

REM 可选：设置日志等级（默认 INFO）。可在运行前自行 set LOG_LEVEL=DEBUG 覆盖。
if not defined LOG_LEVEL set LOG_LEVEL=INFO

echo.
echo 正在启动浏览器...
echo 如果浏览器未自动打开，请手动访问: http://localhost:8501
echo.
echo 提示: 按 Ctrl+C 可以停止服务
echo.

"%STREAMLIT_EXE%" run app.py --server.headless false
endlocal
pause
