@echo off
chcp 65001 >nul
echo ==============================================
echo      🍅 每日督学小助手
echo ==============================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 未检测到Python，请先安装Python 3.8+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查依赖是否安装
python -c "import matplotlib, PIL" >nul 2>&1
if %errorlevel% neq 0 (
    echo ⏳ 正在安装依赖...
    pip install matplotlib Pillow python-docx
    if %errorlevel% neq 0 (
        echo ❌ 依赖安装失败，请手动执行：pip install matplotlib Pillow python-docx
        pause
        exit /b 1
    )
    echo ✅ 依赖安装完成！
)

:: 运行应用
echo 🚀 启动应用...
python app.py

echo.
echo 👋 应用已退出
pause
