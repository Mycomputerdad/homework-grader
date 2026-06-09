@echo off
chcp 65001 >nul
echo.
echo ============================================================
echo   🎓 教师作业批改助手 - 启动脚本
echo   Teacher Homework Grading Assistant - Startup Script
echo ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未找到 Python，请先安装 Python 3.8+
    echo    下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Install dependencies
echo 📦 正在安装依赖...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo ⚠️ 依赖安装失败，尝试使用 pip3...
    pip3 install -r requirements.txt -q
)
echo ✅ 依赖安装完成
echo.

:: Start server
echo 🚀 正在启动服务...
echo 📍 访问地址: http://localhost:5000
echo 📍 按 Ctrl+C 停止服务
echo.
python app.py
pause
