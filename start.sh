#!/bin/bash
echo ""
echo "============================================================"
echo "  教师作业批改助手 - 启动脚本"
echo "  Teacher Homework Grading Assistant - Startup Script"
echo "============================================================"
echo ""

# Detect python command
PYTHON=""
if command -v python3 &> /dev/null; then
    PYTHON=python3
elif command -v python &> /dev/null; then
    PYTHON=python
else
    echo " [ERROR] 未找到 Python，请先安装 Python 3.8+"
    echo "   Ubuntu/Debian: sudo apt install python3 python3-pip"
    echo "   或下载: https://www.python.org/downloads/"
    exit 1
fi

echo "  使用: $($PYTHON --version)"

# Install dependencies
echo " [..] 正在安装依赖..."
$PYTHON -m pip install -r requirements.txt -q 2>/dev/null
if [ $? -ne 0 ]; then
    echo " [..] 尝试使用 pip3..."
    pip3 install -r requirements.txt -q 2>/dev/null
fi
echo " [OK] 依赖安装完成"
echo ""

# Start server
echo " [>>] 正在启动服务..."
echo " [..] 访问地址: http://localhost:5000"
echo " [..] 按 Ctrl+C 停止服务"
echo ""
$PYTHON app.py
