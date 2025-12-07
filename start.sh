#!/bin/bash

# PPT Transfer - 简易启动脚本

cd "$(dirname "$0")"

# 检查并创建虚拟环境
if [ ! -d "venv" ]; then
    echo "首次运行，正在设置环境..."
    python3 -m venv venv
    ./venv/bin/pip install flask python-pptx python-docx werkzeug Pillow
fi

# 启动服务器并打开浏览器
./venv/bin/python server.py
