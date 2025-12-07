#!/bin/bash

# PPT Transfer - 快速启动脚本（开发模式）

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo ""
echo -e "${BLUE}🚀 启动 PPT Transfer (开发模式)${NC}"
echo ""

# 检查并创建虚拟环境
if [ ! -d "venv" ]; then
    echo -e "${YELLOW}📦 创建虚拟环境...${NC}"
    python3 -m venv venv
    echo ""
fi

# 检查依赖
./venv/bin/python -c "import flask" 2>/dev/null || {
    echo -e "${YELLOW}📦 正在安装依赖...${NC}"
    ./venv/bin/pip install flask python-pptx python-docx werkzeug Pillow
    echo ""
}

# 启动服务器
echo -e "${GREEN}✅ 服务器正在启动...${NC}"
echo -e "${BLUE}📡 地址: http://127.0.0.1:5002${NC}"
echo -e "${YELLOW}💡 按 Ctrl+C 停止服务器${NC}"
echo ""

./venv/bin/python server.py
