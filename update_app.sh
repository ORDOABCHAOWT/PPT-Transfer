#!/bin/bash

# PPT Transfer - 快速更新脚本

set -e

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m'

APP_NAME="PPT Transfer"
CURRENT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APPLICATIONS_PATH="/Applications/${APP_NAME}.app"

echo ""
echo -e "${BLUE}🔄 快速更新 ${APP_NAME}...${NC}"
echo ""

# 检查应用是否存在
if [ ! -d "$APPLICATIONS_PATH" ]; then
    echo -e "${YELLOW}⚠️  应用未安装，正在执行完整构建...${NC}"
    ./build_app.sh
    exit 0
fi

# 更新 Python 代码
echo -e "${BLUE}[1/3]${NC} 📋 更新 Python 代码..."
cp "$CURRENT_DIR/server.py" "$APPLICATIONS_PATH/Contents/Resources/"
cp "$CURRENT_DIR/extract_ppt.py" "$APPLICATIONS_PATH/Contents/Resources/"
echo -e "${GREEN}      ✓ Python 代码已更新${NC}"

echo -e "${BLUE}[2/3]${NC} 🎨 更新 Web UI 文件..."
cp "$CURRENT_DIR/templates/index.html" "$APPLICATIONS_PATH/Contents/Resources/templates/"
cp "$CURRENT_DIR/static/style.css" "$APPLICATIONS_PATH/Contents/Resources/static/"
cp "$CURRENT_DIR/static/script.js" "$APPLICATIONS_PATH/Contents/Resources/static/"
echo -e "${GREEN}      ✓ Web UI 文件已更新${NC}"

echo -e "${BLUE}[3/3]${NC} 🔄 刷新应用缓存..."
touch "$APPLICATIONS_PATH"
killall Finder 2>/dev/null || true
killall Dock 2>/dev/null || true
echo -e "${GREEN}      ✓ 应用缓存已刷新${NC}"

echo ""
echo -e "${GREEN}✅ 更新完成！${NC}"
echo ""
echo -e "${YELLOW}💡 提示:${NC} 如果应用正在运行，请重启应用以查看更改"
echo ""
