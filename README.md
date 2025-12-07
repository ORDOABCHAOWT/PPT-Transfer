# PPT Transfer

<div align="center">

![PPT Transfer Icon](icon_256.png)

**macOS Sequoia 风格的 PPT 文案提取工具**

一键提取 PPT 中的所有文案内容，支持导出为 Word 文档

</div>

## ✨ 特性

- 🎨 **macOS Sequoia 设计风格** - 现代化的毛玻璃界面，完美适配 macOS
- 🚀 **一键提取** - 拖放 PPT 文件即可快速提取所有文案
- 📊 **实时进度** - 流畅的进度条显示，实时追踪提取进度
- 📝 **Word 导出** - 自动导出为格式化的 Word 文档
- 🔄 **智能排序** - 支持列优先排序，保持原有结构
- 🎭 **格式保留** - 可选择保留原文本格式（加粗、斜体等）
- ⚡ **自动更新** - 一键更新应用到 Applications 文件夹

## 📦 安装

### 方式一：自动构建（推荐）

```bash
cd /Users/whitney/Downloads/PPT-Transfer
./build_app.sh
```

构建完成后，应用会自动部署到 `/Applications/PPT Transfer.app`

### 方式二：开发模式

```bash
./run.sh
```

开发模式会在 `http://127.0.0.1:5002` 启动服务器

## 🚀 使用方法

### 应用模式

1. 在启动台或 Applications 文件夹中找到 **PPT Transfer**
2. 双击打开，浏览器会自动打开 Web 界面
3. 拖放 PPT 文件到界面中，或点击上传按钮
4. 选择提取选项：
   - **列优先排序**：按列顺序排列文案（推荐）
   - **保留格式**：保留加粗、斜体等文本格式
5. 等待提取完成，Word 文档会自动下载

### 开发模式

```bash
# 启动开发服务器
./run.sh

# 清理旧进程
./clean.sh
```

## 🔄 更新应用

修改代码后，运行快速更新脚本：

```bash
./update_app.sh
```

或执行完整构建：

```bash
./build_app.sh
```

## 📂 项目结构

```
PPT-Transfer/
├── server.py                 # Flask 后端服务
├── extract_ppt.py           # PPT 提取核心逻辑
├── create_icon.py           # 应用图标生成脚本
├── templates/
│   └── index.html          # Web UI 界面
├── static/
│   ├── style.css           # macOS Sequoia 风格样式
│   └── script.js           # 前端交互逻辑
├── build_app.sh            # 自动构建和部署脚本
├── update_app.sh           # 快速更新脚本
├── clean.sh                # 进程清理脚本
└── run.sh                  # 开发模式启动脚本
```

## 🎨 设计特点

- **橙紫渐变图标**：采用 Apple Orange 到 Apple Purple 的渐变色
- **文档提取主题**：图标展示文档到文本的提取过程
- **毛玻璃效果**：使用 backdrop-filter 实现现代化的透明效果
- **流畅动画**：所有交互都配有平滑的过渡动画
- **响应式设计**：完美适配不同尺寸的浏览器窗口

## 🛠 技术栈

- **后端**：Flask + Python 3
- **前端**：HTML5 + CSS3 + Vanilla JavaScript
- **文档处理**：python-pptx + python-docx
- **图标生成**：Pillow (PIL)
- **实时通信**：Server-Sent Events (SSE)

## 📋 系统要求

- macOS 10.15 或更高版本
- Python 3.7 或更高版本
- 依赖包会自动安装：
  - flask
  - python-pptx
  - python-docx
  - werkzeug
  - Pillow

## 🔧 脚本说明

### build_app.sh
完整构建和部署脚本，包含 7 个步骤：
1. 清理旧版本
2. 创建应用包结构
3. 生成应用图标（.icns）
4. 复制应用文件
5. 创建启动脚本
6. 创建 Info.plist
7. 部署到 Applications 文件夹

### update_app.sh
快速更新脚本，仅更新以下文件：
- Python 代码（server.py, extract_ppt.py）
- Web UI 文件（HTML, CSS, JS）

### clean.sh
清理脚本，用于：
- 停止所有 PPT Transfer 相关进程
- 释放端口 5002
- 解决端口冲突问题

### run.sh
开发模式启动脚本：
- 自动检查并安装依赖
- 在 http://127.0.0.1:5002 启动服务器
- 便于快速测试和调试

## 📝 版本历史

### v1.0 (2025-12-03)
- ✨ 首次发布
- 🎨 采用 macOS Sequoia 设计风格
- 📊 实时进度追踪
- 📝 Word 文档导出
- 🔄 自动构建和更新系统

## 📄 许可证

MIT License

## 🙏 鸣谢

基于 [ORDOABCHAOWT/PPT--transfer](https://github.com/ORDOABCHAOWT/PPT--transfer) 项目重构

---

<div align="center">

**🎉 享受使用 PPT Transfer！**

</div>
