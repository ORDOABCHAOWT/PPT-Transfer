# PPT Transfer

<div align="center">

![PPT Transfer Icon](icon_256.png)

**macOS 风格的 PPT 文案提取工具**

一键提取 PowerPoint 演示文稿中的所有文案内容，自动导出为 Word 文档

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![macOS](https://img.shields.io/badge/platform-macOS-lightgrey.svg)](https://www.apple.com/macos/)

</div>

---

## 功能特性

- **智能提取** - 按列从左到右、每列内从上到下自动识别并提取文本
- **格式保留** - 可选保留原始文本格式（加粗、斜体、颜色等）
- **实时进度** - 流畅的进度条显示，实时追踪处理状态
- **现代界面** - macOS Sequoia 风格的毛玻璃效果和优雅动画
- **一键导出** - 自动生成格式化的 Word 文档，即刻下载

## 快速开始

### 安装部署

**构建 macOS 应用（推荐）**

```bash
./build_app.sh
```

构建完成后，应用将自动安装到 `/Applications/PPT Transfer.app`，可从启动台直接打开。

**开发模式运行**

```bash
./run.sh
```

服务将在 `http://127.0.0.1:5002` 启动，浏览器会自动打开界面。

### 使用方法

1. 打开 PPT Transfer 应用或访问开发服务器地址
2. 拖拽 `.pptx` 文件到界面，或点击上传按钮选择文件
3. 选择提取选项：
   - **列优先排序** - 按列顺序组织内容（推荐）
   - **保留文本格式** - 保持原始样式
4. 点击"开始提取"，等待处理完成
5. Word 文档将自动下载

## 项目结构

```
PPT-Transfer/
├── server.py              # Flask Web 服务器
├── extract_ppt.py         # PPT 提取核心引擎
├── templates/
│   └── index.html         # 用户界面
├── static/
│   ├── style.css          # 样式表
│   └── script.js          # 前端逻辑
├── build_app.sh           # 应用构建脚本
├── update_app.sh          # 快速更新脚本
├── run.sh                 # 开发启动脚本
└── clean.sh               # 进程清理脚本
```

## 开发指南

### 系统要求

- macOS 10.15+
- Python 3.7+
- 自动安装依赖：`flask` `python-pptx` `python-docx` `werkzeug` `Pillow`

### 常用命令

```bash
# 快速更新已安装的应用（仅更新代码文件）
./update_app.sh

# 完整重新构建应用
./build_app.sh

# 清理后台进程和端口占用
./clean.sh

# 开发模式运行
./run.sh
```

### 技术栈

| 组件 | 技术 |
|------|------|
| 后端服务 | Flask + Python 3 |
| 前端界面 | HTML5 + CSS3 + JavaScript |
| 文档处理 | python-pptx, python-docx |
| 实时通信 | Server-Sent Events (SSE) |
| UI 设计 | macOS Sequoia 风格 |

## 更新日志

**v1.0** (2025-12-03)
- 首次发布
- 实现智能列优先提取算法
- 支持文本格式保留
- 采用 macOS Sequoia 设计语言
- 提供自动化构建和更新系统

## 许可证

本项目基于 [MIT License](https://opensource.org/licenses/MIT) 开源。

## 致谢

基于 [ORDOABCHAOWT/PPT--transfer](https://github.com/ORDOABCHAOWT/PPT--transfer) 项目重构开发。

---

<div align="center">

**Made with ❤️ for macOS**

</div>
