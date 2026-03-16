# 🐍 Python 效率工具箱 | Python Tools Hub

[![GitHub stars](https://img.shields.io/github/stars/moshidi244262/Python-Tools?style=flat-square&color=blue)](https://github.com/moshidi244262/Python-Tools/stargazers)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-blue?style=flat-square&logo=python)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=flat-square)]()

> **个人独立开发的高效实用 Python 工具集** | 涵盖音视频处理、图片压缩、文本解析等日常需求，打造极致的效率体验。

---

## 📋 项目概览

这是一个由我个人独立开发并维护的 Python 工具集合，旨在解决日常工作和学习中的高频效率痛点。所有工具均采用现代化 GUI 界面设计，支持拖拽操作、批量处理和多线程优化，真正做到**开箱即用、高效便捷**。

**在线演示网站**：[https://moshidi244262.github.io/Python-Tools/](https://moshidi244262.github.io/Python-Tools/)

---

## 🚀 工具列表

| 工具名称 | 核心功能 | 技术栈 | 界面预览 |
| :--- | :--- | :--- | :--- |
| **🎤 音视频转文本** | 基于本地化部署的 Whisper 模型构建的智能转录工具，支持数十种音视频格式，支持拖拽及批量处理。 | Whisper, PyTorch, tkinterdnd2 | ![音视频转文本界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-yw.png) |
| **🖼️ GIF 压缩工具** | 一款基于 Pillow 和 Tkinter 的多线程 GIF 动图压缩工具。支持文件拖拽导入、全局与单文件独立压缩率调节，有效减小体积并保留帧动画。 | Pillow, TkinterDnD, 多线程异步, 自适应调色板 | ![GIF压缩界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-gif.png) |
| **🎭 角色卡提取工具** | 一款强大的酒馆角色卡解析工具。支持拖拽批量处理，通过解析图片元数据及底层二进制扫描，极速提取隐写的 Base64 数据并转换为结构化 JSON。 | Pillow解析, Base64解码, 正则扫描, TkinterDnD | ![角色卡提取界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-jg.png) |
| **📷 图片压缩处理** | 一款功能强大的本地图片批量处理工具。支持格式互转、分辨率缩放、画质压缩，支持鼠标拖拽与多线程极速处理。 | Pillow, TkinterDnD, 多线程, 批量转换 | ![图片压缩界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-ty.png) |
| **🎬 视频音频提取工具** | 一款高效的视频转音频桌面应用，支持批量拖拽处理。一键读取十余种视频格式，并提取为高品质 MP3 或 WAV。 | MoviePy, TkinterDnD, 多线程异步, 批量处理 | ![视频音频提取界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-sy.png) |
| **🎵 极速音频压缩工具** | 基于 FFmpeg 引擎的极速音频批量压缩工具。支持多线程并发、拖拽导入，可自定义比特率、采样率与声道，在体积与音质间寻找完美平衡。 | FFmpeg, 多线程并发, TkinterDnD | ![音频压缩界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-yy.png) |
| **🏷️ 音乐标签清除** | 极简高效的音乐元数据清除工具。支持批量拖拽导入，一键去除音频文件中的广告标签或乱码，生成纯净版音频。 | PySide6, Mutagen, 多线程, 无损处理 | ![音乐标签清除界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-yb.png) |
| **📄 通用文件转 Markdown 工具** | 支持将数十种文档（Word, PDF, Excel等）和代码文件快速转换为 Markdown 格式，支持拖拽与批量处理。 | 文档解析, 多格式转换, 多线程, 自动化排版 | ![文件转Markdown界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-md.png) |
| **🎥 视频批量压缩工具** | 一款基于 FFmpeg 的强大的视频批量压缩桌面端工具。支持硬件加速、拖拽导入、自定义分辨率与帧率，在画质与体积之间取得完美平衡。 | FFmpeg, TkinterDnD, 显卡加速, 多线程异步 | ![视频压缩界面](https://moyermusic.cn/moyer-space-cover/python-cover/py-sys.png) |

---

## ✨ 核心特性

### 🎯 **用户体验优先**
- **现代化 GUI 界面**：所有工具均配备美观易用的桌面应用程序界面
- **拖拽操作支持**：集成 `tkinterdnd2`，支持文件/文件夹一键拖拽导入
- **批量处理能力**：支持同时处理多个文件，大幅提升工作效率
- **实时进度反馈**：多线程处理配合实时进度条，操作过程清晰可见

### ⚡ **性能优化**
- **多线程/异步处理**：避免界面卡顿，充分利用系统资源
- **硬件加速支持**：视频压缩工具支持 NVIDIA GPU 硬件编码
- **智能算法优化**：采用高质量重采样算法，平衡速度与质量

### 🔒 **安全可靠**
- **完全本地运行**：所有处理均在本地完成，保护用户隐私
- **无损处理机制**：自动备份原文件，避免数据丢失风险
- **健壮的错误处理**：完善的异常捕获和用户友好提示

---

## 🛠️ 技术架构

### 前端展示层
- **响应式网站**：基于 HTML5 + Tailwind CSS + JavaScript 构建
- **现代化设计**：玻璃态卡片、渐变文字、平滑动画交互
- **PWA 支持**：可安装为桌面应用，提供原生应用体验

### 后端工具层
- **核心语言**：Python 3.8+
- **GUI 框架**：Tkinter / PySide6
- **图像处理**：Pillow (PIL)
- **音视频处理**：FFmpeg, MoviePy, Whisper
- **文档处理**：python-docx, PyPDF2, python-pptx
- **交互增强**：tkinterdnd2（拖拽支持）

### 工程化部署
- **自动化构建**：GitHub Actions 自动化工作流
- **静态部署**：GitHub Pages 托管演示网站
- **依赖管理**：详细的 `requirements.txt` 和安装指南

---

## 📦 快速开始

### 1. 环境准备
```bash
# 克隆仓库
git clone https://github.com/moshidi244262/Python-Tools.git
cd Python-Tools

# 安装 Python（如未安装）
# 推荐 Python 3.8 或更高版本
```

### 2. 安装通用依赖
```bash
# 基础 GUI 和拖拽支持
pip install Pillow tkinterdnd2

# 根据需要使用特定工具
# 音视频处理
pip install moviepy openai-whisper torch
# 文档处理  
pip install python-docx PyPDF2 python-pptx pandas
# 现代化 GUI
pip install PySide6
```

### 3. 运行工具
```bash
# 进入对应工具目录
cd "工具文件夹名称"

# 运行 Python 脚本
python 主程序文件.py
```

### 4. 使用在线演示
访问 [在线演示网站](https://moshidi244262.github.io/Python-Tools/) 查看所有工具介绍和截图。

---

## 📁 项目结构

```
Python-Tools/
├── index.html                    # 主展示网站
├── README.md                     # 项目说明文档
├── LICENSE                       # 许可证 
├── 音频转文本/                   # 工具1：音视频转文本
│   └── audio-to-txt.py
├── Gif图压缩/                    # 工具2：GIF压缩
│   └── Gif-yasuo.py
├── 酒馆角色卡转JSON/             # 工具3：角色卡提取
│   └── JSON-jiuguan.py
├── 图片压缩/                     # 工具4：图片处理
│   └── comperss-p.py
├── 视频转音频/                   # 工具5：音频提取
│   └── vedio-audio.py
├── 音频压缩/                     # 工具6：音频压缩
│   └── audio-comperss.py
├── 音乐标签去除工具/             # 工具7：标签清除
│   └── audio-label.py
├── 文件转.md/                    # 工具8：文档转换
│   └── file-md.py
└── 视频压缩/                     # 工具9：视频压缩
    └── vedio-compress.py
```

---

## 🎨 设计理念

### 1. **解决真实痛点**
每个工具都源于我个人的实际需求，确保解决的是真实存在的问题，而非"为了技术而技术"。

### 2. **极致用户体验**
- **零学习成本**：直观的界面设计，用户无需阅读复杂文档
- **一键操作**：复杂功能封装为简单按钮点击
- **即时反馈**：所有操作都有明确的进度和结果提示

### 3. **代码质量保证**
- **模块化设计**：每个工具独立可运行，便于维护和扩展
- **详细注释**：关键代码段配有中文注释，方便理解
- **错误处理**：完善的异常捕获和用户友好提示

### 4. **持续迭代优化**
- 根据用户反馈不断改进功能
- 定期更新依赖库版本
- 优化算法提升处理效率

---

## 🔧 开发指南

### 添加新工具
1. 在对应目录创建工具文件夹
2. 实现核心功能脚本
3. 添加必要的依赖说明
4. 更新 `index.html` 中的工具列表
5. 提交 Pull Request

### 代码规范
- 使用有意义的变量和函数名
- 添加必要的代码注释
- 遵循 PEP 8 编码规范
- 编写简单的使用说明

### 测试建议
- 在不同操作系统上测试兼容性
- 使用各种文件格式进行测试
- 测试批量处理的稳定性
- 验证错误处理的健壮性

---

## 🤝 贡献指南

欢迎任何形式的贡献！包括但不限于：

1. **报告 Bug**：在 Issues 中描述遇到的问题
2. **功能建议**：提出新的工具想法或改进建议
3. **代码贡献**：提交 Pull Request 修复问题或添加功能
4. **文档改进**：完善使用说明或添加示例
5. **分享使用**：告诉更多人这个项目的存在

### 提交规范
- 使用清晰描述性的提交信息
- 一个 PR 只解决一个问题
- 确保代码通过基本测试
- 更新相关文档

---

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

---

## 📞 联系与支持

- **GitHub Issues**：[问题反馈](https://github.com/moshidi244262/Python-Tools/issues)
- **在线演示**：[https://moshidi244262.github.io/Python-Tools/](https://moshidi244262.github.io/Python-Tools/)
- **个人主页**：[GitHub Profile](https://github.com/moshidi244262)

---

## 🌟 致谢

感谢所有开源项目的贡献者，特别感谢：
- Python 社区提供的强大生态
- 各个依赖库的维护者
- 测试和反馈的用户们
- AI 辅助编程工具的支持

---

## 📊 项目状态

**当前版本**：v1.0.0  
**最后更新**：2024年3月  
**活跃维护**：✅ 是  
**未来计划**：
- [ ] 添加更多实用工具
- [ ] 开发跨平台安装包
- [ ] 实现云端同步功能
- [ ] 添加插件系统支持
- [ ] 优化移动端适配

---

> **💡 提示**：所有工具均为个人开发作品，旨在分享和学习。如果在使用中遇到问题或有改进建议，欢迎通过 GitHub Issues 反馈。让我们一起打造更好的效率工具！
