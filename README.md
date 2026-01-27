# 🚀 NotebookLM2PPT

> **让 NotebookLM 的演示文稿真正为你所用**
> 从 PDF 到全可编辑 PPT 的智能转换工具


[最新版本 ![](https://img.shields.io/github/release/elliottzheng/NotebookLM2PPT.svg)] | [文档中心](https://elliottzheng.github.io/NotebookLM2PPT) | [下载地址](https://github.com/elliottzheng/NotebookLM2PPT/releases)

---

## 项目简介

**NotebookLM2PPT** 是一款强大的自动化工具，旨在将不可编辑的 PDF 文档（特别是 NotebookLM 生成的演示文稿）转换为**完全可编辑**的 PowerPoint 演示文稿。

### 核心价值
- **🤖 全自动化**：利用微软电脑管家"智能圈选"，自动完成截图、识别、转换和合并。
- **🧠 MinerU 深度优化**：(可选) 集成 MinerU 解析能力，智能重排文本、统一字体、替换高清图片。
- **✨ 智能去水印**：内置针对 NotebookLM 的智能水印去除算法。

---

## 🌟 效果展示

左侧为基础转换（截图识别），右侧为 **MinerU 优化后**（重排版+高清图）：

| 基础转换 PPT | **MinerU 优化后 PPT** |
| :--- | :--- |
| ![Basic](docs/public/page_0004_1_converted.jpg) | ![MinerU](docs/public/page_0004_2_converted.jpg) |
| ![Basic](docs/public/page_0003_1_converted.jpg) | ![MinerU](docs/public/page_0003_2_converted.jpg) |

> 💡 **效果惊人？** 查看 [详细对比](https://elliottzheng.github.io/NotebookLM2PPT/compare.html) 和 [基准测试数据](https://elliottzheng.github.io/NotebookLM2PPT/features.html#%F0%9F%93%8A-%E6%95%88%E6%9E%9C%E8%AF%84%E4%BC%B0)。

---

## 🚀 快速开始

详细教程请查看 [快速开始指南](https://elliottzheng.github.io/NotebookLM2PPT/quickstart.html)。

### 1. 系统要求
- **Windows 10/11**
- **Microsoft PowerPoint** 或 **WPS Office** (v0.6.5+ 支持)
- **[微软电脑管家](https://pcmanager.microsoft.com/)** (版本 $\ge$ 3.17.50.0，必须开启"智能圈选")

### 2. 安装
- **推荐**：直接在 [Releases](https://github.com/elliottzheng/NotebookLM2PPT/releases) 下载 `.exe` 文件运行。
- **开发者**：`pip install notebooklm2ppt -U`

### 3. 使用步骤
1. **启动程序**：运行 exe 或命令行输入 `notebooklm2ppt`。
2. **选择文件**：选择需要转换的 PDF。
3. **校准位置**：**首次使用务必勾选"校准按钮位置"**，根据提示点击屏幕上的"转换为PPT"按钮。
4. **开始转换**：程序将自动接管鼠标完成操作。

---

## 🧠 进阶功能：MinerU 后处理优化

想要获得专业级的排版效果？使用 MinerU 优化功能：

1. 在 [MinerU 官网](https://mineru.net/) 上传 PDF 并下载解析后的 JSON 文件。
2. 在本工具中选择 PDF 时，同时选择对应的 JSON 文件。
3. 程序将在基础转换完成后，自动执行深度优化（文本重排、字体统一、高清图替换）。

👉 [了解 MinerU 优化详情](https://elliottzheng.github.io/NotebookLM2PPT/mineru.html)

---

## ⚠️ 常见问题与注意事项

- **🔴 核心关键：按钮偏移校准**
  本工具依赖模拟点击。如果无法自动点击"转换为PPT"，请务必在界面勾选"校准按钮位置"重新校准。
- **🚫 请勿干扰**
  转换过程中程序会控制鼠标，请不要移动鼠标或操作键盘（按 `ESC` 可紧急停止）。
- **📂 找不到文件？**
  默认情况下，程序会从系统的"下载"文件夹抓取临时文件，请确保下载路径未被修改。

---

## 📚 文档导航

- [快速开始](https://elliottzheng.github.io/NotebookLM2PPT/quickstart.html) - 详细安装和使用教程
- [功能特性](https://elliottzheng.github.io/NotebookLM2PPT/features.html) - 了解所有强大功能
- [MinerU 优化](https://elliottzheng.github.io/NotebookLM2PPT/mineru.html) - 学习如何获得最佳效果
- [实现细节](https://elliottzheng.github.io/NotebookLM2PPT/implementation.html) - 技术原理揭秘
- [更新日志](https://elliottzheng.github.io/NotebookLM2PPT/changelog.html) - 查看版本历史

---

## 📄 开源协议与反馈

本项目基于 [MIT License](LICENSE) 开源。
欢迎提交 [Issues](https://github.com/elliottzheng/NotebookLM2PPT/issues) 或 Pull Request。
