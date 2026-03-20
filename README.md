# PPT2PDF 极速转换器 (Tauri Edition)

这是一个极致小巧、高性能、高还原度的桌面端 PPT 转 PDF 工具。本项目采用 **Tauri + Rust + Vanilla JS** 架构重新重构，将原本庞大的基于 Chromium 的安装包大幅缩减至 **< 10MB**，并保持对本地原生系统环境的高效调用能力。

## ✨ 核心特性

- ⚡️ **极小体积**：不到 10MB 的独立 EXE（依赖于 Windows 系统原生的 WebView2）。
- 🎨 **极简颜值**：现代暗黑极简 UI，体验丝滑，支持文件系统多选拖拽。
- 🔄 **智能队列**：拒绝一次性阻塞宕机，提供多文件“排队装载、批量启动、全量暂停”支持。
- 🔒 **绝对隐私**：纯净离线工具，所有转换在内存与本地临时文件夹闭环，无论断网还是涉密文件皆可安心使用。

## 🎯 转换引擎机制（多重保底）

本工具不“造轮子”解析各类复杂的动画或特殊字体，而是通过 Rust 后台直接静默调度本机的底层组件执行高保真转换。
引擎检测与调用顺序如下：

1. **Microsoft Office (PowerPoint COM)** —— 最优先使用，原厂解析，100% 格式、特效、字体无损还原。速度极快。
2. **WPS Office (KWPP COM)** —— 国内常用备份，兼容绝大多数版式。速度较快。
3. **LibreOffice (Headless 命令行)** —— 无国产及微软套件情况下的自动兜底降级方案。

## 🚀 快速开始

### 开发环境准备
由于项目内核为 Tauri 架构，你必须在系统安装以下环境：
- [Node.js](https://nodejs.org/)
- [Rust & Cargo](https://rustup.rs/) (含 Microsoft C++ Build Tools)

### 运行指令

```bash
# 1. 安装前端依赖插件
npm install

# 2. 启动本地开发窗口
npm run tauri dev

# 3. 编译发布（生成极速版安装程序和独立版 EXE）
npm run tauri build
```

*编译完成的超小体积跨平台分发包可见于 `src-tauri/target/release/bundle/nsis/` 中。*
