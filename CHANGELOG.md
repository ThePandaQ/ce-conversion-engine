# Changelog

## v1.0.0 — 2026-10-17（首个公开版本）

首个公开发布版本，包含命令行工具与 Windows 图形界面。

### 核心功能

- **PDF → Word / Excel / PPT / 图片 / 文本 / 纯图 PDF / 可搜索 PDF**
- **PDF 合并 / 拆分 / 提取页面 / 提取内嵌图片**
- **本地 OCR**（RapidOCR，离线，中文识别）
- **图片 → PDF**
- **PDF ↔ OFD**：`ofd2pdf`、`ofd2docx`、`ofd2xlsx`、`ofd2pptx`、`ofd2png`、`ofd2jpg`、`ofd2txt`、`pdf2ofd`
- **WPS(.wps) ↔ Word / PDF / TXT / OFD**：
  - `.wps` →：`wps2docx`、`wps2pdf`、`wps2txt`、`wps2ofd`
  - → `.wps`：`doc2wps`（doc/docx/rtf/txt）、`pdf2wps`、`ofd2wps`
  - 通过本机 WPS Office COM 自动化（`KWPS.Application`）完成，不联网、不调用 VIP/云端引擎
- 全部转换在本机离线完成（依赖开源引擎，不依赖 WPS，除 `.wps` 格式外）

### Windows GUI

- 深色照片背景（cover 裁切，窗口缩放/最大化始终铺满）
- 自适应控件配色（从背景照片取样统一混色）、深色标题栏
- 28 种转换动作、文件拖放、"添加文档..."文件选择
- 后台异步转换 + 实时进度 + 取消按钮
- 彩色日志（成功绿/失败红/警告橙）、UTF-8 日志落盘、自动换行 + 右侧竖向滚动条、双击打开日志
- 转换成功/失败提示音（可关闭）
- 默认输出目录设置（"设为默认"，配置持久化）
- 状态栏整行显示完整结果

### 其他

- 多尺寸应用图标（16–256，透明圆角）
- 桌面快捷方式自动创建
- `doctor` 自检依赖与 WPS COM 可用性
- 离线回归测试（`tests/run_offline_test.py`，毒化代理下验证不联网）
- WPS 桥接回归测试（`tests/run_wps_test.py`，需要本机 WPS Office）

### 系统要求

- Windows 10 / 11
- Python 3.10+（安装脚本自动创建 venv）
- 可选：WPS Office（仅 `.wps` 转换需要）
- 依赖与许可证见 `NOTICE.md`
