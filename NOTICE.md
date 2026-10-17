# NOTICE / 第三方依赖与许可说明

本项目自身代码以 **GNU Affero General Public License v3.0（AGPL-3.0）** 发布，见 `LICENSE`。

`CE转换引擎（Conversion Engine）` 是独立重建的工具，不包含、不复制 WPS Office 的代码。
以下第三方组件通过包管理器安装或在运行时调用，其版权归各自作者所有，遵循各自许可证：

| 组件 | 用途 | 许可证 |
|---|---|---|
| PyMuPDF (MuPDF) | PDF 渲染/文本/图片 | AGPL-3.0 |
| pypdf | PDF 合并/拆分/提取页 | BSD-3-Clause |
| pdfplumber | PDF 表格/文本提取 | MIT |
| pdf2docx | PDF → DOCX | MIT |
| python-docx | DOCX 读写 | MIT |
| python-pptx | PPTX 读写 | MIT |
| openpyxl | XLSX 读写 | MIT |
| Pillow | 图片处理 | MIT-CMU |
| RapidOCR / onnxruntime | 本地 OCR | Apache-2.0 / MIT |
| easyofd | OFD 解析与生成 | 以该库发布页/仓库声明为准 |
| pywin32 | WPS COM 自动化 | PSF-2.0 |
| WPS Office | `.wps` 格式本地读写（运行时调用，不随本项目分发） | 金山专有软件，需用户自行安装并遵守其许可 |

### 测试样本

- `tests/sample-helloworld.ofd`、`sample-invoice.ofd`、`sample-text.ofd` 来自 [ofdrw](https://github.com/ofdrw/ofdrw) 测试资源（Apache-2.0）。
- `tests/sample-multi.pdf` 等 PDF 由 `tests/make_samples.py` 生成。

### 图像素材

- `desktop/background.jpg` 与 `desktop/icon-source.webp` 为项目所有者提供的图像素材，仅用于本项目的界面与图标。

### 免责声明

本项目按“原样”提供，不附带任何担保。使用第三方格式（尤其 `.wps`）时请确保你拥有相应软件许可与文件处理权限。
