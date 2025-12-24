# Word to Base64 Image Converter

一个简单的 Python 工具，将 **Word 文档（.docx / .doc）** 转换为 **PDF**，再将 PDF 所有页面垂直拼接成一张高清长图，并生成该图片的 Base64 编码。

最终输出到一个以 Word 文件名命名的子文件夹中，结构清晰，便于管理。

## 功能特点

- 支持 `.docx` 和 `.doc` 格式
- 自动创建同名子文件夹
- 生成三个文件：
  - `{文件名}.pdf`        → 转换后的 PDF
  - `img.jpg`             → 所有页面垂直拼接的高清图片（居中对齐）
  - `base64.txt`          → 图片的完整 `data:image/jpeg;base64,...` 字符串（可直接用于 HTML）
- 高质量渲染（默认 300 DPI）
- 优先使用 LibreOffice（推荐，跨平台），Windows 下备选 Microsoft Word

## 生成的文件结构示例

输入文件：`documents/课前MR调配步骤.docx`

输出：