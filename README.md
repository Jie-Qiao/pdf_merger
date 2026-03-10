# 全能文档转PDF合并器

一款简单易用的桌面工具，支持将 Word、Excel、图片、OFD 等多种格式文件**拖拽排序**并**一键合并**为一个 PDF 文件，方便打印或存档。

## ✨ 功能特点

*   **多格式支持**：支持 `.doc`/`.docx`, `.xls`/`.xlsx`, `.jpg`/`.png`/`.bmp` 图片, `.ofd` (电子发票等), 以及直接合并 `.pdf`。
*   **拖拽操作**：文件支持拖拽添加，并可在列表内通过拖拽自由调整合并顺序。
*   **智能转换**：
    *   **Office 文件**：自动适配 Microsoft Office 或 WPS Office 进行转换。
    *   **OFD 文件**：采用社区优化版 `easyofd` 库，转换更稳定（**建议使用**，规避官方版潜在 bug）。
*   **灵活输出**：可自定义 PDF 输出目录，并选择转换后是否自动打开文件进行打印预览。

## 🖼️ 界面预览

![image-20260310102034712](image.png)



## ⚠️ 注意事项

- **Office 依赖**：转换 Word 和 Excel 文件时，系统必须安装 **Microsoft Office** 或 **WPS Office**。
- **OFD 支持**：如需处理 OFD 文件，请确保按照上述要求正确安装 `easyofd` 库。另外easyofd可以考虑非官方社区版本 ([Ian-Jhon/easyofd](https://github.com/Ian-Jhon/easyofd))，修复了部分bug。