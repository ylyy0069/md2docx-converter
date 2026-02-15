# 📄 Markdown 转 Word 转换器 (Converter)

> **Signature**: Coded by Ajin (Gemini) with ❤️

这是一个轻量级的 Python 脚本，专门用来把 **Markdown (.md)** 或 **纯文本 (.txt)** 文件批量转换成排版整洁的 **Word (.docx)** 文档。

## ✨ 功能亮点

- **批量处理**：支持拖入单个、多个文件或整个文件夹进行自动扫描。
- **结构保留**：自动识别 Markdown 标题等级并应用 Word 标题样式。
- **视觉增强**：自动为 `**加粗文字**` 设置深蓝色，突出内容重点。
- **安全鲁棒**：内置 XML 非法字符清洗逻辑，确保转换过程不因特殊控制符中断。

## 🛠️ 环境依赖

使用前请确保安装了 `python-docx` 库：

```bash
pip install python-docx

## 🚀 使用指南

# 

运行脚本：在终端执行 python md2docx_v1.py。

投喂路径：

将 .md 文件或包含文件的文件夹直接拖入命令行窗口。

即时转换：按下回车，转换后的 .docx 文件将即时生成在原文件同级目录下。

Coded with care for efficiency.