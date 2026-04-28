---
name: document-to-markdown
description: 统一文档转换器，自动识别Word/PPT/PDF并转换为Markdown
trigger:
  - 文档转md
  - document转markdown
  - 批量转换文档
  - 文档统一转换
input:
  - file: 文件路径(.doc/.docx/.ppt/.pptx/.pdf)
output:
  - markdown: 转换后的Markdown文件
  - images: 提取的图片列表
  - summary: 转换统计信息
tools:
  - pdfminer.six
  - PyMuPDF
  - LibreOffice
---

# 文档统一转 Markdown 技能 v1.0

## 一、概述

统一入口转换器，根据文件类型自动分发到对应转换器：

| 文件类型 | 对应技能 |
|---------|---------|
| `.doc` / `.docx` | word-to-markdown |
| `.ppt` / `.pptx` | ppt-to-markdown |
| `.pdf` | pdf-to-markdown |

## 二、技术架构

```
document_to_markdown.py
    │
    ├── 检测文件扩展名
    │
    ├── 根据类型动态加载
    │   ├── .doc/.docx → WordToMarkdown
    │   ├── .ppt/.pptx → PPTToMarkdown
    │   └── .pdf       → PDFToMarkdown
    │
    └── 调用对应转换器
```

## 三、依赖

- `pdfminer.six` — PDF 文字提取
- `PyMuPDF` — 图片提取 + 截图
- `LibreOffice` — .doc/.ppt 格式转换

## 四、CLI 使用

```bash
# 自动识别类型
python document_to_markdown.py 文档.docx
python document_to_markdown.py 演示文稿.pptx
python document_to_markdown.py 文档.pdf

# 指定输出目录
python document_to_markdown.py 文档.pdf -o ./output

# 批量混合类型
python document_to_markdown.py 文件1.doc 文件2.pptx 文件3.pdf

# 显示详细信息
python document_to_markdown.py 文档.pdf -v
```

## 五、输出结构

```
源文件目录/
├── 文档.docx_output/
│   ├── 文档.md
│   └── images/
├── 演示文稿.pptx_output/
│   ├── 演示文稿.md
│   ├── images/
│   └── slide_images/
└── 文档.pdf_output/
    ├── 文档.md
    └── images/
```

## 六、与其他技能配合

```
用户上传文档
    ↓
document-to-markdown（统一入口）
    ↓
┌───┴───┐
↓       ↓
Word   PPT   PDF
 ↓      ↓     ↓
word-to-markdown
ppt-to-markdown
pdf-to-markdown
```

---

*本技能通过动态加载调用 word-to-markdown / ppt-to-markdown / pdf-to-markdown 实现*
