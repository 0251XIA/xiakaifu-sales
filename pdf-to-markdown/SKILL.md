---
name: pdf-to-markdown
description: 将PDF文档转换为Markdown格式，支持文字格式提取、表格识别、图片提取和页面截图
trigger:
  - pdf转md
  - PDF转MD
  - pdf转markdown
  - pdf转换
input:
  - file: PDF文件路径(.pdf)
  - mode: standard | enterprise (可选，默认为standard)
output:
  - markdown: 转换后的Markdown文件
  - images: 提取的图片列表
  - summary: 转换统计信息
tools:
  - pdfminer.six (pip install pdfminer.six)
  - PyMuPDF (pip install pymupdf) - 可选，用于图片提取和页面截图
  - camelot-py (可选，pip install camelot-py) - 用于表格提取
---

# PDF 图文转 Markdown 转换技能 v1.0

## 一、适用范围

用于将 `.pdf` PDF 文档转换为结构化 Markdown 文档：

- 文字格式提取（加粗、斜体、标题层级）
- 文本按逻辑顺序排列
- 表格识别与转换（Markdown 表格）
- 内嵌图片提取
- 页面截图导出（可选）

### 典型应用场景

- 培训教材整理
- 产品手册知识库化
- 合同文档标准化
- AI 知识库喂料
- PDF 内容复用

## 二、技术方案

### 2.1 格式分析

PDF 内部结构特点：

| 维度 | 说明 |
|------|------|
| **文本存储** | 实际文本流（文字型PDF）或路径绘制（扫描件） |
| **表格存储** | 无原生表格结构，打散为文字行+间距+边框 |
| **图片存储** | XObject 内联存储，格式 JPEG/PNG |
| **文字型PDF** | 可直接提取文本，无需 OCR |
| **扫描件PDF** | 无文本层，需 OCR 处理 |

### 2.2 核心库选择

```python
# 方案一：PyMuPDF（推荐）
import fitz
# 文本提取 + 图片提取 + 页面截图 + 表格辅助定位

# 方案二：pdfminer.six（细粒度布局分析）
from pdfminer.high_level import extract_text
# 纯 Python，布局信息完整，但不支持图片/表格
```

**推荐方案：PyMuPDF**，原因：
- 功能全面（文本/图片/截图/表格定位）
- 纯 Python wheel，安装简单
- 速度快

### 2.3 表格识别

```python
# 使用 camelot 进行表格提取
import camelot
tables = camelot.read_pdf('input.pdf', pages='1', flavor='stream')
# stream 模式：无边框表，基于间距检测
# lattice 模式：有边框表，基于表格线检测（需 Java）
```

### 2.4 文字格式识别

通过 PyMuPDF 的 `span["flags"]` 判断：

| Flag | 格式 |
|------|------|
| 1 | 斜体 |
| 2 | 加粗 |
| 16+ | 标题（字体大小>14pt+加粗） |

```python
def _format_text(text, page):
    for span in line["spans"]:
        is_bold = span["flags"] & 2 != 0
        is_italic = span["flags"] & 1 != 0
        if is_bold and is_italic:
            text = f"***{span['text']}***"
        elif is_bold:
            text = f"**{span['text']}**"
```

### 2.5 图片提取

```python
for img in page.get_images(full=True):
    xref = img[0]
    base_image = page.parent.extract_image(xref)
    image_bytes = base_image.get('image')
```

### 2.6 页面截图

```python
mat = fitz.Matrix(2.0, 2.0)  # 2x DPI
pix = page.get_pixmap(matrix=mat)
pix.save("page_001.png")
```

---

## 三、输入输出

### 输入

```json
{
  "file_path": "xxx.pdf",
  "mode": "standard",
  "extract_tables": true,
  "extract_images": true,
  "extract_slide_images": false
}
```

### 输出

```json
{
  "md_file": "xxx.md",
  "images": [
    {
      "original_name": "page1_img0.png",
      "saved_name": "img001.png",
      "path": "./images/img001.png"
    }
  ],
  "summary": {
    "total_pages": 15,
    "total_images": 8,
    "total_slide_images": 0
  }
}
```

---

## 四、Markdown 输出格式

```markdown
# 文档标题

## 第 1 页

这是第一段的正文，包含**加粗**和*斜体*。

| 列1 | 列2 | 列3 |
|-----|-----|-----|
| A | B | C |

![第1页图片](images/img001.png)

---

## 第 2 页

文字内容...

**二级标题**

- 要点1
- 要点2

> 备注：内嵌图片展示
![图片描述](images/img002.png)
```

---

## 五、CLI 使用

```bash
# 基础转换
python pdf_to_markdown.py document.pdf

# 指定输出目录
python pdf_to_markdown.py document.pdf -o /path/to/output

# 跳过表格提取
python pdf_to_markdown.py document.pdf --no-tables

# 仅提取图片
python pdf_to_markdown.py document.pdf --images-only

# 导出每页截图
python pdf_to_markdown.py document.pdf --slide-images

# 显示详细信息
python pdf_to_markdown.py document.pdf -v

# 批量转换
python pdf_to_markdown.py file1.pdf file2.pdf file3.pdf
```

---

## 六、与其他技能的配合

```
PPT 转 Markdown
       ↓
Word 转 Markdown
       ↓
PDF 转 Markdown
       ↓
Excel 转知识库
       ↓
   网页转知识库
```

---

*本技能基于 PyMuPDF + camelot（可选）实现*
