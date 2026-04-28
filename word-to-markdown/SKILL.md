---
name: word-to-markdown
description: 将Word图文文档(.doc/.docx)转换为Markdown格式，支持图片提取、图文位置还原、合并单元格表格、页眉页脚清洗
trigger:
  - 上传docx
  - 转知识库
  - Word转MD
  - word转markdown
  - docx转换
  - doc转换
input:
  - file: Word文件路径(.doc/.docx)
  - mode: standard | enterprise (可选，默认为standard)
output:
  - markdown: 转换后的Markdown文件
  - images: 提取的图片列表
  - headers_footers: 页眉页脚列表
  - summary: 转换统计信息
tools:
  - python-docx
  - mammoth
  - pandoc
  - docx2python
  - LibreOffice (用于.doc格式转换)
---

# Word 图文转 Markdown 转换技能 v2.2

## 一、适用范围

用于将 `.doc` / `.docx` / Word 文档中的以下元素自动转换为结构化 Markdown 文档：

- 标题层级（H1-H6）
- 正文内容
- 加粗、斜体、删除线
- 有序列表、无序列表
- 表格（支持合并单元格）
- 图片提取
- 图片位置还原
- **多种图片说明模式**（下图为、如下图、参考图片、图片示例、纯图片、![xxx]）
- 超链接
- **页眉页脚提取**
- **脚注尾注**
- **智能页码清洗**
- 引用块

### 典型应用场景

- 企业知识库整理
- 产品文档迁移
- 培训课件拆解
- AI 知识库喂料
- FAQ 文档标准化
- 合同模板归档

## 二、技能价值

大多数企业资料都以 Word 格式存在：
- 产品手册
- 操作指南
- 培训资料
- 制度文档
- 客户交付文档

AI 难以直接解析 Word，而 Markdown 是 AI 最友好的格式。
因此，本技能是实现 **"知识库入口"** 的关键能力。

## 三、技能流程设计

```
上传 Word
    ↓
解析 docx 结构
    ↓
提取文本层级
    ↓
提取图片
    ↓
图片命名并保存
    ↓
还原图文位置
    ↓
输出 Markdown
    ↓
可直接入知识库 / RAG
```

## 四、输入输出

### 输入

```json
{
  "file_path": "xxx.docx",
  "mode": "standard",
  "extract_images": true,
  "image_prefix": "img",
  "preserve_headers_footers": false
}
```

### 输出

```json
{
  "md_file": "xxx.md",
  "images": [
    {
      "original_name": "image1.png",
      "saved_name": "img001.png",
      "path": "./images/img001.png"
    },
    {
      "original_name": "image2.jpg",
      "saved_name": "img002.jpg",
      "path": "./images/img002.jpg"
    }
  ],
  "summary": {
    "total_pages": 12,
    "total_images": 8,
    "total_headings": 15,
    "total_tables": 3,
    "total paragraphs": 120
  }
}
```

## 五、技术方案

### 核心 Python 库

```python
# 方案一：python-docx（推荐，最可控）
from docx import Document
from docx.shared import Inches, Pt
import zipfile
import os
import re

# 方案二：mammoth（用于纯文本转换）
import mammoth

# 方案三：pandoc（万能转换器）
import subprocess
subprocess.run(['pandoc', 'input.docx', '-o', 'output.md'])
```

### 图片提取核心代码

```python
import zipfile
import os
import re

def extract_images_from_docx(docx_path, output_dir):
    """从 docx 文件中提取所有图片"""
    images = []
    image_counter = 1

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        for file_info in zip_ref.namelist():
            if file_info.startswith('word/media/'):
                # 获取图片扩展名
                ext = os.path.splitext(file_info)[1]
                new_name = f"img{image_counter:03d}{ext}"
                output_path = os.path.join(output_dir, new_name)

                # 提取图片
                zip_ref.extract(file_info, output_dir)
                os.rename(
                    os.path.join(output_dir, file_info),
                    output_path
                )

                images.append({
                    "original_name": os.path.basename(file_info),
                    "saved_name": new_name,
                    "path": output_path
                })
                image_counter += 1

    return images
```

### 标题层级转换

```python
def convert_heading_style(heading_style):
    """将 Word 标题样式转换为 Markdown 标题"""
    style_map = {
        'Heading 1': '#',
        'Heading 2': '##',
        'Heading 3': '###',
        'Heading 4': '####',
        'Heading 5': '#####',
        'Heading 6': '######',
    }
    return style_map.get(heading_style, '')
```

### 表格转换

```python
def convert_table(table):
    """将 Word 表格转换为 Markdown 表格"""
    md_lines = []

    rows = table.rows
    for i, row in enumerate(rows):
        cells = [cell.text.strip() for cell in row.cells]
        md_line = '| ' + ' | '.join(cells) + ' |'
        md_lines.append(md_line)

        # 表头分隔线
        if i == 0:
            md_lines.append('| ' + ' | '.join(['---'] * len(cells)) + ' |')

    return '\n'.join(md_lines)
```

## 六、完整转换器实现

```python
#!/usr/bin/env python3
"""
Word 转 Markdown 转换器
"""

import os
import re
import zipfile
import json
from pathlib import Path
from docx import Document
from docx.shared import Inches


class WordToMarkdown:
    """Word 文档转 Markdown 转换器"""

    def __init__(self, file_path, mode='standard'):
        self.file_path = Path(file_path)
        self.mode = mode
        self.images = []
        self.image_counter = 0
        self.output_dir = self.file_path.parent / f"{self.file_path.stem}_output"
        self.images_dir = self.output_dir / 'images'

    def convert(self):
        """执行转换"""
        # 创建输出目录
        self.output_dir.mkdir(exist_ok=True)
        self.images_dir.mkdir(exist_ok=True)

        # 解析文档
        doc = Document(self.file_path)

        # 提取图片
        self._extract_images()

        # 转换内容
        md_content = self._convert_document(doc)

        # 保存 Markdown 文件
        md_file = self.output_dir / f"{self.file_path.stem}.md"
        md_file.write_text(md_content, encoding='utf-8')

        # 生成报告
        return {
            'md_file': str(md_file),
            'images': self.images,
            'summary': self._generate_summary(doc)
        }

    def _extract_images(self):
        """提取文档中的所有图片"""
        with zipfile.ZipFile(self.file_path, 'r') as zip_ref:
            for file_info in zip_ref.namelist():
                if file_info.startswith('word/media/'):
                    ext = os.path.splitext(file_info)[1]
                    self.image_counter += 1
                    new_name = f"img{self.image_counter:03d}{ext}"
                    output_path = self.images_dir / new_name

                    # 提取并重命名
                    with zip_ref.open(file_info) as source:
                        with open(output_path, 'wb') as target:
                            target.write(source.read())

                    self.images.append({
                        'original_name': os.path.basename(file_info),
                        'saved_name': new_name,
                        'path': str(output_path)
                    })

    def _convert_paragraph(self, para):
        """转换单个段落"""
        # 检查是否是标题
        if para.style.name.startswith('Heading'):
            level = para.style.name.split()[-1]
            text = para.text.strip()
            return f"{'#' * int(level)} {text}\n\n"

        # 普通段落
        text = para.text.strip()
        if not text:
            return '\n'

        # 处理行内格式
        text = self._convert_inline_formatting(para, text)

        return f"{text}\n\n"

    def _convert_inline_formatting(self, para, text):
        """处理行内格式（加粗、斜体等）"""
        # 遍历所有 run
        full_text = ''
        for run in para.runs:
            run_text = run.text

            # 处理超链接
            for hyperlink in para._element.findall('.//w:hyperlink'):
                r_id = hyperlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if r_id:
                    # 获取链接文本
                    link_text = ''.join([r.text for r in hyperlink.findall('.//w:t')])
                    rel = para.part.rels.get(r_id)
                    if rel:
                        url = rel.target_ref
                        run_text = run_text.replace(link_text, f'[{link_text}]({url})')

            # 加粗
            if run.bold:
                run_text = f"**{run_text}**"

            # 斜体
            if run.italic:
                run_text = f"*{run_text}*"

            full_text += run_text

        return full_text

    def _convert_table(self, table):
        """转换表格"""
        md_lines = ['\n']

        rows = table.rows
        for i, row in enumerate(rows):
            cells = []
            for cell in row.cells:
                # 处理单元格内的段落
                cell_text = ''
                for para in cell.paragraphs:
                    cell_text += para.text.strip() + ' '
                cells.append(cell_text.strip())
            md_lines.append('| ' + ' | '.join(cells) + ' |')

            # 表头分隔线
            if i == 0:
                md_lines.append('| ' + ' | '.join(['---'] * len(cells)) + ' |')

        return '\n'.join(md_lines) + '\n\n'

    def _convert_lists(self, element):
        """转换列表"""
        md_lines = []
        for item in element.items:
            text = item.text.strip()
            if item.style.name == 'List Bullet':
                md_lines.append(f"- {text}\n")
            elif item.style.name == 'List Number':
                md_lines.append(f"1. {text}\n")
        return ''.join(md_lines) + '\n'

    def _convert_document(self, doc):
        """转换整个文档"""
        md_content = []

        for element in doc.element.body:
            if element.tag.endswith('p'):  # 段落
                para = element._p
                style_name = para.style.name if para.style else 'Normal'
                if style_name.startswith('Heading'):
                    level = style_name.split()[-1]
                    text = para.text.strip()
                    md_content.append(f"{'#' * int(level)} {text}\n\n")
                else:
                    text = para.text.strip()
                    if text:
                        md_content.append(f"{text}\n\n")

            elif element.tag.endswith('tbl'):  # 表格
                table = element._tbl
                doc_table = None
                for t in doc.tables:
                    if t._tbl == table:
                        doc_table = t
                        break
                if doc_table:
                    md_content.append(self._convert_table(doc_table))

        return ''.join(md_content)

    def _generate_summary(self, doc):
        """生成转换统计"""
        return {
            'total_images': self.image_counter,
            'total_paragraphs': len(doc.paragraphs),
            'total_tables': len(doc.tables)
        }


# 使用示例
if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("用法: python word_to_markdown.py <docx文件路径>")
        sys.exit(1)

    converter = WordToMarkdown(sys.argv[1])
    result = converter.convert()

    print(json.dumps(result, indent=2, ensure_ascii=False))
```

## 七、企业级增强版

在基础转换之上，可选添加：

```python
class EnterpriseWordCleaner(WordToMarkdown):
    """企业级 Word 知识库清洗器"""

    def __init__(self, file_path):
        super().__init__(file_path, mode='enterprise')

    def clean_content(self, md_content):
        """清洗内容"""
        # 删除目录页
        md_content = re.sub(r'^#* 目录.*\n', '', md_content, flags=re.MULTILINE)
        md_content = re.sub(r'^#* Table of Contents.*\n', '', md_content, flags=re.MULTILINE)

        # 删除页码
        md_content = re.sub(r'^\s*\[?\d+\]?\s*$', '', md_content, flags=re.MULTILINE)

        # 删除空白页标记
        md_content = re.sub(r'^---.*page break.*---\n', '', md_content, flags=re.IGNORECASE)

        # 合并分段标题
        md_content = re.sub(r'\n(#+) ([^\n]+)\n\1 ', r'\n\1 \2\n', md_content)

        return md_content

    def split_sections(self, md_content):
        """自动切分章节"""
        sections = re.split(r'^(#{1,2}) ', md_content, flags=re.MULTILINE)
        # 返回章节列表
        return sections

    def generate_faq(self, content):
        """从内容自动生成 FAQ"""
        # 简单的 FAQ 生成逻辑
        faq = []
        # ... FAQ 生成逻辑
        return faq
```

## 八、最佳实践

1. **保持目录结构**
   ```
   /output_dir/
   ├── document.md
   └── /images/
       ├── img001.png
       └── img002.jpg
   ```

2. **图片命名规范**
   - 使用序号：`img001.png`, `img002.jpg`
   - 避免原始文件名（含中文、空格）

3. **Markdown 格式**
   - 标题层级清晰
   - 表格有表头分隔线
   - 列表使用 `-` 或 `1.`

4. **错误处理**
   - 文档损坏时给出明确提示
   - 图片提取失败时继续处理文本

## 九、CLI 使用

```bash
# 基础转换（支持 .doc 和 .docx）
python word_to_markdown.py document.docx
python word_to_markdown.py document.doc   # 自动转 docx 再处理

# 指定输出目录
python word_to_markdown.py document.docx -o /path/to/output

# 指定 LibreOffice 路径（用于 .doc 转换）
python word_to_markdown.py document.doc --libreoffice /Applications/LibreOffice.app/Contents/MacOS/soffice

# 仅提取图片
python word_to_markdown.py document.docx --images-only

# 跳过内容清洗
python word_to_markdown.py document.docx --no-clean

# 显示详细信息
python word_to_markdown.py document.docx -v

# 批量转换（支持多个文件）
python word_to_markdown.py file1.docx file2.docx file3.docx

# 批量转换指定输出目录
python word_to_markdown.py file1.docx file2.docx -o /path/to/output

# 批量转换混合格式（.doc 和 .docx）
python word_to_markdown.py doc1.doc doc2.docx doc3.doc
```

## 十、与其他技能的配合

```
Word 转 Markdown
       ↓
  PDF 转 Markdown
       ↓
 Excel 转知识库
       ↓
 PPT 转培训资料
       ↓
  网页转知识库
```

---

*本技能基于 python-docx、mammoth、pandoc 等库实现*
