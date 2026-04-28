---
name: ppt-to-markdown
description: 将PPT演示文稿(.ppt/.pptx)转换为Markdown格式，支持幻灯片提取、文本框逻辑重组、图片提取、演讲者备注、SmartArt和图表支持
trigger:
  - 上传ppt
  - PPT转MD
  - ppt转markdown
  - pptx转换
  - ppt转换
input:
  - file: PPT文件路径(.ppt/.pptx)
  - mode: standard | enterprise (可选，默认为standard)
output:
  - markdown: 转换后的Markdown文件
  - images: 提取的图片列表
  - notes: 演讲者备注列表
  - summary: 转换统计信息
tools:
  - python-pptx (可选，非必须)
  - python-pptx
  - LibreOffice (用于.ppt格式转换)
---

# PPT 图文转 Markdown 转换技能 v1.0

## 一、适用范围

用于将 `.ppt` / `.pptx` PowerPoint 演示文稿转换为结构化 Markdown 文档：

- 幻灯片标题提取
- 文本框内容按逻辑顺序重组
- 图片提取与位置还原
- 演讲者备注提取
- SmartArt 图形转嵌套列表
- 图表处理（表格化或截图）
- 列表识别（有序/无序）

### 典型应用场景

- 培训课件整理
- 产品演示文稿知识库化
- 会议纪要标准化
- AI 知识库喂料
- PPT 内容复用

## 二、技术方案

### 2.1 格式分析

`.pptx` 是 ZIP + XML 格式，内部结构：

```
.pptx (ZIP包)
├── ppt/slides/slide1.xml      # 幻灯片内容
├── ppt/slides/slide2.xml
├── ppt/slides/_rels/          # 幻灯片关系
├── ppt/media/                  # 媒体文件
├── ppt/layouts/               # 布局
├── ppt/theme/                 # 主题
├── ppt/notesSlides/           # 演讲者备注
├── ppt/notesSlides/_rels/
├── ppt/charts/                # 图表
├── ppt/smartArts/             # SmartArt
└── [Content_Types].xml
```

### 2.2 核心库选择

```python
# 方案一：python-pptx（需要安装）
from pptx import Presentation

# 方案二：纯 Python XML 解析（无需额外依赖）
import zipfile
import xml.etree.ElementTree as ET
```

**推荐方案二**，原因与 word-to-markdown 相同：减少依赖，更精细控制。

### 2.3 逻辑顺序重组

文本框不还原位置，按逻辑顺序排列：

1. 提取每页所有文本框的文本内容
2. 按出现的自然顺序（XML 中的位置）重组
3. 识别标题（字号最大或占位符形状）
4. 识别列表项（相同前缀如 "1."、"- "）

```python
def extract_slide_content(slide_xml):
    """按逻辑顺序提取幻灯片内容"""
    texts = []
    for shape in slide_xml.findall('.//p:sp'):
        # 跳过标题占位符
        # 提取文本
        # 按 XML 顺序加入列表
    return texts
```

### 2.4 SmartArt 转换

```python
def convert_smartart(smartart_path):
    """SmartArt 转为嵌套列表"""
    # 解析 diagram 目录下的 XML
    # 提取节点层级关系
    # 输出嵌套列表
```

### 2.5 图表处理

```python
def convert_chart(chart_path):
    """图表处理策略"""
    # 方案A: 保留为图片引用
    # 方案B: 提取数据为 Markdown 表格
    # 方案C: 生成图表描述
```

### 2.6 演讲者备注提取

```python
def extract_notes(notes_slide_xml):
    """提取演讲者备注"""
    # 从 notesSlides/slide*.xml 提取
    # 转换为脚注格式
```

---

## 三、技能流程设计

```
上传 PPT
    ↓
解析 pptx 结构
    ↓
遍历每张幻灯片
    ↓
提取标题和文本框
    ↓
按逻辑顺序重组
    ↓
提取图片
    ↓
处理 SmartArt
    ↓
处理图表
    ↓
提取演讲者备注
    ↓
输出 Markdown
```

---

## 四、输入输出

### 输入

```json
{
  "file_path": "xxx.pptx",
  "mode": "standard",
  "extract_images": true,
  "extract_notes": true,
  "image_prefix": "img"
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
    }
  ],
  "notes": [
    {"slide": 1, "note": "这是备注内容..."}
  ],
  "summary": {
    "total_slides": 15,
    "total_images": 8,
    "total_notes": 10,
    "total_smartarts": 2
  }
}
```

---

## 五、Markdown 输出格式

```markdown
# 演示文稿标题

## 第 1 页：课程介绍

这是第一页的正文内容。

- 要点1
- 要点2
  - 子要点2.1
  - 子要点2.2

![图片描述](images/img001.png)

> 备注：这里是演讲者备注内容[^1]

## 第 2 页：操作步骤

1. 步骤一
2. 步骤二
3. 步骤三

![操作图](images/img002.png)

[^1]: 第1页备注：这里要重点强调
```

---

## 六、CLI 使用

```bash
# 基础转换（支持 .ppt 和 .pptx）
python ppt_to_markdown.py presentation.pptx
python ppt_to_markdown.py presentation.ppt   # 自动转 pptx 再处理

# 指定输出目录
python ppt_to_markdown.py presentation.pptx -o /path/to/output

# 指定 LibreOffice 路径
python ppt_to_markdown.py presentation.ppt --libreoffice /Applications/LibreOffice.app/Contents/MacOS/soffice

# 仅提取图片
python ppt_to_markdown.py presentation.pptx --images-only

# 跳过备注提取
python ppt_to_markdown.py presentation.pptx --no-notes

# 跳过 SmartArt 处理
python ppt_to_markdown.py presentation.pptx --no-smartart

# 导出每页幻灯片截图（需要 LibreOffice）
python ppt_to_markdown.py presentation.pptx --slide-images

# 显示详细信息
python ppt_to_markdown.py presentation.pptx -v

# 批量转换
python ppt_to_markdown.py file1.pptx file2.pptx file3.pptx
```

---

## 七、与其他技能的配合

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

*本技能基于纯 Python XML 解析和 python-pptx（可选）实现*
