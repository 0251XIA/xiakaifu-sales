#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 转 Markdown 转换器 v2.1
支持 .doc/.docx 格式，增强表格处理

依赖:
    - LibreOffice (用于转换 .doc 格式)
    - Python 内置库 (无需安装额外依赖)

用法:
    python word_to_markdown.py <doc文件路径> [选项]

选项:
    -o, --output DIR      指定输出目录
    --libreoffice PATH    LibreOffice 命令路径 (默认: /Applications/LibreOffice.app/Contents/MacOS/soffice)
    --images-only         仅提取图片
    --md-only            仅生成 Markdown（不提取图片）
    --no-clean           跳过内容清洗
    -v, --verbose        显示详细信息

示例:
    python word_to_markdown.py 文档.docx
    python word_to_markdown.py 文档.doc      # 自动转 docx
    python word_to_markdown.py 文档.docx -o ./output
    python word_to_markdown.py 文档.doc --libreoffice /Applications/LibreOffice.app/Contents/MacOS/soffice
"""

import os
import sys
import zipfile
import json
import re
import argparse
import subprocess
import shutil
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime


# XML 命名空间
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}

# 注册命名空间以避免 ns0 前缀
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


class WordToMarkdown:
    """Word 文档转 Markdown 转换器"""

    def __init__(self, file_path, output_dir=None, extract_images=True,
                 clean_content=True, verbose=False):
        self.file_path = Path(file_path).resolve()
        self.extract_images = extract_images
        self.clean_content = clean_content
        self.verbose = verbose

        # 输出目录
        if output_dir:
            self.output_dir = Path(output_dir).resolve()
        else:
            self.output_dir = self.file_path.parent / f"{self.file_path.stem}_output"

        self.images_dir = self.output_dir / 'images'
        self.md_file = self.output_dir / f"{self.file_path.stem}.md"

        self.images = []
        self.image_counter = 0
        self.heading_count = 0
        self.paragraph_count = 0
        self.table_count = 0
        self.headers_footers = []

    def convert(self):
        """执行转换"""
        self._log(f"📂 输入文件: {self.file_path}")
        self._log(f"📁 输出目录: {self.output_dir}")

        # 创建输出目录
        self.output_dir.mkdir(parents=True, exist_ok=True)
        if self.extract_images:
            self.images_dir.mkdir(exist_ok=True)

        # 提取图片
        if self.extract_images:
            self._extract_images()

        # 解析并转换文档
        md_content = self._parse_document()

        # 插入图片引用
        if self.extract_images:
            md_content = self._insert_image_references(md_content)

        # 内容清洗
        if self.clean_content:
            md_content = self._clean_content(md_content)

        # 保存 Markdown
        self.md_file.write_text(md_content, encoding='utf-8')

        # 生成报告
        result = {
            'success': True,
            'md_file': str(self.md_file),
            'images_dir': str(self.images_dir) if self.extract_images else None,
            'images': [img['saved_name'] for img in self.images],
            'headers_footers': self.headers_footers if self.headers_footers else None,
            'summary': {
                'total_images': self.image_counter,
                'total_headings': self.heading_count,
                'total_paragraphs': self.paragraph_count,
                'total_tables': self.table_count,
                'total_headers_footers': len(self.headers_footers),
            }
        }

        return result

    def _log(self, msg):
        """打印日志"""
        if self.verbose:
            print(msg)

    def _extract_images(self):
        """提取文档中的所有图片"""
        self._log("🖼️  正在提取图片...")

        with zipfile.ZipFile(self.file_path, 'r') as z:
            for fname in z.namelist():
                if fname.startswith('word/media/'):
                    ext = os.path.splitext(fname)[1] or '.png'
                    self.image_counter += 1
                    new_name = f"img{self.image_counter:03d}{ext}"
                    out_path = self.images_dir / new_name

                    with z.open(fname) as src:
                        with open(out_path, 'wb') as dst:
                            dst.write(src.read())

                    self.images.append({
                        'original_name': fname.split('/')[-1],
                        'saved_name': new_name,
                        'path': str(out_path)
                    })
                    self._log(f"   提取: {new_name}")

    def _parse_document(self):
        """解析 Word 文档"""
        self._log("📄 正在解析文档...")

        with zipfile.ZipFile(self.file_path, 'r') as z:
            xml_content = z.read('word/document.xml')

            # 提取页眉页脚
            self._extract_headers_footers(z)

        root = ET.fromstring(xml_content)
        body = root.find('.//w:body', NS)

        md_lines = []
        for elem in body:
            tag = elem.tag.split('}')[-1]

            if tag == 'p':  # 段落
                para_md = self._convert_paragraph(elem)
                if para_md:
                    md_lines.append(para_md)
                    self.paragraph_count += 1

            elif tag == 'tbl':  # 表格
                tbl_md = self._convert_table(elem)
                md_lines.append(tbl_md)
                self.table_count += 1

        return '\n'.join(md_lines)

    def _extract_headers_footers(self, zip_ref):
        """提取页眉页脚"""
        self._log("📑 正在提取页眉页脚...")

        # 查找所有页眉页脚文件
        for fname in zip_ref.namelist():
            if fname.startswith('word/header') or fname.startswith('word/footer'):
                try:
                    xml_content = zip_ref.read(fname)
                    root = ET.fromstring(xml_content)
                    body = root.find('.//w:body', NS)
                    if body is not None:
                        for elem in body:
                            if elem.tag.endswith('p'):
                                text = self._get_paragraph_text(elem).strip()
                                if text and len(text) < 200:  # 过滤掉正文内容
                                    self.headers_footers.append(text)
                                    self._log(f"   提取: {fname} - {text[:30]}...")
                except (ET.ParseError, KeyError, zipfile.BadZipFile, AttributeError) as e:
                    self._log(f"   ⚠️  无法解析 {fname}: {e}")

        # 提取脚注和尾注
        for fname in zip_ref.namelist():
            if 'footnotes' in fname or 'endnotes' in fname:
                try:
                    xml_content = zip_ref.read(fname)
                    root = ET.fromstring(xml_content)
                    for elem in root.iter():
                        if elem.tag.endswith('}t'):
                            text = elem.text
                            if text and len(text) > 5:
                                self.headers_footers.append(f"[脚注] {text}")
                except (ET.ParseError, KeyError, zipfile.BadZipFile) as e:
                    self._log(f"   ⚠️  无法解析 {fname}: {e}")

    def _convert_paragraph(self, para):
        """转换单个段落"""
        style = self._get_paragraph_style(para)
        text = self._get_paragraph_text(para)

        if not text:
            return ''

        # 标题
        if style.startswith('Heading'):
            try:
                level = int(style.replace('Heading', ''))
                self.heading_count += 1
                return f"{'#' * level} {text}\n"
            except ValueError:
                return f"## {text}\n"

        # 列表项
        if style == 'ListParagraph':
            numPr = para.find('.//w:numPr', NS)
            if numPr is not None:
                return f"- {text}\n"

        # 引用
        if style == 'Quote':
            return f"> {text}\n"

        return f"{text}\n"

    def _get_paragraph_style(self, para):
        """获取段落样式"""
        pPr = para.find('w:pPr', NS)
        if pPr is not None:
            pStyle = pPr.find('w:pStyle', NS)
            if pStyle is not None:
                return pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'Normal')
        return 'Normal'

    def _get_paragraph_text(self, para):
        """获取段落文本，处理加粗、斜体等"""
        # 先获取所有超链接
        hyperlinks = {}
        for hl in para.findall('.//w:hyperlink', NS):
            r_id = hl.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if r_id:
                text = ''.join([t.text or '' for t in hl.findall('.//w:t', NS)])
                hyperlinks[id(hl)] = (r_id, text)

        parts = []
        for elem in para.iter():
            # 文本
            if elem.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t':
                text = elem.text or ''
                # 检查是否在超链接中
                parent = self._find_parent(elem, para)
                if parent is not None and parent.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink':
                    r_id = parent.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if r_id:
                        text = f'[{text}]({self._get_url_by_id(r_id)})'
                parts.append(text)

            # 制表符
            elif elem.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab':
                parts.append('\t')

            # 换行
            elif elem.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br':
                parts.append('\n')

        # 处理行内格式
        full_text = ''.join(parts)
        full_text = self._process_inline_formatting(para, full_text)

        return full_text.strip()

    def _find_parent(self, elem, root):
        """查找元素的父元素"""
        parents = {}
        for parent in root.iter():
            for child in parent:
                parents[id(child)] = parent
        return parents.get(id(elem))

    def _get_url_by_id(self, r_id):
        """通过关系ID获取URL"""
        try:
            with zipfile.ZipFile(self.file_path, 'r') as z:
                rels_content = z.read('word/_rels/document.xml.rels')
            rels_root = ET.fromstring(rels_content)
            for rel in rels_root:
                if rel.get('Id') == r_id:
                    return rel.get('Target', '')
        except (ET.ParseError, KeyError, IOError, OSError) as e:
            self._log(f"   ⚠️  解析超链接失败: {e}")
        return '[BROKEN_LINK]'

    def _process_inline_formatting(self, para, text):
        """处理加粗、斜体等行内格式"""
        # 检查是否整段都是加粗（可能是标题）
        is_entirely_bold = self._is_entirely_bold_paragraph(para)

        if is_entirely_bold and len(text) > 0 and len(text) < 100:
            # 整段加粗且较短，可能是标题
            # 根据长度判断级别
            if len(text) < 10:
                level = 1
            elif len(text) < 20:
                level = 2
            else:
                level = 3
            # 去掉末尾冒号（如果有）
            clean_text = text.rstrip('：:').strip()
            return f"{'#' * level} {clean_text}"

        # 正常处理行内格式
        result = []
        current_bold = False
        current_italic = False

        for run in para.findall('.//w:r', NS):
            # 检查格式
            rPr = run.find('w:rPr', NS)
            if rPr is not None:
                if rPr.find('w:b', NS) is not None:
                    current_bold = True
                if rPr.find('w:i', NS) is not None:
                    current_italic = True

            # 获取文本
            for t in run.findall('w:t', NS):
                run_text = t.text or ''
                if current_bold:
                    run_text = f"**{run_text}**"
                if current_italic:
                    run_text = f"*{run_text}*"
                result.append(run_text)

        if result:
            return ''.join(result)
        return text

    def _is_entirely_bold_paragraph(self, para):
        """检查段落是否全部加粗"""
        has_bold = False
        has_normal = False

        for run in para.findall('.//w:r', NS):
            rPr = run.find('w:rPr', NS)
            if rPr is not None:
                if rPr.find('w:b', NS) is not None:
                    has_bold = True
                else:
                    has_normal = True
            else:
                has_normal = True

        return has_bold and not has_normal

    def _convert_table(self, tbl):
        """转换表格，支持合并单元格"""
        md_lines = ['\n']

        # 获取所有行
        rows = tbl.findall('.//w:tr', NS)

        # 分析合并单元格信息
        merged_cells_info = self._analyze_table_merges(tbl, rows)

        for i, row in enumerate(rows):
            cells = []
            cell_idx = 0

            for cell in row.findall('w:tc', NS):
                # 检查是否是被合并的单元格（需要跳过）
                grid_span = cell.find('.//w:gridSpan', NS)
                if grid_span is not None:
                    colspan = int(grid_span.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 1))
                else:
                    colspan = 1

                # 获取垂直合并信息
                vMerge = cell.find('.//w:vMerge', NS)
                vMerge_val = None
                if vMerge is not None:
                    vMerge_val = vMerge.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')

                # 获取单元格文本
                cell_text = ''
                for p in cell.findall('.//w:p', NS):
                    cell_text += self._get_paragraph_text(p) + ' '
                cell_text = cell_text.strip()

                # 处理水平合并
                if colspan > 1:
                    cells.append(f"{cell_text}")
                    for _ in range(colspan - 1):
                        cells.append("")  # 空单元格用于 colspan
                elif vMerge_val == 'continue':
                    # 垂直合并的继续单元格，跳过
                    continue
                else:
                    cells.append(cell_text)

            # 构建 Markdown 表格行
            if any(cells):  # 跳过空行
                md_lines.append('| ' + ' | '.join(cells) + ' |')

                # 表头分隔线（第一行之后）
                if i == 0:
                    # 计算实际列数
                    col_count = len([c for c in cells if c])
                    md_lines.append('| ' + ' | '.join(['---'] * col_count) + ' |')

        md_lines.append('')
        return '\n'.join(md_lines)

    def _analyze_table_merges(self, tbl, rows):
        """分析表格的合并单元格信息"""
        merges = []
        for i, row in enumerate(rows):
            for cell in row.findall('w:tc', NS):
                # 水平合并
                gridSpan = cell.find('.//w:gridSpan', NS)
                if gridSpan is not None:
                    colspan = int(gridSpan.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 1))

                # 垂直合并
                vMerge = cell.find('.//w:vMerge', NS)
                if vMerge is not None:
                    val = vMerge.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    if val == 'continue':
                        merges.append(('v', i, 'continue'))

        return merges

    def _clean_content(self, md_content):
        """清洗内容（页眉、页脚、页码等）"""
        self._log("🧹 正在清洗内容...")

        # 删除目录标题
        md_content = re.sub(r'^#+\s*(目录|Table of Contents|TOC)\s*$',
                           '', md_content, flags=re.MULTILINE)

        # 删除孤立的多余空行（超过2个连续空行）
        md_content = re.sub(r'\n{3,}', '\n\n', md_content)

        # 清理行首行尾空白
        lines = [line.strip() for line in md_content.split('\n')]
        md_content = '\n'.join(lines)

        # 删除只有页码的行（多种格式）
        page_patterns = [
            r'^[\[\]【】\d\.\)]+\s*$',           # [1] 或 1. 或 1) 或 【1】
            r'^第\s*\d+\s*页\s*(/\s*共\s*\d+\s*页)?\s*$',  # 第 1 页 / 共 10 页
            r'^Page\s+\d+\s*(of\s+\d+)?\s*$',     # Page 1 of 10
            r'^\d+\s*/\s*\d+\s*$',                # 1/10
            r'^-\s*(\d+)\s*-\s*$',                # - 1 -
            r'^页眉\s*:.*$',                       # 页眉: xxx
            r'^Footer\s*:.*$',                    # Footer: xxx
            r'^Header\s*:.*$',                    # Header: xxx
        ]
        for pattern in page_patterns:
            md_content = re.sub(pattern, '', md_content, flags=re.IGNORECASE | re.MULTILINE)

        # 删除页眉页脚残留标记
        header_footer_patterns = [
            r'^\{.*?\}$',                         # { xxx }
            r'^<.*?>$',                           # < xxx >
            r'^\[DOCPROPERTY.*?\]$',             # [DOCPROPERTY ...]
            r'^Skimlinks.*$',                    # Skimlinks 残留
        ]
        for pattern in header_footer_patterns:
            md_content = re.sub(pattern, '', md_content, flags=re.MULTILINE)

        # 删除公司名称/文档标题类页眉（常见格式）
        company_patterns = [
            r'^[^\u4e00-\u9fa5]*公司[^\u4e00-\u9fa5]*$',  # xxx公司xxx
            r'^[^\u4e00-\u9fa5]*制度[^\u4e00-\u9fa5]*$',  # xxx制度xxx
            r'^[^\u4e00-\u9fa5]*手册[^\u4e00-\u9fa5]*$',  # xxx手册xxx
            r'^内部资料\s*$',                         # 内部资料
            r'^机密\s*$',                             # 机密
            r'^绝密\s*$',                             # 绝密
        ]
        for pattern in company_patterns:
            md_content = re.sub(pattern, '', md_content, flags=re.MULTILINE)

        # 清理多余空行
        md_content = re.sub(r'\n{3,}', '\n\n', md_content)

        return md_content.strip() + '\n'

    def _insert_image_references(self, md_content):
        """在图片说明位置插入实际图片引用"""
        self._log("🖼️  正在插入图片引用...")

        # 图片说明匹配模式（多种格式）
        caption_patterns = [
            (r'（下图为([^）]+)）', 1),           # （下图为xxx）
            (r'（如下图([^）]+)）', 1),            # （如下图xxx）
            (r'（参考图片：([^）]+)）', 1),        # （参考图片：xxx）
            (r'（图片示例：([^）]+)）', 1),        # （图片示例：xxx）
            (r'（纯图片）', 0),                     # （纯图片）无说明
            (r'!\[([^\]]*)\]', 1),                  # ![xxx] Markdown图片语法
        ]

        # 用于记录下一个要使用的图片索引
        image_index = 0

        def make_replacement(match, group_idx=1):
            nonlocal image_index
            if image_index < len(self.images):
                # 获取图片说明文字
                if group_idx > 0 and group_idx <= len(match.groups()):
                    caption_text = match.group(group_idx).strip().replace('**', '')
                else:
                    caption_text = f"图片{image_index + 1}"

                if not caption_text:
                    caption_text = f"图片{image_index + 1}"

                img_name = self.images[image_index]['saved_name']
                img_path = f"images/{img_name}"
                image_index += 1
                self._log(f"   插入图片: {img_name} - {caption_text}")
                return f'\n![{caption_text}]({img_path})\n'
            return match.group(0)

        # 依次应用所有模式
        for pattern, group_idx in caption_patterns:
            def make_repl(m, gi=group_idx):
                return make_replacement(m, gi)
            md_content = re.sub(pattern, make_repl, md_content)

        # 清理孤立的 **（不成对的加粗标记）
        lines = md_content.split('\n')
        cleaned_lines = []
        for line in lines:
            bold_count = line.count('**')
            if bold_count % 2 != 0 and bold_count > 0:
                line = line.replace('**', '')
            cleaned_lines.append(line)

        md_content = '\n'.join(cleaned_lines)

        # 清理多余空行
        md_content = re.sub(r'\n{3,}', '\n\n', md_content)

        # 如果还有未使用的图片，追加到结尾
        if image_index < len(self.images):
            remaining = self.images[image_index:]
            md_content += '\n\n---\n\n## 图片\n\n'
            for img in remaining:
                img_path = f"images/{img['saved_name']}"
                md_content += f"![{img['saved_name']}]({img_path})\n"
            self._log(f"   追加 {len(remaining)} 张未关联图片到文档末尾")

        return md_content

    def extract_images_only(self):
        """仅提取图片"""
        self._log("🖼️  仅提取图片模式")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.images_dir.mkdir(exist_ok=True)
        self._extract_images()

        return {
            'success': True,
            'images_dir': str(self.images_dir),
            'images': [img['saved_name'] for img in self.images],
            'summary': {
                'total_images': self.image_counter,
            }
        }


def main():
    parser = argparse.ArgumentParser(
        description='Word 转 Markdown 转换器',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument('files', nargs='+', help='Word 文档路径 (.doc/.docx)，支持多个文件批量转换')
    parser.add_argument('-o', '--output', help='输出目录（默认：源文件同级目录的 output 文件夹）')
    parser.add_argument('--libreoffice', default='/Applications/LibreOffice.app/Contents/MacOS/soffice',
                       help='LibreOffice 命令路径 (默认: /Applications/LibreOffice.app/Contents/MacOS/soffice)')
    parser.add_argument('--images-only', action='store_true',
                       help='仅提取图片')
    parser.add_argument('--md-only', action='store_true',
                       help='仅生成 Markdown（不提取图片）')
    parser.add_argument('--no-clean', action='store_true',
                       help='跳过内容清洗')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='显示详细信息')

    args = parser.parse_args()

    # 批量转换模式
    is_batch = len(args.files) > 1
    total_files = len(args.files)
    success_count = 0
    fail_count = 0
    results_summary = []

    for idx, file_path_str in enumerate(args.files, 1):
        file_path = Path(file_path_str)

        if is_batch:
            print(f"\n{'='*50}")
            print(f"📦 处理文件 [{idx}/{total_files}]: {file_path.name}")
            print('='*50)

        # 检查文件
        if not os.path.exists(file_path):
            print(f"❌ 文件不存在: {file_path}")
            fail_count += 1
            results_summary.append({'file': str(file_path), 'success': False, 'error': '文件不存在'})
            continue

        original_file = file_path

        # .doc 格式需要先转换为 .docx
        if file_path.suffix.lower() == '.doc':
            print("📄 检测到 .doc 格式，正在转换为 .docx...")
            try:
                file_path = convert_doc_to_docx(file_path, args.libreoffice, args.verbose)
                print(f"✅ 转换成功: {file_path}")
            except FileNotFoundError:
                print("❌ 未找到 LibreOffice，无法转换 .doc 格式")
                print("")
                print("💡 解决方案（二选一）：")
                print("   1. 安装 LibreOffice:")
                print("      macOS: brew install --cask libreoffice")
                print("      Ubuntu: sudo apt install libreoffice")
                print("   2. 将 .doc 文档在 Word/WPS 中另存为 .docx 格式")
                print("")
                print("📝 .docx 格式无需额外依赖，可直接转换。")
                fail_count += 1
                results_summary.append({'file': str(file_path), 'success': False, 'error': 'LibreOffice 未安装'})
                continue
            except Exception as e:
                print(f"❌ .doc 转换失败: {e}")
                print("💡 提示: 请将文档另存为 .docx 格式后重试")
                fail_count += 1
                results_summary.append({'file': str(file_path), 'success': False, 'error': str(e)})
                continue

        elif file_path.suffix.lower() != '.docx':
            print(f"❌ 请提供 .doc 或 .docx 文件: {file_path}")
            fail_count += 1
            results_summary.append({'file': str(file_path), 'success': False, 'error': '不支持的文件格式'})
            continue

        # 创建转换器
        converter = WordToMarkdown(
            file_path=str(file_path),
            output_dir=args.output,
            extract_images=not args.md_only,
            clean_content=not args.no_clean,
            verbose=args.verbose
        )

        # 执行转换
        try:
            if args.images_only:
                result = converter.extract_images_only()
            else:
                result = converter.convert()

            success_count += 1
            results_summary.append({
                'file': str(file_path),
                'success': True,
                'md_file': result['md_file'],
                'images_dir': result.get('images_dir', ''),
                'summary': result.get('summary', {})
            })

            if is_batch:
                print(f"✅ 转换完成: {result['md_file']}")

        except Exception as e:
            fail_count += 1
            results_summary.append({'file': str(file_path), 'success': False, 'error': str(e)})
            print(f"❌ 转换失败: {e}")
            import traceback
            if args.verbose:
                traceback.print_exc()

    # 汇总报告
    print("\n" + "=" * 50)
    if is_batch:
        print(f"📊 批量转换完成: {success_count}/{total_files} 成功")
        print("=" * 50)
        for r in results_summary:
            status = "✅" if r['success'] else "❌"
            print(f"{status} {Path(r['file']).name}")
            if r['success'] and 'md_file' in r:
                print(f"   → {r['md_file']}")
            else:
                print(f"   → {r.get('error', '未知错误')}")
    else:
        # 单文件模式（保持原有输出格式）
        if success_count == 1:
            result = results_summary[0]
            print("✅ 转换完成!")
            print("=" * 50)
            print(f"📄 Markdown: {result['md_file']}")
            if result.get('images_dir'):
                print(f"📁 图片目录: {result['images_dir']}")
            s = result.get('summary', {})
            print(f"🖼️  图片数量: {s.get('total_images', 0)}")
            print(f"📝 段落数量: {s.get('total_paragraphs', 'N/A')}")
            print(f"📋 表格数量: {s.get('total_tables', 'N/A')}")
            print(f"📌 标题数量: {s.get('total_headings', 'N/A')}")
        else:
            print("❌ 转换失败")
            sys.exit(1)


def convert_doc_to_docx(doc_path, libreoffice_cmd='/Applications/LibreOffice.app/Contents/MacOS/soffice', verbose=False, timeout=120):
    """使用 LibreOffice 将 .doc 转换为 .docx"""
    doc_path = Path(doc_path).resolve()

    # 安全检查：验证 libreoffice_cmd 是合法路径
    # 防止命令注入风险（只允许绝对路径或已知可执行文件名）
    cmd_path = Path(libreoffice_cmd)
    if not cmd_path.is_absolute():
        # 如果是相对路径，只允许已知的安全命令名
        allowed_commands = {'lowriter', 'soffice', 'libreoffice'}
        if cmd_path.name not in allowed_commands:
            raise ValueError(f"不支持的 LibreOffice 命令: {libreoffice_cmd}")
    elif not cmd_path.exists():
        raise FileNotFoundError(f"LibreOffice 路径不存在: {libreoffice_cmd}")

    # 检查 LibreOffice 是否存在
    if not shutil.which(libreoffice_cmd):
        raise FileNotFoundError(f"LibreOffice 未安装或路径不正确: {libreoffice_cmd}")

    # 创建临时目录用于转换
    with tempfile.TemporaryDirectory() as tmpdir:
        # 复制文件到临时目录（避免路径问题）
        tmp_input = Path(tmpdir) / doc_path.name
        tmp_input.write_bytes(doc_path.read_bytes())

        # 调用 LibreOffice 转换
        cmd = [
            libreoffice_cmd,
            '--headless',
            '--convert-to', 'docx',
            '--outdir', tmpdir,
            str(tmp_input)
        ]

        if verbose:
            print(f"🔄 执行: {' '.join(cmd)}")

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired:
            raise TimeoutError(f"LibreOffice 转换超时（{timeout}秒），文件可能已损坏")

        if result.returncode != 0:
            if verbose:
                print(f"❌ LibreOffice 错误: {result.stderr}")
            raise RuntimeError(f"LibreOffice 转换失败 (code {result.returncode})")

        # 查找转换后的文件
        expected_output = Path(tmpdir) / f"{doc_path.stem}.docx"
        if expected_output.exists():
            # 复制回原目录
            output_path = doc_path.parent / f"{doc_path.stem}_converted.docx"
            output_path.write_bytes(expected_output.read_bytes())
            return output_path
        else:
            # 尝试查找任何 .docx 文件
            docx_files = list(Path(tmpdir).glob('*.docx'))
            if docx_files:
                output_path = doc_path.parent / f"{doc_path.stem}_converted.docx"
                output_path.write_bytes(docx_files[0].read_bytes())
                return output_path

            raise RuntimeError("LibreOffice 未能生成 .docx 文件")


if __name__ == '__main__':
    main()
