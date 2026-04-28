#!/usr/bin/env python3
"""
PDF 转 Markdown 转换器 v1.0

将 PDF 文档转换为结构化 Markdown，保留文字格式和基本表格。

依赖:
    - pdfminer.six (pip install pdfminer.six)
    - Python 内置库

可选依赖:
    - PyMuPDF (pip install pymupdf) - 用于图片提取和页面截图
    - camelot-py (pip install camelot-py) - 用于表格提取

用法:
    python pdf_to_markdown.py <pdf文件> [选项]

选项:
    -o, --output DIR      指定输出目录
    --no-tables           跳过表格提取
    --no-images           跳过图片提取
    --images-only         仅提取图片
    --slide-images        导出每页截图（需要 PyMuPDF）
    -v, --verbose         显示详细信息
"""

import os
import re
import sys
import json
import shutil
import argparse
import tempfile
from pathlib import Path
from typing import Any

# pdfminer.six - 文字提取（必需）
try:
    from pdfminer.high_level import extract_text
    from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTChar, LTAnno
    from pdfminer.pdfpage import PDFPage
    from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
    from pdfminer.converter import PDFPageAggregator
    PDFMINER_AVAILABLE = True
except ImportError:
    print("❌ 缺少 pdfminer.six，请先安装：pip install pdfminer.six")
    sys.exit(1)

# PyMuPDF - 图片提取+截图（可选）
try:
    import fitz
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# camelot - 表格提取（可选）
try:
    import camelot
    CAMELOT_AVAILABLE = True
except ImportError:
    CAMELOT_AVAILABLE = False


class PDFToMarkdown:
    """PDF 文档转 Markdown 转换器"""

    def __init__(
        self,
        file_path: str,
        output_dir: str | None = None,
        extract_tables: bool = True,
        extract_images: bool = True,
        extract_slide_images: bool = False,
        verbose: bool = False
    ):
        self.file_path = Path(file_path)
        self.extract_tables = extract_tables
        self.extract_images = extract_images and PYMUPDF_AVAILABLE
        self.extract_slide_images = extract_slide_images and PYMUPDF_AVAILABLE
        self.verbose = verbose

        # 设置输出目录
        if output_dir:
            self.output_dir = Path(output_dir)
        else:
            self.output_dir = self.file_path.parent / f"{self.file_path.stem}_output"

        self.images_dir = self.output_dir / "images"
        self.slide_images_dir = self.output_dir / "slide_images"

        # 统计
        self.image_counter = 0
        self.images = []
        self.slide_images_by_page = {}
        self.page_count = 0
        self.title = ""

    def _log(self, msg: str):
        if self.verbose:
            print(msg)

    def convert(self) -> dict[str, Any]:
        """执行转换"""
        self._log(f"📂 输入文件: {self.file_path}")
        self._log(f"📁 输出目录: {self.output_dir}")
        if not self.extract_images:
            self._log("⚠️  PyMuPDF 未安装，图片提取不可用")
        if not self.extract_tables or not CAMELOT_AVAILABLE:
            self._log("⚠️  camelot 未安装，表格提取不可用")

        self.output_dir.mkdir(parents=True, exist_ok=True)
        if self.extract_images:
            self.images_dir.mkdir(exist_ok=True)
        if self.extract_slide_images:
            self.slide_images_dir.mkdir(exist_ok=True)

        md_content = self._parse_pdf()

        md_file = self.output_dir / f"{self.file_path.stem}.md"
        md_file.write_text(md_content, encoding='utf-8')

        return {
            'success': True,
            'md_file': str(md_file),
            'images_dir': str(self.images_dir) if self.extract_images else '',
            'slide_images_dir': str(self.slide_images_dir) if self.extract_slide_images else '',
            'images': [img['saved_name'] for img in self.images],
            'summary': {
                'total_pages': self.page_count,
                'total_images': self.image_counter,
                'total_slide_images': len(self.slide_images_by_page),
            }
        }

    def _parse_pdf(self) -> str:
        """解析 PDF 文件"""
        self._log("🔍 正在解析 PDF...")

        # 1. 提取文档标题（从元数据）
        self.title = self._extract_title()
        if not self.title:
            self.title = self.file_path.stem.replace('_', ' ').replace('-', ' ')

        md_lines = [f"# {self.title}\n"]

        # 2. 使用 pdfminer 提取每页文本
        with open(self.file_path, 'rb') as f:
            pages = list(PDFPage.get_pages(f))

        self.page_count = len(pages)

        for page_num, page in enumerate(pages, 1):
            self._log(f"   处理第 {page_num} 页")
            page_md = self._parse_page(page, page_num)
            md_lines.append(page_md)

        # 3. 使用 PyMuPDF 提取内嵌图片（如可用）
        if self.extract_images and PYMUPDF_AVAILABLE:
            self._extract_images_with_pymupdf()

        # 4. 导出幻灯片截图（如可用）
        if self.extract_slide_images and PYMUPDF_AVAILABLE:
            self._export_slide_images()

        return '\n'.join(md_lines)

    def _extract_title(self) -> str:
        """从 PDF 元数据提取标题"""
        try:
            with open(self.file_path, 'rb') as f:
                pages = list(PDFPage.get_pages(f))
                if not pages:
                    return ""

                # 尝试从第一页提取大字体文本作为标题
                from pdfminer.high_level import extract_pages
                from pdfminer.layout import LTTextContainer

                for page in extract_pages(self.file_path, page_numbers=[0]):
                    for element in page:
                        if isinstance(element, LTTextContainer):
                            text = element.get_text().strip()
                            if 2 < len(text) < 100:
                                # 检查是否有大字体
                                return text.split('\n')[0][:50]
        except Exception as e:
            self._log(f"   ⚠️ 标题提取失败: {e}")
        return ""

    def _parse_page(self, page, page_num: int) -> str:
        """解析单页 PDF"""
        md_lines = [f"\n## 第 {page_num} 页\n"]

        # 1. 提取文本
        page_text = extract_text(self.file_path, page_numbers=[page_num - 1])
        if page_text.strip():
            formatted_text = self._format_text(page_text)
            md_lines.append(formatted_text)
            md_lines.append("\n")

        # 2. 提取表格
        if self.extract_tables and CAMELOT_AVAILABLE:
            tables = self._extract_tables(page_num)
            for table_md in tables:
                md_lines.append(table_md)
                md_lines.append("\n")

        # 3. 使用 PyMuPDF 提取该页内嵌图片（如可用）
        if self.extract_images and PYMUPDF_AVAILABLE:
            page_images = self._extract_page_images_with_pymupdf(page_num)
            for img in page_images:
                img_path = f"images/{img['saved_name']}"
                md_lines.append(f"![第{page_num}页图片]({img_path})\n\n")

        # 4. 插入页面截图（如果已导出）
        if self.extract_slide_images and PYMUPDF_AVAILABLE and page_num in self.slide_images_by_page:
            slide_img_path = self.slide_images_by_page[page_num]
            slide_img_name = Path(slide_img_path).name
            md_lines.append(f"\n![第{page_num}页截图](slide_images/{slide_img_name})\n\n")

        md_lines.append("\n---\n")
        return ''.join(md_lines)

    def _format_text(self, text: str) -> str:
        """格式化文本，识别基本列表"""
        lines = text.split('\n')
        result_lines = []
        in_list = False
        list_type = None
        current_list = []

        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue

            # 识别无序列表
            if re.match(r'^[\-\*\•]\s+', stripped):
                if list_type != 'ul' or not in_list:
                    if current_list and in_list:
                        result_lines.extend(self._format_list(current_list, list_type))
                    current_list = [stripped]
                    list_type = 'ul'
                    in_list = True
                else:
                    current_list.append(stripped)

            # 识别有序列表
            elif re.match(r'^\d+[\.\、]\s+', stripped):
                if list_type != 'ol' or not in_list:
                    if current_list and in_list:
                        result_lines.extend(self._format_list(current_list, list_type))
                    current_list = [stripped]
                    list_type = 'ol'
                    in_list = True
                else:
                    current_list.append(stripped)

            # 非列表项
            else:
                if in_list:
                    result_lines.extend(self._format_list(current_list, list_type))
                    current_list = []
                    in_list = False
                    list_type = None
                result_lines.append(stripped)

        if current_list and in_list:
            result_lines.extend(self._format_list(current_list, list_type))

        return '\n'.join(result_lines)

    def _format_list(self, items: list[str], list_type: str) -> list[str]:
        lines = []
        for item in items:
            clean_item = re.sub(r'^[\-\*\•]\s+', '', item)
            clean_item = re.sub(r'^\d+[\.\、]\s+', '', clean_item)
            if list_type == 'ul':
                lines.append(f"- {clean_item}")
            else:
                lines.append(f"1. {clean_item}")
        lines.append("")
        return lines

    def _extract_tables(self, page_num: int) -> list[str]:
        """提取表格"""
        table_md = []

        if not CAMELOT_AVAILABLE:
            return table_md

        try:
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                tmp_path = tmp.name

            # 复制 PDF（camelot 需要文件）
            shutil.copy(self.file_path, tmp_path)

            try:
                tables = camelot.read_pdf(tmp_path, pages=str(page_num), flavor='stream')
                if tables.n > 0:
                    for i, table in enumerate(tables):
                        df = table.df
                        table_lines = self._dataframe_to_md_table(df)
                        table_md.append(table_lines)
                        self._log(f"   ✅ 第{page_num}页表格 {i+1}: {len(df)}行")
            except Exception as e:
                self._log(f"   ⚠️ 表格提取失败: {e}")
        except Exception as e:
            self._log(f"   ⚠️ 表格处理失败: {e}")
        finally:
            try:
                os.unlink(tmp_path)
            except:
                pass

        return table_md

    def _dataframe_to_md_table(self, df) -> str:
        lines = []
        header = df.columns.tolist()
        lines.append('| ' + ' | '.join(str(h) for h in header) + ' |')
        lines.append('| ' + ' | '.join(['---'] * len(header)) + ' |')

        for _, row in df.iterrows():
            cells = []
            for cell in row:
                cell_str = str(cell).strip().replace('\n', ' ').replace('|', '\\|')
                cells.append(cell_str)
            lines.append('| ' + ' | '.join(cells) + ' |')

        return '\n'.join(lines)

    def _extract_page_images_with_pymupdf(self, page_num: int) -> list:
        """使用 PyMuPDF 提取单页图片"""
        if not PYMUPDF_AVAILABLE:
            return []

        page_images = []
        try:
            doc = fitz.open(self.file_path)
            page = doc[page_num - 1]

            for img_index, img in enumerate(page.get_images(full=True)):
                xref = img[0]
                base_image = doc.extract_image(xref)
                ext = base_image.get('ext', 'png')
                original_name = f"page{page_num}_img{img_index}.{ext}"
                self.image_counter += 1
                new_name = f"img{self.image_counter:03d}.{ext}"
                output_path = self.images_dir / new_name

                try:
                    image_bytes = base_image.get('image')
                    if image_bytes:
                        with open(output_path, 'wb') as f:
                            f.write(image_bytes)
                        img_info = {
                            'original_name': original_name,
                            'saved_name': new_name,
                            'path': str(output_path),
                            'page_num': page_num
                        }
                        self.images.append(img_info)
                        page_images.append(img_info)
                        self._log(f"   ✅ 提取图片: {new_name}")
                except Exception as e:
                    self._log(f"   ⚠️ 图片提取失败: {e}")

            doc.close()
        except Exception as e:
            self._log(f"   ⚠️ PyMuPDF 图片提取失败: {e}")

        return page_images

    def _extract_images_with_pymupdf(self):
        """使用 PyMuPDF 提取所有图片"""
        if not PYMUPDF_AVAILABLE:
            return

        try:
            doc = fitz.open(self.file_path)
            for page_num, page in enumerate(doc, 1):
                self._extract_page_images_with_pymupdf(page_num)
            doc.close()
        except Exception as e:
            self._log(f"   ⚠️ 图片提取失败: {e}")

    def _export_slide_images(self):
        """使用 PyMuPDF 导出每页截图"""
        if not PYMUPDF_AVAILABLE:
            return

        self._log("🖼️  正在导出页面截图...")

        try:
            doc = fitz.open(self.file_path)
            for page_num, page in enumerate(doc, 1):
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat)
                dest_name = f"page_{page_num:03d}.png"
                dest_path = self.slide_images_dir / dest_name
                pix.save(str(dest_path))
                self.slide_images_by_page[page_num] = str(dest_path)
                self._log(f"   ✅ 第 {page_num} 页截图")
            doc.close()
        except Exception as e:
            self._log(f"   ⚠️ 截图导出失败: {e}")

    def extract_images_only(self) -> dict[str, Any]:
        """仅提取图片"""
        self._log("🖼️  仅提取图片模式")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.images_dir.mkdir(exist_ok=True)

        if PYMUPDF_AVAILABLE:
            self._extract_images_with_pymupdf()
        else:
            self._log("⚠️  PyMuPDF 未安装，无法提取图片")

        return {
            'success': True,
            'images_dir': str(self.images_dir),
            'images': [img['saved_name'] for img in self.images],
            'summary': {'total_images': self.image_counter}
        }


def main():
    parser = argparse.ArgumentParser(
        description='PDF 转 Markdown 转换器',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('files', nargs='+', help='PDF 文件路径')
    parser.add_argument('-o', '--output', help='输出目录')
    parser.add_argument('--no-tables', action='store_true', help='跳过表格提取')
    parser.add_argument('--no-images', action='store_true', help='跳过图片提取')
    parser.add_argument('--images-only', action='store_true', help='仅提取图片')
    parser.add_argument('--slide-images', action='store_true', help='导出每页截图（需要 PyMuPDF）')
    parser.add_argument('-v', '--verbose', action='store_true', help='显示详细信息')

    args = parser.parse_args()

    is_batch = len(args.files) > 1
    total_files = len(args.files)
    success_count = 0
    fail_count = 0
    results_summary = []

    for idx, file_path_str in enumerate(args.files, 1):
        file_path = Path(file_path_str)

        if is_batch:
            print(f"\n{'='*50}")
            print(f"📦 [{idx}/{total_files}]: {file_path.name}")
            print('='*50)

        if not os.path.exists(file_path):
            print(f"❌ 文件不存在: {file_path}")
            fail_count += 1
            continue

        if file_path.suffix.lower() != '.pdf':
            print(f"❌ 请提供 PDF 文件: {file_path}")
            fail_count += 1
            continue

        try:
            converter = PDFToMarkdown(
                file_path=str(file_path),
                output_dir=args.output,
                extract_tables=not args.no_tables,
                extract_images=not args.no_images,
                extract_slide_images=args.slide_images,
                verbose=args.verbose
            )

            if args.images_only:
                result = converter.extract_images_only()
            else:
                result = converter.convert()

            success_count += 1
            results_summary.append({
                'file': str(file_path),
                'success': True,
                'md_file': result['md_file'],
                'summary': result.get('summary', {})
            })

            if is_batch:
                print(f"✅ 完成: {result['md_file']}")

        except Exception as e:
            fail_count += 1
            print(f"❌ 失败: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()

    # 汇总
    print("\n" + "=" * 50)
    if is_batch:
        print(f"📊 {success_count}/{total_files} 成功")
    else:
        if success_count == 1:
            result = results_summary[0]
            print("✅ 转换完成!")
            print(f"📄 {result['md_file']}")
            s = result.get('summary', {})
            print(f"📑 {s.get('total_pages', 0)} 页 | 🖼️ {s.get('total_images', 0)} 图")
        else:
            print("❌ 转换失败")
            sys.exit(1)


if __name__ == '__main__':
    main()
