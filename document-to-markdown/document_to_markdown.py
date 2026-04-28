#!/usr/bin/env python3
"""
文档统一转 Markdown 转换器

自动识别文件类型并调用对应转换器：
- .doc/.docx → word-to-markdown
- .ppt/.pptx → ppt-to-markdown
- .pdf → pdf-to-markdown

依赖:
    - pdfminer.six
    - PyMuPDF (pip install pymupdf)
    - LibreOffice (用于 .doc/.ppt 转换)
"""

import os
import sys
import argparse
from pathlib import Path

# 三个转换器目录
SKILL_DIR = Path(__file__).parent.parent
WORD_DIR = SKILL_DIR / 'word-to-markdown'
PPT_DIR = SKILL_DIR / 'ppt-to-markdown'
PDF_DIR = SKILL_DIR / 'pdf-to-markdown'


def _load_converter(name: str):
    """动态加载转换器"""
    if name == 'word':
        sys.path.insert(0, str(WORD_DIR))
        from word_to_markdown import WordToMarkdown
        return WordToMarkdown
    elif name == 'ppt':
        sys.path.insert(0, str(PPT_DIR))
        from ppt_to_markdown import PPTToMarkdown
        return PPTToMarkdown
    elif name == 'pdf':
        sys.path.insert(0, str(PDF_DIR))
        from pdf_to_markdown import PDFToMarkdown
        return PDFToMarkdown


class DocumentRouter:
    """文档路由分发器"""

    HANDLERS = {
        '.doc': 'word',
        '.docx': 'word',
        '.ppt': 'ppt',
        '.pptx': 'ppt',
        '.pdf': 'pdf',
    }

    def __init__(self, file_path: str, output_dir: str | None = None, verbose: bool = False):
        self.file_path = Path(file_path)
        self.output_dir = Path(output_dir) if output_dir else None
        self.verbose = verbose
        self.supported = list(self.HANDLERS.keys())

    def _log(self, msg: str):
        if self.verbose:
            print(msg)

    def detect_type(self) -> str:
        """检测文件类型"""
        ext = self.file_path.suffix.lower()
        handler = self.HANDLERS.get(ext)
        if not handler:
            raise ValueError(f"不支持的文件格式: {ext}，支持: {self.supported}")
        return handler

    def convert(self) -> dict:
        """执行转换"""
        file_type = self.detect_type()
        self._log(f"📂 检测到文件类型: {file_type.upper()} ({self.file_path.name})")

        if file_type == 'word':
            return self._convert_word()
        elif file_type == 'ppt':
            return self._convert_ppt()
        elif file_type == 'pdf':
            return self._convert_pdf()

    def _convert_word(self) -> dict:
        """转换 Word 文档"""
        if not self.output_dir:
            self.output_dir = self.file_path.parent / f"{self.file_path.stem}_output"
        cls = _load_converter('word')
        converter = cls(file_path=str(self.file_path), output_dir=str(self.output_dir), verbose=self.verbose)
        return converter.convert()

    def _convert_ppt(self) -> dict:
        """转换 PPT 演示文稿"""
        if not self.output_dir:
            self.output_dir = self.file_path.parent / f"{self.file_path.stem}_output"
        cls = _load_converter('ppt')
        converter = cls(file_path=str(self.file_path), output_dir=str(self.output_dir), verbose=self.verbose)
        return converter.convert()

    def _convert_pdf(self) -> dict:
        """转换 PDF 文档"""
        if not self.output_dir:
            self.output_dir = self.file_path.parent / f"{self.file_path.stem}_output"
        cls = _load_converter('pdf')
        converter = cls(file_path=str(self.file_path), output_dir=str(self.output_dir), verbose=self.verbose)
        return converter.convert()


def main():
    parser = argparse.ArgumentParser(
        description='文档统一转 Markdown 转换器',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
支持格式:
  .doc/.docx  Word 文档
  .ppt/.pptx  PowerPoint 演示文稿
  .pdf       PDF 文档

示例:
  python document_to_markdown.py 文档.docx
  python document_to_markdown.py 演示文稿.pptx
  python document_to_markdown.py 文档.pdf
  python document_to_markdown.py 文件1.doc 文件2.pptx 文件3.pdf
        """
    )
    parser.add_argument('files', nargs='+', help='要转换的文件路径')
    parser.add_argument('-o', '--output', help='输出目录（默认：源文件同级目录的 _output 文件夹）')
    parser.add_argument('-v', '--verbose', action='store_true', help='显示详细信息')

    args = parser.parse_args()

    is_batch = len(args.files) > 1
    total = len(args.files)
    success, failed = 0, 0

    for idx, file_path_str in enumerate(args.files, 1):
        file_path = Path(file_path_str)

        if is_batch:
            print(f"\n{'='*50}")
            print(f"📦 [{idx}/{total}] {file_path.name}")
            print('='*50)

        if not file_path.exists():
            print(f"❌ 文件不存在: {file_path}")
            failed += 1
            continue

        ext = file_path.suffix.lower()
        if ext not in DocumentRouter.HANDLERS:
            print(f"❌ 不支持格式: {ext}")
            failed += 1
            continue

        try:
            router = DocumentRouter(
                file_path=str(file_path),
                output_dir=args.output,
                verbose=args.verbose
            )
            result = router.convert()
            success += 1
            if is_batch:
                print(f"✅ 完成: {result['md_file']}")

        except Exception as e:
            failed += 1
            print(f"❌ 失败: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()

    # 汇总
    print("\n" + "=" * 50)
    if is_batch:
        print(f"📊 批量完成: {success}/{total} 成功")
    else:
        if success == 1:
            print("✅ 转换完成!")
            print(f"📄 {Path(args.files[0]).stem}_output/")
        else:
            print("❌ 转换失败")
            sys.exit(1)


if __name__ == '__main__':
    main()
