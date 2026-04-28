#!/usr/bin/env python3
"""
PPT 转 Markdown 转换器 v1.0

将 PowerPoint 演示文稿转换为结构化 Markdown，
保留幻灯片结构、文本逻辑顺序、图片和演讲者备注。
"""

import os
import re
import sys
import json
import zipfile
import tempfile
import shutil
import subprocess
import argparse
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

__doc__ = """
PPT 转 Markdown 转换器

用法:
    python ppt_to_markdown.py <pptx文件> [选项]

示例:
    python ppt_to_markdown.py presentation.pptx
    python ppt_to_markdown.py file1.pptx file2.pptx -o ./output
"""

# XML 命名空间
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'n': 'http://schemas.openxmlformats.org/officeDocument/2006/notes',
}

# 注册命名空间
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


class PPTToMarkdown:
    """PPT 演示文稿转 Markdown 转换器"""

    def __init__(
        self,
        file_path: str,
        output_dir: str | None = None,
        extract_images: bool = True,
        extract_notes: bool = True,
        extract_slide_images: bool = False,  # 新增：导出每页幻灯片截图
        image_prefix: str = "img",
        verbose: bool = False
    ):
        self.file_path = Path(file_path)
        self.extract_images = extract_images
        self.extract_notes = extract_notes
        self.extract_slide_images = extract_slide_images
        self.image_prefix = image_prefix
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
        self.images_by_slide = {}  # {slide_num: [image_info, ...]}
        self.slide_images_by_slide = {}  # {slide_num: image_path}
        self.notes = []
        self.slide_count = 0
        self.title = ""

    def _log(self, msg: str):
        """输出日志"""
        if self.verbose:
            print(msg)

    def convert(self) -> dict[str, Any]:
        """执行转换"""
        self._log(f"📂 输入文件: {self.file_path}")
        self._log(f"📁 输出目录: {self.output_dir}")

        # 创建输出目录
        self.output_dir.mkdir(parents=True, exist_ok=True)
        if self.extract_images:
            self.images_dir.mkdir(exist_ok=True)
        if self.extract_slide_images:
            self.slide_images_dir.mkdir(exist_ok=True)

        # 如果启用幻灯片截图，先导出
        if self.extract_slide_images:
            self._export_slide_images()

        # 解析 PPT
        md_content = self._parse_pptx()

        # 保存 Markdown
        md_file = self.output_dir / f"{self.file_path.stem}.md"
        md_file.write_text(md_content, encoding='utf-8')

        return {
            'success': True,
            'md_file': str(md_file),
            'images_dir': str(self.images_dir) if self.extract_images else '',
            'slide_images_dir': str(self.slide_images_dir) if self.extract_slide_images else '',
            'images': [img['saved_name'] for img in self.images],
            'notes': self.notes,
            'summary': {
                'total_slides': self.slide_count,
                'total_images': self.image_counter,
                'total_slide_images': len(self.slide_images_by_slide),
                'total_notes': len(self.notes),
            }
        }

    def _export_slide_images(self):
        """使用 LibreOffice 将每页幻灯片导出为图片"""
        self._log("🖼️  正在导出幻灯片截图（依赖 LibreOffice）...")

        # 检查 LibreOffice 是否可用
        libreoffice_cmd = 'soffice'
        if not shutil.which(libreoffice_cmd):
            # 尝试常见路径
            for path in [
                '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                '/usr/bin/soffice',
                '/usr/local/bin/soffice',
            ]:
                if Path(path).exists():
                    libreoffice_cmd = path
                    break
            else:
                self._log("   ⚠️  LibreOffice 未安装，跳过幻灯片截图功能")
                self._log("   💡 安装 LibreOffice: brew install --cask libreoffice")
                return

        try:
            # 使用临时目录
            with tempfile.TemporaryDirectory() as tmpdir:
                # 复制文件
                tmp_input = Path(tmpdir) / self.file_path.name
                tmp_input.write_bytes(self.file_path.read_bytes())

                # 调用 LibreOffice 转换为 PNG
                cmd = [
                    libreoffice_cmd,
                    '--headless',
                    '--convert-to', 'png',
                    '--outdir', tmpdir,
                    str(tmp_input)
                ]

                if self.verbose:
                    self._log(f"   🔄 执行: {' '.join(cmd)}")

                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=120
                )

                if result.returncode != 0:
                    self._log(f"   ⚠️  LibreOffice 导出失败: {result.stderr}")
                    return

                # 查找生成的 PNG 文件
                tmp_output = Path(tmpdir)
                png_files = sorted(tmp_output.glob(f"{self.file_path.stem}*.png"))

                if not png_files:
                    # LibreOffice 可能生成不同命名的文件
                    png_files = sorted(tmp_output.glob("*.png"))

                for png_file in png_files:
                    # 从文件名提取幻灯片编号
                    # 格式：presentationName1.png, presentationName2.png, ...
                    match = re.search(r'(\d+)\.png$', png_file.name)
                    if match:
                        slide_num = int(match.group(1))
                    else:
                        # 按修改时间排序，第几个就是第几页
                        slide_num = png_files.index(png_file) + 1

                    # 复制到输出目录
                    dest_name = f"slide_{slide_num:03d}.png"
                    dest_path = self.slide_images_dir / dest_name
                    shutil.copy(png_file, dest_path)

                    self.slide_images_by_slide[slide_num] = str(dest_path)
                    self._log(f"   ✅ 导出幻灯片 {slide_num}: {dest_name}")

        except subprocess.TimeoutExpired:
            self._log("   ⚠️  LibreOffice 导出超时")
        except Exception as e:
            self._log(f"   ⚠️  导出失败: {e}")

    def _parse_pptx(self) -> str:
        """解析 PPTX 文件"""
        with zipfile.ZipFile(self.file_path, 'r') as z:
            # 1. 提取演示文稿标题
            self._extract_title(z)

            # 2. 遍历幻灯片（图片在解析每页时提取）
            slide_files = sorted([
                f for f in z.namelist()
                if re.match(r'ppt/slides/slide\d+\.xml', f)
            ])

            md_lines = [f"# {self.title}\n"]

            for idx, slide_file in enumerate(slide_files, 1):
                self._log(f"   处理幻灯片 {idx}: {slide_file}")
                slide_md = self._parse_slide(z, slide_file, idx)
                md_lines.append(slide_md)
                self.slide_count = idx

            # 4. 提取演讲者备注
            if self.extract_notes:
                self._extract_notes(z)

            # 添加备注脚注
            if self.notes:
                md_lines.append("\n---\n\n## 演讲者备注\n")
                for i, note_data in enumerate(self.notes, 1):
                    md_lines.append(f"[^slide{i}]: {note_data['note']}\n")

            return '\n'.join(md_lines)

    def _extract_title(self, z: zipfile.ZipFile):
        """提取演示文稿标题"""
        try:
            # 尝试从 presentation.xml 获取标题
            pres_xml = z.read('ppt/presentation.xml')
            root = ET.fromstring(pres_xml)

            # 查找第一个幻灯片ID对应的标题
            sld_id_list = root.find('.//{*}sldIdLst')
            if sld_id_list is not None:
                first_sld_id = sld_id_list.find('{*}sldId')
                if first_sld_id is not None:
                    r_id = first_sld_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    # 从关系文件中查找
                    rels = z.read('ppt/_rels/presentation.xml.rels')
                    rels_root = ET.fromstring(rels)
                    for rel in rels_root:
                        if rel.get('Id') == r_id:
                            target = rel.get('Target', '')
                            if 'slides/slide' in target:
                                slide_num = re.search(r'slide(\d+)', target)
                                if slide_num:
                                    slide_file = f"ppt/slides/slide{slide_num.group(1)}.xml"
                                    slide_xml = z.read(slide_file)
                                    slide_root = ET.fromstring(slide_xml)
                                    # 尝试获取标题
                                    title = self._extract_slide_title(slide_root)
                                    if title:
                                        self.title = title
                                        return
        except Exception as e:
            self._log(f"   ⚠️  无法提取标题: {e}")

        # 默认标题
        self.title = self.file_path.stem.replace('_', ' ').replace('-', ' ')

    def _extract_slide_title(self, slide_root: ET.Element) -> str:
        """从幻灯片中提取标题"""
        # 查找标题占位符
        for sp in slide_root.findall('.//p:sp', NS):
            nvSpPr = sp.find('.//p:nvSpPr', NS)
            if nvSpPr is not None:
                nvPr = nvSpPr.find('.//p:nvPr', NS)
                if nvPr is not None:
                    type_val = nvPr.find('.//p:ph', NS)
                    if type_val is not None and type_val.get('type') == 'title':
                        # 获取标题文本
                        texts = []
                        for t in sp.findall('.//a:t', NS):
                            if t.text:
                                texts.append(t.text)
                        if texts:
                            return ''.join(texts)
        return ""

    def _parse_slide(self, z: zipfile.ZipFile, slide_file: str, idx: int) -> str:
        """解析单张幻灯片"""
        slide_xml = z.read(slide_file)
        root = ET.fromstring(slide_xml)

        md_lines = [f"\n## 第 {idx} 页\n"]

        # 1. 提取标题
        title = self._extract_slide_title(root)
        if title:
            md_lines.append(f"**{title}**\n\n")

        # 2. 提取文本内容（按逻辑顺序）
        texts = self._extract_slide_texts(root)
        for text in texts:
            if text.strip():
                md_lines.append(f"{text}\n\n")

        # 3. 提取并插入该页的图片
        slide_images = self._extract_slide_images(z, slide_file, idx)
        for img in slide_images:
            img_path = f"images/{img['saved_name']}"
            md_lines.append(f"![第{idx}页图片]({img_path})\n\n")

        # 4. 检查是否有 SmartArt
        self._handle_smartart(md_lines, slide_file, z)

        # 5. 检查是否有图表
        self._handle_chart(md_lines, slide_file, z)

        # 6. 插入幻灯片截图（如果已导出）
        if self.extract_slide_images and idx in self.slide_images_by_slide:
            slide_img_path = self.slide_images_by_slide[idx]
            slide_img_name = Path(slide_img_path).name
            md_lines.append(f"\n![第{idx}页截图](slide_images/{slide_img_name})\n\n")

        # 添加分隔线
        md_lines.append("\n---\n")

        return ''.join(md_lines)

    def _extract_slide_images(self, z: zipfile.ZipFile, slide_file: str, slide_idx: int) -> list:
        """提取单张幻灯片的图片，返回该页的图片列表"""
        slide_images = []

        # 获取幻灯片关系文件
        rels_file = slide_file.replace('slides/', 'slides/_rels/') + '.rels'
        try:
            rels_xml = z.read(rels_file)
            rels_root = ET.fromstring(rels_xml)

            # 查找图片关系
            for rel in rels_root:
                rel_type = rel.get('Type', '')
                if 'image' in rel_type:
                    target = rel.get('Target', '')

                    if not target:
                        continue

                    # 解析相对路径：../media/image1.jpeg → ppt/media/image1.jpeg
                    # rels 文件在 ppt/slides/_rels/，target 是 ../media/xxx
                    if target.startswith('../'):
                        # 去掉 ../ 得到 media/image1.jpeg
                        relative = target[3:]
                        media_path = f"ppt/{relative}"
                    else:
                        media_path = f"ppt/{target}"

                    # 查找原始文件名
                    original_name = Path(media_path).name

                    # 检查图片是否已提取
                    existing = None
                    for img in self.images:
                        if img['original_name'] == original_name:
                            existing = img
                            break

                    if not existing:
                        # 提取图片
                        if media_path in z.namelist():
                            self.image_counter += 1
                            ext = Path(original_name).suffix.lower()
                            new_name = f"{self.image_prefix}{self.image_counter:03d}{ext}"
                            output_path = self.images_dir / new_name

                            try:
                                with z.open(media_path) as source:
                                    with open(output_path, 'wb') as target_file:
                                        target_file.write(source.read())

                                existing = {
                                    'original_name': original_name,
                                    'saved_name': new_name,
                                    'path': str(output_path),
                                    'slide_idx': slide_idx
                                }
                                self.images.append(existing)
                                self._log(f"   ✅ 提取: {original_name} → {new_name} (第{slide_idx}页)")
                            except Exception as e:
                                self._log(f"   ⚠️  提取失败 {media_path}: {e}")

                    if existing:
                        slide_images.append(existing)

        except KeyError:
            pass

        return slide_images

    def _extract_slide_texts(self, slide_root: ET.Element) -> list[str]:
        """按逻辑顺序提取幻灯片文本 - 按形状分组"""
        results = []
        current_list = []
        in_list = False
        list_type = None  # 'ul' or 'ol'

        # 查找所有形状（文本框）
        shapes = slide_root.findall('.//p:sp', NS)

        for shape in shapes:
            # 跳过标题占位符（已单独处理）
            nvSpPr = shape.find('.//p:nvSpPr', NS)
            if nvSpPr is not None:
                nvPr = nvSpPr.find('.//p:nvPr', NS)
                if nvPr is not None:
                    ph = nvPr.find('.//p:ph', NS)
                    if ph is not None and ph.get('type') == 'title':
                        continue  # 标题已单独提取

            # 提取形状中的所有文本
            shape_texts = []
            for t in shape.findall('.//a:t', NS):
                if t.text:
                    shape_texts.append(t.text)

            if not shape_texts:
                continue

            # 合并同一形状的文本
            full_text = ''.join(shape_texts).strip()
            if not full_text:
                continue

            # 判断是否是列表项
            is_list_item, item_type = self._is_list_item(full_text)

            if is_list_item:
                if list_type != item_type or not in_list:
                    if current_list and in_list:
                        results.extend(self._format_list(current_list, list_type))
                        current_list = []
                    current_list.append(full_text)
                    list_type = item_type
                    in_list = True
                else:
                    current_list.append(full_text)
            else:
                # 非列表项
                if in_list:
                    results.extend(self._format_list(current_list, list_type))
                    current_list = []
                    in_list = False
                    list_type = None

                # 跳过无意义文本（过短或只有特殊字符）
                if len(full_text) > 1 and not self._is_noise_text(full_text):
                    results.append(full_text)

        # 处理最后残留的列表
        if current_list and in_list:
            results.extend(self._format_list(current_list, list_type))

        return results

    def _is_noise_text(self, text: str) -> bool:
        """判断是否是无意义文本"""
        # 纯数字、纯符号等
        if re.match(r'^[\d\.\-\—\–\→\←\↑\↓\s]+$', text):
            return True
        # 过短
        if len(text) < 2:
            return True
        return False

    def _is_list_item(self, text: str) -> tuple[bool, str | None]:
        """判断文本是否是列表项"""
        # 无序列表
        if re.match(r'^[\-\*\•]\s+', text):
            return True, 'ul'
        # 有序列表
        if re.match(r'^\d+[\.\、]\s+', text):
            return True, 'ol'
        return False, None

    def _format_list(self, items: list[str], list_type: str) -> list[str]:
        """格式化列表"""
        lines = []
        for item in items:
            if list_type == 'ul':
                lines.append(f"- {item}")
            else:
                lines.append(f"1. {item}")
        lines.append("")  # 空行
        return lines

    def _handle_smartart(self, md_lines: list, slide_file: str, z: zipfile.ZipFile):
        """处理 SmartArt 图形"""
        # SmartArt 数据存储在 ppt/diagrams/ 目录
        # 这里简化处理，输出提示信息
        try:
            diagrams_dir = 'ppt/diagrams/'
            for name in z.namelist():
                if name.startswith(diagrams_dir) and name.endswith('.xml'):
                    md_lines.append(f"\n> [SmartArt 图形: {Path(name).stem}]\n\n")
                    # 尝试解析 SmartArt 结构
                    try:
                        diagram_xml = z.read(name)
                        root = ET.fromstring(diagram_xml)
                        # 简化：提取文本节点
                        texts = []
                        for t in root.findall('.//a:t', NS):
                            if t.text:
                                texts.append(t.text)
                        if texts:
                            for text in texts[:5]:  # 只取前5个
                                md_lines.append(f"- {text}\n")
                    except Exception:
                        pass
        except Exception:
            pass

    def _handle_chart(self, md_lines: list, slide_file: str, z: zipfile.ZipFile):
        """处理图表"""
        # 检查是否有图表关系
        try:
            rels_file = slide_file.replace('slides/', 'slides/_rels/') + '.rels'
            rels_xml = z.read(rels_file)
            rels_root = ET.fromstring(rels_xml)

            for rel in rels_root:
                rel_type = rel.get('Type', '')
                if 'chart' in rel_type:
                    target = rel.get('Target', '')
                    chart_name = Path(target).stem
                    md_lines.append(f"\n> [图表: {chart_name}]\n\n")
        except KeyError:
            pass

    def _extract_images(self, z: zipfile.ZipFile):
        """提取所有图片"""
        self._log("🖼️  提取图片...")

        for name in z.namelist():
            if name.startswith('ppt/media/'):
                ext = Path(name).suffix.lower()
                if ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg']:
                    self.image_counter += 1
                    new_name = f"{self.image_prefix}{self.image_counter:03d}{ext}"
                    output_path = self.images_dir / new_name

                    try:
                        with z.open(name) as source:
                            with open(output_path, 'wb') as target:
                                target.write(source.read())

                        self.images.append({
                            'original_name': Path(name).name,
                            'saved_name': new_name,
                            'path': str(output_path)
                        })
                        self._log(f"   ✅ 提取: {Path(name).name} → {new_name}")
                    except Exception as e:
                        self._log(f"   ⚠️  提取失败 {name}: {e}")

    def _extract_notes(self, z: zipfile.ZipFile):
        """提取所有演讲者备注"""
        self._log("📝 提取演讲者备注...")

        for name in z.namelist():
            if re.match(r'ppt/notesSlides/notesSlide\d+\.xml', name):
                try:
                    notes_xml = z.read(name)
                    root = ET.fromstring(notes_xml)

                    # 提取备注文本
                    texts = []
                    for t in root.findall('.//a:t', NS):
                        if t.text and t.text.strip():
                            texts.append(t.text.strip())

                    if texts:
                        # 从文件名提取幻灯片编号
                        slide_num = int(re.search(r'notesSlide(\d+)', name).group(1))
                        note_text = ' '.join(texts)

                        self.notes.append({
                            'slide': slide_num,
                            'note': note_text
                        })
                        self._log(f"   ✅ 幻灯片 {slide_num} 备注: {note_text[:30]}...")
                except Exception as e:
                    self._log(f"   ⚠️  无法解析备注 {name}: {e}")

    def extract_images_only(self) -> dict[str, Any]:
        """仅提取图片"""
        self._log("🖼️  仅提取图片模式")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.images_dir.mkdir(exist_ok=True)

        with zipfile.ZipFile(self.file_path, 'r') as z:
            self._extract_images(z)

        return {
            'success': True,
            'images_dir': str(self.images_dir),
            'images': [img['saved_name'] for img in self.images],
            'summary': {'total_images': self.image_counter}
        }


def main():
    parser = argparse.ArgumentParser(
        description='PPT 转 Markdown 转换器',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument('files', nargs='+', help='PPT 文件路径 (.ppt/.pptx)，支持多个文件批量转换')
    parser.add_argument('-o', '--output', help='输出目录')
    parser.add_argument('--libreoffice', default='/Applications/LibreOffice.app/Contents/MacOS/soffice',
                       help='LibreOffice 命令路径 (默认: lowriter)')
    parser.add_argument('--images-only', action='store_true',
                       help='仅提取图片')
    parser.add_argument('--no-notes', action='store_true',
                       help='跳过演讲者备注提取')
    parser.add_argument('--no-smartart', action='store_true',
                       help='跳过 SmartArt 处理')
    parser.add_argument('--slide-images', action='store_true',
                       help='导出每页幻灯片截图（需要 LibreOffice）')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='显示详细信息')

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
            print(f"📦 处理文件 [{idx}/{total_files}]: {file_path.name}")
            print('='*50)

        if not os.path.exists(file_path):
            print(f"❌ 文件不存在: {file_path}")
            fail_count += 1
            results_summary.append({'file': str(file_path), 'success': False, 'error': '文件不存在'})
            continue

        original_file = file_path

        # .ppt 格式需要先转换为 .pptx
        if file_path.suffix.lower() == '.ppt':
            print("📄 检测到 .ppt 格式，正在转换为 .pptx...")
            try:
                file_path = convert_ppt_to_pptx(file_path, args.libreoffice, args.verbose)
                print(f"✅ 转换成功: {file_path}")
            except FileNotFoundError:
                print("❌ 未找到 LibreOffice，无法转换 .ppt 格式")
                print("")
                print("💡 解决方案（二选一）：")
                print("   1. 安装 LibreOffice:")
                print("      macOS: brew install --cask libreoffice")
                print("      Ubuntu: sudo apt install libreoffice")
                print("   2. 将 .ppt 文档在 PowerPoint 中另存为 .pptx 格式")
                fail_count += 1
                results_summary.append({'file': str(file_path), 'success': False, 'error': 'LibreOffice 未安装'})
                continue
            except Exception as e:
                print(f"❌ .ppt 转换失败: {e}")
                fail_count += 1
                results_summary.append({'file': str(file_path), 'success': False, 'error': str(e)})
                continue

        elif file_path.suffix.lower() != '.pptx':
            print(f"❌ 请提供 .ppt 或 .pptx 文件: {file_path}")
            fail_count += 1
            results_summary.append({'file': str(file_path), 'success': False, 'error': '不支持的文件格式'})
            continue

        # 创建转换器
        converter = PPTToMarkdown(
            file_path=str(file_path),
            output_dir=args.output,
            extract_images=not args.images_only,
            extract_notes=not args.no_notes,
            extract_slide_images=args.slide_images,
            verbose=args.verbose
        )

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
            if args.verbose:
                import traceback
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
        if success_count == 1:
            result = results_summary[0]
            print("✅ 转换完成!")
            print("=" * 50)
            print(f"📄 Markdown: {result['md_file']}")
            if result.get('images_dir'):
                print(f"📁 图片目录: {result['images_dir']}")
            s = result.get('summary', {})
            print(f"📑 幻灯片: {s.get('total_slides', 0)}")
            print(f"🖼️  图片数量: {s.get('total_images', 0)}")
            print(f"📝 备注数量: {s.get('total_notes', 0)}")
        else:
            print("❌ 转换失败")
            sys.exit(1)


def convert_ppt_to_pptx(ppt_path, libreoffice_cmd='/Applications/LibreOffice.app/Contents/MacOS/soffice', verbose=False, timeout=120):
    """使用 LibreOffice 将 .ppt 转换为 .pptx"""
    ppt_path = Path(ppt_path).resolve()

    # 安全检查
    cmd_path = Path(libreoffice_cmd)
    if not cmd_path.is_absolute():
        allowed_commands = {'lowriter', 'soffice', 'libreoffice'}
        if cmd_path.name not in allowed_commands:
            raise ValueError(f"不支持的 LibreOffice 命令: {libreoffice_cmd}")
    elif not cmd_path.exists():
        raise FileNotFoundError(f"LibreOffice 路径不存在: {libreoffice_cmd}")

    if not shutil.which(libreoffice_cmd):
        raise FileNotFoundError(f"LibreOffice 未安装或路径不正确: {libreoffice_cmd}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_input = Path(tmpdir) / ppt_path.name
        tmp_input.write_bytes(ppt_path.read_bytes())

        cmd = [
            libreoffice_cmd,
            '--headless',
            '--convert-to', 'pptx',
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

        expected_output = Path(tmpdir) / f"{ppt_path.stem}.pptx"
        if expected_output.exists():
            output_path = ppt_path.parent / f"{ppt_path.stem}_converted.pptx"
            output_path.write_bytes(expected_output.read_bytes())
            return output_path

        docx_files = list(Path(tmpdir).glob('*.pptx'))
        if docx_files:
            output_path = ppt_path.parent / f"{ppt_path.stem}_converted.pptx"
            output_path.write_bytes(docx_files[0].read_bytes())
            return output_path

        raise RuntimeError("LibreOffice 未能生成 .pptx 文件")


if __name__ == '__main__':
    main()
