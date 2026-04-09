"""
PPT生成核心模块 - 使用python-pptx生成演示文稿
支持自动分页、排版和背景模板应用
"""

import os
import re
import urllib.request
from typing import List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from .template_config import TemplateConfig
from .content_parser import SlideContent, PresentationContent, TableData


def hex_to_rgb(hex_color: str) -> RGBColor:
    """将十六进制颜色转换为RGBColor对象"""
    hex_color = hex_color.lstrip('#')
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


def safe_color(color_str: str, default: str = "#333333") -> RGBColor:
    """
    安全地解析颜色字符串
    如果不支持则使用默认值
    """
    if not color_str:
        return hex_to_rgb(default)

    # 处理rgba/rgb格式 - 提取主要颜色或使用默认值
    if color_str.startswith('rgba') or color_str.startswith('rgb'):
        # 对于QuintenStyle,白色文本使用白色
        if '255,255,255' in color_str or '255, 255, 255' in color_str:
            return RGBColor(255, 255, 255)
        return hex_to_rgb(default)

    # 标准hex颜色
    if color_str.startswith('#') and len(color_str) == 7:
        try:
            return hex_to_rgb(color_str)
        except (ValueError, IndexError):
            return hex_to_rgb(default)

    return hex_to_rgb(default)


class PPTGenerator:
    """PPT生成器"""

    def __init__(self, config: TemplateConfig):
        self.config = config
        self.prs = Presentation()
        # 设置页面尺寸
        self.prs.slide_width = Inches(config.slide_width)
        self.prs.slide_height = Inches(config.slide_height)

    def set_background(self, slide):
        """为幻灯片设置背景"""
        # 优先使用背景图片
        if self.config.background_image and os.path.exists(self.config.background_image):
            left = top = Inches(0)
            width = self.prs.slide_width
            height = self.prs.slide_height
            slide.shapes.add_picture(
                self.config.background_image, left, top,
                width=width, height=height
            )
        # QuintenStyle特殊处理:自动生成背景图片
        elif self.config.use_gradient_background and self.config.style_name == "QuintenStyle":
            # 尝试生成背景图片
            bg_path = self._generate_quinten_background()
            if bg_path and os.path.exists(bg_path):
                left = top = Inches(0)
                width = self.prs.slide_width
                height = self.prs.slide_height
                slide.shapes.add_picture(
                    bg_path, left, top,
                    width=width, height=height
                )
            else:
                # 降级为纯色
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = hex_to_rgb(self.config.background_color)
        else:
            # 使用纯色背景
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = hex_to_rgb(self.config.background_color)

    def _generate_quinten_background(self):
        """生成或获取QuintenStyle背景图片"""
        # 检查templates目录
        templates_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates')
        bg_path = os.path.join(templates_dir, 'quinten_background.png')

        # 如果已存在,直接返回
        if os.path.exists(bg_path):
            return bg_path

        # 尝试生成
        try:
            from generate_quinten_background import generate_quinten_background
            os.makedirs(templates_dir, exist_ok=True)
            generate_quinten_background(1920, 1080, bg_path)
            return bg_path
        except Exception as e:
            print(f"警告: 无法生成QuintenStyle背景图片: {e}")
            return None

    def _add_table(self, slide, table_data: TableData, left, top, width):
        """
        添加表格到幻灯片

        Args:
            slide: 幻灯片对象
            table_data: 表格数据
            left: 左边距
            top: 顶部位置
            width: 表格宽度
        """
        rows = len(table_data.rows) + 1  # +1 for header
        cols = len(table_data.headers)

        if rows == 0 or cols == 0:
            return

        # 计算行高
        row_height = Inches(0.5)
        height = row_height * rows

        # 添加表格
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table

        # 设置表头
        for i, header in enumerate(table_data.headers):
            cell = table.cell(0, i)
            cell.text = header

            # 设置表头样式
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(self.config.body_size - 2)
                paragraph.font.bold = True
                paragraph.font.color.rgb = safe_color("#1A4D4D", "#FFFFFF")
                paragraph.alignment = PP_ALIGN.CENTER

            # 设置表头背景（半透明）
            fill = cell.fill
            fill.solid()
            fill.fore_color.rgb = safe_color("rgba(180, 205, 190, 0.5)", "#B0C5BE")
            fill.transparency = 0.5  # 50%透明度

            # 设置表头边框（更粗）
            for border in [cell.border_top, cell.border_bottom, cell.border_left, cell.border_right]:
                border.width = Pt(2.5)  # 加粗边框
                border.color.rgb = safe_color("#1A4D4D", "#666666")

        # 设置数据行
        for row_idx, row_data in enumerate(table_data.rows, 1):
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < cols:
                    cell = table.cell(row_idx, col_idx)
                    cell.text = cell_data

                    # 设置单元格样式
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(self.config.body_size - 2)
                        paragraph.font.color.rgb = safe_color("#1A4D4D", "#333333")

                        # 设置对齐方式
                        if col_idx < len(table_data.alignment):
                            align = table_data.alignment[col_idx]
                            if align == 'center':
                                paragraph.alignment = PP_ALIGN.CENTER
                            elif align == 'right':
                                paragraph.alignment = PP_ALIGN.RIGHT
                            else:
                                paragraph.alignment = PP_ALIGN.LEFT

                    # 数据单元格背景透明
                    fill = cell.fill
                    fill.background()  # 设置为背景色（透明）

                    # 设置数据单元格边框（较细）
                    for border in [cell.border_top, cell.border_bottom, cell.border_left, cell.border_right]:
                        border.width = Pt(1.5)
                        border.color.rgb = safe_color("#2A5D5D", "#999999")

        # 设置表格整体边框（最外层加粗）
        for row_idx in range(rows):
            for col_idx in range(cols):
                cell = table.cell(row_idx, col_idx)
                # 第一行和最后一行的上下边框加粗
                if row_idx == 0:
                    cell.border_top.width = Pt(3.0)
                if row_idx == rows - 1:
                    cell.border_bottom.width = Pt(3.0)
                # 第一列和最后一列的左右边框加粗
                if col_idx == 0:
                    cell.border_left.width = Pt(3.0)
                if col_idx == cols - 1:
                    cell.border_right.width = Pt(3.0)

    def _download_image(self, url: str) -> Optional[str]:
        """
        下载图片到临时目录

        Args:
            url: 图片URL

        Returns:
            本地文件路径或None
        """
        try:
            # 创建缓存目录
            cache_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'image_cache')
            os.makedirs(cache_dir, exist_ok=True)

            # 生成文件名
            filename = re.sub(r'[^\w\-_\.]', '_', url.split('/')[-1]) or 'image.png'
            local_path = os.path.join(cache_dir, filename)

            # 如果已存在,直接返回
            if os.path.exists(local_path):
                return local_path

            # 下载图片
            print(f"  下载图片: {url}")
            urllib.request.urlretrieve(url, local_path)
            return local_path

        except Exception as e:
            print(f"  警告: 图片下载失败 {url}: {e}")
            return None

    def _add_image(self, slide, img_path: str, left, top, max_width) -> float:
        """
        添加图片到幻灯片

        Args:
            slide: 幻灯片对象
            img_path: 图片路径或URL
            left: 左边距
            top: 顶部位置
            max_width: 最大宽度

        Returns:
            图片实际高度
        """
        # 检查是否为URL
        if img_path.startswith('http://') or img_path.startswith('https://'):
            local_path = self._download_image(img_path)
            if not local_path:
                return 0.0
            img_path = local_path

        # 检查文件是否存在
        if not os.path.exists(img_path):
            print(f"  警告: 图片不存在: {img_path}")
            return 0.0

        try:
            # 添加图片
            pic = slide.shapes.add_picture(img_path, left, top)

            # 计算缩放比例
            scale = min(max_width / pic.width, 1.0)
            new_width = int(pic.width * scale)
            new_height = int(pic.height * scale)

            # 调整图片大小
            pic.width = new_width
            pic.height = new_height

            # 返回英寸高度
            height_inches = new_height / 914400  # EMU to inches
            return height_inches

        except Exception as e:
            print(f"  警告: 图片插入失败: {e}")
            return 0.0

    def add_title_slide(self, title: str, subtitle: str = ""):
        """添加标题页"""
        slide_layout = self.prs.slide_layouts[6]  # 空白布局
        slide = self.prs.slides.add_slide(slide_layout)

        self.set_background(slide)

        # 计算可用区域
        margin_left = Inches(self.config.margin_left)
        margin_right = Inches(self.config.margin_right)
        margin_top = Inches(self.config.margin_top)
        available_width = self.prs.slide_width - margin_left - margin_right
        available_height = self.prs.slide_height - margin_top - Inches(self.config.margin_bottom)

        # 添加标题
        title_height = Inches(1.5)
        title_box = slide.shapes.add_textbox(
            margin_left, margin_top,
            available_width, title_height
        )
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        p = title_frame.paragraphs[0]
        p.text = title
        p.font.name = self.config.title_font
        p.font.size = Pt(self.config.title_size)
        p.font.bold = True
        p.font.color.rgb = safe_color(self.config.primary_color, "#FFFFFF")
        p.alignment = PP_ALIGN.CENTER

        # 添加副标题
        if subtitle:
            subtitle_top = margin_top + title_height + Inches(0.3)
            subtitle_height = Inches(1.0)
            subtitle_box = slide.shapes.add_textbox(
                margin_left, subtitle_top,
                available_width, subtitle_height
            )
            sub_frame = subtitle_box.text_frame
            sub_frame.word_wrap = True

            p = sub_frame.paragraphs[0]
            p.text = subtitle
            p.font.name = self.config.content_font
            p.font.size = Pt(self.config.subtitle_size)
            p.font.color.rgb = safe_color(self.config.secondary_color, "#E0E0E0")
            p.alignment = PP_ALIGN.CENTER

    def add_content_slide(self, slide_content: SlideContent):
        """添加内容页"""
        slide_layout = self.prs.slide_layouts[6]  # 空白布局
        slide = self.prs.slides.add_slide(slide_layout)

        self.set_background(slide)

        # 计算可用区域
        margin_left = Inches(self.config.margin_left)
        margin_right = Inches(self.config.margin_right)
        margin_top = Inches(self.config.margin_top)
        margin_bottom = Inches(self.config.margin_bottom)
        available_width = self.prs.slide_width - margin_left - margin_right
        available_height = self.prs.slide_height - margin_top - margin_bottom

        current_top = margin_top

        # 添加标题
        title_height = Inches(0.8)
        title_box = slide.shapes.add_textbox(
            margin_left, current_top,
            available_width, title_height
        )
        title_frame = title_box.text_frame
        title_frame.word_wrap = True

        p = title_frame.paragraphs[0]
        p.text = slide_content.title
        p.font.name = self.config.title_font
        p.font.size = Pt(self.config.title_size)
        p.font.bold = True
        p.font.color.rgb = safe_color(self.config.primary_color, "#FFFFFF")
        p.alignment = PP_ALIGN.LEFT

        current_top += title_height + Inches(0.2)

        # 添加副标题
        if slide_content.subtitle:
            subtitle_height = Inches(0.5)
            subtitle_box = slide.shapes.add_textbox(
                margin_left, current_top,
                available_width, subtitle_height
            )
            sub_frame = subtitle_box.text_frame
            sub_frame.word_wrap = True

            p = sub_frame.paragraphs[0]
            p.text = slide_content.subtitle
            p.font.name = self.config.content_font
            p.font.size = Pt(self.config.subtitle_size - 4)
            p.font.color.rgb = safe_color(self.config.secondary_color, "#E0E0E0")
            p.alignment = PP_ALIGN.LEFT

            current_top += subtitle_height + Inches(0.2)

        # 添加图片
        if slide_content.images:
            for img_path in slide_content.images:
                img_height = self._add_image(slide, img_path, margin_left, current_top, available_width)
                current_top += img_height + Inches(0.3)

        # 添加表格
        if slide_content.table:
            table_top = current_top
            self._add_table(slide, slide_content.table, margin_left, table_top, available_width)
            current_top += Inches(3.0)  # 表格占用高度

        # 添加项目符号列表
        if slide_content.bullet_points:
            remaining_height = available_height - (current_top - margin_top)
            bullet_box = slide.shapes.add_textbox(
                margin_left, current_top,
                available_width, remaining_height
            )
            bullet_frame = bullet_box.text_frame
            bullet_frame.word_wrap = True
            bullet_frame.line_spacing = self.config.line_spacing

            for i, point in enumerate(slide_content.bullet_points):
                if i == 0:
                    p = bullet_frame.paragraphs[0]
                else:
                    p = bullet_frame.add_paragraph()

                p.text = point
                p.font.name = self.config.content_font
                p.font.size = Pt(self.config.body_size)
                p.font.color.rgb = safe_color(self.config.text_color, "#FFFFFF")
                p.level = 0
                p.space_before = Pt(6)
                p.space_after = Pt(6)

        # 添加普通文本内容
        if slide_content.content_lines:
            remaining_height = available_height - (current_top - margin_top)
            content_box = slide.shapes.add_textbox(
                margin_left, current_top,
                available_width, remaining_height
            )
            content_frame = content_box.text_frame
            content_frame.word_wrap = True
            content_frame.line_spacing = self.config.line_spacing

            # 合并内容行并按字符数换行
            full_text = ' '.join(slide_content.content_lines)

            # 智能分段
            paragraphs = self._split_text_into_paragraphs(full_text)

            for i, paragraph in enumerate(paragraphs):
                if i == 0:
                    p = content_frame.paragraphs[0]
                else:
                    p = content_frame.add_paragraph()

                p.text = paragraph
                p.font.name = self.config.content_font
                p.font.size = Pt(self.config.body_size)
                p.font.color.rgb = safe_color(self.config.text_color, "#FFFFFF")
                p.space_before = Pt(6)
                p.space_after = Pt(6)

    def _split_text_into_paragraphs(self, text: str) -> List[str]:
        """
        将长文本智能分割为多个段落
        考虑每行最大字符数和每页最大行数
        """
        max_chars = self.config.max_chars_per_line
        max_lines = self.config.max_lines_per_slide

        # 按句子分割 - 修复Python 3.6兼容性,避免空模式匹配
        # 使用非空断言确保正则表达式有效
        sentences = re.split(r'(?<=[。！？.!?])\s+', text) if text else []

        paragraphs = []
        current_paragraph = ""
        line_count = 0

        for sentence in sentences:
            if not sentence.strip():
                continue

            # 计算这句话需要的行数
            lines_needed = (len(sentence) + max_chars - 1) // max_chars

            # 如果当前段落加上新句子会超过限制,开始新段落
            if line_count + lines_needed > max_lines and current_paragraph:
                paragraphs.append(current_paragraph)
                current_paragraph = sentence
                line_count = lines_needed
            else:
                if current_paragraph:
                    current_paragraph += " " + sentence
                else:
                    current_paragraph = sentence
                line_count += lines_needed

        # 添加最后一个段落
        if current_paragraph:
            paragraphs.append(current_paragraph)

        return paragraphs if paragraphs else [text]

    def generate(self, content: PresentationContent, output_path: str):
        """
        生成PPT文件

        Args:
            content: 结构化的演示文稿内容
            output_path: 输出文件路径
        """
        # 清空现有幻灯片
        while len(self.prs.slides) > 0:
            rId = self.prs.slides._sldIdLst[0].rId
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[0]

        total_slides = len(content.slides)

        # 添加标题页
        if content.slides:
            first_slide = content.slides[0]
            self.add_title_slide(first_slide.title, first_slide.subtitle)
            # 添加页脚到标题页(如果启用)
            if self.config.show_footer_on_title_slide:
                self._add_footer(self.prs.slides[0], 1, total_slides)

            # 添加内容页
            for idx, slide_content in enumerate(content.slides[1:], 2):
                self.add_content_slide(slide_content)
                # 添加页脚
                self._add_footer(self.prs.slides[idx - 1], idx, total_slides)

        # 保存文件
        self.prs.save(output_path)
        print(f"PPT已生成: {output_path}")
        print(f"共 {len(self.prs.slides)} 页")

    def _add_footer(self, slide, current_page: int, total_pages: int):
        """
        添加页脚和页码到幻灯片

        Args:
            slide: 幻灯片对象
            current_page: 当前页码
            total_pages: 总页数
        """
        if not self.config.show_page_number and not self.config.footer_text:
            return

        # 构建页脚文本
        footer_parts = []

        # 添加自定义页脚文本
        if self.config.footer_text:
            footer_parts.append(self.config.footer_text)

        # 添加页码
        if self.config.show_page_number:
            page_text = self.config.page_number_format.format(
                current=current_page,
                total=total_pages
            )
            footer_parts.append(page_text)

        if not footer_parts:
            return

        footer_text = " | ".join(footer_parts)

        # 计算页脚位置
        margin_bottom = Inches(self.config.margin_bottom)
        footer_height = Inches(0.4)
        footer_top = self.prs.slide_height - margin_bottom - footer_height

        # 创建页脚文本框
        footer_box = slide.shapes.add_textbox(
            Inches(0.5), footer_top,
            self.prs.slide_width - Inches(1.0), footer_height
        )
        footer_frame = footer_box.text_frame
        footer_frame.word_wrap = True

        p = footer_frame.paragraphs[0]
        p.text = footer_text
        p.font.name = self.config.content_font
        p.font.size = Pt(self.config.footer_font_size)
        p.font.color.rgb = safe_color(self.config.footer_color, "#666666")
        p.alignment = PP_ALIGN.CENTER


# 需要导入re模块
import re
