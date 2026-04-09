"""
智能布局模块 - 内容预扫描 + 智能分页

解决渲染重叠问题，将渲染流程从单阶段改为两阶段：
- 阶段1: 内容预扫描，计算每块实际高度
- 阶段2: 智能分页，根据高度分配到页面
- 阶段3: 按分页结果渲染
"""

import os
from dataclasses import dataclass, field
from typing import List, Optional, Any
from pptx.util import Inches, Pt

from .content_parser import SlideContent, PresentationContent, ChartData, TableData, CodeBlock
from .template_config import TemplateConfig


@dataclass
class ContentBlock:
    """内容块 - 预扫描的基本单元"""
    type: str  # 'title', 'subtitle', 'text', 'bullet', 'image', 'chart', 'table', 'code', 'mermaid'
    content: Any  # 实际内容数据
    height: float = 0.0  # 预计占用高度（英寸）
    page: int = 0  # 分配到的页码
    min_height: float = 0.0  # 最小高度限制


def estimate_text_height(text: str, font_size_pt: float, line_spacing: float = 1.2, 
                         max_width_inches: float = 10.0) -> float:
    """
    估算文本占用高度（英寸）

    Args:
        text: 文本内容
        font_size_pt: 字体大小（Pt）
        line_spacing: 行间距倍数
        max_width_inches: 最大宽度（英寸）

    Returns:
        文本高度（英寸）
    """
    if not text:
        return 0.0

    # 计算字符数和行数
    chars_per_line = int(max_width_inches * 10)  # 粗略估算
    lines = len(text) / max(chars_per_line, 1)
    actual_lines = max(1, int(lines) + text.count('\n'))

    # 计算高度：行数 * (字体大小 / 72 * 行间距)
    line_height = (font_size_pt / 72.0) * line_spacing
    total_height = actual_lines * line_height

    # 添加段落间距
    total_height += 0.1 * actual_lines

    return total_height


def estimate_image_height_inches(image_path: str, max_width_inches: float) -> float:
    """
    估算图片高度（英寸）

    Args:
        image_path: 图片路径
        max_width_inches: 最大宽度（英寸）

    Returns:
        图片高度（英寸）
    """
    if not image_path or not os.path.exists(image_path):
        return 0.0

    try:
        from PIL import Image
        with Image.open(image_path) as img:
            width, height = img.size
            # 计算宽高比
            aspect_ratio = height / width if width > 0 else 1.0
            return max_width_inches * aspect_ratio
    except Exception:
        return 3.0  # 默认高度


def scan_slide_content(slide_content: SlideContent, config: TemplateConfig, 
                       available_width: float) -> List[ContentBlock]:
    """
    预扫描单页内容，计算每个内容块的高度

    Args:
        slide_content: 幻灯片内容
        config: 模板配置
        available_width: 可用宽度（英寸）

    Returns:
        内容块列表
    """
    blocks = []

    # 1. 标题
    if slide_content.title:
        blocks.append(ContentBlock(
            type='title',
            content=slide_content.title,
            height=0.8  # 标题固定 0.8 英寸
        ))

    # 2. 副标题
    if slide_content.subtitle:
        blocks.append(ContentBlock(
            type='subtitle',
            content=slide_content.subtitle,
            height=0.5  # 副标题固定 0.5 英寸
        ))

    # 3. 图片
    for img_path in slide_content.images:
        img_height = estimate_image_height_inches(img_path, available_width)
        blocks.append(ContentBlock(
            type='image',
            content={'path': img_path},
            height=img_height + 0.3  # 图片 + 间距
        ))

    # 4. 表格
    if slide_content.table:
        row_count = len(slide_content.table.rows) + 1  # +1 for header
        table_height = row_count * 0.5  # 每行 0.5 英寸
        blocks.append(ContentBlock(
            type='table',
            content=slide_content.table,
            height=table_height + 0.3  # 表格 + 间距
        ))

    # 5. 项目符号列表
    if slide_content.bullet_points:
        list_height = len(slide_content.bullet_points) * (config.body_size / 72.0 * config.line_spacing + 0.1)
        blocks.append(ContentBlock(
            type='bullet',
            content=slide_content.bullet_points,
            height=list_height
        ))

    # 6. 普通文本
    if slide_content.content_lines:
        full_text = ' '.join(slide_content.content_lines)
        text_height = estimate_text_height(
            full_text,
            config.body_size,
            config.line_spacing,
            available_width
        )
        blocks.append(ContentBlock(
            type='text',
            content=full_text,
            height=text_height
        ))

    # 7. 代码块 / Mermaid 图表
    for code_block in slide_content.code_blocks:
        if code_block.language.lower() in ['mermaid', 'mmd']:
            # Mermaid 图表（将转换为图片）
            # 预估高度：标题 + 图片
            estimated_img_height = 3.0  # 默认图片高度
            blocks.append(ContentBlock(
                type='mermaid',
                content={
                    'code': code_block.content,
                    'language': code_block.language
                },
                height=0.8 + 0.3 + estimated_img_height + 0.3,  # 标题 + 间距 + 图片 + 间距
                min_height=2.0
            ))
        else:
            # 普通代码块
            code_lines_count = len(code_block.content.split('\n'))
            code_height = code_lines_count * (12 / 72.0 * 1.1) + 0.3
            blocks.append(ContentBlock(
                type='code',
                content=code_block,
                height=code_height
            ))

    # 8. 原生图表
    for chart in slide_content.charts:
        chart_height = max(config.default_chart_height, 3.0)  # 最小 3 英寸
        blocks.append(ContentBlock(
            type='chart',
            content=chart,
            height=chart_height + 0.3,  # 图表 + 间距
            min_height=3.0
        ))

    return blocks


def paginate_content_blocks(blocks: List[ContentBlock], 
                            max_height_per_page: float = 6.0) -> int:
    """
    智能分页 - 根据内容高度分配到不同页面

    Args:
        blocks: 内容块列表
        max_height_per_page: 每页最大高度（英寸）

    Returns:
        总页数
    """
    if not blocks:
        return 0

    current_page = 1
    current_height = 0

    for block in blocks:
        # 检查当前页是否能容纳此块
        if current_height + block.height > max_height_per_page and current_height > 0:
            # 需要新页
            current_page += 1
            current_height = 0

        block.page = current_page
        current_height += block.height

    return current_page


def create_layout_plan(original_slides: List[SlideContent], 
                       config: TemplateConfig) -> List[List[ContentBlock]]:
    """
    创建完整布局计划

    Args:
        original_slides: 原始幻灯片内容列表
        config: 模板配置

    Returns:
        每页的内容块列表 [[page1_blocks], [page2_blocks], ...]
    """
    all_blocks = []

    # 计算可用宽度
    available_width = config.slide_width - config.margin_left - config.margin_right
    available_height = config.slide_height - config.margin_top - config.margin_bottom

    # 预扫描所有内容
    for slide_content in original_slides:
        slide_blocks = scan_slide_content(slide_content, config, available_width)
        all_blocks.extend(slide_blocks)

    # 智能分页
    total_pages = paginate_content_blocks(all_blocks, available_height * 0.9)

    # 按页分组
    layout_plan = []
    for page_num in range(1, total_pages + 1):
        page_blocks = [b for b in all_blocks if b.page == page_num]
        layout_plan.append(page_blocks)

    return layout_plan
