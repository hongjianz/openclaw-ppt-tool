"""
智能分页模块 - 基于语义单元的PPT分页算法

改进点:
1. 按语义单元分页(完整段落、列表),避免内容截断
2. 检测内容溢出并自动创建新页
3. 支持"续页"标记(如"第 1 页/共 2 页")
4. 考虑PPT页面实际可容纳的内容量
"""

import re
from typing import List, Tuple
from dataclasses import dataclass

from .content_parser import SlideContent, PresentationContent


@dataclass
class ContentBlock:
    """内容块 - 最小的不可分割的内容单元"""
    block_type: str  # 'paragraph', 'list', 'table', 'image'
    content: any
    estimated_lines: int  # 预估占用的行数


def estimate_lines_for_text(text: str, chars_per_line: int = 40) -> int:
    """
    估算文本需要的行数

    Args:
        text: 文本内容
        chars_per_line: 每行字符数

    Returns:
        预估行数
    """
    if not text:
        return 0
    # 中文字符算1个,英文单词算1个
    lines = 0
    current_line_chars = 0

    for char in text:
        if char == '\n':
            lines += 1
            current_line_chars = 0
        else:
            current_line_chars += 1
            if current_line_chars >= chars_per_line:
                lines += 1
                current_line_chars = 0

    if current_line_chars > 0:
        lines += 1

    return max(1, lines)


def extract_content_blocks(slide: SlideContent, config) -> List[ContentBlock]:
    """
    从SlideContent中提取内容块

    Args:
        slide: 幻灯片内容
        config: 模板配置

    Returns:
        内容块列表
    """
    blocks = []

    # 表格作为一个整体块
    if slide.table:
        # 表格行数 = 表头 + 数据行 + 1(边距)
        table_rows = len(slide.table.rows) + 2
        blocks.append(ContentBlock(
            block_type='table',
            content=slide.table,
            estimated_lines=table_rows
        ))

    # 图片作为独立块
    for img_path in slide.images:
        blocks.append(ContentBlock(
            block_type='image',
            content=img_path,
            estimated_lines=8  # 图片预估占用8行
        ))

    # 列表项 - 每个列表项可以独立
    if slide.bullet_points:
        for point in slide.bullet_points:
            lines = estimate_lines_for_text(point, config.max_chars_per_line)
            blocks.append(ContentBlock(
                block_type='list',
                content=point,
                estimated_lines=lines
            ))

    # 普通文本段落 - 按段落分割
    if slide.content_lines:
        # 尝试按句子或段落分割
        full_text = '\n'.join(slide.content_lines)
        # 按空行分割段落
        paragraphs = re.split(r'\n\s*\n', full_text)

        for para in paragraphs:
            if para.strip():
                lines = estimate_lines_for_text(para, config.max_chars_per_line)
                blocks.append(ContentBlock(
                    block_type='paragraph',
                    content=para.strip(),
                    estimated_lines=lines
                ))

    return blocks


def smart_split_slide(slide: SlideContent, config, max_lines_per_slide: int = None) -> List[SlideContent]:
    """
    智能分割单个幻灯片内容

    Args:
        slide: 原始幻灯片内容
        config: 模板配置
        max_lines_per_slide: 每页最大行数

    Returns:
        分割后的幻灯片列表
    """
    if max_lines_per_slide is None:
        max_lines_per_slide = config.max_lines_per_slide

    # 提取内容块
    blocks = extract_content_blocks(slide, config)

    if not blocks:
        return [slide]

    # 计算可用行数(减去标题和副标题占用的空间)
    title_lines = 2  # 标题占用约2行
    subtitle_lines = 1 if slide.subtitle else 0
    available_lines = max_lines_per_slide - title_lines - subtitle_lines - 2  # 2行边距

    # 如果所有内容能放在一页,直接返回
    total_lines = sum(block.estimated_lines for block in blocks)
    if total_lines <= available_lines:
        return [slide]

    # 需要分页 - 按语义单元分配
    result_slides = []
    current_blocks = []
    current_lines = 0
    page_number = 1

    for block in blocks:
        # 检查添加这个块是否会溢出
        if current_lines + block.estimated_lines > available_lines and current_blocks:
            # 创建当前页
            new_slide = _create_slide_from_blocks(
                slide, current_blocks, page_number,
                total_pages=None  # 稍后更新
            )
            result_slides.append(new_slide)

            # 开始新页
            current_blocks = [block]
            current_lines = block.estimated_lines
            page_number += 1
        else:
            current_blocks.append(block)
            current_lines += block.estimated_lines

    # 添加最后一页
    if current_blocks:
        new_slide = _create_slide_from_blocks(
            slide, current_blocks, page_number,
            total_pages=None
        )
        result_slides.append(new_slide)

    # 更新总页数标记
    total_pages = len(result_slides)
    for i, slide_page in enumerate(result_slides, 1):
        if total_pages > 1:
            # 在标题中添加页码标记
            original_title = slide.title
            slide_page.title = f"{original_title} (第 {i}/{total_pages} 页)"

    return result_slides


def _create_slide_from_blocks(original_slide: SlideContent,
                               blocks: List[ContentBlock],
                               page_number: int,
                               total_pages: int = None) -> SlideContent:
    """
    从内容块创建新的幻灯片

    Args:
        original_slide: 原始幻灯片
        blocks: 内容块列表
        page_number: 当前页码
        total_pages: 总页数

    Returns:
        新的幻灯片对象
    """
    new_slide = SlideContent(
        title=original_slide.title,
        subtitle=original_slide.subtitle if page_number == 1 else "",
        content_lines=[],
        bullet_points=[],
        table=None,
        images=[]
    )

    for block in blocks:
        if block.block_type == 'paragraph':
            new_slide.content_lines.append(block.content)
        elif block.block_type == 'list':
            new_slide.bullet_points.append(block.content)
        elif block.block_type == 'table':
            new_slide.table = block.content
        elif block.block_type == 'image':
            new_slide.images.append(block.content)

    return new_slide


def smart_paginate(presentation: PresentationContent, config) -> PresentationContent:
    """
    对整个演示文稿进行智能分页

    Args:
        presentation: 原始演示文稿
        config: 模板配置

    Returns:
        分页后的演示文稿
    """
    new_presentation = PresentationContent(title=presentation.title)

    for slide in presentation.slides:
        # 对每个幻灯片进行智能分割
        split_slides = smart_split_slide(slide, config)
        new_presentation.slides.extend(split_slides)

    return new_presentation
