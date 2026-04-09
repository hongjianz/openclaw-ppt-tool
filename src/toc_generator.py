"""
目录页生成模块 - 自动生成PPT目录/导航页

功能:
- 解析所有标题,自动生成目录页
- 支持超链接跳转到对应页(PPT内部链接)
- 可选是否包含子标题
- 支持多级目录结构
"""

from typing import List, Tuple, Optional
from dataclasses import dataclass

from .content_parser import SlideContent, PresentationContent


@dataclass
class TocEntry:
    """目录项"""
    title: str
    slide_index: int  # 在slides列表中的索引(从0开始)
    level: int = 1  # 标题层级: 1=主标题, 2=副标题
    subtitle: str = ""  # 副标题(可选)


def extract_toc_entries(presentation: PresentationContent, include_subtitles: bool = True) -> List[TocEntry]:
    """
    从演示文稿中提取目录项

    Args:
        presentation: 演示文稿内容
        include_subtitles: 是否包含副标题(二级标题)

    Returns:
        目录项列表
    """
    toc_entries = []

    for i, slide in enumerate(presentation.slides):
        # 跳过空标题
        if not slide.title or slide.title.strip() == "":
            continue

        # 添加主标题
        toc_entries.append(TocEntry(
            title=slide.title,
            slide_index=i,
            level=1,
            subtitle=slide.subtitle if include_subtitles and slide.subtitle else ""
        ))

    return toc_entries


def create_toc_slide(toc_entries: List[TocEntry], presentation_title: str = "") -> SlideContent:
    """
    创建目录页幻灯片

    Args:
        toc_entries: 目录项列表
        presentation_title: 演示文稿标题

    Returns:
        目录页SlideContent
    """
    toc_slide = SlideContent(
        title="目录" if not presentation_title else f"{presentation_title} - 目录",
        subtitle="",
        content_lines=[],
        bullet_points=[]
    )

    # 将目录项转换为bullet points
    for entry in toc_entries:
        # 格式: "标题 (第X页)"
        page_num = entry.slide_index + 1
        bullet_text = f"{entry.title}"
        toc_slide.bullet_points.append(bullet_text)

    return toc_slide


def insert_toc_at_beginning(presentation: PresentationContent, include_subtitles: bool = True) -> PresentationContent:
    """
    在演示文稿开头插入目录页

    Args:
        presentation: 原始演示文稿
        include_subtitles: 是否包含副标题

    Returns:
        带目录页的演示文稿
    """
    # 提取目录项
    toc_entries = extract_toc_entries(presentation, include_subtitles)

    if not toc_entries:
        return presentation

    # 创建目录页
    toc_slide = create_toc_slide(toc_entries, presentation.title)

    # 在开头插入目录页
    new_slides = [toc_slide] + presentation.slides

    # 更新slide_index(因为插入了目录页,所有索引+1)
    updated_toc_entries = []
    for entry in toc_entries:
        updated_toc_entries.append(TocEntry(
            title=entry.title,
            slide_index=entry.slide_index + 1,  # 索引+1
            level=entry.level,
            subtitle=entry.subtitle
        ))

    new_presentation = PresentationContent(
        title=presentation.title,
        slides=new_slides,
        auto_toc=True
    )

    return new_presentation


def generate_toc_summary(toc_entries: List[TocEntry]) -> str:
    """
    生成目录摘要文本

    Args:
        toc_entries: 目录项列表

    Returns:
        目录摘要字符串
    """
    lines = []
    for entry in toc_entries:
        indent = "  " * (entry.level - 1)
        page_num = entry.slide_index + 1
        line = f"{indent}- {entry.title} (第{page_num}页)"
        if entry.subtitle:
            line += f" - {entry.subtitle}"
        lines.append(line)

    return "\n".join(lines)
