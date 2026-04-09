"""
文稿解析模块 - 解析Markdown或纯文本内容为结构化PPT页面数据
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class SlideContent:
    """单页PPT内容"""
    title: str
    subtitle: str = ""
    content_lines: List[str] = field(default_factory=list)
    bullet_points: List[str] = field(default_factory=list)
    image_path: Optional[str] = None


@dataclass
class PresentationContent:
    """完整演示文稿内容"""
    title: str = ""
    slides: List[SlideContent] = field(default_factory=list)


def parse_markdown(content: str) -> PresentationContent:
    """
    解析Markdown格式的文稿内容

    支持以下Markdown语法:
    - # 标题 -> PPT页面标题
    - ## 副标题 -> PPT页面副标题
    - - 或 * 列表项 -> 项目符号
    - 普通文本 -> 正文内容
    - --- 分隔符 -> 分页
    """
    presentation = PresentationContent()
    current_slide = None

    lines = content.split('\n')

    for line in lines:
        line = line.strip()

        # 跳过空行
        if not line:
            continue

        # 水平分隔线 - 新页面
        if line.startswith('---'):
            if current_slide:
                presentation.slides.append(current_slide)
            current_slide = SlideContent(title="", content_lines=[], bullet_points=[])
            continue

        # 一级标题 - 新页面标题
        if line.startswith('# ') and not line.startswith('## '):
            if current_slide is None:
                current_slide = SlideContent(title="", content_lines=[], bullet_points=[])

            # 如果当前页面已有内容,先保存
            if current_slide.title and (current_slide.content_lines or current_slide.bullet_points):
                presentation.slides.append(current_slide)
                current_slide = SlideContent(title="", content_lines=[], bullet_points=[])

            current_slide.title = line[2:].strip()
            continue

        # 二级标题 - 副标题
        if line.startswith('## '):
            if current_slide is None:
                current_slide = SlideContent(title="", content_lines=[], bullet_points=[])
            current_slide.subtitle = line[3:].strip()
            continue

        # 列表项
        if line.startswith('- ') or line.startswith('* '):
            if current_slide is None:
                current_slide = SlideContent(title="", content_lines=[], bullet_points=[])
            current_slide.bullet_points.append(line[2:].strip())
            continue

        # 普通文本
        if current_slide is None:
            current_slide = SlideContent(title="", content_lines=[], bullet_points=[])

        # 如果还没有标题,将第一行作为标题
        if not current_slide.title:
            current_slide.title = line
        else:
            current_slide.content_lines.append(line)

    # 添加最后一页
    if current_slide:
        presentation.slides.append(current_slide)

    # 设置演示文稿标题
    if presentation.slides:
        presentation.title = presentation.slides[0].title

    return presentation


def parse_plain_text(content: str, chars_per_page: int = 500) -> PresentationContent:
    """
    解析纯文本格式的文稿内容

    按字符数自动分页,每页约chars_per_page个字符
    """
    presentation = PresentationContent()

    # 清理文本
    lines = [line.strip() for line in content.split('\n') if line.strip()]

    if not lines:
        return presentation

    # 第一行作为标题
    presentation.title = lines[0]

    # 按字符数分页
    current_chars = 0
    current_slide = SlideContent(
        title=lines[0],
        content_lines=[],
        bullet_points=[]
    )

    for i, line in enumerate(lines[1:], 1):
        current_chars += len(line) + 1  # +1 for newline

        # 达到每页字符限制,创建新页面
        if current_chars >= chars_per_page and current_slide.content_lines:
            presentation.slides.append(current_slide)
            current_slide = SlideContent(
                title=f"第{len(presentation.slides) + 1}页",
                content_lines=[],
                bullet_points=[]
            )
            current_chars = len(line)

        # 如果没有标题,使用第一行
        if not current_slide.title and i == 1:
            current_slide.title = line
        else:
            current_slide.content_lines.append(line)

    # 添加最后一页
    if current_slide.content_lines or current_slide.title:
        presentation.slides.append(current_slide)

    return presentation


def smart_parse(content: str, format_hint: str = "auto") -> PresentationContent:
    """
    智能解析文稿内容

    Args:
        content: 文稿内容
        format_hint: 格式提示 ("markdown", "plain", "auto")

    Returns:
        PresentationContent: 结构化的演示文稿内容
    """
    if format_hint == "auto":
        # 检测是否包含Markdown语法
        if any(line.strip().startswith('#') or line.strip().startswith('- ')
               or line.strip().startswith('* ') or '---' in line
               for line in content.split('\n')):
            return parse_markdown(content)
        else:
            return parse_plain_text(content)
    elif format_hint == "markdown":
        return parse_markdown(content)
    else:
        return parse_plain_text(content)
