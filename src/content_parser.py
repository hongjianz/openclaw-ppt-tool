"""
文稿解析模块 - 解析Markdown或纯文本内容为结构化PPT页面数据
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class TableCell:
    """表格单元格"""
    content: str
    is_header: bool = False


@dataclass
class TableData:
    """表格数据"""
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    alignment: List[str] = field(default_factory=list)  # 'left', 'center', 'right'


@dataclass
class CodeBlock:
    """代码块"""
    content: str
    language: str = "text"  # python, java, text, flowchart, matrix等
    is_diagram: bool = False  # 是否为图表（流程图/矩阵图）


@dataclass
class SlideContent:
    """单页PPT内容"""
    title: str
    subtitle: str = ""
    content_lines: List[str] = field(default_factory=list)
    bullet_points: List[str] = field(default_factory=list)
    table: Optional[TableData] = None  # 新增表格支持
    image_path: Optional[str] = None
    images: List[str] = field(default_factory=list)  # 支持多张图片
    code_blocks: List[CodeBlock] = field(default_factory=list)  # 代码块


@dataclass
class PresentationContent:
    """完整演示文稿内容"""
    title: str = ""
    slides: List[SlideContent] = field(default_factory=list)
    auto_toc: bool = False  # 是否自动生成目录


def parse_markdown(content: str) -> PresentationContent:
    """
    解析Markdown格式的文稿内容

    支持以下Markdown语法:
    - # 标题 -> PPT页面标题
    - ## 副标题 -> PPT页面副标题
    - - 或 * 列表项 -> 项目符号
    - | 表格 | 语法 | -> PPT表格对象
    - ![alt](url) -> 图片
    - ``` 代码块 -> 等宽字体文本框或图表
    - 普通文本 -> 正文内容
    - --- 分隔符 -> 分页
    """
    presentation = PresentationContent()
    current_slide = None
    in_table = False
    table_headers = []
    table_rows = []
    table_alignment = []
    prev_line_was_separator = False
    
    # 代码块相关
    in_code_block = False
    code_lines = []
    code_language = "text"

    lines = content.split('\n')
    i = 0

    while i < len(lines):
        line = lines[i].strip()
        i += 1

        # 跳过空行
        if not line:
            # 如果之前在解析表格,空行表示表格结束
            if in_table and table_headers:
                current_slide.table = TableData(
                    headers=table_headers,
                    rows=table_rows,
                    alignment=table_alignment
                )
                in_table = False
                table_headers = []
                table_rows = []
                table_alignment = []
                prev_line_was_separator = False
            continue

        # 水平分隔线 - 新页面
        if line.startswith('---'):
            if current_slide:
                # 如果有未保存的表格
                if in_table and table_headers:
                    current_slide.table = TableData(
                        headers=table_headers,
                        rows=table_rows,
                        alignment=table_alignment
                    )
                    in_table = False
                    table_headers = []
                    table_rows = []
                    table_alignment = []
                
                # 如果有未保存的代码块
                if in_code_block and code_lines:
                    code_content = '\n'.join(code_lines)
                    is_diagram = _is_diagram(code_content)
                    current_slide.code_blocks.append(CodeBlock(
                        content=code_content,
                        language=code_language,
                        is_diagram=is_diagram
                    ))
                    in_code_block = False
                    code_lines = []
                    
                presentation.slides.append(current_slide)
            current_slide = SlideContent(title="", content_lines=[], bullet_points=[])
            continue

        # 检测代码块开始/结束
        if line.startswith('```'):
            if in_code_block:
                # 结束代码块
                code_content = '\n'.join(code_lines)
                is_diagram = _is_diagram(code_content)
                if current_slide:
                    current_slide.code_blocks.append(CodeBlock(
                        content=code_content,
                        language=code_language,
                        is_diagram=is_diagram
                    ))
                in_code_block = False
                code_lines = []
                code_language = "text"
            else:
                # 开始代码块
                in_code_block = True
                code_language = line[3:].strip() or "text"
            continue
        
        # 如果在代码块内，收集代码行
        if in_code_block:
            code_lines.append(line)
            continue

        # 检测Markdown表格
        if '|' in line and not line.startswith('#') and not line.startswith('- ') and not line.startswith('* '):
            # 检查是否是表头分隔行 (|---|---| 或 |:---|:---:)
            separator_match = re.match(r'^[\s|:\-]+$', line)
            if separator_match and all(c in ' |-:' for c in line):
                # 这是表格分隔行,解析对齐方式
                parts = [p.strip() for p in line.split('|')]
                # 移除首尾空元素
                if parts and parts[0] == '':
                    parts = parts[1:]
                if parts and parts[-1] == '':
                    parts = parts[:-1]
                
                table_alignment = []
                for part in parts:
                    if part.startswith(':') and part.endswith(':'):
                        table_alignment.append('center')
                    elif part.endswith(':'):
                        table_alignment.append('right')
                    else:
                        table_alignment.append('left')
                prev_line_was_separator = True
                continue

            # 表格行
            parts = [p.strip() for p in line.split('|')]
            # 移除首尾空元素
            if parts and parts[0] == '':
                parts = parts[1:]
            if parts and parts[-1] == '':
                parts = parts[:-1]

            if not prev_line_was_separator and not in_table:
                # 这是表头
                table_headers = parts
                in_table = True
            elif in_table:
                # 这是数据行
                table_rows.append(parts)

            prev_line_was_separator = False
            continue

        # 如果之前在解析表格,现在结束了
        if in_table and not ('|' in line):
            # 保存表格到当前slide
            if current_slide and table_headers:
                current_slide.table = TableData(
                    headers=table_headers,
                    rows=table_rows,
                    alignment=table_alignment
                )
            in_table = False
            table_headers = []
            table_rows = []
            table_alignment = []
            prev_line_was_separator = False

        # 一级标题 - 新页面标题
        if line.startswith('# ') and not line.startswith('## '):
            if current_slide is None:
                current_slide = SlideContent(title="", content_lines=[], bullet_points=[])

            # 如果当前页面已有内容,先保存
            if current_slide and current_slide.title and (current_slide.content_lines or current_slide.bullet_points or current_slide.table):
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

        # 图片语法 ![alt](url)
        img_match = re.match(r'!\[([^\]]*)\]\(([^)]+)\)', line)
        if img_match:
            if current_slide is None:
                current_slide = SlideContent(title="", content_lines=[], bullet_points=[])
            current_slide.images.append(img_match.group(2))
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

    # 处理最后的表格
    if current_slide and in_table and table_headers:
        current_slide.table = TableData(
            headers=table_headers,
            rows=table_rows,
            alignment=table_alignment
        )

    # 处理最后未关闭的代码块
    if current_slide and in_code_block and code_lines:
        code_content = '\n'.join(code_lines)
        is_diagram = _is_diagram(code_content)
        current_slide.code_blocks.append(CodeBlock(
            content=code_content,
            language=code_language,
            is_diagram=is_diagram
        ))

    # 添加最后一页
    if current_slide:
        presentation.slides.append(current_slide)

    # 设置演示文稿标题
    if presentation.slides:
        presentation.title = presentation.slides[0].title

    return presentation


def _is_diagram(code_content: str) -> bool:
    """
    检测代码内容是否为图表（流程图或矩阵图）

    Args:
        code_content: 代码块内容

    Returns:
        是否为图表类型
    """
    lines = code_content.strip().split('\n')

    # 检测ASCII流程图特征
    flowchart_patterns = [
        r'[┌┬┐└┴┘├┤┼─│]',  # 框线字符
        r'-->',              # 箭头
        r'=>',               # 双箭头
        r'<->',              # 双向箭头
        r'\[.*?\]--',       # [节点]--
        r'--\[.*?\]',       # --[节点]
    ]

    # 检测矩阵图特征
    matrix_patterns = [
        r'^[\s|\-+]*$',      # 纯线条行
        r'\d+\s+\d+',        # 数字矩阵
    ]

    for line in lines:
        for pattern in flowchart_patterns:
            if re.search(pattern, line):
                return True

    # 检查是否有多行包含框线字符（可能是矩阵）
    box_char_count = sum(1 for line in lines if re.search(r'[┌┬┐└┴┘├┤┼─│+|]', line))
    if box_char_count >= 2:
        return True

    return False


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
