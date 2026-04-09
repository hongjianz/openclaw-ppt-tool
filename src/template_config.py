"""
模板配置模块 - 处理PPT背景模板和样式配置
支持图片背景和CSS样式配置
"""

import json
import os
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class TemplateConfig:
    """PPT模板配置"""

    # 页面尺寸 (英寸)
    slide_width: float = 13.333  # 16:9 默认宽度
    slide_height: float = 7.5     # 16:9 默认高度

    # 背景配置
    background_image: Optional[str] = None  # 背景图片路径

    # 颜色配置
    primary_color: str = "#1f4e79"      # 主色调
    secondary_color: str = "#2e75b6"    # 次要色
    text_color: str = "#333333"         # 文本颜色
    background_color: str = "#ffffff"   # 背景色

    # 字体配置
    title_font: str = "Microsoft YaHei"     # 标题字体
    content_font: str = "Microsoft YaHei"   # 正文字体
    title_size: int = 32                    # 标题字号
    subtitle_size: int = 24                 # 副标题字号
    body_size: int = 18                     # 正文字号

    # 边距配置 (英寸)
    margin_left: float = 0.8
    margin_right: float = 0.8
    margin_top: float = 0.8
    margin_bottom: float = 0.8

    # 布局配置
    max_chars_per_line: int = 40        # 每行最大字符数
    max_lines_per_slide: int = 12       # 每页最大行数
    line_spacing: float = 1.2           # 行间距倍数

    # 特殊样式标记
    style_name: str = ""                # 样式名称(如 QuintenStyle)
    use_gradient_background: bool = False  # 是否使用渐变背景

    @classmethod
    def from_json(cls, config_path: str) -> "TemplateConfig":
        """从JSON配置文件加载配置"""
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return cls(**{k: v for k, v in data.items() if hasattr(cls, k)})

    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'slide_width': self.slide_width,
            'slide_height': self.slide_height,
            'background_image': self.background_image,
            'primary_color': self.primary_color,
            'secondary_color': self.secondary_color,
            'text_color': self.text_color,
            'background_color': self.background_color,
            'title_font': self.title_font,
            'content_font': self.content_font,
            'title_size': self.title_size,
            'subtitle_size': self.subtitle_size,
            'body_size': self.body_size,
            'margin_left': self.margin_left,
            'margin_right': self.margin_right,
            'margin_top': self.margin_top,
            'margin_bottom': self.margin_bottom,
            'max_chars_per_line': self.max_chars_per_line,
            'max_lines_per_slide': self.max_lines_per_slide,
            'line_spacing': self.line_spacing,
        }


def parse_css_to_config(css_content: str) -> TemplateConfig:
    """
    解析CSS样式到模板配置
    支持基本的CSS属性映射和特殊样式识别
    """
    config = TemplateConfig()

    # 检测QuintenStyle
    if 'QuintenStyle' in css_content or '#267878' in css_content:
        config.style_name = "QuintenStyle"
        config.background_color = "#267878"
        config.text_color = "#FFFFFF"
        config.primary_color = "#FFFFFF"
        config.secondary_color = "#E0E0E0"
        # QuintenStyle需要使用背景图片来实现复杂渐变
        config.use_gradient_background = True

    # 简单的CSS解析
    css_rules = {}
    current_selector = None

    for line in css_content.split('\n'):
        line = line.strip()
        if not line or line.startswith('/*'):
            continue

        if '{' in line:
            current_selector = line.split('{')[0].strip()
            css_rules[current_selector] = {}
        elif '}' in line:
            current_selector = None
        elif current_selector and ':' in line:
            prop, value = line.split(':', 1)
            css_rules[current_selector][prop.strip()] = value.strip().rstrip(';')

    # 映射CSS属性到配置
    for selector, props in css_rules.items():
        # 跳过已处理的QuintenStyle特殊标记
        if config.style_name == "QuintenStyle":
            # 可以提取其他非颜色相关配置
            if 'font-family' in props:
                font = props['font-family'].split(',')[0].strip().strip('"\'')
                config.title_font = font
                config.content_font = font
            if 'font-size' in props:
                try:
                    size = int(props['font-size'].replace('px', '').replace('pt', ''))
                    config.body_size = size
                except ValueError:
                    pass
            continue

        if 'color' in props:
            color_val = props['color']
            # 处理 !important
            color_val = color_val.replace('!important', '').strip()
            # 跳过rgba等PPT不支持的格式
            if color_val.startswith('#') and len(color_val) == 7:
                config.text_color = color_val
        if 'background-color' in props:
            bg_val = props['background-color'].replace('!important', '').strip()
            if bg_val.startswith('#') and len(bg_val) == 7:
                config.background_color = bg_val
        if 'font-family' in props:
            font = props['font-family'].split(',')[0].strip().strip('"\'')
            config.title_font = font
            config.content_font = font
        if 'font-size' in props:
            try:
                size = int(props['font-size'].replace('px', '').replace('pt', ''))
                config.body_size = size
            except ValueError:
                pass

    return config
