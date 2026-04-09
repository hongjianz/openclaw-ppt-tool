#!/usr/bin/env python3
"""
PPT自动生成工具 - 主入口脚本

用法:
    python main.py --input <文稿文件> --output <输出文件.pptx> [选项]

选项:
    --input, -i       输入文稿文件路径 (Markdown或纯文本)
    --output, -o      输出PPT文件路径
    --config, -c      模板配置文件路径 (JSON格式)
    --background, -b  背景图片路径
    --css             CSS样式文件路径
    --format          文稿格式 (markdown/plain/auto, 默认auto)
    --width           幻灯片宽度 (英寸, 默认13.333)
    --height          幻灯片高度 (英寸, 默认7.5)
"""

import argparse
import os
import sys

# 添加src目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.template_config import TemplateConfig, parse_css_to_config
from src.content_parser import smart_parse
from src.ppt_generator import PPTGenerator


def main():
    parser = argparse.ArgumentParser(
        description='PPT自动生成工具 - 根据文稿内容自动生成演示文稿',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 使用默认配置生成PPT
  python main.py -i content.md -o presentation.pptx

  # 使用自定义背景图片
  python main.py -i content.md -o presentation.pptx -b background.png

  # 使用CSS样式配置
  python main.py -i content.txt -o presentation.pptx --css styles.css

  # 使用JSON配置文件
  python main.py -i content.md -o presentation.pptx -c config.json
        """
    )

    parser.add_argument('-i', '--input', required=True, help='输入文稿文件路径')
    parser.add_argument('-o', '--output', required=True, help='输出PPT文件路径')
    parser.add_argument('-c', '--config', help='模板配置文件路径 (JSON格式)')
    parser.add_argument('-b', '--background', help='背景图片路径')
    parser.add_argument('--css', help='CSS样式文件路径')
    parser.add_argument('--format', choices=['markdown', 'plain', 'auto'],
                       default='auto', help='文稿格式 (默认: auto)')
    parser.add_argument('--width', type=float, default=13.333,
                       help='幻灯片宽度 (英寸, 默认: 13.333)')
    parser.add_argument('--height', type=float, default=7.5,
                       help='幻灯片高度 (英寸, 默认: 7.5)')

    args = parser.parse_args()

    # 检查输入文件是否存在
    if not os.path.exists(args.input):
        print(f"错误: 输入文件不存在: {args.input}")
        sys.exit(1)

    # 加载配置
    config = None

    # 优先使用JSON配置文件
    if args.config and os.path.exists(args.config):
        print(f"加载配置文件: {args.config}")
        config = TemplateConfig.from_json(args.config)
    else:
        config = TemplateConfig()

    # 应用命令行参数覆盖
    if args.background:
        if os.path.exists(args.background):
            config.background_image = args.background
            print(f"设置背景图片: {args.background}")
        else:
            print(f"警告: 背景图片不存在: {args.background}")

    # 加载CSS样式
    if args.css and os.path.exists(args.css):
        print(f"加载CSS样式: {args.css}")
        with open(args.css, 'r', encoding='utf-8') as f:
            css_content = f.read()
        css_config = parse_css_to_config(css_content)
        # 合并CSS配置
        for key, value in vars(css_config).items():
            if value != getattr(TemplateConfig(), key):
                setattr(config, key, value)

    # 设置页面尺寸
    config.slide_width = args.width
    config.slide_height = args.height

    # 读取文稿内容
    print(f"读取文稿: {args.input}")
    with open(args.input, 'r', encoding='utf-8') as f:
        content_text = f.read()

    # 解析文稿
    print(f"解析文稿内容 (格式: {args.format})...")
    content = smart_parse(content_text, args.format)
    print(f"解析完成: {len(content.slides)} 页内容")

    # 生成PPT
    print("生成PPT...")
    generator = PPTGenerator(config)
    generator.generate(content, args.output)

    print("完成!")


if __name__ == '__main__':
    main()
