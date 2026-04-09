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
    --smart-paginate  启用智能分页(按语义单元而非字符数)
    --max-lines       每页最大行数 (默认12)
"""

import argparse
import os
import sys

# 添加src目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.template_config import TemplateConfig, parse_css_to_config
from src.content_parser import smart_parse
from src.smart_pagination import smart_paginate
from src.toc_generator import insert_toc_at_beginning
from src.ppt_generator import PPTGenerator
from src.error_handler import (
    validate_input_file, validate_output_path,
    check_dependencies, print_dependency_help,
    safe_generate_ppt, graceful_exit
)
from src.output_callback import FeishuUploader, create_upload_callback, create_chat_send_callback


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
    parser.add_argument('--smart-paginate', action='store_true',
                       help='启用智能分页(按语义单元而非字符数)')
    parser.add_argument('--smart-layout', action='store_true',
                       help='启用智能布局模式(预扫描+智能分页，避免重叠)')
    parser.add_argument('--max-lines', type=int, default=None,
                       help='每页最大行数 (默认使用配置文件)')
    parser.add_argument('--toc', action='store_true',
                       help='自动生成目录页')
    parser.add_argument('--no-subtitles-in-toc', action='store_true',
                       help='目录中不包含副标题')
    parser.add_argument('--footer', type=str, default='',
                       help='页脚文本')
    parser.add_argument('--no-page-number', action='store_true',
                       help='不显示页码')
    parser.add_argument('--page-format', type=str, default=None,
                       help='页码格式 (默认: 第 {current} / {total} 页)')
    parser.add_argument('--footer-on-title', action='store_true',
                       help='在标题页显示页脚')
    parser.add_argument('--content', type=str, default=None,
                       help='直接传入文稿内容(无需文件)')
    parser.add_argument('--feishu-doc', type=str, default=None,
                       help='飞书文档URL或ID')
    parser.add_argument('--upload-to-feishu', action='store_true',
                       help='上传到飞书云空间')
    parser.add_argument('--feishu-folder', type=str, default=None,
                       help='飞书文件夹token')
    parser.add_argument('--send-to-chat', type=str, default=None,
                       help='发送到飞书聊天(chat_id)')
    parser.add_argument('--chat-message', type=str, default=None,
                       help='聊天消息文本')

    args = parser.parse_args()

    # 检查依赖
    missing_deps = check_dependencies()
    if missing_deps:
        print_dependency_help(missing_deps)
        graceful_exit("请先安装缺失的依赖", 1)

    # 验证输出路径
    if not validate_output_path(args.output):
        graceful_exit("输出路径验证失败", 1)

    # 处理输入来源
    content_text = None

    # 优先级1: --content 参数直接传入内容
    if args.content:
        print("使用命令行传入的内容...")
        content_text = args.content

    # 优先级2: --feishu-doc 从飞书文档读取
    elif args.feishu_doc:
        print(f"从飞书文档读取: {args.feishu_doc}")
        try:
            from src.feishu_reader import FeishuDocReader
            reader = FeishuDocReader()
            content_text = reader.read_document(args.feishu_doc)
            print(f"成功读取飞书文档 ({len(content_text)} 字符)")
        except ImportError:
            graceful_exit("错误: 飞书集成模块未安装\n运行: pip install requests", 1)
        except Exception as e:
            graceful_exit(f"读取飞书文档失败: {e}", 1)

    # 优先级3: --input 从文件读取
    elif args.input:
        if not validate_input_file(args.input):
            graceful_exit("输入文件验证失败", 1)
        print(f"读取文稿: {args.input}")
        with open(args.input, 'r', encoding='utf-8') as f:
            content_text = f.read()
    else:
        graceful_exit("错误: 请指定 --input、--content 或 --feishu-doc", 1)

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

    # 设置最大行数
    if args.max_lines:
        config.max_lines_per_slide = args.max_lines

    # 配置页脚和页码
    if args.footer:
        config.footer_text = args.footer
    if args.no_page_number:
        config.show_page_number = False
    if args.page_format:
        config.page_number_format = args.page_format
    if args.footer_on_title:
        config.show_footer_on_title_slide = True

    # 读取文稿内容
    print(f"读取文稿: {args.input}")
    with open(args.input, 'r', encoding='utf-8') as f:
        content_text = f.read()

    # 解析文稿
    print(f"解析文稿内容 (格式: {args.format})...")
    content = smart_parse(content_text, args.format)
    print(f"解析完成: {len(content.slides)} 页内容")

    # 智能分页
    if args.smart_paginate:
        print("执行智能分页优化...")
        content = smart_paginate(content, config)
        print(f"分页完成: {len(content.slides)} 页内容")

    # 生成目录页
    if args.toc:
        print("生成目录页...")
        include_subtitles = not args.no_subtitles_in_toc
        content = insert_toc_at_beginning(content, include_subtitles)
        print(f"目录页已添加: 共 {len(content.slides)} 页")

    # 生成PPT
    print("生成PPT...")
    try:
        generator = PPTGenerator(config)
        
        # 选择生成模式
        if args.smart_layout:
            # 使用智能布局模式
            print("使用智能布局模式（预扫描 + 智能分页）...")
            generator.generate_with_smart_layout(content, args.output)
        else:
            # 使用传统模式
            success = safe_generate_ppt(generator, content, args.output)
            
            if not success:
                graceful_exit("PPT生成失败,请检查错误信息", 1)

        # 执行输出回调
        if args.upload_to_feishu or args.send_to_chat:
            try:
                uploader = FeishuUploader()

                # 上传到云空间
                if args.upload_to_feishu:
                    callback = create_upload_callback(uploader, args.feishu_folder)
                    callback(args.output)

                # 发送到聊天
                if args.send_to_chat:
                    callback = create_chat_send_callback(
                        uploader,
                        args.send_to_chat,
                        args.chat_message
                    )
                    callback(args.output)

            except Exception as e:
                print(f"警告: 回调执行失败: {e}")
                print("但PPT文件已成功生成,可手动上传")

        graceful_exit("完成!", 0)

    except Exception as e:
        logger.error(f"未知错误: {e}")
        import traceback
        traceback.print_exc()
        graceful_exit("程序异常退出", 1)


if __name__ == '__main__':
    main()
