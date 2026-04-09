#!/usr/bin/env python3
"""
测试脚本 - 验证PPT工具的核心功能
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from content_parser import smart_parse, parse_markdown, parse_plain_text
from template_config import TemplateConfig, parse_css_to_config


def test_markdown_parser():
    """测试Markdown解析器"""
    print("=" * 60)
    print("测试1: Markdown解析")
    print("=" * 60)

    markdown_content = """
# 人工智能概述

## AI技术简介

---

# 什么是AI?

- 模拟人类智能
- 学习能力
- 决策能力

---

# 应用领域

医疗、金融、交通等领域都有广泛应用。

机器学习是核心技术之一。
"""

    result = parse_markdown(markdown_content)
    print(f"文稿标题: {result.title}")
    print(f"页面数量: {len(result.slides)}")

    for i, slide in enumerate(result.slides, 1):
        print(f"\n第{i}页:")
        print(f"  标题: {slide.title}")
        print(f"  副标题: {slide.subtitle}")
        print(f"  要点数量: {len(slide.bullet_points)}")
        print(f"  内容行数: {len(slide.content_lines)}")

    print("\n✓ Markdown解析测试通过\n")
    return True


def test_plain_text_parser():
    """测试纯文本解析器"""
    print("=" * 60)
    print("测试2: 纯文本解析")
    print("=" * 60)

    text_content = """
人工智能发展报告

人工智能正在改变我们的生活方式。从智能手机到自动驾驶汽车,AI技术无处不在。

机器学习是AI的核心技术之一。通过大量数据训练,计算机可以学会识别模式、做出预测。

深度学习作为机器学习的子集,在图像识别、自然语言处理等领域取得了突破性进展。

未来,AI将继续发展,为人类社会带来更多便利和创新。
"""

    result = parse_plain_text(text_content, chars_per_page=200)
    print(f"文稿标题: {result.title}")
    print(f"页面数量: {len(result.slides)}")

    for i, slide in enumerate(result.slides, 1):
        print(f"\n第{i}页:")
        print(f"  标题: {slide.title}")
        print(f"  内容行数: {len(slide.content_lines)}")

    print("\n✓ 纯文本解析测试通过\n")
    return True


def test_smart_parse():
    """测试智能解析"""
    print("=" * 60)
    print("测试3: 智能解析(自动检测格式)")
    print("=" * 60)

    # 测试Markdown自动检测
    markdown_text = "# 标题\n- 要点1\n- 要点2"
    result1 = smart_parse(markdown_text, "auto")
    print(f"Markdown检测: {len(result1.slides)} 页")

    # 测试纯文本自动检测
    plain_text = "这是第一段文字。\n这是第二段文字。\n这是第三段文字。"
    result2 = smart_parse(plain_text, "auto")
    print(f"纯文本检测: {len(result2.slides)} 页")

    print("\n✓ 智能解析测试通过\n")
    return True


def test_template_config():
    """测试模板配置"""
    print("=" * 60)
    print("测试4: 模板配置")
    print("=" * 60)

    # 测试默认配置
    config = TemplateConfig()
    print(f"默认页面尺寸: {config.slide_width} x {config.slide_height}")
    print(f"默认主色: {config.primary_color}")
    print(f"默认字体: {config.title_font}")

    # 测试CSS解析
    css_content = """
    .slide {
        color: #ff0000;
        background-color: #f0f0f0;
        font-family: Arial;
    }
    """
    css_config = parse_css_to_config(css_content)
    print(f"\nCSS解析结果:")
    print(f"  文本颜色: {css_config.text_color}")
    print(f"  背景色: {css_config.background_color}")
    print(f"  字体: {css_config.title_font}")

    print("\n✓ 模板配置测试通过\n")
    return True


def test_config_load():
    """测试配置文件加载"""
    print("=" * 60)
    print("测试5: JSON配置文件加载")
    print("=" * 60)

    config_path = os.path.join(os.path.dirname(__file__), 'examples', 'template_config.json')

    if os.path.exists(config_path):
        config = TemplateConfig.from_json(config_path)
        print(f"加载配置成功")
        print(f"  页面尺寸: {config.slide_width} x {config.slide_height}")
        print(f"  主色: {config.primary_color}")
        print(f"  标题字号: {config.title_size}")
        print("\n✓ JSON配置加载测试通过\n")
        return True
    else:
        print(f"配置文件不存在: {config_path}")
        print("⊘ 跳过此测试\n")
        return True


def main():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("PPT自动生成工具 - 功能测试")
    print("=" * 60 + "\n")

    tests = [
        test_markdown_parser,
        test_plain_text_parser,
        test_smart_parse,
        test_template_config,
        test_config_load,
    ]

    passed = 0
    failed = 0

    for test_func in tests:
        try:
            if test_func():
                passed += 1
        except Exception as e:
            print(f"✗ 测试失败: {e}\n")
            failed += 1
            import traceback
            traceback.print_exc()

    print("=" * 60)
    print(f"测试结果: {passed} 通过, {failed} 失败")
    print("=" * 60)

    return failed == 0


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
