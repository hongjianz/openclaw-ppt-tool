#!/usr/bin/env python3
"""
表格渲染问题诊断工具

使用方法:
    python diagnose_table.py <input_file.md>

这个脚本会详细显示表格解析的每一步，帮助定位问题。
"""

import sys
import os

if len(sys.argv) < 2:
    print("用法: python diagnose_table.py <input_file.md>")
    print("\n示例:")
    print("  python diagnose_table.py examples/table_example.md")
    sys.exit(1)

input_file = sys.argv[1]

if not os.path.exists(input_file):
    print(f"错误: 文件不存在 - {input_file}")
    sys.exit(1)

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from content_parser import parse_markdown
from template_config import TemplateConfig
from ppt_generator import PPTGenerator

print("=" * 80)
print("表格渲染问题诊断工具")
print("=" * 80)

# 读取文件
with open(input_file, 'r', encoding='utf-8') as f:
    content = f.read()

print(f"\n文件: {input_file}")
print(f"大小: {len(content)} 字符")
print(f"行数: {len(content.splitlines())} 行")

# 解析
print("\n" + "=" * 80)
print("步骤1: 解析Markdown内容")
print("=" * 80)

result = parse_markdown(content)

print(f"\n文稿标题: {result.title}")
print(f"页面数量: {len(result.slides)}")

table_slides = []
for i, slide in enumerate(result.slides, 1):
    has_table = slide.table is not None
    status = "✓ 有表格" if has_table else "✗ 无表格"
    print(f"\n第{i}页: {slide.title} - {status}")

    if has_table:
        table_slides.append(i)
        print(f"  表头: {slide.table.headers}")
        print(f"  列数: {len(slide.table.headers)}")
        print(f"  数据行数: {len(slide.table.rows)}")
        print(f"  对齐方式: {slide.table.alignment}")

        for j, row in enumerate(slide.table.rows, 1):
            print(f"  行{j}: {row}")

print("\n" + "=" * 80)
print(f"总结: 共 {len(table_slides)} 页包含表格 (页码: {table_slides})")
print("=" * 80)

if not table_slides:
    print("\n⚠ 警告: 没有找到任何表格!")
    print("可能的原因:")
    print("  1. Markdown表格语法不正确")
    print("  2. 表格前后缺少空行")
    print("  3. 分隔行格式不正确")
    print("\n正确的Markdown表格格式:")
    print("""
| 列1 | 列2 | 列3 |
|:---:|:---:|:---:|
| A   | B   | C   |
| D   | E   | F   |
    """)
else:
    print("\n✓ 表格解析成功!")
    print("\n现在尝试生成PPT...")

    # 生成PPT
    output_file = input_file.replace('.md', '_test.pptx').replace('.txt', '_test.pptx')
    if output_file == input_file:
        output_file = 'output_test.pptx'

    try:
        config = TemplateConfig()
        generator = PPTGenerator(config)
        generator.generate(result, output_file)

        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"\n✓ PPT生成成功: {output_file}")
            print(f"  文件大小: {file_size:,} bytes ({file_size/1024:.1f} KB)")
            print(f"\n请打开 {output_file} 查看表格是否正确显示")
        else:
            print(f"\n✗ PPT文件未生成")
    except Exception as e:
        print(f"\n✗ PPT生成失败: {e}")
        import traceback
        traceback.print_exc()

print("\n" + "=" * 80)
