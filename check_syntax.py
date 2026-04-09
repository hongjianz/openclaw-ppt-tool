#!/usr/bin/env python3
"""检查Python文件语法"""
import py_compile
import sys

files = [
    'src/content_parser.py',
    'src/ppt_generator.py',
    'src/toc_generator.py',
    'test_ppt_tool.py',
    'main.py'
]

errors = []
for filepath in files:
    try:
        py_compile.compile(filepath, doraise=True)
        print(f"✓ {filepath}")
    except py_compile.PyCompileError as e:
        print(f"✗ {filepath}: {e}")
        errors.append((filepath, str(e)))

if errors:
    print(f"\n{len(errors)} 个文件有语法错误")
    sys.exit(1)
else:
    print("\n所有文件语法检查通过")
    sys.exit(0)
