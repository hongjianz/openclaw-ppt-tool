#!/usr/bin/env python3
"""
图表渲染功能测试脚本

测试内容:
1. 原生图表解析和生成
2. Mermaid图表检测和降级
3. 混合内容渲染
"""

import sys
import os

# 添加src目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.content_parser import parse_markdown, ChartData, ChartSeries
from src.ppt_generator import PPTGenerator
from src.template_config import TemplateConfig


def test_chart_parsing():
    """测试图表解析功能"""
    print("=" * 60)
    print("测试1: 图表解析")
    print("=" * 60)

    # 测试柱状图解析
    bar_chart_content = """
# 图表测试

:::chart{type="bar" title="季度销售数据"}
| 季度 | Q1 | Q2 | Q3 | Q4 |
|------|----|----|----|----|
| 销售额 | 19.2 | 21.4 | 16.7 | 22.1 |
:::
"""

    print("\n测试柱状图解析...")
    result = parse_markdown(bar_chart_content)
    assert len(result.slides) > 0, "解析失败: 没有幻灯片"
    
    slide = result.slides[0]
    assert len(slide.charts) > 0, "解析失败: 没有图表"
    
    chart = slide.charts[0]
    assert chart.chart_type == "bar", f"图表类型错误: {chart.chart_type}"
    assert chart.title == "季度销售数据", f"图表标题错误: {chart.title}"
    assert len(chart.categories) == 4, f"分类数量错误: {len(chart.categories)}"
    assert len(chart.series) == 1, f"数据系列数量错误: {len(chart.series)}"
    assert chart.series[0].name == "销售额", f"数据系列名称错误: {chart.series[0].name}"
    
    print("✓ 柱状图解析成功")
    print(f"  - 类型: {chart.chart_type}")
    print(f"  - 标题: {chart.title}")
    print(f"  - 分类: {chart.categories}")
    print(f"  - 数据: {chart.series[0].values}")

    # 测试饼图解析
    pie_chart_content = """
# 饼图测试

:::chart{type="pie" title="市场份额"}
- 产品A: 35
- 产品B: 25
- 产品C: 20
- 产品D: 20
:::
"""

    print("\n测试饼图解析...")
    result = parse_markdown(pie_chart_content)
    slide = result.slides[0]
    assert len(slide.charts) > 0, "解析失败: 没有图表"
    
    chart = slide.charts[0]
    assert chart.chart_type == "pie", f"图表类型错误: {chart.chart_type}"
    assert len(chart.categories) == 4, f"分类数量错误: {len(chart.categories)}"
    
    print("✓ 饼图解析成功")
    print(f"  - 类型: {chart.chart_type}")
    print(f"  - 标题: {chart.title}")
    print(f"  - 分类: {chart.categories}")

    # 测试多系列图表解析
    multi_series_content = """
# 多系列测试

:::chart{type="column" title="产品对比"}
| 产品 | 2023年 | 2024年 | 2025年 |
|------|--------|--------|--------|
| 产品A | 120 | 150 | 180 |
| 产品B | 90 | 110 | 140 |
:::
"""

    print("\n测试多系列图表解析...")
    result = parse_markdown(multi_series_content)
    slide = result.slides[0]
    chart = slide.charts[0]
    
    assert len(chart.series) == 2, f"数据系列数量错误: {len(chart.series)}"
    assert chart.series[0].name == "产品A", f"数据系列名称错误: {chart.series[0].name}"
    assert chart.series[1].name == "产品B", f"数据系列名称错误: {chart.series[1].name}"
    
    print("✓ 多系列图表解析成功")
    print(f"  - 系列1: {chart.series[0].name} = {chart.series[0].values}")
    print(f"  - 系列2: {chart.series[1].name} = {chart.series[1].values}")

    print("\n✓ 所有图表解析测试通过!\n")


def test_mermaid_detection():
    """测试Mermaid检测功能"""
    print("=" * 60)
    print("测试2: Mermaid检测")
    print("=" * 60)

    mermaid_content = """
# Mermaid测试

```mermaid
graph TD
    A[开始] --> B{判断}
    B -->|是| C[处理]
    B -->|否| D[结束]
```
"""

    print("\n测试Mermaid代码块解析...")
    result = parse_markdown(mermaid_content)
    slide = result.slides[0]
    
    assert len(slide.code_blocks) > 0, "解析失败: 没有代码块"
    code_block = slide.code_blocks[0]
    
    assert code_block.language == "mermaid", f"语言检测错误: {code_block.language}"
    assert "graph TD" in code_block.content, "代码内容错误"
    
    print("✓ Mermaid代码块解析成功")
    print(f"  - 语言: {code_block.language}")
    print(f"  - 内容长度: {len(code_block.content)} 字符")

    print("\n✓ Mermaid检测测试通过!\n")


def test_ppt_generation():
    """测试PPT生成功能"""
    print("=" * 60)
    print("测试3: PPT生成")
    print("=" * 60)

    test_content = """
# 图表生成测试

## 柱状图

:::chart{type="bar" title="销售数据"}
| 季度 | Q1 | Q2 | Q3 | Q4 |
|------|----|----|----|----|
| 销售额 | 100 | 150 | 120 | 180 |
:::

## 饼图

:::chart{type="pie" title="市场份额"}
- 产品A: 35
- 产品B: 25
- 产品C: 20
- 产品D: 20
:::
"""

    print("\n解析测试内容...")
    content = parse_markdown(test_content)
    print(f"✓ 解析完成: {len(content.slides)} 页")

    print("\n配置PPT生成器...")
    config = TemplateConfig()
    config.default_chart_width = 8.0
    config.default_chart_height = 4.0
    
    generator = PPTGenerator(config)

    output_path = os.path.join(os.path.dirname(__file__), 'test_output_charts.pptx')
    print(f"生成PPT文件: {output_path}")

    try:
        generator.generate(content, output_path)
        print(f"✓ PPT生成成功: {output_path}")
        
        # 检查文件是否存在
        assert os.path.exists(output_path), "PPT文件不存在"
        file_size = os.path.getsize(output_path)
        print(f"  - 文件大小: {file_size / 1024:.2f} KB")
        
        return True
    except Exception as e:
        print(f"✗ PPT生成失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_mixed_content():
    """测试混合内容渲染"""
    print("=" * 60)
    print("测试4: 混合内容渲染")
    print("=" * 60)

    mixed_content = """
# 混合内容测试

## 文本和图表

这是一段测试文本，用于验证混合内容渲染。

:::chart{type="bar" title="数据可视化"}
| 类别 | A | B | C |
|------|---|---|---|
| 数值 | 10 | 20 | 15 |
:::

## 多个图表

:::chart{type="pie" title="饼图示例"}
- 选项A: 40
- 选项B: 30
- 选项C: 30
:::

:::chart{type="line" title="折线图示例"}
| 时间 | 1月 | 2月 | 3月 |
|------|-----|-----|-----|
| 数据 | 100 | 150 | 120 |
:::
"""

    print("\n解析混合内容...")
    result = parse_markdown(mixed_content)
    
    print(f"✓ 解析完成: {len(result.slides)} 页")
    
    for i, slide in enumerate(result.slides, 1):
        chart_count = len(slide.charts)
        print(f"  - 幻灯片 {i}: {chart_count} 个图表")
    
    # 统计总图表数
    total_charts = sum(len(slide.charts) for slide in result.slides)
    print(f"\n✓ 总共解析出 {total_charts} 个图表")
    
    return total_charts > 0


def main():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("图表渲染功能测试")
    print("=" * 60 + "\n")

    tests = [
        ("图表解析测试", test_chart_parsing),
        ("Mermaid检测测试", test_mermaid_detection),
        ("PPT生成测试", test_ppt_generation),
        ("混合内容测试", test_mixed_content),
    ]

    passed = 0
    failed = 0

    for name, test_func in tests:
        try:
            result = test_func()
            if result is None or result:
                passed += 1
            else:
                failed += 1
                print(f"\n✗ {name} 失败")
        except Exception as e:
            failed += 1
            print(f"\n✗ {name} 异常: {e}")
            import traceback
            traceback.print_exc()

    print("\n" + "=" * 60)
    print(f"测试总结: {passed} 通过, {failed} 失败")
    print("=" * 60)

    if failed == 0:
        print("\n✓ 所有测试通过!")
        return 0
    else:
        print(f"\n✗ {failed} 个测试失败")
        return 1


if __name__ == '__main__':
    sys.exit(main())
