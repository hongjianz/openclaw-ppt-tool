#!/usr/bin/env python3
"""
QuintenStyle背景图片生成器
将CSS渐变转换为PPT可用的背景图片
"""

import sys
import os
from PIL import Image, ImageDraw


def generate_quinten_background(width=1920, height=1080, output_path="templates/quinten_background.png"):
    """
    生成QuintenStyle背景图片

    Args:
        width: 图片宽度(像素)
        height: 图片高度(像素)
        output_path: 输出文件路径
    """
    # 创建图片
    img = Image.new('RGB', (width, height))
    draw = ImageDraw.Draw(img)

    # 基础渐变色: #267878 -> #B0C5B4
    # 简化版:使用线性渐变
    for y in range(height):
        # 计算渐变比例
        ratio = y / height

        # 插值颜色
        r = int(38 + (176 - 38) * ratio)  # 0x26 -> 0xB0
        g = int(120 + (197 - 120) * ratio)  # 0x78 -> 0xC5
        b = int(120 + (180 - 120) * ratio)  # 0x78 -> 0xB4

        draw.line([(0, y), (width, y)], fill=(r, g, b))

    # 添加装饰性圆形渐变效果(简化版)
    # 右上角光晕
    center_x = int(width * 1.2)
    center_y = int(height * 0.2)
    radius = int(width * 0.4)

    for r in range(radius, 0, -2):
        alpha_ratio = 1 - (r / radius)
        if alpha_ratio > 0.6:
            continue

        color_intensity = int(45 * alpha_ratio * 0.25)  # rgba(45, 134, 136, 0.25)
        color = (30 + color_intensity, 95 + color_intensity * 2, 95 + color_intensity * 2)

        # 绘制圆弧
        x1 = center_x - r
        y1 = center_y - r
        x2 = center_x + r
        y2 = center_y + r

        if x1 < width and y1 < height:
            draw.ellipse([x1, y1, x2, y2], outline=color)

    # 右下角白色光晕
    center_x2 = int(width * 1.25)
    center_y2 = int(height * 0.85)
    radius2 = int(width * 0.35)

    for r in range(radius2, 0, -2):
        alpha_ratio = 1 - (r / radius2)
        if alpha_ratio > 0.55:
            continue

        intensity = int(255 * alpha_ratio * 0.35)
        color = (intensity, intensity, intensity)

        x1 = center_x2 - r
        y1 = center_y2 - r
        x2 = center_x2 + r
        y2 = center_y2 + r

        if x1 < width and y1 < height:
            draw.ellipse([x1, y1, x2, y2], outline=color)

    # 保存
    os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else '.', exist_ok=True)
    img.save(output_path, 'PNG')
    print(f"QuintenStyle背景图片已生成: {output_path}")
    print(f"尺寸: {width}x{height}")

    return output_path


if __name__ == '__main__':
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("错误: 需要安装Pillow库")
        print("运行: pip install Pillow")
        sys.exit(1)

    # 默认参数
    w = 1920
    h = 1080
    out = "templates/quinten_background.png"

    # 支持命令行参数
    if len(sys.argv) > 1:
        out = sys.argv[1]
    if len(sys.argv) > 3:
        w = int(sys.argv[2])
        h = int(sys.argv[3])

    generate_quinten_background(w, h, out)
