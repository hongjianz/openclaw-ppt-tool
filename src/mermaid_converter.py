"""
Mermaid图表转换模块 - 将Mermaid语法转换为图片

依赖: Node.js + @mermaid-js/mermaid-cli
安装: npm install -g @mermaid-js/mermaid-cli
"""

import subprocess
import os
import tempfile
from typing import Optional


def convert_mermaid_to_image(mmd_content: str, output_format: str = "png",
                             width: int = 1920, height: int = 1080) -> Optional[str]:
    """
    将Mermaid图表代码转换为图片

    Args:
        mmd_content: Mermaid语法内容
        output_format: 输出格式 ('png' 或 'svg')
        width: 图片宽度（像素）
        height: 图片高度（像素）

    Returns:
        生成的图片路径，失败返回None
    """
    try:
        # 检查mmdc是否可用
        try:
            subprocess.run(['mmdc', '--version'], capture_output=True, check=True)
        except FileNotFoundError:
            print("警告: Mermaid CLI未安装。请运行: npm install -g @mermaid-js/mermaid-cli")
            return None

        # 创建临时文件
        with tempfile.NamedTemporaryFile(mode='w', suffix='.mmd', delete=False, encoding='utf-8') as f:
            f.write(mmd_content)
            mmd_path = f.name

        output_path = mmd_path.replace('.mmd', f'.{output_format}')

        # 构建命令
        cmd = [
            'mmdc',
            '-i', mmd_path,
            '-o', output_path,
            '-w', str(width),
            '-H', str(height),
            '-b', 'transparent',  # 透明背景
            '-t', 'default'       # 主题
        ]

        # 执行转换
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

        # 清理临时文件
        os.unlink(mmd_path)

        if result.returncode == 0 and os.path.exists(output_path):
            return output_path
        else:
            print(f"Mermaid转换失败: {result.stderr}")
            return None

    except subprocess.TimeoutExpired:
        print("Mermaid转换超时")
        return None
    except Exception as e:
        print(f"Mermaid转换错误: {e}")
        return None
