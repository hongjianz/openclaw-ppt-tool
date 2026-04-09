"""
错误处理模块 - 提供友好的错误提示和降级处理

功能:
- 文件不存在时友好提示
- 字体缺失时自动降级
- 生成失败时保留部分成果
- 统一的错误日志记录
"""

import os
import sys
import logging
from typing import Optional, List
from functools import wraps

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s: %(message)s'
)
logger = logging.getLogger('ppt_generator')


class PPTGenerationError(Exception):
    """PPT生成错误"""
    def __init__(self, message: str, partial_result=None):
        super().__init__(message)
        self.message = message
        self.partial_result = partial_result


def handle_file_errors(func):
    """文件操作错误处理装饰器"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            logger.error(f"文件未找到: {e.filename}")
            logger.info("提示: 请检查文件路径是否正确")
            sys.exit(1)
        except PermissionError as e:
            logger.error(f"权限错误: 无法访问 {e.filename}")
            logger.info("提示: 请检查文件权限")
            sys.exit(1)
        except Exception as e:
            logger.error(f"文件操作失败: {e}")
            sys.exit(1)
    return wrapper


def validate_input_file(filepath: str) -> bool:
    """
    验证输入文件

    Args:
        filepath: 文件路径

    Returns:
        是否有效
    """
    if not filepath:
        logger.error("错误: 未指定输入文件")
        return False

    if not os.path.exists(filepath):
        logger.error(f"错误: 文件不存在: {filepath}")
        logger.info("提示: 请检查文件路径是否正确")
        return False

    if not os.path.isfile(filepath):
        logger.error(f"错误: 不是有效文件: {filepath}")
        return False

    # 检查文件大小
    file_size = os.path.getsize(filepath)
    if file_size == 0:
        logger.warning(f"警告: 文件为空: {filepath}")
        return False

    if file_size > 10 * 1024 * 1024:  # 10MB
        logger.warning(f"警告: 文件较大 ({file_size / 1024 / 1024:.1f}MB), 处理可能需要较长时间")

    return True


def validate_output_path(filepath: str) -> bool:
    """
    验证输出路径

    Args:
        filepath: 输出文件路径

    Returns:
        是否有效
    """
    if not filepath:
        logger.error("错误: 未指定输出文件")
        return False

    # 检查扩展名
    if not filepath.lower().endswith('.pptx'):
        logger.warning(f"警告: 输出文件建议以 .pptx 结尾")

    # 检查目录是否存在
    output_dir = os.path.dirname(filepath)
    if output_dir and not os.path.exists(output_dir):
        logger.info(f"创建输出目录: {output_dir}")
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            logger.error(f"错误: 无法创建目录: {e}")
            return False

    # 检查是否可写
    if os.path.exists(filepath):
        if not os.access(filepath, os.W_OK):
            logger.error(f"错误: 无法写入文件, 权限不足: {filepath}")
            return False

    return True


def get_fallback_font(preferred_font: str) -> str:
    """
    获取降级字体

    Args:
        preferred_font: 首选字体

    Returns:
        可用字体
    """
    # 常见中文字体列表
    chinese_fonts = [
        "Microsoft YaHei",  # 微软雅黑 (Windows)
        "SimHei",           # 黑体 (Windows)
        "SimSun",           # 宋体 (Windows)
        "PingFang SC",      # 苹方 (macOS)
        "STHeiti",          # 华文黑体 (macOS)
        "WenQuanYi Micro Hei",  # 文泉驿微米黑 (Linux)
        "Droid Sans Fallback",   # Android
        "Arial Unicode MS",      # 通用
        "Arial",                  # 英文备选
    ]

    # 如果首选字体在列表中,返回它
    if preferred_font in chinese_fonts:
        return preferred_font

    # 否则返回第一个可用字体
    return chinese_fonts[0]


def safe_generate_ppt(generator, content, output_path: str) -> bool:
    """
    安全地生成PPT,失败时保留部分成果

    Args:
        generator: PPT生成器实例
        content: 演示文稿内容
        output_path: 输出路径

    Returns:
        是否成功
    """
    try:
        generator.generate(content, output_path)
        logger.info(f"成功生成PPT: {output_path}")
        return True

    except Exception as e:
        logger.error(f"PPT生成失败: {e}")
        
        # 打印详细堆栈跟踪用于调试
        import traceback
        logger.debug(traceback.format_exc())

        # 尝试保存部分成果
        partial_path = output_path.replace('.pptx', '_partial.pptx')
        try:
            if hasattr(generator, 'prs') and len(generator.prs.slides) > 0:
                generator.prs.save(partial_path)
                logger.info(f"已保存部分成果: {partial_path}")
                logger.info(f"共生成 {len(generator.prs.slides)} 页")
        except Exception as save_error:
            logger.error(f"无法保存部分成果: {save_error}")

        return False


def check_dependencies() -> List[str]:
    """
    检查依赖项

    Returns:
        缺失的依赖列表
    """
    missing = []

    # 检查python-pptx
    try:
        import pptx
    except ImportError:
        missing.append("python-pptx")

    # 检查Pillow
    try:
        from PIL import Image
    except ImportError:
        missing.append("Pillow")

    return missing


def print_dependency_help(missing: List[str]):
    """打印依赖安装帮助"""
    if not missing:
        return

    logger.info("\n" + "="*60)
    logger.info("缺少以下依赖:")
    for dep in missing:
        logger.info(f"  - {dep}")

    logger.info("\n安装命令:")
    logger.info(f"  pip install {' '.join(missing)}")
    logger.info("="*60 + "\n")


def graceful_exit(message: str, exit_code: int = 1):
    """
    优雅退出

    Args:
        message: 退出消息
        exit_code: 退出码
    """
    if exit_code == 0:
        logger.info(message)
    else:
        logger.error(message)
    sys.exit(exit_code)
