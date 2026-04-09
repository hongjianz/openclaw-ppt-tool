#!/bin/bash
# OpenClaw Skill - PPT生成工具
# 用法: 在OpenClaw中调用此脚本来生成PPT

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 检查Python是否可用
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "错误: 未找到Python,请先安装Python 3.x"
    exit 1
fi

# 检查依赖是否已安装
if ! $PYTHON_CMD -c "import pptx" 2>/dev/null; then
    echo "正在安装依赖..."
    $PYTHON_CMD -m pip install -r requirements.txt
fi

# 执行PPT生成
$PYTHON_CMD main.py "$@"
