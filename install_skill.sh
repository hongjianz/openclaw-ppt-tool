#!/bin/bash
# OpenClaw Skill安装脚本
# 用法: ./install_skill.sh [openclaw_skills_directory]

set -e

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 默认OpenClaw skills目录
SKILLS_DIR="${1:-$HOME/.openclaw/skills}"

echo -e "${GREEN}================================${NC}"
echo -e "${GREEN}OpenClaw PPT Generator Skill 安装程序${NC}"
echo -e "${GREEN}================================${NC}\n"

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}错误: 未找到Python3,请先安装Python 3.x${NC}"
    exit 1
fi

echo -e "${YELLOW}✓${NC} Python3已安装: $(python3 --version)"

# 创建skills目录
mkdir -p "$SKILLS_DIR/ppt-generator"

# 复制文件
echo -e "\n${YELLOW}正在安装Skill...${NC}"
cp -r "$SCRIPT_DIR"/* "$SKILLS_DIR/ppt-generator/"

# 设置执行权限
chmod +x "$SKILLS_DIR/ppt-generator/generate_ppt.sh"
chmod +x "$SKILLS_DIR/ppt-generator/generate_quinten_background.py"

# 安装Python依赖
echo -e "\n${YELLOW}正在安装Python依赖...${NC}"
cd "$SKILLS_DIR/ppt-generator"
pip3 install -r requirements.txt

echo -e "\n${GREEN}================================${NC}"
echo -e "${GREEN}安装完成!${NC}"
echo -e "${GREEN}================================${NC}\n"

echo -e "Skill已安装到: ${YELLOW}$SKILLS_DIR/ppt-generator${NC}\n"

echo -e "${YELLOW}使用方法:${NC}"
echo -e "1. 在OpenClaw中调用:"
echo -e "   ${GREEN}@ppt-generator --input content.md --output presentation.pptx${NC}\n"

echo -e "2. 使用QuintenStyle模板:"
echo -e "   ${GREEN}@ppt-generator --input content.md --output out.pptx --css examples/quinten_style.css${NC}\n"

echo -e "3. 查看帮助:"
echo -e "   ${GREEN}@ppt-generator --help${NC}\n"

echo -e "${YELLOW}提示:${NC}"
echo -e "- 首次使用QuintenStyle时会自动生成背景图片"
echo -e "- 示例文稿位于: examples/sample_content.md"
echo -e "- 更多文档请查看: README.md 和 QUICKSTART.md\n"
