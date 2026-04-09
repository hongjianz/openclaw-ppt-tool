# PPT自动生成工具 - OpenClaw Skill

一个基于Python的智能PPT生成工具,可作为OpenClaw Skill使用。根据文稿内容(Markdown或纯文本)自动生成美观的演示文稿,支持自定义背景模板和样式。

## ✨ 核心功能

- 📝 **智能文稿解析**: 支持Markdown和纯文本格式,自动识别和分页
- 🎨 **模板系统**: 支持背景图片、CSS样式和JSON配置
- 🏢 **QuintenStyle**: 内置专业公司模板(青绿色渐变背景)
- 📐 **自动排版**: 智能调整字体、边距和行间距
- 🔧 **灵活配置**: 页面尺寸、颜色、字体完全可定制
- 🤖 **OpenClaw集成**: 可直接作为Skill在OpenClaw中调用

## 🚀 快速开始

### 安装

```bash
# 克隆仓库
git clone https://github.com/YOUR_USERNAME/openclaw-ppt-tool.git
cd openclaw-ppt-tool

# 安装依赖
pip install -r requirements.txt
```

### 基本使用

```bash
# 使用示例文稿生成PPT
python main.py -i examples/sample_content.md -o output/presentation.pptx

# 使用QuintenStyle公司模板
python main.py -i content.md -o output.pptx --css examples/quinten_style.css
```

### OpenClaw中使用

```bash
# 在OpenClaw中调用
@ppt-generator 帮我根据以下文稿生成PPT:

# 这里是你的文稿内容...
```

## 📖 文档

- [快速开始指南](QUICKSTART.md) - 详细安装和使用说明
- [配置说明](README.md) - 完整的配置选项
- [示例文稿](examples/sample_content.md) - Markdown格式示例

## 🎯 使用场景

1. **快速制作汇报PPT**: 编写Markdown文稿,一键生成演示文稿
2. **统一公司风格**: 使用QuintenStyle或其他自定义模板
3. **批量生成PPT**: 自动化处理多个文稿文件
4. **OpenClaw工作流**: 在对话中直接生成PPT

## 🛠️ 技术栈

- Python 3.x
- python-pptx - PPT生成库
- Pillow - 图像处理
- OpenClaw - AI Agent框架

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交Issue和Pull Request!

---

**提示**: 将 `YOUR_USERNAME` 替换为你的GitHub用户名后再推送仓库。
