# GitHub部署指南

## 推送到GitHub仓库

### 1. 创建GitHub仓库

访问 https://github.com/new 创建新仓库:
- 仓库名: `openclaw-ppt-tool`
- 描述: `Auto PPT Generator for OpenClaw - Generate beautiful presentations from Markdown or text`
- 设为公开(Public)或私有(Private)
- **不要**初始化README(我们已有)

### 2. 本地Git初始化

```bash
cd openclaw-ppt-tool

# 初始化Git仓库
git init

# 添加所有文件
git add .

# 创建首次提交
git commit -m "Initial commit: PPT generator with QuintenStyle support

Features:
- Auto PPT generation from Markdown/text
- QuintenStyle company template
- CSS and JSON configuration
- OpenClaw Skill integration
- Smart pagination and layout"
```

### 3. 关联远程仓库

```bash
# 替换YOUR_USERNAME为你的GitHub用户名
git remote add origin https://github.com/YOUR_USERNAME/openclaw-ppt-tool.git

# 推送到GitHub
git branch -M main
git push -u origin main
```

### 4. 验证推送

访问你的仓库页面确认文件已上传:
`https://github.com/YOUR_USERNAME/openclaw-ppt-tool`

## OpenClaw中使用

### 方法1: 直接克隆使用

```bash
# 在服务器上
cd /opt
git clone https://github.com/YOUR_USERNAME/openclaw-ppt-tool.git
cd openclaw-ppt-tool
pip install -r requirements.txt

# 测试
python main.py -i examples/sample_content.md -o output/test.pptx
```

### 方法2: 安装为OpenClaw Skill

```bash
# 运行安装脚本
./install_skill.sh ~/.openclaw/skills

# 或在OpenClaw中直接使用
@ppt-generator --input content.md --output presentation.pptx
```

### 方法3: 从GitHub直接安装

```bash
# 克隆到OpenClaw skills目录
cd ~/.openclaw/skills
git clone https://github.com/YOUR_USERNAME/openclaw-ppt-tool.git ppt-generator
cd ppt-generator
pip install -r requirements.txt
```

## 更新维护

### 发布新版本

```bash
# 修改版本号 (在.openclaw-skill/skill.yml中)
# version: 1.0.1

git add .
git commit -m "Update to v1.0.1: [changes description]"
git tag v1.0.1
git push origin main --tags
```

### 创建Release

1. 访问: `https://github.com/YOUR_USERNAME/openclaw-ppt-tool/releases/new`
2. 选择标签: `v1.0.1`
3. 填写发布说明
4. 点击 "Publish release"

## CI/CD自动化

项目已包含GitHub Actions配置,会自动:
- 运行测试
- 验证PPT生成功能
- 检查依赖

查看测试结果: `https://github.com/YOUR_USERNAME/openclaw-ppt-tool/actions`

## 常见问题

### Q: 如何修改仓库可见性?
A: 在GitHub仓库 Settings -> General -> Change visibility

### Q: 如何添加协作者?
A: Settings -> Collaborators -> Add people

### Q: 如何设置仓库主题?
A: About区域 -> Gear图标 -> Topics添加: `ppt`, `openclaw`, `python`, `presentation`

### Q: 如何启用Issues?
A: Settings -> General -> Features -> 勾选"Issues"

## 推广建议

1. **添加Topics**: ppt, openclaw, python, presentation, generator
2. **完善README**: 添加截图和演示GIF
3. **编写博客**: 介绍工具使用方法
4. **分享到社区**: OpenClaw Discord, Reddit等
5. **示例展示**: 在examples中添加更多样例输出

## 许可证提醒

本项目使用MIT许可证,确保:
- 保留LICENSE文件
- 使用者需注明原作者
- 允许商业使用和修改
