# 推送到GitHub - 操作指南

## 当前状态
✅ Git仓库已初始化
✅ 所有文件已提交(24个文件)
✅ 首次commit已完成

## 下一步操作

### 1. 创建GitHub仓库

访问 https://github.com/new

填写信息:
- **Repository name**: `openclaw-ppt-tool`
- **Description**: `Auto PPT Generator for OpenClaw - Generate beautiful presentations from Markdown or text with QuintenStyle support`
- **Visibility**: Public (推荐,便于分享) 或 Private
- **不要勾选** "Initialize this repository with a README"

点击 "Create repository"

### 2. 关联远程仓库并推送

在终端执行以下命令(替换YOUR_USERNAME为你的GitHub用户名):

```bash
# 添加远程仓库
git remote add origin https://github.com/YOUR_USERNAME/openclaw-ppt-tool.git

# 重命名分支为main
git branch -M main

# 推送到GitHub
git push -u origin main
```

### 3. 验证推送

访问你的仓库页面:
```
https://github.com/YOUR_USERNAME/openclaw-ppt-tool
```

确认所有文件已上传。

### 4. 完善仓库信息

#### 添加Topics(标签)
在仓库首页 "About" 区域点击齿轮图标,添加:
- `ppt`
- `openclaw`
- `python`
- `presentation`
- `generator`
- `markdown`

#### 设置网站(可选)
如果有演示页面,可在 Settings -> Pages 中启用GitHub Pages

#### 启用Issues
Settings -> General -> Features -> 勾选 "Issues"

### 5. 测试CI/CD

推送后,GitHub Actions会自动运行测试。

查看: `https://github.com/YOUR_USERNAME/openclaw-ppt-tool/actions`

## OpenClaw中使用

### 方法1: 直接克隆

```bash
# 在阿里云服务器上
cd /opt
git clone https://github.com/YOUR_USERNAME/openclaw-ppt-tool.git
cd openclaw-ppt-tool
pip3 install -r requirements.txt

# 测试
python3 main.py -i examples/sample_content.md -o output/test.pptx
```

### 方法2: 安装为Skill

```bash
# 运行安装脚本
cd /opt/openclaw-ppt-tool
./install_skill.sh ~/.openclaw/skills
```

### 方法3: 在OpenClaw中调用

```
@ppt-generator --input content.md --output presentation.pptx --css examples/quinten_style.css
```

## 常见问题

### Q: 推送时要求登录?
A: 使用GitHub账号密码或Personal Access Token(PAT)

创建PAT: https://github.com/settings/tokens

### Q: 如何配置SSH密钥?
```bash
# 生成SSH密钥
ssh-keygen -t ed25519 -C "your_email@example.com"

# 添加到GitHub
# 复制公钥: cat ~/.ssh/id_ed25519.pub
# 粘贴到: https://github.com/settings/keys

# 使用SSH地址
git remote set-url origin git@github.com:YOUR_USERNAME/openclaw-ppt-tool.git
```

### Q: 推送失败,提示权限?
```bash
# 检查远程仓库地址
git remote -v

# 重新设置(使用HTTPS)
git remote set-url origin https://github.com/YOUR_USERNAME/openclaw-ppt-tool.git

# 或使用SSH
git remote set-url origin git@github.com:YOUR_USERNAME/openclaw-ppt-tool.git
```

## 后续更新

```bash
# 修改代码后
git add .
git commit -m "描述修改内容"
git push

# 发布新版本
git tag v1.0.1
git push origin --tags
```

## 完整命令示例

```bash
# 假设你的GitHub用户名是 example-user
git remote add origin https://github.com/example-user/openclaw-ppt-tool.git
git branch -M main
git push -u origin main
```

推送成功后,你就可以在GitHub上分享这个工具,并在OpenClaw中使用了!
