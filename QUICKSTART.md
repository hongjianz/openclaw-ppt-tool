# 快速开始指南

## 在阿里云服务器上部署

### 1. 上传文件到服务器

```bash
# 使用scp上传
scp -r openclaw-ppt-tool root@your-server-ip:/opt/

# 或使用git
git clone <repository-url> /opt/openclaw-ppt-tool
```

### 2. 安装Python和依赖

```bash
# 连接到服务器
ssh root@your-server-ip

# 安装Python 3 (如果未安装)
yum install python3 python3-pip -y    # CentOS/RHEL
# 或
apt-get install python3 python3-pip -y  # Ubuntu/Debian

# 进入项目目录
cd /opt/openclaw-ppt-tool

# 安装Python依赖
pip3 install -r requirements.txt
```

### 3. 测试工具

```bash
# 运行测试脚本
python3 test_ppt_tool.py

# 使用示例文稿生成PPT
python3 main.py -i examples/sample_content.md -o output/test.pptx
```

### 4. 集成到OpenClaw

```bash
# 创建OpenClaw Skill目录
mkdir -p ~/.openclaw/skills/ppt-generator

# 复制Skill配置
cp .openclaw-skill/skill.yml ~/.openclaw/skills/ppt-generator/
cp generate_ppt.sh ~/.openclaw/skills/ppt-generator/

# 修改skill.yml中的路径指向实际项目目录
# 编辑 ~/.openclaw/skills/ppt-generator/skill.yml
# 将 main 路径改为: /opt/openclaw-ppt-tool/generate_ppt.sh
```

## 本地开发环境(Windows)

### 1. 安装Python

从 https://www.python.org/downloads/ 下载并安装Python 3.8+

安装时勾选 "Add Python to PATH"

### 2. 安装依赖

```bash
cd openclaw-ppt-tool
pip install -r requirements.txt
```

### 3. 运行测试

```bash
python test_ppt_tool.py
```

### 4. 生成PPT

```bash
python main.py -i examples/sample_content.md -o output/my_presentation.pptx
```

## 使用示例

### 示例1: 基本用法

```bash
python3 main.py -i my_content.md -o presentation.pptx
```

### 示例2: 添加背景图片

```bash
# 准备背景图片(建议尺寸: 1920x1080)
python3 main.py -i content.md -o output.pptx -b templates/background.png
```

### 示例3: 使用自定义配置

```bash
python3 main.py -i content.md -o output.pptx -c examples/template_config.json
```

### 示例4: 使用CSS样式

```bash
python3 main.py -i content.md -o output.pptx --css examples/styles.css
```

### 示例5: 自定义页面尺寸

```bash
# 4:3 比例
python3 main.py -i content.md -o output.pptx --width 10 --height 7.5

# 16:10 比例
python3 main.py -i content.md -o output.pptx --width 13.333 --height 8.333
```

## 文稿编写指南

### Markdown格式

```markdown
# 演示文稿标题

## 副标题或演讲者信息

---

# 第一页: 介绍

- 要点一
- 要点二
- 要点三

---

# 第二页: 详细内容

这里是正文段落,可以包含多行文字。

系统会自动进行分页和排版。

- 也可以混合使用列表
- 和普通文本

---

# 第三页: 总结

感谢聆听!
```

### 纯文本格式

直接编写文本内容,系统会按字符数自动分页:

```
这是我的演示文稿标题

这是第一段内容,介绍主题的背景和重要性。

这是第二段内容,详细阐述主要观点。

(继续编写更多内容...)
```

## 故障排除

### 问题1: 找不到python3命令

```bash
# 检查Python安装
which python3
python3 --version

# 如果未安装,安装Python
yum install python3 -y    # CentOS
apt-get install python3 -y  # Ubuntu
```

### 问题2: pip安装失败

```bash
# 尝试使用国内镜像源
pip3 install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 问题3: 中文字体显示问题

```bash
# CentOS/RHEL 安装中文字体
yum install wqy-microhei-fonts -y

# Ubuntu/Debian 安装中文字体
apt-get install fonts-wqy-microhei -y
```

### 问题4: 权限错误

```bash
# 给脚本添加执行权限
chmod +x generate_ppt.sh

# 确保输出目录存在且有写权限
mkdir -p output
chmod 755 output
```

## 下一步

- 查看 `examples/` 目录中的示例文件
- 阅读 `README.md` 了解详细配置选项
- 根据需要修改 `template_config.json` 自定义样式
