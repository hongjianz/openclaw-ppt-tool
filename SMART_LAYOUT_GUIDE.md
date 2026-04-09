# 智能布局功能文档

## 概述

智能布局功能解决了传统PPT生成中的内容重叠和布局混乱问题。通过**两阶段渲染**方法，确保生成的PPT页面布局合理、内容不重叠。

## 问题背景

### 传统方法的问题

```
解析 → 渲染 → ❌ 内容重叠
```

在传统单阶段渲染中：
- 内容按顺序直接渲染到页面
- 不预先计算内容高度
- 导致图表、表格、文本互相重叠
- 页面底部内容被截断

### 智能布局的解决方案

```
解析 → 预扫描计算 → 智能分页 → 渲染 → ✅ 完美布局
```

## 使用方法

### 命令行参数

```bash
# 使用智能布局模式生成PPT
python main.py -i test_smart_layout.md -o output/smart_layout_demo.pptx --smart-layout

# 结合其他功能使用
python main.py -i test_smart_layout.md -o output/demo.pptx --smart-layout --toc --footer "公司内部资料"
```

### 参数说明

- `--smart-layout`: 启用智能布局模式
  - 自动预扫描所有内容块高度
  - 智能分配到多个页面
  - 避免内容重叠

## 技术实现

### 三阶段流程

#### 阶段 1: 内容预扫描

系统会扫描每个内容块并计算高度：

- **标题**: 固定 0.8 英寸
- **副标题**: 固定 0.5 英寸
- **图片**: 根据宽高比计算
- **表格**: 行数 × 0.5 英寸
- **项目列表**: 项数 × 行高
- **文本**: 根据字符数和宽度计算
- **代码块**: 行数 × 行高
- **Mermaid图表**: 预估图片高度
- **原生图表**: 配置的默认高度

#### 阶段 2: 智能分页

根据内容高度智能分配到页面：

```python
# 示例逻辑
当前页高度 = 0
for 内容块 in 所有内容块:
    if 当前页高度 + 内容块高度 > 页面最大高度:
        创建新页面()
        当前页高度 = 0
    
    内容块.页码 = 当前页码
    当前页高度 += 内容块高度
```

#### 阶段 3: 渲染输出

按分页结果逐个渲染页面：
- 根据内容块类型选择渲染方法
- 使用预计算的高度定位
- 确保不重叠

### 核心模块

```
src/smart_layout.py
├── ContentBlock           # 内容块数据结构
├── scan_slide_content()   # 预扫描单页内容
├── paginate_content_blocks()  # 智能分页算法
└── create_layout_plan()   # 创建完整布局计划
```

## 支持的内容类型

| 类型 | 描述 | 高度计算方式 |
|------|------|-------------|
| title | 标题 | 固定 0.8" |
| subtitle | 副标题 | 固定 0.5" |
| text | 普通文本 | 字符数估算 |
| bullet | 项目列表 | 项数 × 行高 |
| image | 图片 | 图片宽高比 |
| table | 表格 | 行数 × 0.5" |
| code | 代码块 | 行数 × 行高 |
| mermaid | Mermaid图表 | 预估图片高度 |
| chart | 原生图表 | 配置默认高度 |

## 使用场景

### 场景 1: 大量图表混合

当幻灯片包含多个图表时：

```markdown
# 数据分析报告

:::chart{type="bar" title="销售趋势"}
| ... |
:::

:::chart{type="pie" title="市场份额"}
- ...
:::

详细说明文字...
```

**传统模式**: 图表可能重叠  
**智能布局**: 自动分到多页

### 场景 2: Mermaid图表 + 文本

```markdown
# 系统架构

架构说明文字...

```mermaid
graph TD
    A --> B
```

更多说明文字...
```

**传统模式**: Mermaid转换后高度不确定  
**智能布局**: 预估高度并合理分页

### 场景 3: 长文档自动生成

```bash
python main.py -i long_report.md -o output.pptx --smart-layout
```

**效果**: 长文档自动智能分页，每页布局合理

## 对比测试

### 测试文件

使用 `test_smart_layout.md` 进行测试，包含：
- 5个原始幻灯片
- 4个原生图表
- 3个Mermaid图表
- 2个表格
- 大量文本

### 测试结果

```bash
# 传统模式
python main.py -i test_smart_layout.md -o output/traditional.pptx
# 结果: 部分内容重叠，图表位置混乱

# 智能布局模式
python main.py -i test_smart_layout.md -o output/smart.pptx --smart-layout
# 结果: 无重叠，布局清晰，自动分页合理
```

## 配置选项

在 `src/template_config.py` 中可调整：

```python
@dataclass
class TemplateConfig:
    # 图表默认高度
    default_chart_width: float = 8.0   # 英寸
    default_chart_height: float = 4.0  # 英寸
    
    # 页面尺寸
    slide_width: float = 13.333   # 16:9
    slide_height: float = 7.5
    
    # 边距
    margin_top: float = 0.8
    margin_bottom: float = 0.8
```

## 限制和注意事项

1. **预估高度**: Mermaid图表高度为预估值，实际可能不同
2. **分页策略**: 当前按高度顺序分页，未来会支持语义单元分页
3. **性能**: 预扫描阶段会额外消耗少量时间
4. **兼容性**: 完全兼容传统模式，可通过参数切换

## 未来改进

- [ ] 支持语义单元分页（不按高度硬性分割）
- [ ] Mermaid图表实际高度预计算
- [ ] 内容块优先级设置（重要内容不分页）
- [ ] 可视化布局预览
- [ ] 智能留白和美观优化

## 故障排除

### Q: 智能布局和传统模式有什么区别？

A: 传统模式直接渲染，可能重叠；智能布局预先计算并分页，保证不重叠。

### Q: 何时使用智能布局？

A: 当内容包含多个图表、表格或大段文本时建议使用。

### Q: 智能布局会改变我的内容吗？

A: 不会。只调整分页和位置，内容完全保留。

### Q: 可以控制分页位置吗？

A: 当前自动计算，未来版本会支持手动控制分页点。

## 示例命令

```bash
# 基本使用
python main.py -i content.md -o output.pptx --smart-layout

# 添加目录
python main.py -i content.md -o output.pptx --smart-layout --toc

# 添加页脚
python main.py -i content.md -o output.pptx --smart-layout --footer "机密文档"

# 完整示例
python main.py -i test_smart_layout.md \
    -o output/demo.pptx \
    --smart-layout \
    --toc \
    --footer "内部资料" \
    --width 13.333 \
    --height 7.5
```

## 相关文档

- [图表渲染文档](./README.md#图表支持)
- [Mermaid使用指南](./README.md#mermaid安装指南)
- [配置说明](./README.md#配置说明)
