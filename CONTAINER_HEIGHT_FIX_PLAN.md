# 容器高度计算修复计划

## 当前问题总结

### 根本原因
1. **硬编码依赖**: 所有容器高度使用固定值估算（220px, 240px, 80px等）
2. **忽略CSS约束**: 未读取min-height/max-height/padding/gap等CSS规则
3. **内容动态性缺失**: 未根据实际文本行数、元素数量计算高度
4. **高度传递错误**: 部分方法不返回实际使用的Y坐标

### 影响
- 容器间距不一致
- 布局与HTML视觉效果偏差
- 多行文本时高度计算错误

---

## CSS约束参考

从slide01.html提取的实际CSS规则：

```css
.stat-card {
    padding: 20px;
    min-height: 200px;
    max-height: 300px;
}

.stat-box {
    padding: 20px;
    min-height: 200px;
    max-height: 300px;
    display: flex;
}

.stats-container {
    display: grid;
    grid-template-columns: repeat(2, 1fr);  /* 注意：slide01.html用inline style覆盖为repeat(3, 1fr) */
    gap: 20px;  /* 行列间距都是20px，不是30px！ */
    margin-bottom: 30px;
}

.data-card {
    padding-left: 15px;
    min-height: 200px;
    max-height: 300px;
}

.strategy-card {
    padding: 10px;
    min-height: 200px;
    max-height: 300px;
    margin-bottom: 5px;
}

.action-item {
    display: flex;
    margin-bottom: 15px;  /* Tailwind mb-15 = 60px? 需确认 */
}
```

---

## 修复计划（按优先级）

### 阶段1：立即修复（高优先级）

#### 1.1 修正 `_convert_stats_container` 返回值

**文件**: `src/main.py` (line ~305)

**当前代码**:
```python
# 计算下一个元素的Y坐标
num_rows = (num_boxes + num_columns - 1) // num_columns
return y_start + num_rows * (box_height + gap) + 30  # ❌ box_height=220硬编码，+30无依据
```

**修改为**:
```python
# 计算下一个元素的Y坐标
num_rows = (num_boxes + num_columns - 1) // num_columns
gap = 20  # 从CSS读取：stats-container的gap

# 注意：这里计算的是所有stat-box渲染完毕后的Y坐标
# 每一行占用：box_height + gap（除了最后一行没有gap）
# 正确公式：y_start + num_rows * box_height + (num_rows - 1) * gap
actual_height = num_rows * box_height + (num_rows - 1) * gap

return y_start + actual_height
```

**改进点**:
- 移除神秘的`+30`
- 使用正确的行间距计算公式（最后一行不需要gap）
- 添加注释说明计算逻辑

---

#### 1.2 修正 stat-card 高度估算（包含stats-container分支）

**文件**: `src/main.py` (line ~323)

**当前代码**:
```python
num_rows = (num_boxes + num_columns - 1) // num_columns
card_height = num_rows * 240 + 60  # ❌ 完全估算
```

**修改为**:
```python
# 从CSS读取约束
stat_card_padding_top = 20
stat_card_padding_bottom = 20
stats_container_gap = 20
stat_box_height = 220  # TODO阶段2: 改为动态计算

# 计算stats-container的实际高度
num_rows = (num_boxes + num_columns - 1) // num_columns
stats_container_height = num_rows * stat_box_height + (num_rows - 1) * stats_container_gap

# 计算stat-card总高度（包括自身padding）
# stat-card = padding-top + (可选标题35px) + stats-container + padding-bottom
has_title = card.find('p', class_='primary-color') is not None
title_height = 35 if has_title else 0

card_height = stat_card_padding_top + title_height + stats_container_height + stat_card_padding_bottom
```

**改进点**:
- 基于CSS padding值计算
- 分离标题高度和内容高度
- 使用正确的gap公式

---

#### 1.3 修正 stat-card（canvas分支）高度

**文件**: `src/main.py` (line ~454)

**当前代码**:
```python
# 估算高度并添加stat-card背景
card_height = 300  # ❌ 固定值
```

**修改为**:
```python
# 从CSS读取约束
stat_card_padding_top = 20
stat_card_padding_bottom = 20

# 标题高度
has_title = card.find('p', class_='primary-color') is not None
title_height = 35 if has_title else 0

# canvas高度（固定220px，这是convert_chart传入的height）
canvas_height = 220

# stat-card总高度
card_height = stat_card_padding_top + title_height + canvas_height + stat_card_padding_bottom
```

---

#### 1.4 data-card 返回实际Y坐标

**文件**: `src/main.py` (line ~735+)

**当前问题**: 方法签名是`-> int`但实际未返回值

**修改**:
```python
def _convert_data_card(self, card, pptx_slide, shape_converter, y_start: int) -> int:
    logger.info("处理data-card")
    x_base = 80

    # 添加左边框
    shape_converter.add_border_left(x_base, y_start, 280, 4)

    # 标题
    p_elem = card.find('p', class_='primary-color')
    progress_y = y_start + 10  # 初始化

    if p_elem:
        # ... 渲染标题 ...
        progress_y = y_start + 50  # 标题后的位置

    # 进度条
    progress_bars = card.find_all('div', class_='progress-container')
    for progress in progress_bars:
        # ... 渲染进度条 ...
        progress_y += 60

    # 列表项
    bullet_points = card.find_all('div', class_='bullet-point')
    for bullet in bullet_points:
        # ... 渲染bullet ...
        progress_y += 40  # 每个bullet约40px

    # ✅ 返回实际使用的Y坐标
    return progress_y
```

**改进点**:
- 明确初始化progress_y
- 累加每个元素的实际高度
- 返回最终Y坐标

---

#### 1.5 strategy-card 动态高度计算

**文件**: `src/main.py` (line ~597)

**当前代码**:
```python
action_items = card.find_all('div', class_='action-item')
card_height = len(action_items) * 80 + 80  # ❌ 估算
```

**修改为**:
```python
action_items = card.find_all('div', class_='action-item')

# 从CSS读取约束
strategy_card_padding = 10  # top + bottom = 20
action_item_margin_bottom = 15  # CSS中的margin-bottom

# 标题高度
has_title = card.find('p', class_='primary-color') is not None
title_height = 40 if has_title else 0

# 每个action-item的高度组成：
# - 圆形图标: 28px
# - 标题(action-title): 18px字体 × 1.5 = 27px
# - 描述(p): 16px字体 × 1.5 × 行数（估算2行）= 48px
# - margin-bottom: 15px
# 总计：28 + 27 + 48 + 15 = 118px

# 简化估算（TODO阶段2：根据实际文本行数计算）
single_action_item_height = 118

# strategy-card总高度
# = padding-top + title + (action-items × height) + padding-bottom
card_height = (strategy_card_padding + title_height +
               len(action_items) * single_action_item_height +
               strategy_card_padding)

# 限制在max-height范围内
card_height = min(card_height, 300)
```

**改进点**:
- 基于CSS margin-bottom计算
- 分离各部分高度
- 应用max-height约束

---

### 阶段2：中期优化（后续实施）

#### 2.1 实现 stat-box 动态内容高度计算

**目标**: 根据实际文本行数和元素数量计算stat-box高度

**实现思路**:
```python
def _calculate_stat_box_height(self, box) -> int:
    """
    动态计算stat-box的实际高度

    stat-box结构：
    - padding-top: 20px
    - 图标(36px font-size × 1.5): 54px
    - 图标与标题间距: 45px
    - 标题(18px × 1.5): 27px
    - 标题与h2间距: 30px
    - h2(36px × 1.5): 54px
    - h2与p间距: 50px
    - p(18px × 1.5 × 行数): 27px × 行数
    - padding-bottom: 20px
    """
    padding = 20  # top + bottom

    # 图标
    icon = box.find('i')
    icon_height = 54 if icon else 0

    # 标题
    title = box.find('div', class_='stat-title')
    title_height = 27 if title else 0

    # h2
    h2 = box.find('h2')
    h2_height = 54 if h2 else 0

    # p文本（需要估算行数）
    p = box.find('p')
    if p:
        text = p.get_text(strip=True)
        # 粗略估算：每行约15个中文字符，box宽度约400px
        chars_per_line = 15
        num_lines = max(1, len(text) // chars_per_line)
        p_height = 27 * num_lines
    else:
        p_height = 0

    # 间距（根据HTML实际结构）
    spacing = 45 + 30 + 50 if all([icon, title, h2, p]) else 0

    content_height = icon_height + title_height + h2_height + p_height + spacing
    total_height = padding * 2 + content_height

    # 应用CSS min/max约束
    total_height = max(200, min(total_height, 300))

    return total_height
```

**调用位置**: 在`_convert_stats_container`中替换固定的`box_height = 220`

---

#### 2.2 从 CSS Parser 读取容器约束

**新增方法**: `src/parser/css_parser.py`

```python
def get_height_constraints(self, selector: str) -> dict:
    """
    获取元素的高度约束

    Returns:
        {
            'min_height': int,  # px
            'max_height': int,  # px
            'padding_top': int,
            'padding_bottom': int,
            'padding_left': int,
            'padding_right': int,
        }
    """
    style = self.get_style(selector)
    if not style:
        return {}

    result = {}

    # min-height
    if 'min-height' in style:
        result['min_height'] = self._parse_size(style['min-height'])

    # max-height
    if 'max-height' in style:
        result['max_height'] = self._parse_size(style['max-height'])

    # padding
    if 'padding' in style:
        # 简化：假设padding是单一值
        padding = self._parse_size(style['padding'])
        result['padding_top'] = padding
        result['padding_bottom'] = padding
        result['padding_left'] = padding
        result['padding_right'] = padding

    return result

def _parse_size(self, size_str: str) -> int:
    """解析尺寸字符串（如'200px', '20px'）为整数"""
    import re
    match = re.search(r'(\d+)px', size_str)
    return int(match.group(1)) if match else 0
```

**使用示例**:
```python
constraints = self.css_parser.get_height_constraints('.stat-card')
min_height = constraints.get('min_height', 200)
max_height = constraints.get('max_height', 300)
padding_top = constraints.get('padding_top', 20)
```

---

#### 2.3 实现内容高度超出警告

**位置**: 每个容器转换方法的结尾

```python
# 检查是否超出max-height
if calculated_height > max_height:
    logger.warning(
        f"{card_type}内容高度({calculated_height}px)超出max-height({max_height}px)，"
        f"将被截断或溢出"
    )
```

---

### 阶段3：长期优化（可选）

#### 3.1 文本行数精确计算
- 使用PIL/Pillow测量文本实际宽度
- 根据字体大小和容器宽度计算准确行数

#### 3.2 自动高度调整
- 当内容超出max-height时，自动调整字体大小或截断文本
- 提供选项控制溢出行为

#### 3.3 单元测试
- 为每个容器类型编写高度计算测试
- 验证CSS约束是否正确应用

---

## 验证清单

修复完成后，检查以下项：

### 视觉验证
- [ ] slide01.html第一个容器与标题间距正常
- [ ] slide01.html容器之间间距一致（40px）
- [ ] slidewithtable.html所有容器正常显示
- [ ] 多行文本的stat-box高度正确

### 代码验证
- [ ] 所有`_convert_*`方法都返回实际使用的Y坐标
- [ ] 移除所有硬编码的高度值（220, 240, 80等）
- [ ] gap和padding从CSS读取而非假设
- [ ] 列数从inline style优先读取

### 日志验证
添加详细日志：
```python
logger.info(f"容器高度计算: {card_type}, 估算={estimated_height}px, 实际={actual_height}px")
```

---

## 实施顺序建议

1. **第一步**: 修复阶段1.1-1.5（返回值修正）
   - 预期效果：容器间距改善
   - 测试文件：slide01.html, slidewithtable.html

2. **第二步**: 实施阶段2.1（stat-box动态高度）
   - 预期效果：stat-box高度更准确
   - 测试文件：slide01.html（第一个容器）

3. **第三步**: 实施阶段2.2（CSS约束读取）
   - 预期效果：所有容器基于CSS规则
   - 测试：新增不同padding/height的HTML

4. **第四步**: 根据实际效果决定是否实施阶段3

---

## 注意事项

1. **渐进式修改**: 每次修改后立即测试，避免引入新bug
2. **保持降级逻辑**: 当CSS读取失败时使用合理默认值
3. **添加注释**: 每个高度计算都注释说明来源（CSS规则或计算公式）
4. **日志输出**: 关键计算步骤输出日志，便于调试

---

## 相关文件

- `src/main.py`: 主要容器转换逻辑
- `src/parser/css_parser.py`: CSS规则解析
- `src/converters/shape_converter.py`: 背景和边框渲染
- `slide01.html`: 测试文件（3列stat-box + strategy-card）
- `slidewithtable.html`: 测试文件（4列stat-box + canvas）

---

最后更新：2025-10-13
状态：待实施
