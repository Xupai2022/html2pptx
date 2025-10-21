# 容器坐标计算逻辑全面重构计划

## 问题诊断

### 当前致命缺陷

1. **硬编码高度问题**
   - `box_height = 220` (固定值)
   - `card_height = 180/220/240/300` (估算值)
   - `estimated_height = 80/100/120` (粗略估计)
   - 这些硬编码值无法适应不同内容量的容器

2. **缺少CSS约束读取**
   - 未读取`min-height/max-height`
   - 未读取`padding-top/padding-bottom`
   - 未读取`margin-bottom`
   - 未读取`gap`值（部分方法使用固定20px）

3. **内容高度计算缺失**
   - 未根据实际文本行数计算高度
   - 未根据子元素数量累加高度
   - 未考虑换行和文本溢出

4. **坐标传递错误**
   - 部分方法不返回正确的`y_offset`
   - 导致下一个容器位置错误，产生重叠或间距过大

5. **背景和内容不同步**
   - 背景矩形使用估算高度
   - 内容实际渲染高度不同
   - 导致背景过长/过短，文字溢出或紧贴边缘

### 具体表现

根据用户描述和代码分析，问题包括：
- ✗ 上下容器过窄，背景样式过长导致重合
- ✗ 背景样式过窄导致文字溢出
- ✗ 文字紧贴容器上边缘，缺少padding
- ✗ 两行文字挨得很近，行间距不正确
- ✗ 容器之间间距不一致

---

## 修复方案

### 阶段1：CSS解析器增强

#### 1.1 添加CSS约束提取方法

在`src/parser/css_parser.py`中添加：

```python
def get_height_constraints(self, selector: str) -> dict:
    """
    获取元素的高度相关约束
    
    Returns:
        {
            'min_height': int,  # px
            'max_height': int,  # px
            'padding_top': int,
            'padding_bottom': int,
            'padding_left': int,
            'padding_right': int,
            'margin_bottom': int,
            'margin_top': int
        }
    """
    style = self.get_style(selector)
    if not style:
        return {}
    
    result = {}
    
    # 解析min-height
    if 'min-height' in style:
        result['min_height'] = self._parse_size(style['min-height'])
    
    # 解析max-height
    if 'max-height' in style:
        result['max_height'] = self._parse_size(style['max-height'])
    
    # 解析padding (简化版，假设是单一值)
    if 'padding' in style:
        padding = self._parse_size(style['padding'])
        result['padding_top'] = padding
        result['padding_bottom'] = padding
        result['padding_left'] = padding
        result['padding_right'] = padding
    
    # 解析具体的padding-*
    if 'padding-top' in style:
        result['padding_top'] = self._parse_size(style['padding-top'])
    if 'padding-bottom' in style:
        result['padding_bottom'] = self._parse_size(style['padding-bottom'])
    if 'padding-left' in style:
        result['padding_left'] = self._parse_size(style['padding-left'])
    if 'padding-right' in style:
        result['padding_right'] = self._parse_size(style['padding-right'])
    
    # 解析margin
    if 'margin-bottom' in style:
        result['margin_bottom'] = self._parse_size(style['margin-bottom'])
    if 'margin-top' in style:
        result['margin_top'] = self._parse_size(style['margin-top'])
    
    return result

def _parse_size(self, size_str: str) -> int:
    """解析尺寸字符串（如'200px', '20px'）为整数"""
    import re
    match = re.search(r'(\d+)px', size_str)
    return int(match.group(1)) if match else 0
```

### 阶段2：内容高度计算器

#### 2.1 创建ContentHeightCalculator

创建`src/utils/content_height_calculator.py`：

```python
"""
内容高度计算器
精确计算容器内容的实际高度
"""

class ContentHeightCalculator:
    """内容高度计算器"""
    
    def __init__(self, css_parser, style_computer):
        self.css_parser = css_parser
        self.style_computer = style_computer
    
    def calculate_stat_box_height(self, box, box_width: int) -> int:
        """
        动态计算stat-box的实际高度
        
        stat-box结构：
        - padding-top: 20px
        - 图标(36px font-size)
        - 标题(stat-title, 18px)
        - h2(36px)
        - p(可能多行)
        - padding-bottom: 20px
        
        Args:
            box: stat-box元素
            box_width: 容器宽度（用于计算文本换行）
            
        Returns:
            实际高度（px）
        """
        # 获取CSS约束
        constraints = self.css_parser.get_height_constraints('.stat-box')
        padding_top = constraints.get('padding_top', 20)
        padding_bottom = constraints.get('padding_bottom', 20)
        min_height = constraints.get('min_height', 200)
        max_height = constraints.get('max_height', 300)
        
        content_height = 0
        
        # 1. 图标高度
        icon = box.find('i')
        if icon:
            icon_font_size = 36  # 从CSS获取
            content_height += icon_font_size
            content_height += 20  # margin-right作为间距
        
        # 2. 标题高度
        title = box.find('div', class_='stat-title')
        if title:
            title_font_size_pt = self.style_computer.get_font_size_pt(title)
            title_height = int(title_font_size_pt * 1.5)  # 1.5倍行高
            content_height += title_height
            content_height += 5  # margin-bottom
        
        # 3. h2高度
        h2 = box.find('h2')
        if h2:
            h2_font_size_pt = self.style_computer.get_font_size_pt(h2)
            h2_height = int(h2_font_size_pt * 1.5)
            content_height += h2_height
            content_height += 10  # margin-bottom
        
        # 4. p标签高度（可能有多个，可能多行）
        p_tags = box.find_all('p')
        for p in p_tags:
            p_text = p.get_text(strip=True)
            if p_text:
                p_font_size_pt = self.style_computer.get_font_size_pt(p)
                # 计算文本行数
                # 简化：假设中文字符宽度=字体大小，英文=0.6*字体大小
                # 可用宽度 = box_width - padding_left - padding_right
                available_width = box_width - 40  # 左右padding各20px
                num_lines = self._calculate_text_lines(p_text, p_font_size_pt, available_width)
                p_height = num_lines * int(p_font_size_pt * 1.5)
                content_height += p_height
                content_height += 8  # margin-bottom
        
        # 5. 计算总高度
        total_height = padding_top + content_height + padding_bottom
        
        # 6. 应用min/max约束
        total_height = max(min_height, min(total_height, max_height))
        
        return total_height
    
    def calculate_data_card_height(self, card, card_width: int) -> int:
        """
        计算data-card的实际高度
        
        data-card结构：
        - padding-top: 15px (CSS中的padding-left暗示有内边距)
        - h3标题(可选)
        - bullet-point列表 或 risk-item列表
        - padding-bottom: 15px
        """
        # 从CSS获取约束
        constraints = self.css_parser.get_height_constraints('.data-card')
        padding_top = constraints.get('padding_top', 15)
        padding_bottom = constraints.get('padding_bottom', 15)
        min_height = constraints.get('min_height', 200)
        max_height = constraints.get('max_height', 300)
        
        content_height = 0
        
        # 1. h3标题
        h3 = card.find('h3')
        if h3:
            h3_font_size_pt = self.style_computer.get_font_size_pt(h3)
            h3_height = int(h3_font_size_pt * 1.5)
            content_height += h3_height
            content_height += 12  # margin-bottom
        
        # 2. bullet-point
        bullet_points = card.find_all('div', class_='bullet-point')
        if bullet_points:
            for bp in bullet_points:
                # bullet-point高度 = 图标(20px) + 文字行数
                p_elem = bp.find('p')
                if p_elem:
                    p_font_size_pt = self.style_computer.get_font_size_pt(p_elem)
                    available_width = card_width - 40 - 20  # 减去padding和图标
                    p_text = p_elem.get_text(strip=True)
                    num_lines = self._calculate_text_lines(p_text, p_font_size_pt, available_width)
                    bp_height = max(20, num_lines * int(p_font_size_pt * 1.5))
                    content_height += bp_height
                    content_height += 8  # margin-bottom
        
        # 3. risk-item
        risk_items = card.find_all('div', class_='risk-item')
        if risk_items:
            for ri in risk_items:
                # risk-item = 标题行(22px) + 描述(可能多行) + margin-bottom(12px)
                content_height += 22  # 标题
                content_height += 4   # 间距
                desc_p = ri.find('p', class_='text-sm')
                if desc_p:
                    desc_text = desc_p.get_text(strip=True)
                    desc_font_size_pt = 14  # text-sm
                    available_width = card_width - 60
                    num_lines = self._calculate_text_lines(desc_text, desc_font_size_pt, available_width)
                    content_height += num_lines * int(desc_font_size_pt * 1.6)
                content_height += 12  # margin-bottom
        
        # 4. 计算总高度
        total_height = padding_top + content_height + padding_bottom
        
        # 5. 应用约束
        total_height = max(min_height, min(total_height, max_height))
        
        return total_height
    
    def calculate_strategy_card_height(self, card) -> int:
        """
        计算strategy-card的实际高度
        
        strategy-card结构：
        - padding: 10px
        - 标题(可选)
        - action-item列表
        """
        constraints = self.css_parser.get_height_constraints('.strategy-card')
        padding = constraints.get('padding_top', 10)
        min_height = constraints.get('min_height', 200)
        max_height = constraints.get('max_height', 300)
        
        content_height = 0
        
        # 1. 标题
        title = card.find('p', class_='primary-color')
        if title:
            title_font_size_pt = self.style_computer.get_font_size_pt(title)
            content_height += int(title_font_size_pt * 1.5)
            content_height += 10
        
        # 2. action-item列表
        action_items = card.find_all('div', class_='action-item')
        for ai in action_items:
            # action-item高度 = 圆形图标(28px) + 标题(18px*1.5) + 描述(16px*1.5*行数) + margin-bottom(15px)
            action_height = 28  # 圆形图标
            
            action_title = ai.find('div', class_='action-title')
            if action_title:
                action_height += 27  # 18px * 1.5
            
            action_p = ai.find('p')
            if action_p:
                # 假设描述文字2行
                action_height += 48  # 16px * 1.5 * 2行
            
            action_height += 15  # margin-bottom
            content_height += action_height
        
        # 3. 计算总高度
        total_height = padding * 2 + content_height
        total_height = max(min_height, min(total_height, max_height))
        
        return total_height
    
    def _calculate_text_lines(self, text: str, font_size_pt: float, available_width: int) -> int:
        """
        计算文本的行数
        
        Args:
            text: 文本内容
            font_size_pt: 字体大小（点）
            available_width: 可用宽度（像素）
            
        Returns:
            行数
        """
        if not text:
            return 1
        
        # 将pt转换为px (1pt ≈ 0.75px)
        font_size_px = int(font_size_pt * 0.75)
        
        # 累计当前行宽度
        current_line_width = 0
        lines = 1
        
        for char in text:
            # 中文字符宽度 ≈ 字体大小
            if '\u4e00' <= char <= '\u9fff':
                char_width = font_size_px
            # 英文字母和数字宽度 ≈ 0.6 * 字体大小
            elif char.isalnum() or char in '.,;:!?\'"()[]{}-+/\\=_@#%&*':
                char_width = int(font_size_px * 0.6)
            # 空格宽度 ≈ 0.3 * 字体大小
            elif char == ' ':
                char_width = int(font_size_px * 0.3)
            else:
                char_width = font_size_px
            
            current_line_width += char_width
            
            # 超出可用宽度，换行
            if current_line_width > available_width:
                lines += 1
                current_line_width = char_width
        
        return lines
```

### 阶段3：容器转换方法重构

#### 3.1 重构_convert_stats_container

```python
def _convert_stats_container(self, container, pptx_slide, y_start: int) -> int:
    """
    转换统计卡片容器 (.stats-container)
    """
    stat_boxes = container.find_all('div', class_='stat-box')
    num_boxes = len(stat_boxes)
    
    if num_boxes == 0:
        return y_start
    
    # 动态获取列数
    num_columns = self._get_grid_columns(container)
    
    # 从CSS获取gap
    gap = self.css_parser.get_gap_size('.stats-container')
    if gap == 20:  # 如果是默认值，尝试从inline style获取
        inline_style = container.get('style', '')
        if 'gap' in inline_style:
            import re
            gap_match = re.search(r'gap:\s*(\d+)px', inline_style)
            if gap_match:
                gap = int(gap_match.group(1))
    
    logger.info(f"stats-container: {num_boxes}个box, {num_columns}列, gap={gap}px")
    
    # 计算布局
    total_width = 1760
    box_width = int((total_width - (num_columns - 1) * gap) / num_columns)
    x_start = 80
    
    # 初始化高度计算器
    height_calculator = ContentHeightCalculator(self.css_parser, self.style_computer)
    
    # 记录每一行的最大高度
    row_heights = []
    current_row = -1
    
    # 第一遍：计算所有box的高度
    box_heights = []
    for box in stat_boxes:
        box_height = height_calculator.calculate_stat_box_height(box, box_width)
        box_heights.append(box_height)
    
    # 第二遍：渲染所有box
    current_y = y_start
    for idx, box in enumerate(stat_boxes):
        col = idx % num_columns
        row = idx // num_columns
        
        # 新行开始
        if row != current_row:
            if current_row >= 0:
                # 移动到下一行
                current_y += row_heights[current_row] + gap
            current_row = row
            # 计算当前行的最大高度
            row_start_idx = row * num_columns
            row_end_idx = min(row_start_idx + num_columns, num_boxes)
            row_max_height = max(box_heights[row_start_idx:row_end_idx])
            row_heights.append(row_max_height)
        
        x = x_start + col * (box_width + gap)
        y = current_y
        box_height = box_heights[idx]
        
        # 添加背景（使用精确计算的高度）
        shape_converter = ShapeConverter(pptx_slide, self.css_parser)
        shape_converter.add_stat_box_background(x, y, box_width, box_height)
        
        # 渲染内容（传入精确的box_height）
        self._render_stat_box_content(box, pptx_slide, x, y, box_width, box_height)
    
    # 返回下一个元素的Y坐标
    # 计算总高度 = 所有行的高度 + 行间距
    total_height = sum(row_heights) + (len(row_heights) - 1) * gap if row_heights else 0
    
    return y_start + total_height
```

#### 3.2 重构_convert_stat_card

```python
def _convert_stat_card(self, card, pptx_slide, y_start: int) -> int:
    """
    转换stat-card（带背景的统计卡片）
    """
    logger.info("处理stat-card")
    
    # 获取CSS约束
    constraints = self.css_parser.get_height_constraints('.stat-card')
    padding_top = constraints.get('padding_top', 20)
    padding_bottom = constraints.get('padding_bottom', 20)
    
    # 检查是否包含stats-container
    stats_container = card.find('div', class_='stats-container')
    
    if stats_container:
        # stat-card包含stats-container的情况
        # 结构：padding-top + (可选标题) + stats-container + padding-bottom
        
        # 1. 计算标题高度
        title_height = 0
        title_elem = card.find('p', class_='primary-color')
        if title_elem:
            title_font_size_pt = self.style_computer.get_font_size_pt(title_elem)
            title_height = int(title_font_size_pt * 1.5) + 12  # 加上margin-bottom
        
        # 2. 计算stats-container的高度
        stat_boxes = stats_container.find_all('div', class_='stat-box')
        num_boxes = len(stat_boxes)
        num_columns = self._get_grid_columns(stats_container)
        gap = self.css_parser.get_gap_size('.stats-container')
        
        # 获取box宽度
        total_width = 1760  # 整个幻灯片宽度 - 左右边距
        box_width = int((total_width - (num_columns - 1) * gap) / num_columns)
        
        # 计算每个box的高度
        height_calculator = ContentHeightCalculator(self.css_parser, self.style_computer)
        box_heights = []
        for box in stat_boxes:
            box_height = height_calculator.calculate_stat_box_height(box, box_width)
            box_heights.append(box_height)
        
        # 计算stats-container的实际高度
        num_rows = (num_boxes + num_columns - 1) // num_columns
        row_heights = []
        for row in range(num_rows):
            row_start = row * num_columns
            row_end = min(row_start + num_columns, num_boxes)
            row_max_height = max(box_heights[row_start:row_end])
            row_heights.append(row_max_height)
        
        stats_container_height = sum(row_heights) + (num_rows - 1) * gap if row_heights else 0
        
        # 3. 计算stat-card总高度
        card_height = padding_top + title_height + stats_container_height + padding_bottom
        
        # 4. 添加stat-card背景
        from pptx.enum.shapes import MSO_SHAPE
        bg_color_str = self.css_parser.get_background_color('.stat-card')
        if bg_color_str:
            bg_shape = pptx_slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                UnitConverter.px_to_emu(80),
                UnitConverter.px_to_emu(y_start),
                UnitConverter.px_to_emu(1760),
                UnitConverter.px_to_emu(card_height)
            )
            bg_shape.fill.solid()
            bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
            if bg_rgb and alpha < 1.0:
                bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
            bg_shape.fill.fore_color.rgb = bg_rgb
            bg_shape.line.fill.background()
        
        # 5. 渲染标题
        current_y = y_start + padding_top
        if title_elem:
            title_text = title_elem.get_text(strip=True)
            # 渲染标题...
            current_y += title_height
        
        # 6. 渲染stats-container
        self._convert_stats_container(stats_container, pptx_slide, current_y)
        
        # 7. 返回下一个元素的Y坐标
        return y_start + card_height
    
    elif card.find('canvas'):
        # stat-card包含canvas（图表）的情况
        # 结构：padding-top + 标题 + canvas(220px) + padding-bottom
        
        title_height = 0
        title_elem = card.find('p', class_='primary-color')
        if title_elem:
            title_font_size_pt = self.style_computer.get_font_size_pt(title_elem)
            title_height = int(title_font_size_pt * 1.5) + 12
        
        canvas_height = 220  # canvas固定高度
        card_height = padding_top + title_height + canvas_height + padding_bottom
        
        # 添加背景和内容...
        
        return y_start + card_height
    
    else:
        # 其他类型的stat-card，降级处理
        # 使用ContentHeightCalculator计算
        logger.warning("stat-card未匹配到已知模式，使用降级逻辑")
        
        # 估算内容高度
        content_height = self._estimate_generic_content_height(card, 1760)
        card_height = padding_top + content_height + padding_bottom
        
        # 添加背景和内容...
        
        return y_start + card_height
```

#### 3.3 重构_convert_data_card

```python
def _convert_data_card(self, card, pptx_slide, shape_converter, y_start: int) -> int:
    """
    转换data-card（带左边框的数据卡片）
    """
    logger.info("处理data-card")
    
    # 使用高度计算器
    height_calculator = ContentHeightCalculator(self.css_parser, self.style_computer)
    card_width = 1760
    card_height = height_calculator.calculate_data_card_height(card, card_width)
    
    x_base = 80
    
    # 1. 添加背景（使用精确计算的高度）
    bg_color_str = self.css_parser.get_background_color('.data-card')
    if bg_color_str:
        from pptx.enum.shapes import MSO_SHAPE
        bg_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            UnitConverter.px_to_emu(x_base),
            UnitConverter.px_to_emu(y_start),
            UnitConverter.px_to_emu(card_width),
            UnitConverter.px_to_emu(card_height)
        )
        bg_shape.fill.solid()
        bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
        if bg_rgb and alpha < 1.0:
            bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
        bg_shape.fill.fore_color.rgb = bg_rgb
        bg_shape.line.fill.background()
    
    # 2. 添加左边框（使用精确的高度）
    shape_converter.add_border_left(x_base, y_start, card_height, 4)
    
    # 3. 渲染内容
    constraints = self.css_parser.get_height_constraints('.data-card')
    padding_top = constraints.get('padding_top', 15)
    current_y = y_start + padding_top
    
    # 渲染h3、bullet-point、risk-item等...
    # （保持现有逻辑，但使用padding_top而不是硬编码的10px）
    
    # 4. 返回下一个元素的Y坐标
    return y_start + card_height
```

#### 3.4 重构_convert_strategy_card

```python
def _convert_strategy_card(self, card, pptx_slide, y_start: int) -> int:
    """
    转换strategy-card（策略卡片）
    """
    logger.info("处理strategy-card")
    
    # 使用高度计算器
    height_calculator = ContentHeightCalculator(self.css_parser, self.style_computer)
    card_height = height_calculator.calculate_strategy_card_height(card)
    
    x_base = 80
    card_width = 1760
    
    # 1. 添加背景
    bg_color_str = self.css_parser.get_background_color('.strategy-card')
    if bg_color_str:
        from pptx.enum.shapes import MSO_SHAPE
        bg_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            UnitConverter.px_to_emu(x_base),
            UnitConverter.px_to_emu(y_start),
            UnitConverter.px_to_emu(card_width),
            UnitConverter.px_to_emu(card_height)
        )
        bg_shape.fill.solid()
        bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
        if bg_rgb and alpha < 1.0:
            bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
        bg_shape.fill.fore_color.rgb = bg_rgb
        bg_shape.line.fill.background()
    
    # 2. 添加左边框
    shape_converter = ShapeConverter(pptx_slide, self.css_parser)
    shape_converter.add_border_left(x_base, y_start, card_height, 4)
    
    # 3. 渲染内容
    constraints = self.css_parser.get_height_constraints('.strategy-card')
    padding = constraints.get('padding_top', 10)
    current_y = y_start + padding
    
    # 渲染标题和action-item...
    
    # 4. 返回下一个元素的Y坐标
    return y_start + card_height
```

### 阶段4：辅助方法

#### 4.1 添加_get_grid_columns方法

```python
def _get_grid_columns(self, container) -> int:
    """
    动态获取grid的列数
    优先级：inline style > CSS规则 > 默认值
    """
    num_columns = 4  # 默认值
    
    # 1. 检查inline style
    inline_style = container.get('style', '')
    if 'grid-template-columns' in inline_style:
        import re
        repeat_match = re.search(r'repeat\((\d+),', inline_style)
        if repeat_match:
            num_columns = int(repeat_match.group(1))
            logger.debug(f"从inline style获取列数: {num_columns}")
            return num_columns
        
        fr_count = len(re.findall(r'1fr', inline_style))
        if fr_count > 0:
            num_columns = fr_count
            logger.debug(f"从inline style获取列数: {num_columns}")
            return num_columns
    
    # 2. 从CSS规则获取
    container_classes = container.get('class', [])
    for cls in container_classes:
        if cls.startswith('grid-cols-'):
            num_columns = self.css_parser.get_grid_columns(f'.{cls}')
            logger.debug(f"从CSS类获取列数: {num_columns}")
            return num_columns
    
    # 3. 检查CSS选择器
    num_columns = self.css_parser.get_grid_columns('.stats-container')
    logger.debug(f"从CSS选择器获取列数: {num_columns}")
    
    return num_columns
```

#### 4.2 添加_estimate_generic_content_height方法

```python
def _estimate_generic_content_height(self, container, container_width: int) -> int:
    """
    估算通用容器的内容高度
    当无法使用专门的计算器时的降级方法
    """
    content_height = 0
    
    # 查找所有文本元素
    text_elements = []
    for elem in container.descendants:
        if hasattr(elem, 'name') and elem.name in ['p', 'h1', 'h2', 'h3', 'h4', 'div', 'span']:
            # 只提取没有子块级元素的文本节点
            if not elem.find_all(['div', 'p', 'h1', 'h2', 'h3']):
                text = elem.get_text(strip=True)
                if text and len(text) > 2:
                    text_elements.append(elem)
    
    # 计算每个元素的高度
    for elem in text_elements[:10]:  # 最多处理10个元素
        font_size_pt = self.style_computer.get_font_size_pt(elem)
        elem_height = int(font_size_pt * 1.5)  # 简单估算
        content_height += elem_height + 10  # 加上间距
    
    return content_height if content_height > 0 else 100  # 最小100px
```

---

## 实施步骤

### Step 1: 增强CSS解析器 (优先级: 🔴 最高)

文件：`src/parser/css_parser.py`

- [ ] 添加`get_height_constraints()`方法
- [ ] 添加`_parse_size()`辅助方法
- [ ] 测试CSS约束提取功能

### Step 2: 创建内容高度计算器 (优先级: 🔴 最高)

文件：`src/utils/content_height_calculator.py`（新建）

- [ ] 实现`ContentHeightCalculator`类
- [ ] 实现`calculate_stat_box_height()`
- [ ] 实现`calculate_data_card_height()`
- [ ] 实现`calculate_strategy_card_height()`
- [ ] 实现`_calculate_text_lines()`文本行数计算
- [ ] 单元测试

### Step 3: 重构main.py容器转换方法 (优先级: 🔴 最高)

文件：`src/main.py`

- [ ] 重构`_convert_stats_container()`
- [ ] 重构`_convert_stat_card()`
- [ ] 重构`_convert_data_card()`
- [ ] 重构`_convert_strategy_card()`
- [ ] 重构`_convert_grid_data_card()`
- [ ] 重构`_convert_grid_stat_card()`
- [ ] 重构`_convert_grid_risk_card()`
- [ ] 添加`_get_grid_columns()`辅助方法
- [ ] 添加`_estimate_generic_content_height()`辅助方法

### Step 4: 修复所有硬编码值 (优先级: 🟡 高)

在`src/main.py`中全局搜索并替换：

- [ ] 搜索`220`，替换为动态计算
- [ ] 搜索`240`，替换为动态计算
- [ ] 搜索`180`，替换为动态计算
- [ ] 搜索`80`，替换为动态计算（容器高度相关）
- [ ] 搜索`100`，替换为动态计算（容器高度相关）
- [ ] 搜索`gap = 20`，替换为`gap = self.css_parser.get_gap_size(...)`

### Step 5: 增强ShapeConverter (优先级: 🟡 高)

文件：`src/converters/shape_converter.py`

- [ ] 修改`add_stat_box_background()`，接受高度参数
- [ ] 修改`add_border_left()`，确保高度参数正确使用

### Step 6: 测试验证 (优先级: 🔴 最高)

- [ ] 测试`slide01.html`（包含3列stat-box + strategy-card）
- [ ] 测试`slide02.html`
- [ ] 测试`slidewithtable.html`（包含4列stat-box + canvas）
- [ ] 测试所有其他slide*.html文件
- [ ] 检查容器间距是否一致
- [ ] 检查背景与内容是否完美对齐
- [ ] 检查文字是否有足够的padding
- [ ] 检查是否有重叠或溢出

### Step 7: 日志和调试 (优先级: 🟢 中)

在关键位置添加详细日志：

```python
logger.info(f"容器: {container_type}, 计算高度={calculated_height}px, 实际高度={actual_height}px")
logger.info(f"  - padding_top={padding_top}px, padding_bottom={padding_bottom}px")
logger.info(f"  - content_height={content_height}px")
logger.info(f"  - 约束: min={min_height}px, max={max_height}px")
```

### Step 8: 文档更新 (优先级: 🟢 中)

- [ ] 更新`PROJECT_SUMMARY.md`，记录此次重构
- [ ] 更新`CONTAINER_HEIGHT_FIX_PLAN.md`状态为已完成
- [ ] 创建`LAYOUT_CALCULATION_GUIDE.md`开发者指南

---

## 验证清单

### 视觉验证

- [ ] slide01.html第一个容器与标题间距正常（应为40px，space-y-10的间距）
- [ ] slide01.html容器之间间距一致（40px）
- [ ] slidewithtable.html所有容器正常显示
- [ ] 多行文本的stat-box高度正确，文字不溢出
- [ ] data-card的左边框与背景高度完全一致
- [ ] strategy-card的action-item排列整齐，无重叠
- [ ] 所有容器的padding正确（文字不紧贴边缘）
- [ ] 行间距合理（1.5倍行高）

### 代码验证

- [ ] 所有`_convert_*`方法都返回实际使用的Y坐标
- [ ] 移除所有硬编码的高度值（220, 240, 80等）
- [ ] gap和padding从CSS读取而非假设
- [ ] 列数从inline style优先读取
- [ ] 所有容器使用ContentHeightCalculator计算高度
- [ ] 背景矩形高度 = 内容实际高度

### 性能验证

- [ ] 转换时间没有显著增加（应在原有基础上 ±10%）
- [ ] 内存占用正常
- [ ] 无死循环或递归过深

---

## 注意事项

### 1. 渐进式修改

- **不要一次性修改所有方法**
- 按照Step 1 → Step 2 → Step 3的顺序逐步实施
- 每完成一个Step，立即测试

### 2. 保持降级逻辑

- 当CSS读取失败时使用合理默认值
- 当无法使用ContentHeightCalculator时使用`_estimate_generic_content_height()`

### 3. 添加详细注释

每个高度计算都注释说明来源：

```python
# 从CSS读取: .stat-box { padding: 20px; }
padding_top = constraints.get('padding_top', 20)

# 根据字体大小计算: font-size * 1.5 (行高)
title_height = int(title_font_size_pt * 1.5)

# 应用CSS约束: min-height: 200px, max-height: 300px
total_height = max(200, min(total_height, 300))
```

### 4. 处理边界情况

- 空容器（没有内容）
- 单行文本 vs 多行文本
- 非常短的文本 vs 超长文本
- 不同列数的grid (2列、3列、4列)

### 5. 兼容性

确保修改后的代码仍然兼容：
- 所有现有的HTML文件
- template.txt定义的所有容器类型
- 不同的inline style写法

---

## 预期效果

修复完成后，应达到以下效果：

1. ✅ **布局精确**：PPTX布局与HTML视觉效果完全一致
2. ✅ **间距统一**：所有容器间距符合CSS定义（space-y-10 = 40px）
3. ✅ **背景完美**：背景矩形高度与内容高度完全匹配
4. ✅ **文字合理**：文字有足够的padding，行间距正确
5. ✅ **无溢出**：所有文字和元素都在容器范围内
6. ✅ **无重叠**：容器之间不会重叠
7. ✅ **可维护**：代码清晰，易于理解和扩展

---

## 相关文件

- `src/main.py`: 主要容器转换逻辑
- `src/parser/css_parser.py`: CSS规则解析
- `src/utils/content_height_calculator.py`: 内容高度计算器（新建）
- `src/converters/shape_converter.py`: 背景和边框渲染
- `slide01.html`: 测试文件（3列stat-box + strategy-card）
- `slidewithtable.html`: 测试文件（4列stat-box + canvas）
- `template.txt`: HTML模板定义

---

最后更新：2025-10-21  
状态：待实施
