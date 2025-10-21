"""
内容高度计算器
精确计算容器内容的实际高度，避免硬编码
"""

from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ContentHeightCalculator:
    """
    内容高度计算器
    
    根据实际内容（文本、图标、子元素等）动态计算容器高度
    考虑CSS约束（min-height, max-height, padding等）
    """
    
    def __init__(self, css_parser, style_computer):
        """
        初始化内容高度计算器
        
        Args:
            css_parser: CSS解析器实例
            style_computer: 样式计算器实例
        """
        self.css_parser = css_parser
        self.style_computer = style_computer
    
    def calculate_stat_box_height(self, box, box_width: int) -> int:
        """
        动态计算stat-box的实际高度
        
        stat-box结构：
        - padding-top: 20px
        - 图标(36px font-size) + margin-right(20px)
        - 标题(stat-title, 18px) + margin-bottom(5px)
        - h2(36px) + margin-bottom(10px)
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
        # 移除硬编码默认值
        min_height = constraints.get('min_height', None)
        max_height = constraints.get('max_height', None)
        
        logger.debug(f"stat-box CSS约束: padding={padding_top}/{padding_bottom}, min={min_height}, max={max_height}")
        
        content_height = 0
        
        # 1. 图标高度
        icon = box.find('i')
        if icon:
            icon_font_size = 36  # 从CSS .stat-icon获取
            content_height += icon_font_size
            logger.debug(f"  + 图标: {icon_font_size}px")
        
        # 2. 标题高度
        title = box.find('div', class_='stat-title')
        if title:
            title_font_size_pt = self.style_computer.get_font_size_pt(title)
            title_height = int(title_font_size_pt * 1.5)  # 1.5倍行高
            content_height += title_height
            content_height += 5  # margin-bottom from CSS
            logger.debug(f"  + 标题: {title_height}px + 5px margin")
        
        # 3. h2高度
        h2 = box.find('h2')
        if h2:
            h2_font_size_pt = self.style_computer.get_font_size_pt(h2)
            h2_height = int(h2_font_size_pt * 1.5)
            content_height += h2_height
            content_height += 10  # margin-bottom估算
            logger.debug(f"  + h2: {h2_height}px + 10px margin")
        
        # 4. p标签高度（可能有多个，可能多行）
        p_tags = box.find_all('p')
        for p in p_tags:
            p_text = p.get_text(strip=True)
            if p_text:
                p_font_size_pt = self.style_computer.get_font_size_pt(p)
                # 计算文本行数
                # 可用宽度 = box_width - padding_left - padding_right - icon_width - margin
                available_width = box_width - 40  # 左右padding各20px
                if icon:
                    available_width -= 56  # icon(36px) + margin-right(20px)
                
                num_lines = self._calculate_text_lines(p_text, p_font_size_pt, available_width)
                p_height = num_lines * int(p_font_size_pt * 1.5)
                content_height += p_height
                content_height += 8  # margin-bottom估算
                logger.debug(f"  + p标签: {num_lines}行 × {int(p_font_size_pt * 1.5)}px = {p_height}px")
        
        # 5. 计算总高度
        total_height = padding_top + content_height + padding_bottom
        
        # 6. 应用min/max约束（只在CSS明确定义时才应用）
        original_height = total_height
        if min_height is not None:
            total_height = max(min_height, total_height)
        if max_height is not None:
            total_height = min(total_height, max_height)
        
        if total_height != original_height:
            logger.debug(f"应用约束: {original_height}px → {total_height}px")
        
        logger.info(f"stat-box高度计算完成: {total_height}px (内容={content_height}px, padding={padding_top+padding_bottom}px)")
        
        return total_height
    
    def calculate_data_card_height(self, card, card_width: int) -> int:
        """
        计算data-card的实际高度
        
        data-card结构：
        - padding-left导致的顶部间距
        - h3标题(可选) + margin-bottom(12px)
        - bullet-point列表 或 risk-item列表
        - padding-bottom
        
        Args:
            card: data-card元素
            card_width: 卡片宽度
            
        Returns:
            实际高度（px）
        """
        # 从CSS获取约束
        constraints = self.css_parser.get_height_constraints('.data-card')
        # data-card的padding-left=15px，我们假设上下也有类似的间距
        padding_top = constraints.get('padding_top', 15)
        padding_bottom = constraints.get('padding_bottom', 15)
        # 移除硬编码默认值，如果CSS没有定义则为None
        min_height = constraints.get('min_height', None)
        max_height = constraints.get('max_height', None)
        
        logger.debug(f"data-card CSS约束: padding={padding_top}/{padding_bottom}, min={min_height}, max={max_height}")
        
        content_height = 0
        
        # 1. h3标题
        h3 = card.find('h3')
        if h3:
            h3_font_size_pt = self.style_computer.get_font_size_pt(h3)
            h3_height = int(h3_font_size_pt * 1.5)
            content_height += h3_height
            content_height += 12  # margin-bottom
            logger.debug(f"  + h3标题: {h3_height}px + 12px margin")
        
        # 2. bullet-point
        bullet_points = card.find_all('div', class_='bullet-point')
        # 同时检查space-y-3容器内的flex items-start结构
        if not bullet_points:
            space_y_containers = card.find_all('div', class_='space-y-3')
            for container in space_y_containers:
                flex_items = container.find_all('div', class_='flex')
                for flex_item in flex_items:
                    if flex_item.find('i') and flex_item.find('p'):
                        bullet_points.append(flex_item)
        
        if bullet_points:
            logger.debug(f"  发现{len(bullet_points)}个bullet-point")
            for idx, bp in enumerate(bullet_points):
                # bullet-point高度 = max(图标高度, 文字高度)
                icon_height = 20  # 默认图标高度
                
                p_elem = bp.find('p')
                if p_elem:
                    p_font_size_pt = self.style_computer.get_font_size_pt(p_elem)
                    available_width = card_width - 60  # 减去padding和图标+间距
                    p_text = p_elem.get_text(strip=True)
                    num_lines = self._calculate_text_lines(p_text, p_font_size_pt, available_width)
                    text_height = num_lines * int(p_font_size_pt * 1.5)
                    bp_height = max(icon_height, text_height)
                    content_height += bp_height
                    # space-y-3意味着每个元素间距12px，但最后一个不加
                    if idx < len(bullet_points) - 1:
                        content_height += 12
                    logger.debug(f"    bullet-point #{idx+1}: {bp_height}px ({num_lines}行文字)")
        
        # 3. risk-item
        risk_items = card.find_all('div', class_='risk-item')
        if risk_items:
            logger.debug(f"  发现{len(risk_items)}个risk-item")
            for idx, ri in enumerate(risk_items):
                # risk-item = 标题行(22px) + 间距(4px) + 描述(多行) + margin-bottom(12px)
                ri_height = 22  # 标题行
                ri_height += 4   # 间距
                
                desc_p = ri.find('p', class_='text-sm')
                if desc_p:
                    desc_text = desc_p.get_text(strip=True)
                    desc_font_size_pt = 14  # text-sm通常是14px
                    available_width = card_width - 60
                    num_lines = self._calculate_text_lines(desc_text, desc_font_size_pt, available_width)
                    desc_height = num_lines * int(desc_font_size_pt * 1.6)
                    ri_height += desc_height
                
                # 最后一个risk-item不加margin-bottom
                if idx < len(risk_items) - 1:
                    ri_height += 12
                
                content_height += ri_height
                logger.debug(f"    risk-item #{idx+1}: {ri_height}px")
        
        # 4. 如果既没有bullet-point也没有risk-item，处理其他内容
        if not bullet_points and not risk_items:
            # 查找所有直接子元素
            direct_children = []
            for child in card.children:
                if hasattr(child, 'name') and child.name:
                    if child.name != 'h3':  # h3已经计算过
                        text = child.get_text(strip=True)
                        if text and len(text) > 2:
                            direct_children.append(child)
            
            # 每个元素约35px高度
            if direct_children:
                content_height += len(direct_children) * 35
                logger.debug(f"  其他内容: {len(direct_children)}个元素 × 35px")
        
        # 5. 确保最小内容高度
        if content_height < 50:
            content_height = 50
            logger.debug("  内容过少，应用最小内容高度50px")
        
        # 6. 计算总高度
        total_height = padding_top + content_height + padding_bottom
        
        # 7. 应用约束（只在CSS明确定义时才应用）
        original_height = total_height
        if min_height is not None:
            total_height = max(min_height, total_height)
        if max_height is not None:
            total_height = min(total_height, max_height)
        
        if total_height != original_height:
            logger.debug(f"应用约束: {original_height}px → {total_height}px")
        
        logger.info(f"data-card高度计算完成: {total_height}px (内容={content_height}px, padding={padding_top+padding_bottom}px)")
        
        return total_height
    
    def calculate_strategy_card_height(self, card) -> int:
        """
        计算strategy-card的实际高度
        
        strategy-card结构：
        - padding: 10px (上下左右)
        - 标题(可选) + margin-bottom
        - action-item列表 (每个action-item包含圆形图标、标题、描述)
        
        Args:
            card: strategy-card元素
            
        Returns:
            实际高度（px）
        """
        constraints = self.css_parser.get_height_constraints('.strategy-card')
        padding = constraints.get('padding_top', 10)  # strategy-card使用统一padding
        # 移除硬编码默认值
        min_height = constraints.get('min_height', None)
        max_height = constraints.get('max_height', None)
        
        logger.debug(f"strategy-card CSS约束: padding={padding}, min={min_height}, max={max_height}")
        
        content_height = 0
        
        # 1. 标题
        title = card.find('p', class_='primary-color')
        if title:
            title_font_size_pt = self.style_computer.get_font_size_pt(title)
            title_height = int(title_font_size_pt * 1.5)
            content_height += title_height
            content_height += 10  # margin-bottom估算
            logger.debug(f"  + 标题: {title_height}px + 10px margin")
        
        # 2. action-item列表
        action_items = card.find_all('div', class_='action-item')
        if action_items:
            logger.debug(f"  发现{len(action_items)}个action-item")
            for idx, ai in enumerate(action_items):
                # action-item高度组成：
                # - 圆形图标: 28px (action-number)
                # - 标题(action-title): 18px × 1.5 = 27px
                # - 描述(p): 16px × 1.5 × 行数（估算2行）= 48px
                # - margin-bottom: 15px (CSS定义)
                
                action_height = 28  # 圆形图标（固定）
                
                # 标题
                action_title = ai.find('div', class_='action-title')
                if action_title:
                    action_height = max(action_height, 27)  # 至少27px
                
                # 描述文字
                action_p = ai.find('p')
                if action_p:
                    p_text = action_p.get_text(strip=True)
                    # 简化：假设描述文字2-3行
                    estimated_lines = 2 if len(p_text) < 50 else 3
                    action_height += estimated_lines * 24  # 16px × 1.5
                
                # 最后一个action-item不加margin-bottom
                if idx < len(action_items) - 1:
                    action_height += 15  # margin-bottom from CSS
                
                content_height += action_height
                logger.debug(f"    action-item #{idx+1}: {action_height}px")
        
        # 3. 计算总高度
        total_height = padding * 2 + content_height  # 上下都有padding
        
        # 4. 应用约束（只在CSS明确定义时才应用）
        original_height = total_height
        if min_height is not None:
            total_height = max(min_height, total_height)
        if max_height is not None:
            total_height = min(total_height, max_height)
        
        if total_height != original_height:
            logger.debug(f"应用约束: {original_height}px → {total_height}px")
        
        logger.info(f"strategy-card高度计算完成: {total_height}px (内容={content_height}px, padding={padding*2}px)")
        
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
        
        logger.debug(f"    文本行数计算: '{text[:30]}...' → {lines}行 (字体={font_size_pt}pt, 宽度={available_width}px)")
        
        return lines
