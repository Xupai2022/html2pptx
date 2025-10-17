"""
样式计算器
处理CSS级联、继承和最终样式计算
"""

from typing import Dict, Optional, List
from bs4 import BeautifulSoup, Tag

from src.utils.logger import setup_logger
from src.utils.font_size_extractor import FontSizeExtractor

logger = setup_logger(__name__)


class StyleComputer:
    """
    样式计算器

    功能：
    1. 计算元素的最终样式（考虑级联和继承）
    2. 处理CSS优先级
    3. 实现样式继承机制
    4. 提供统一的样式获取接口
    """

    # CSS选择器优先级权重
    SELECTOR_WEIGHTS = {
        'inline': 1000,     # 内联样式
        'id': 100,         # ID选择器
        'tailwind': 100,   # Tailwind类（更高优先级）
        'class': 10,       # 类选择器
        'tag': 1,          # 标签选择器
        'universal': 0,    # 通用选择器
    }

    # 默认样式值
    DEFAULT_STYLES = {
        'font-size': '16px',
        'font-family': 'serif',
        'color': '#000000',
        'font-weight': 'normal',
        'line-height': '1.2',
        'text-align': 'left',
    }

    def __init__(self, css_parser):
        """
        初始化样式计算器

        Args:
            css_parser: CSS解析器实例
        """
        self.css_parser = css_parser
        self.font_size_extractor = FontSizeExtractor(css_parser)
        self._style_cache = {}  # 样式缓存

    def compute_computed_style(self, element: Tag, parent_element: Tag = None) -> Dict[str, str]:
        """
        计算元素的最终样式（computed style）

        处理顺序：
        1. 收集所有适用的CSS规则
        2. 按优先级排序
        3. 应用继承的样式
        4. 应用默认样式

        Args:
            element: 目标元素
            parent_element: 父元素（用于继承）

        Returns:
            计算后的样式字典
        """
        if not element:
            return {}

        # 生成缓存键
        element_id = id(element)
        parent_id = id(parent_element) if parent_element else None
        cache_key = f"{element_id}_{parent_id}"

        if cache_key in self._style_cache:
            return self._style_cache[cache_key]

        # 1. 收集样式规则
        style_rules = self._collect_style_rules(element)

        # 2. 计算优先级并排序
        sorted_rules = self._sort_rules_by_specificity(style_rules)

        # 3. 应用样式规则
        computed_style = self._apply_style_rules(sorted_rules)

        # 4. 继承父元素样式
        if parent_element:
            parent_style = self.compute_computed_style(parent_element)
            computed_style = self._inherit_parent_style(computed_style, parent_style)

        # 5. 应用默认样式
        computed_style = self._apply_default_styles(computed_style)

        # 缓存结果
        self._style_cache[cache_key] = computed_style

        return computed_style

    def get_font_size_pt(self, element: Tag, parent_element: Tag = None) -> int:
        """
        获取元素的字体大小（pt）

        现在返回真正的pt值，而不是px值

        Args:
            element: 目标元素
            parent_element: 父元素

        Returns:
            字体大小(pt)
        """
        if not element:
            from src.utils.unit_converter import UnitConverter
            return UnitConverter.font_size_px_to_pt(16)  # 默认16px -> 12pt

        # 计算父元素的字体大小
        parent_font_size_px = None
        if parent_element:
            parent_computed_style = self.compute_computed_style(parent_element)
            parent_font_size_str = parent_computed_style.get('font-size', '16px')
            parent_font_size_px = self.font_size_extractor._parse_font_size_value(parent_font_size_str)
            logger.debug(f"父元素字体大小: {parent_font_size_str} → {parent_font_size_px}px")

        # 使用字体大小提取器
        font_size_px = self.font_size_extractor.extract_font_size(element, parent_font_size_px)

        # 转换为pt
        from src.utils.unit_converter import UnitConverter
        font_size_pt = UnitConverter.font_size_px_to_pt(font_size_px)

        # 获取元素信息用于调试
        element_info = element.name
        if element.get('class'):
            element_info += f".{'.'.join(element.get('class', []))}"
        text_preview = element.get_text(strip=True)[:20]

        logger.debug(f"元素 {element_info} 字体大小: {font_size_px}px → {font_size_pt}pt (文本: {text_preview})")

        return font_size_pt

    def _collect_style_rules(self, element: Tag) -> List[Dict]:
        """
        收集适用于元素的所有CSS规则

        Args:
            element: 目标元素

        Returns:
            样式规则列表
        """
        rules = []

        # 1. 内联样式（优先级最高）
        inline_style = element.get('style', '')
        if inline_style:
            rules.append({
                'type': 'inline',
                'specificity': self.SELECTOR_WEIGHTS['inline'],
                'style': self._parse_inline_style(inline_style),
                'source': 'inline'
            })

        # 2. ID选择器
        element_id = element.get('id')
        if element_id:
            id_style = self.css_parser.get_style(f"#{element_id}")
            if id_style:
                rules.append({
                    'type': 'id',
                    'specificity': self.SELECTOR_WEIGHTS['id'],
                    'style': id_style,
                    'source': f"#{element_id}"
                })

        # 3. 类选择器（包括Tailwind CSS类）
        classes = element.get('class', [])
        if classes:
            # 复合类选择器
            compound_class = f".{'.'.join(classes)}"
            compound_style = self.css_parser.get_style(compound_class)
            if compound_style:
                rules.append({
                    'type': 'class',
                    'specificity': self.SELECTOR_WEIGHTS['class'] * len(classes),
                    'style': compound_style,
                    'source': compound_class
                })

            # 单个类选择器（包括Tailwind CSS类）
            for cls in classes:
                class_style = self.css_parser.get_style(f".{cls}")
                if class_style:
                    rules.append({
                        'type': 'class',
                        'specificity': self.SELECTOR_WEIGHTS['class'],
                        'style': class_style,
                        'source': f".{cls}"
                    })
                else:
                    # 尝试解析Tailwind CSS类
                    tailwind_style = self._parse_tailwind_class(cls)
                    if tailwind_style:
                        rules.append({
                            'type': 'tailwind',
                            'specificity': self.SELECTOR_WEIGHTS['tailwind'],
                            'style': tailwind_style,
                            'source': f"tailwind.{cls}"
                        })

        # 4. 标签选择器
        tag_name = element.name
        if tag_name:
            tag_style = self.css_parser.get_style(tag_name)
            if tag_style:
                rules.append({
                    'type': 'tag',
                    'specificity': self.SELECTOR_WEIGHTS['tag'],
                    'style': tag_style,
                    'source': tag_name
                })

        return rules

    def _sort_rules_by_specificity(self, rules: List[Dict]) -> List[Dict]:
        """
        按优先级排序样式规则

        Args:
            rules: 样式规则列表

        Returns:
            排序后的样式规则列表
        """
        return sorted(rules, key=lambda rule: rule['specificity'], reverse=True)

    def _apply_style_rules(self, sorted_rules: List[Dict]) -> Dict[str, str]:
        """
        应用排序后的样式规则

        Args:
            sorted_rules: 排序后的样式规则列表

        Returns:
            应用后的样式字典
        """
        computed_style = {}

        for rule in sorted_rules:
            style_dict = rule['style']
            for property_name, value in style_dict.items():
                # 后面的规则覆盖前面的规则
                computed_style[property_name] = value

        return computed_style

    def _inherit_parent_style(self, computed_style: Dict[str, str], parent_style: Dict[str, str]) -> Dict[str, str]:
        """
        从父元素继承样式

        Args:
            computed_style: 当前计算出的样式
            parent_style: 父元素样式

        Returns:
            继承后的样式字典
        """
        # 可继承的CSS属性
        inheritable_properties = {
            'font-size', 'font-family', 'color', 'font-weight', 'line-height',
            'text-align', 'font-style', 'text-decoration', 'letter-spacing',
            'word-spacing', 'text-indent', 'white-space'
        }

        for property_name, value in parent_style.items():
            if property_name in inheritable_properties and property_name not in computed_style:
                computed_style[property_name] = value

        return computed_style

    def _apply_default_styles(self, computed_style: Dict[str, str]) -> Dict[str, str]:
        """
        应用默认样式

        Args:
            computed_style: 当前计算出的样式

        Returns:
            应用默认样式后的样式字典
        """
        for property_name, default_value in self.DEFAULT_STYLES.items():
            if property_name not in computed_style:
                computed_style[property_name] = default_value

        return computed_style

    def _parse_tailwind_class(self, class_name: str) -> Optional[Dict[str, str]]:
        """
        解析Tailwind CSS类，返回对应的CSS属性

        Args:
            class_name: Tailwind CSS类名

        Returns:
            CSS样式字典，如果不是Tailwind类则返回None
        """
        # 使用CSS解析器的Tailwind映射
        if hasattr(self.css_parser, 'parse_element_classes'):
            return self.css_parser.parse_element_classes({'class': [class_name]})

        # 手动解析一些常用的Tailwind类
        style = {}

        # 字体大小类
        if class_name in ['text-xs', 'text-sm', 'text-base', 'text-lg', 'text-xl', 'text-2xl', 'text-3xl', 'text-4xl', 'text-5xl', 'text-6xl']:
            if hasattr(self.css_parser, 'tailwind_font_sizes'):
                font_size = self.css_parser.tailwind_font_sizes.get(class_name)
                if font_size:
                    style['font-size'] = font_size

        # 颜色类
        elif class_name.startswith('text-'):
            if hasattr(self.css_parser, 'tailwind_colors'):
                color = self.css_parser.tailwind_colors.get(class_name)
                if color:
                    style['color'] = color

        # 网格列类
        elif class_name.startswith('grid-cols-'):
            if hasattr(self.css_parser, 'tailwind_grid_columns'):
                columns = self.css_parser.tailwind_grid_columns.get(class_name)
                if columns:
                    style['grid-template-columns'] = f"repeat({columns}, 1fr)"

        # 间距类
        elif class_name.startswith('gap-'):
            if hasattr(self.css_parser, 'tailwind_spacing'):
                gap = self.css_parser.tailwind_spacing.get(class_name)
                if gap:
                    style['gap'] = gap

        # 文本对齐类
        elif class_name in ['text-left', 'text-center', 'text-right', 'text-justify']:
            alignment = class_name.replace('text-', '')
            style['text-align'] = alignment

        # 字体粗细类
        elif class_name in ['font-light', 'font-normal', 'font-medium', 'font-semibold', 'font-bold']:
            weight_map = {
                'font-light': '300',
                'font-normal': '400',
                'font-medium': '500',
                'font-semibold': '600',
                'font-bold': '700'
            }
            if class_name in weight_map:
                style['font-weight'] = weight_map[class_name]

        return style if style else None

    def _parse_inline_style(self, style_str: str) -> Dict[str, str]:
        """
        解析内联样式

        Args:
            style_str: 内联样式字符串

        Returns:
            样式字典
        """
        style_dict = {}

        for declaration in style_str.split(';'):
            if ':' not in declaration:
                continue

            property_name, property_value = declaration.split(':', 1)
            style_dict[property_name.strip().lower()] = property_value.strip()

        return style_dict

    def clear_cache(self):
        """清除缓存"""
        self._style_cache.clear()
        self.font_size_extractor.clear_cache()
        logger.debug("样式计算器缓存已清除")

    def get_cache_stats(self) -> Dict[str, int]:
        """获取缓存统计信息"""
        return {
            'style_cache_size': len(self._style_cache),
            'font_size_cache_size': len(self.font_size_extractor._font_size_cache),
        }


# 全局样式计算器实例（延迟初始化）
_style_computer_instance = None


def get_style_computer(css_parser=None) -> StyleComputer:
    """
    获取全局样式计算器实例

    Args:
        css_parser: CSS解析器（首次调用时必须提供）

    Returns:
        StyleComputer实例
    """
    global _style_computer_instance

    if _style_computer_instance is None:
        if css_parser is None:
            raise ValueError("首次调用时必须提供css_parser参数")
        _style_computer_instance = StyleComputer(css_parser)
    elif css_parser is not None:
        # 更新CSS解析器
        _style_computer_instance.css_parser = css_parser
        _style_computer_instance.font_size_extractor.css_parser = css_parser

    return _style_computer_instance