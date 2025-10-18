"""
字体大小提取器
通用的HTML元素字体大小提取系统
支持CSS选择器、继承、单位转换等复杂场景
"""

import re
from typing import Optional, Dict, List, Set
from bs4 import BeautifulSoup, Tag
from pptx.util import Pt

from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class FontSizeExtractor:
    """
    字体大小提取器

    功能：
    1. 从CSS样式表中提取字体大小
    2. 处理CSS选择器匹配
    3. 实现字体大小继承
    4. 单位转换(px/pt/em/rem → pptx Pt)
    5. 处理内联样式和样式优先级
    """

    # 默认字体大小（浏览器默认值）
    DEFAULT_FONT_SIZE = 16  # px

    # 单位转换比例 (统一转换为px)
    UNIT_CONVERSION = {
        'px': 1.0,          # px直接使用
        'pt': 0.75,         # 1pt = 0.75px (96 DPI标准)
        'em': 16.0,         # 1em = 16px (基于默认字体大小)
        'rem': 16.0,        # 1rem = 16px (基于根字体大小)
        '%': 0.16,          # 1% = 0.16px (基于默认16px)
    }

    def __init__(self, css_parser):
        """
        初始化字体大小提取器

        Args:
            css_parser: CSS解析器实例
        """
        self.css_parser = css_parser
        self._font_size_cache = {}  # 字体大小缓存
        self._html_file_id = None  # HTML文件标识，用于缓存键

    def set_html_file_id(self, html_file_path: str):
        """
        设置HTML文件标识，用于缓存键

        Args:
            html_file_path: HTML文件路径
        """
        import os
        self._html_file_id = os.path.basename(html_file_path)
        # 清理缓存，因为文件标识已改变
        self.clear_cache()

    def extract_font_size(self, element: Tag, parent_font_size: int = None) -> Optional[int]:
        """
        提取元素的最终字体大小（px）

        处理顺序：
        1. 内联样式优先级最高
        2. CSS类选择器
        3. CSS标签选择器
        4. 继承父元素
        5. 浏览器默认值

        Args:
            element: BeautifulSoup元素
            parent_font_size: 父元素字体大小(px)

        Returns:
            字体大小(px)，如果无法确定则返回None
        """
        if not element:
            return None

        # 生成缓存键，包含HTML文件标识以避免跨文件冲突
        element_id = id(element)
        file_id = self._html_file_id or 'unknown'
        cache_key = f"{file_id}_{element_id}_{parent_font_size}"

        if cache_key in self._font_size_cache:
            return self._font_size_cache[cache_key]

        font_size_px = None

        # 1. 检查内联样式
        inline_style = element.get('style', '')
        if inline_style:
            font_size_px = self._extract_from_inline_style(inline_style, parent_font_size)

        # 2. 检查CSS选择器
        if font_size_px is None:
            font_size_px = self._extract_from_css_selectors(element, parent_font_size)

        # 3. 继承父元素
        if font_size_px is None and parent_font_size:
            font_size_px = parent_font_size

        # 4. 使用默认值
        if font_size_px is None:
            font_size_px = self.DEFAULT_FONT_SIZE

        # 直接缓存px值，不转换为Pt
        self._font_size_cache[cache_key] = font_size_px

        # 获取元素信息用于调试
        element_info = element.name
        if element.get('class'):
            element_info += f".{'.'.join(element.get('class', []))}"
        text_preview = element.get_text(strip=True)[:20]

        logger.debug(f"FontSizeExtractor: 元素 {element_info} 字体大小: {font_size_px}px (文本: {text_preview})")
        return font_size_px

    def _extract_from_inline_style(self, style_str: str, parent_font_size: int = None) -> Optional[int]:
        """
        从内联样式提取字体大小

        Args:
            style_str: 内联样式字符串
            parent_font_size: 父元素字体大小(px)

        Returns:
            字体大小(px)，如果未找到则返回None
        """
        # 解析内联样式
        style_dict = self._parse_inline_style(style_str)
        font_size_str = style_dict.get('font-size')

        if not font_size_str:
            return None

        return self._parse_font_size_value(font_size_str, parent_font_size)

    def _extract_from_css_selectors(self, element: Tag, parent_font_size: int = None) -> Optional[int]:
        """
        从CSS选择器提取字体大小

        按CSS优先级顺序：
        1. ID选择器 (#id)
        2. 类选择器 (.class1.class2)
        3. 标签选择器 (tag)
        4. Tailwind CSS类
        5. 通用选择器 (*)

        Args:
            element: BeautifulSoup元素
            parent_font_size: 父元素字体大小(px)

        Returns:
            字体大小(px)，如果未找到则返回None
        """
        # 按优先级收集选择器
        selectors = self._generate_selectors(element)

        for selector in selectors:
            font_size_str = self.css_parser.get_font_size(selector)
            if font_size_str:
                font_size_px = self._parse_font_size_value(font_size_str, parent_font_size)
                if font_size_px:
                    logger.debug(f"从CSS选择器 {selector} 提取字体大小: {font_size_str} → {font_size_px}px")
                    return font_size_px

        # 检查Tailwind CSS字体大小类
        classes = element.get('class', [])
        for cls in classes:
            if cls.startswith('text-'):
                px_size = self.get_tailwind_font_size(cls)
                if px_size:
                    logger.debug(f"从Tailwind类 {cls} 提取字体大小: {px_size}px")
                    return px_size

        return None

    def _generate_selectors(self, element: Tag) -> List[str]:
        """
        生成CSS选择器列表（按优先级排序）

        Args:
            element: BeautifulSoup元素

        Returns:
            选择器列表
        """
        selectors = []

        # 1. ID选择器（优先级最高）
        element_id = element.get('id')
        if element_id:
            selectors.append(f"#{element_id}")

        # 2. 类选择器（复合类选择器）
        classes = element.get('class', [])
        if classes:
            # 单个类选择器
            for cls in classes:
                selectors.append(f".{cls}")
            # 复合类选择器
            if len(classes) > 1:
                selectors.append(f".{'.'.join(classes)}")

        # 3. 标签选择器
        tag_name = element.name
        if tag_name:
            selectors.append(tag_name)

        # 4. 标签+类选择器
        if tag_name and classes:
            selectors.append(f"{tag_name}.{'.'.join(classes)}")

        return selectors

    def _parse_font_size_value(self, font_size_str: str, parent_font_size: int = None) -> Optional[int]:
        """
        解析字体大小值字符串

        支持格式：
        - "18px"
        - "1.2em"
        - "120%"
        - "1rem"
        - "14pt"

        Args:
            font_size_str: 字体大小字符串
            parent_font_size: 父元素字体大小(px)，用于相对单位计算

        Returns:
            字体大小(px)，如果解析失败则返回None
        """
        if not font_size_str:
            return None

        # 移除空白字符
        font_size_str = font_size_str.strip()

        # 正则表达式匹配数字和单位
        match = re.match(r'^([\d.]+)\s*(px|pt|em|rem|%)?$', font_size_str.lower())
        if not match:
            logger.warning(f"无法解析字体大小: {font_size_str}")
            return None

        value = float(match.group(1))
        unit = match.group(2) or 'px'  # 默认单位为px

        # 单位转换
        if unit == 'em' and parent_font_size:
            # em相对于父元素
            return int(value * parent_font_size)
        elif unit == 'rem':
            # rem相对于根元素（假设为16px）
            return int(value * self.DEFAULT_FONT_SIZE)
        elif unit == '%':
            # 百分比相对于父元素
            base_size = parent_font_size or self.DEFAULT_FONT_SIZE
            return int(value * base_size / 100)
        else:
            # px, pt等绝对单位
            conversion_factor = self.UNIT_CONVERSION.get(unit, 1.0)
            return int(value * conversion_factor)

    def _parse_inline_style(self, style_str: str) -> Dict[str, str]:
        """
        解析内联样式字符串

        Args:
            style_str: 内联样式字符串

        Returns:
            样式字典
        """
        style_dict = {}

        # 分割样式属性
        for declaration in style_str.split(';'):
            if ':' not in declaration:
                continue

            property_name, property_value = declaration.split(':', 1)
            style_dict[property_name.strip().lower()] = property_value.strip()

        return style_dict

    def _px_to_pt(self, px_size: int) -> int:
        """
        将px转换为Pt

        使用标准转换：1px = 0.75pt (96 DPI)
        确保字体大小转换的一致性

        Args:
            px_size: 像素大小

        Returns:
            Pt大小
        """
        pt_size = max(1, int(px_size * 0.75))
        logger.debug(f"字体大小转换: {px_size}px → {pt_size}Pt")
        return pt_size

    def get_tailwind_font_size(self, class_name: str) -> Optional[int]:
        """
        获取Tailwind CSS类的字体大小

        支持的Tailwind CSS字体大小类：
        - text-xs: 12px
        - text-sm: 14px
        - text-base: 16px
        - text-lg: 18px
        - text-xl: 20px
        - text-2xl: 24px
        - text-3xl: 30px
        - text-4xl: 36px
        - text-5xl: 48px

        Args:
            class_name: Tailwind CSS类名

        Returns:
            字体大小(px)，如果未找到则返回None
        """
        tailwind_sizes = {
            'text-xs': 12,     # 12px
            'text-sm': 14,     # 14px
            'text-base': 16,   # 16px
            'text-lg': 18,     # 18px
            'text-xl': 20,     # 20px
            'text-2xl': 24,    # 24px
            'text-3xl': 30,    # 30px
            'text-4xl': 36,    # 36px
            'text-5xl': 48,    # 48px
            'text-6xl': 60,    # 60px
            'text-7xl': 72,    # 72px
            'text-8xl': 96,    # 96px
            'text-9xl': 128,   # 128px
        }

        px_size = tailwind_sizes.get(class_name)
        if px_size:
            logger.debug(f"Tailwind类 {class_name}: {px_size}px")
            return px_size

        return None

    def clear_cache(self):
        """清除缓存"""
        self._font_size_cache.clear()
        logger.debug("字体大小缓存已清除")

    def get_cache_stats(self) -> Dict[str, int]:
        """获取缓存统计信息"""
        return {
            'cache_size': len(self._font_size_cache),
            'default_font_size_px': self.DEFAULT_FONT_SIZE,
        }


# 全局字体大小提取器实例（延迟初始化）
_font_size_extractor_instance = None


def get_font_size_extractor(css_parser=None) -> FontSizeExtractor:
    """
    获取全局字体大小提取器实例

    Args:
        css_parser: CSS解析器（首次调用时必须提供）

    Returns:
        FontSizeExtractor实例
    """
    global _font_size_extractor_instance

    if _font_size_extractor_instance is None:
        if css_parser is None:
            raise ValueError("首次调用时必须提供css_parser参数")
        _font_size_extractor_instance = FontSizeExtractor(css_parser)
    elif css_parser is not None:
        # 更新CSS解析器
        _font_size_extractor_instance.css_parser = css_parser

    return _font_size_extractor_instance