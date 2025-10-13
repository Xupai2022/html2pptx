"""
CSS样式解析器
"""

import re
from typing import Dict, Optional
from bs4 import BeautifulSoup

from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class CSSParser:
    """CSS解析器"""

    def __init__(self, soup: BeautifulSoup):
        """
        初始化CSS解析器

        Args:
            soup: BeautifulSoup对象
        """
        self.soup = soup
        self.style_rules = {}
        self._parse_styles()

    def _parse_styles(self):
        """解析<style>标签中的CSS规则"""
        style_tags = self.soup.find_all('style')

        for style_tag in style_tags:
            css_text = style_tag.string
            if not css_text:
                continue

            # 简单的CSS规则提取(不使用cssutils以避免依赖问题)
            self._extract_rules(css_text)

        logger.info(f"解析了 {len(self.style_rules)} 条CSS规则")

    def _extract_rules(self, css_text: str):
        """
        提取CSS规则

        Args:
            css_text: CSS文本
        """
        # 移除注释
        css_text = re.sub(r'/\*.*?\*/', '', css_text, flags=re.DOTALL)

        # 提取规则: selector { property: value; }
        pattern = r'([^{]+)\{([^}]+)\}'
        matches = re.findall(pattern, css_text)

        for selector, properties in matches:
            selector = selector.strip()
            prop_dict = self._parse_properties(properties)
            self.style_rules[selector] = prop_dict

    def _parse_properties(self, properties: str) -> Dict[str, str]:
        """
        解析CSS属性

        Args:
            properties: CSS属性文本

        Returns:
            属性字典
        """
        prop_dict = {}
        items = properties.split(';')

        for item in items:
            if ':' not in item:
                continue

            key, value = item.split(':', 1)
            prop_dict[key.strip()] = value.strip()

        return prop_dict

    def get_style(self, selector: str) -> Optional[Dict[str, str]]:
        """
        获取指定选择器的样式

        Args:
            selector: CSS选择器

        Returns:
            样式字典
        """
        return self.style_rules.get(selector, {})

    def get_class_style(self, class_name: str) -> Optional[Dict[str, str]]:
        """
        获取class样式

        Args:
            class_name: 类名(不含点号)

        Returns:
            样式字典
        """
        return self.get_style(f'.{class_name}')

    def get_element_style(self, element_name: str) -> Optional[Dict[str, str]]:
        """
        获取元素样式

        Args:
            element_name: 元素名称(如'h1', 'p')

        Returns:
            样式字典
        """
        return self.get_style(element_name)

    def get_font_size(self, selector: str) -> Optional[str]:
        """
        获取字体大小

        Args:
            selector: 选择器

        Returns:
            字体大小字符串(如'48px')
        """
        style = self.get_style(selector)
        return style.get('font-size') if style else None

    def get_color(self, selector: str) -> Optional[str]:
        """
        获取颜色

        Args:
            selector: 选择器

        Returns:
            颜色字符串
        """
        style = self.get_style(selector)
        return style.get('color') if style else None

    def get_grid_columns(self, selector: str) -> int:
        """
        从grid-template-columns提取列数

        Args:
            selector: CSS选择器

        Returns:
            列数，默认4列
        """
        style = self.get_style(selector)
        if not style:
            return 4  # 默认4列

        grid_template = style.get('grid-template-columns', '')
        if not grid_template:
            return 4

        # 解析 "repeat(3, 1fr)" 格式
        repeat_match = re.match(r'repeat\((\d+),', grid_template)
        if repeat_match:
            return int(repeat_match.group(1))

        # 解析 "1fr 1fr 1fr" 格式
        fr_count = len(re.findall(r'1fr', grid_template))
        if fr_count > 0:
            return fr_count

        # 解析其他格式，计算空格分隔的项数
        items = [item.strip() for item in grid_template.split() if item.strip()]
        if items:
            return len(items)

        return 4  # 默认4列

    def get_background_color(self, selector: str) -> Optional[str]:
        """
        获取背景颜色

        Args:
            selector: 选择器

        Returns:
            背景颜色字符串
        """
        style = self.get_style(selector)
        return style.get('background-color') if style else None

    def merge_styles(self, *selectors) -> Dict[str, str]:
        """
        合并多个选择器的样式(后面的覆盖前面的)

        Args:
            *selectors: 选择器列表

        Returns:
            合并后的样式字典
        """
        merged = {}
        for selector in selectors:
            style = self.get_style(selector)
            if style:
                merged.update(style)
        return merged
