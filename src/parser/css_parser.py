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

        # 重要修复：从整个HTML文档解析样式，而不是只从slide中
        # 如果传入的是slide-container，需要找到完整的soup对象
        if soup.name == 'div' and soup.get('class') and 'slide-container' in soup.get('class', []):
            # 这是一个slide-container，需要找到根soup
            root_soup = soup.find_parent('html') or soup.find_parent('body') or soup
            # 如果还是找不到，使用原始soup（可能已经是完整的HTML）
            if root_soup and root_soup != soup:
                self._parse_styles_from_soup(root_soup)
            else:
                # 尝试从原始HTML文件重新解析
                self._parse_styles_from_soup(soup)
        else:
            # 假设传入的是完整的HTML soup或包含head的元素
            self._parse_styles_from_soup(soup)

        # 初始化Tailwind CSS映射
        self._init_tailwind_mappings()

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

    def _parse_styles_from_soup(self, soup: BeautifulSoup):
        """从指定的soup对象解析样式"""
        # 先在当前soup中查找style标签
        style_tags = soup.find_all('style')

        # 如果在当前soup中没有找到，且传入的是slide-container
        # 尝试通过其他方式获取完整的HTML内容
        if not style_tags and soup.name == 'div' and soup.get('class') and 'slide-container' in soup.get('class', []):
            # 这是一个slide-container，但没有找到style标签
            # 可能需要重新解析原始HTML文件
            # 但由于我们没有文件路径，这里暂时记录警告
            logger.warning("在slide-container中未找到style标签，可能无法解析内部样式")

        for style_tag in style_tags:
            css_text = style_tag.string
            if not css_text:
                continue

            # 简单的CSS规则提取(不使用cssutils以避免依赖问题)
            self._extract_rules(css_text)

        logger.info(f"解析了 {len(self.style_rules)} 条CSS规则")

    def _init_tailwind_mappings(self):
        """初始化Tailwind CSS字体大小和颜色映射"""
        self.tailwind_font_sizes = {
            'text-xs': '12px',
            'text-sm': '14px',
            'text-base': '16px',
            'text-lg': '18px',
            'text-xl': '20px',
            'text-2xl': '24px',
            'text-3xl': '30px',
            'text-4xl': '36px',
            'text-5xl': '48px',
            'text-6xl': '60px',
        }

        # 新增Tailwind CSS颜色映射
        self.tailwind_colors = {
            # 红色系
            'text-red-50': '#fef2f2',
            'text-red-100': '#fee2e2',
            'text-red-200': '#fecaca',
            'text-red-300': '#fca5a5',
            'text-red-400': '#f87171',
            'text-red-500': '#ef4444',
            'text-red-600': '#dc2626',
            'text-red-700': '#b91c1c',
            'text-red-800': '#991b1b',
            'text-red-900': '#7f1d1d',

            # 绿色系
            'text-green-50': '#f0fdf4',
            'text-green-100': '#dcfce7',
            'text-green-200': '#bbf7d0',
            'text-green-300': '#86efac',
            'text-green-400': '#4ade80',
            'text-green-500': '#22c55e',
            'text-green-600': '#16a34a',
            'text-green-700': '#15803d',
            'text-green-800': '#166534',
            'text-green-900': '#14532d',

            # 蓝色系
            'text-blue-50': '#eff6ff',
            'text-blue-100': '#dbeafe',
            'text-blue-200': '#bfdbfe',
            'text-blue-300': '#93c5fd',
            'text-blue-400': '#60a5fa',
            'text-blue-500': '#3b82f6',
            'text-blue-600': '#2563eb',
            'text-blue-700': '#1d4ed8',
            'text-blue-800': '#1e40af',
            'text-blue-900': '#1e3a8a',

            # 灰色系
            'text-gray-50': '#f9fafb',
            'text-gray-100': '#f3f4f6',
            'text-gray-200': '#e5e7eb',
            'text-gray-300': '#d1d5db',
            'text-gray-400': '#9ca3af',
            'text-gray-500': '#6b7280',
            'text-gray-600': '#4b5563',
            'text-gray-700': '#374151',
            'text-gray-800': '#1f2937',
            'text-gray-900': '#111827',

            # 橙色系
            'text-orange-50': '#fff7ed',
            'text-orange-100': '#ffedd5',
            'text-orange-200': '#fed7aa',
            'text-orange-300': '#fdba74',
            'text-orange-400': '#fb923c',
            'text-orange-500': '#f97316',
            'text-orange-600': '#ea580c',
            'text-orange-700': '#c2410c',
            'text-orange-800': '#9a3412',
            'text-orange-900': '#7c2d12',

            # 紫色系
            'text-purple-50': '#faf5ff',
            'text-purple-100': '#f3e8ff',
            'text-purple-200': '#e9d5ff',
            'text-purple-300': '#d8b4fe',
            'text-purple-400': '#c084fc',
            'text-purple-500': '#a855f7',
            'text-purple-600': '#9333ea',
            'text-purple-700': '#7c3aed',
            'text-purple-800': '#6b21a8',
            'text-purple-900': '#581c87',
        }

        # 网格布局映射
        self.tailwind_grid_columns = {
            'grid-cols-1': 1,
            'grid-cols-2': 2,
            'grid-cols-3': 3,
            'grid-cols-4': 4,
            'grid-cols-5': 5,
            'grid-cols-6': 6,
            'grid-cols-7': 7,
            'grid-cols-8': 8,
            'grid-cols-9': 9,
            'grid-cols-10': 10,
            'grid-cols-11': 11,
            'grid-cols-12': 12,
        }

        # 间距映射
        self.tailwind_spacing = {
            'gap-1': '0.25rem',  # 4px
            'gap-2': '0.5rem',   # 8px
            'gap-3': '0.75rem',  # 12px
            'gap-4': '1rem',     # 16px
            'gap-5': '1.25rem',  # 20px
            'gap-6': '1.5rem',   # 24px
            'gap-8': '2rem',     # 32px
            'gap-10': '2.5rem',  # 40px
            'gap-12': '3rem',    # 48px
            'gap-16': '4rem',    # 64px
            'gap-20': '5rem',    # 80px
        }

        logger.debug(f"初始化 {len(self.tailwind_font_sizes)} 个Tailwind字体大小映射")
        logger.debug(f"初始化 {len(self.tailwind_colors)} 个Tailwind颜色映射")
        logger.debug(f"初始化 {len(self.tailwind_grid_columns)} 个Tailwind网格列映射")
        logger.debug(f"初始化 {len(self.tailwind_spacing)} 个Tailwind间距映射")

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

        支持的选择器类型：
        - 标签选择器: h1, p, div
        - 类选择器: .class-name
        - ID选择器: #id-name
        - 复合选择器: div.class-name
        - 多类选择器: .class1.class2
        - Tailwind CSS类: text-xl, text-lg等

        Args:
            selector: 选择器

        Returns:
            字体大小字符串(如'48px')
        """
        # 直接匹配
        style = self.get_style(selector)
        if style and 'font-size' in style:
            return style.get('font-size')

        # 检查Tailwind CSS字体大小类
        if selector.startswith('.text-'):
            class_name = selector[1:]  # 移除点号
            if class_name in self.tailwind_font_sizes:
                font_size = self.tailwind_font_sizes[class_name]
                logger.debug(f"Tailwind CSS类 {selector}: {font_size}")
                return font_size

        # 如果没有直接匹配，尝试模糊匹配
        return self._match_font_size_fallback(selector)

    def _match_font_size_fallback(self, selector: str) -> Optional[str]:
        """
        模糊匹配字体大小

        当直接匹配失败时，尝试更宽松的匹配规则

        Args:
            selector: 选择器

        Returns:
            字体大小字符串
        """
        # 1. 如果是复合选择器，尝试各个部分
        if '.' in selector and '#' not in selector:
            parts = selector.split('.')
            # 尝试主选择器 + 单个类
            base = parts[0]
            for i in range(1, len(parts)):
                partial_selector = f"{base}.{'.'.join(parts[1:i+1])}"
                style = self.get_style(partial_selector)
                if style and 'font-size' in style:
                    return style.get('font-size')

            # 尝试单个类
            for i in range(1, len(parts)):
                class_selector = f".{parts[i]}"
                style = self.get_style(class_selector)
                if style and 'font-size' in style:
                    return style.get('font-size')

        # 2. 如果是标签选择器，没有匹配，尝试通用标签
        if selector in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'div', 'span']:
            # 检查是否有通用标签样式
            for tag_style in self.style_rules:
                if tag_style == selector and 'font-size' in self.style_rules[tag_style]:
                    return self.style_rules[tag_style]['font-size']

        # 3. 如果是类选择器，检查是否有相关的类匹配
        if selector.startswith('.'):
            class_name = selector[1:]
            # 寻找包含此类的复合选择器
            for css_selector, style in self.style_rules.items():
                if 'font-size' in style and class_name in css_selector:
                    # 检查是否包含此类（避免部分匹配）
                    class_parts = css_selector.split('.')
                    if class_name in class_parts:
                        return style['font-size']

        return None

    def list_font_size_rules(self) -> Dict[str, str]:
        """
        列出所有包含字体大小的CSS规则

        Returns:
            字体大小规则字典 {选择器: 字体大小}
        """
        font_size_rules = {}
        for selector, style in self.style_rules.items():
            if 'font-size' in style:
                font_size_rules[selector] = style['font-size']
        return font_size_rules

    def get_color(self, selector: str) -> Optional[str]:
        """
        获取颜色

        Args:
            selector: 选择器

        Returns:
            颜色字符串
        """
        # 首先检查CSS样式
        style = self.get_style(selector)
        if style and 'color' in style:
            return style.get('color')

        # 检查Tailwind CSS颜色类
        if selector.startswith('.text-'):
            class_name = selector[1:]  # 移除点号
            if class_name in self.tailwind_colors:
                color = self.tailwind_colors[class_name]
                logger.debug(f"Tailwind CSS颜色类 {selector}: {color}")
                return color

        return None

    def get_grid_columns(self, selector: str) -> int:
        """
        从grid-template-columns提取列数或从Tailwind CSS类获取列数

        Args:
            selector: CSS选择器

        Returns:
            列数，默认4列
        """
        # 首先检查CSS样式中的grid-template-columns
        style = self.get_style(selector)
        if style:
            grid_template = style.get('grid-template-columns', '')
            if grid_template:
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

        # 检查Tailwind CSS网格列类
        if selector.startswith('.grid-cols-'):
            class_name = selector[1:]  # 移除点号
            if class_name in self.tailwind_grid_columns:
                columns = self.tailwind_grid_columns[class_name]
                logger.debug(f"Tailwind CSS网格列类 {selector}: {columns}列")
                return columns

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

    def get_font_family(self, selector: str) -> Optional[str]:
        """
        获取字体族

        Args:
            selector: 选择器

        Returns:
            字体族字符串（可能包含多个字体，用逗号分隔）
        """
        style = self.get_style(selector)
        return style.get('font-family') if style else None

    def get_gap_size(self, selector: str) -> int:
        """
        获取gap间距大小（返回像素值）

        Args:
            selector: CSS选择器

        Returns:
            间距大小（像素），默认20px
        """
        # 首先检查CSS样式
        style = self.get_style(selector)
        if style:
            gap = style.get('gap', '')
            if gap:
                # 简单转换rem/px到px
                if gap.endswith('px'):
                    return int(gap[:-2])
                elif gap.endswith('rem'):
                    rem_value = float(gap[:-3])
                    return int(rem_value * 16)  # 1rem = 16px

        # 检查Tailwind CSS间距类
        if selector.startswith('.gap-'):
            class_name = selector[1:]  # 移除点号
            if class_name in self.tailwind_spacing:
                gap_value = self.tailwind_spacing[class_name]
                # 转换rem到px
                if gap_value.endswith('rem'):
                    rem_value = float(gap_value[:-3])
                    px_value = int(rem_value * 16)
                    logger.debug(f"Tailwind CSS间距类 {selector}: {px_value}px")
                    return px_value

        return 20  # 默认20px

    def parse_element_classes(self, element) -> Dict[str, str]:
        """
        解析元素的所有CSS类，返回样式信息

        Args:
            element: BeautifulSoup元素对象或字典

        Returns:
            样式信息字典，包含字体大小、颜色等
        """
        styles = {}

        # 处理不同类型的输入
        if hasattr(element, 'get'):
            classes = element.get('class', [])
        elif isinstance(element, dict):
            classes = element.get('class', [])
        else:
            return styles

        for class_name in classes:
            # 解析字体大小
            if class_name in self.tailwind_font_sizes:
                font_size = self.tailwind_font_sizes[class_name]
                styles['font-size'] = font_size

            # 解析颜色
            elif class_name in self.tailwind_colors:
                color = self.tailwind_colors[class_name]
                styles['color'] = color

            # 解析网格列数
            elif class_name in self.tailwind_grid_columns:
                columns = self.tailwind_grid_columns[class_name]
                styles['grid-template-columns'] = f"repeat({columns}, 1fr)"

            # 解析间距
            elif class_name in self.tailwind_spacing:
                gap_value = self.tailwind_spacing[class_name]
                styles['gap'] = gap_value

            # 解析文本对齐
            elif class_name in ['text-left', 'text-center', 'text-right', 'text-justify']:
                alignment = class_name.replace('text-', '')
                styles['text-align'] = alignment

            # 解析字体粗细
            elif class_name in ['font-light', 'font-normal', 'font-medium', 'font-semibold', 'font-bold']:
                weight_map = {
                    'font-light': '300',
                    'font-normal': '400',
                    'font-medium': '500',
                    'font-semibold': '600',
                    'font-bold': '700'
                }
                if class_name in weight_map:
                    styles['font-weight'] = weight_map[class_name]

        return styles

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
