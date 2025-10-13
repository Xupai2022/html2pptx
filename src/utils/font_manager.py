"""
字体管理器
统一管理PPTX字体选择逻辑
"""

import re
from typing import Optional, List
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class FontManager:
    """
    字体管理器

    负责：
    1. 从CSS解析font-family
    2. 处理字体回退链
    3. 映射Web字体到Windows字体
    """

    # Windows系统常见中文字体映射
    FONT_MAPPING = {
        # 微软字体
        'Microsoft YaHei': 'Microsoft YaHei',
        '微软雅黑': 'Microsoft YaHei',
        'Microsoft JhengHei': 'Microsoft JhengHei',
        '微软正黑体': 'Microsoft JhengHei',

        # 苹果字体映射到Windows
        'PingFang SC': 'Microsoft YaHei',  # 苹方 → 微软雅黑
        '苹方-简': 'Microsoft YaHei',
        'Heiti SC': 'SimHei',  # 黑体 → 黑体
        '黑体-简': 'SimHei',

        # 通用字体族
        'sans-serif': 'Microsoft YaHei',  # 默认无衬线
        'serif': 'SimSun',  # 默认衬线
        'monospace': 'Consolas',  # 等宽字体

        # 其他中文字体
        'SimSun': 'SimSun',  # 宋体
        '宋体': 'SimSun',
        'SimHei': 'SimHei',  # 黑体
        '黑体': 'SimHei',
        'KaiTi': 'KaiTi',  # 楷体
        '楷体': 'KaiTi',
        'FangSong': 'FangSong',  # 仿宋
        '仿宋': 'FangSong',

        # 英文字体
        'Arial': 'Arial',
        'Calibri': 'Calibri',
        'Times New Roman': 'Times New Roman',
        'Verdana': 'Verdana',
    }

    # 默认字体规则（当HTML未指定字体时使用）
    DEFAULT_FONTS = {
        'h1': 'Source Han Sans CN Bold',      # 主标题：思源黑体 Bold
        'h2': 'Source Han Sans CN',            # 副标题：思源黑体
        'p': 'DengXian',                       # 正文：等线
        'body': 'DengXian',                    # body默认：等线
        'default': 'Microsoft YaHei',          # 最终回退
    }

    def __init__(self, css_parser=None):
        """
        初始化字体管理器

        Args:
            css_parser: CSS解析器实例
        """
        self.css_parser = css_parser
        self._cached_fonts = {}  # 缓存已解析的字体

    def get_font(self, selector: str = 'body', element_style: dict = None) -> str:
        """
        获取字体名称

        优先级：
        1. 元素的inline style中的font-family
        2. CSS选择器中的font-family
        3. body的font-family
        4. 默认字体

        Args:
            selector: CSS选择器（如'h1', '.stat-title'）
            element_style: 元素的inline style字典

        Returns:
            PPTX支持的字体名称
        """
        # 1. 检查缓存
        cache_key = f"{selector}_{str(element_style)}"
        if cache_key in self._cached_fonts:
            return self._cached_fonts[cache_key]

        # 2. 尝试从inline style获取
        if element_style and 'font-family' in element_style:
            font = self._parse_font_family(element_style['font-family'])
            if font:
                self._cached_fonts[cache_key] = font
                logger.debug(f"使用inline字体: {font} (选择器: {selector})")
                return font

        # 3. 尝试从CSS选择器获取
        if self.css_parser:
            css_font_family = self.css_parser.get_font_family(selector)
            if css_font_family:
                font = self._parse_font_family(css_font_family)
                if font:
                    self._cached_fonts[cache_key] = font
                    logger.debug(f"使用CSS字体: {font} (选择器: {selector})")
                    return font

            # 4. 尝试从body获取
            if selector != 'body':
                body_font_family = self.css_parser.get_font_family('body')
                if body_font_family:
                    font = self._parse_font_family(body_font_family)
                    if font:
                        self._cached_fonts[cache_key] = font
                        logger.debug(f"使用body字体: {font} (选择器: {selector})")
                        return font

        # 5. 使用默认字体规则
        default_font = self._get_default_font(selector)
        self._cached_fonts[cache_key] = default_font
        logger.info(f"使用默认字体规则: {selector} → {default_font}")
        return default_font

    def _get_default_font(self, selector: str) -> str:
        """
        获取默认字体（当HTML未指定时）

        Args:
            selector: CSS选择器

        Returns:
            默认字体名称
        """
        # 直接匹配
        if selector in self.DEFAULT_FONTS:
            return self.DEFAULT_FONTS[selector]

        # 类选择器尝试提取标签名 (.stat-title → 默认)
        # 最终回退
        return self.DEFAULT_FONTS['default']

    def _parse_font_family(self, font_family_str: str) -> Optional[str]:
        """
        解析font-family字符串，返回第一个可用字体

        处理格式：
        - "Microsoft YaHei", sans-serif
        - 'PingFang SC', 'Helvetica Neue', Arial
        - Microsoft YaHei

        Args:
            font_family_str: font-family字符串

        Returns:
            映射后的Windows字体名称，如果都不可用则返回None
        """
        if not font_family_str:
            return None

        # 分割字体列表（处理引号和逗号）
        # 移除换行和多余空格
        font_family_str = ' '.join(font_family_str.split())

        # 分割字体（支持引号包裹的字体名）
        fonts = self._split_font_family(font_family_str)

        # 尝试每个字体
        for font_name in fonts:
            # 清理字体名（移除引号和空格）
            font_name = font_name.strip().strip('"').strip("'").strip()

            if not font_name:
                continue

            # 查找映射
            mapped_font = self.FONT_MAPPING.get(font_name)
            if mapped_font:
                logger.info(f"字体映射: {font_name} → {mapped_font}")
                return mapped_font

            # 如果没有映射但看起来是Windows字体，直接使用
            if self._is_likely_windows_font(font_name):
                logger.info(f"使用未映射的字体: {font_name}")
                return font_name

        return None

    def _split_font_family(self, font_family_str: str) -> List[str]:
        """
        分割font-family字符串为字体列表

        支持：
        - "Font One", "Font Two", sans-serif
        - 'Font One', 'Font Two'
        - Font One, Font Two

        Args:
            font_family_str: font-family字符串

        Returns:
            字体名称列表
        """
        fonts = []
        current_font = []
        in_quotes = False
        quote_char = None

        for char in font_family_str:
            if char in ('"', "'"):
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
                    quote_char = None
            elif char == ',' and not in_quotes:
                font = ''.join(current_font).strip()
                if font:
                    fonts.append(font)
                current_font = []
                continue

            current_font.append(char)

        # 添加最后一个字体
        font = ''.join(current_font).strip()
        if font:
            fonts.append(font)

        return fonts

    def _is_likely_windows_font(self, font_name: str) -> bool:
        """
        判断是否可能是Windows字体

        启发式规则：
        - 包含常见Windows字体关键词
        - 不是通用字体族

        Args:
            font_name: 字体名称

        Returns:
            是否可能是Windows字体
        """
        # 通用字体族不算
        generic_families = ['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui']
        if font_name.lower() in generic_families:
            return False

        # 包含常见Windows字体关键词
        windows_keywords = ['Microsoft', 'Sim', 'Ming', 'Kai', 'Fang', 'Arial', 'Calibri', 'Times', 'Verdana']
        for keyword in windows_keywords:
            if keyword in font_name:
                return True

        # 如果包含中文字符，可能是中文字体
        if re.search(r'[\u4e00-\u9fff]', font_name):
            return True

        return False


# 全局字体管理器实例（延迟初始化）
_font_manager_instance = None


def get_font_manager(css_parser=None) -> FontManager:
    """
    获取全局字体管理器实例

    Args:
        css_parser: CSS解析器（首次调用时必须提供）

    Returns:
        FontManager实例
    """
    global _font_manager_instance

    if _font_manager_instance is None:
        _font_manager_instance = FontManager(css_parser)
    elif css_parser is not None:
        # 更新CSS解析器
        _font_manager_instance.css_parser = css_parser

    return _font_manager_instance
