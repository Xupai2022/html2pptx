"""
颜色解析工具
支持rgb()、rgba()、十六进制等格式
"""

import re
from typing import Tuple, Optional
from pptx.dml.color import RGBColor


class ColorParser:
    """颜色解析器"""

    # 主题色定义
    PRIMARY_COLOR = RGBColor(10, 66, 117)  # rgb(10, 66, 117)
    TEXT_DEFAULT = RGBColor(51, 51, 51)  # #333
    WHITE = RGBColor(255, 255, 255)
    BLACK = RGBColor(0, 0, 0)

    @staticmethod
    def parse_color(color_str: str) -> Optional[RGBColor]:
        """
        解析颜色字符串

        支持格式:
        - rgb(10, 66, 117)
        - rgba(10, 66, 117, 0.8)
        - #0A4275
        - #333

        Args:
            color_str: 颜色字符串

        Returns:
            RGBColor对象,解析失败返回None
        """
        if not color_str:
            return None

        color_str = color_str.strip().lower()

        # rgb/rgba格式
        rgb_match = re.match(r'rgba?\((\d+),\s*(\d+),\s*(\d+)', color_str)
        if rgb_match:
            r, g, b = map(int, rgb_match.groups())
            return RGBColor(r, g, b)

        # 十六进制格式
        hex_match = re.match(r'#([0-9a-f]{3,6})', color_str)
        if hex_match:
            hex_color = hex_match.group(1)
            if len(hex_color) == 3:
                # #abc -> #aabbcc
                hex_color = ''.join([c * 2 for c in hex_color])
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            return RGBColor(r, g, b)

        # 命名颜色
        named_colors = {
            'white': ColorParser.WHITE,
            'black': ColorParser.BLACK,
        }
        return named_colors.get(color_str)

    @staticmethod
    def parse_rgba(color_str: str) -> Tuple[Optional[RGBColor], float]:
        """
        解析rgba颜色,返回颜色和透明度

        Args:
            color_str: rgba颜色字符串

        Returns:
            (RGBColor, alpha) 其中alpha为0-1之间的浮点数
        """
        if not color_str:
            return None, 1.0

        color_str = color_str.strip().lower()

        # rgba格式
        rgba_match = re.match(r'rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)', color_str)
        if rgba_match:
            r, g, b, a = rgba_match.groups()
            return RGBColor(int(r), int(g), int(b)), float(a)

        # 其他格式默认不透明
        color = ColorParser.parse_color(color_str)
        return color, 1.0

    @staticmethod
    def get_primary_color() -> RGBColor:
        """获取主题色"""
        return ColorParser.PRIMARY_COLOR

    @staticmethod
    def get_text_color() -> RGBColor:
        """获取默认文本颜色"""
        return ColorParser.TEXT_DEFAULT

    @staticmethod
    def blend_with_white(color: RGBColor, alpha: float) -> RGBColor:
        """
        将颜色与白色混合(模拟透明度效果)

        Args:
            color: 原始颜色
            alpha: 透明度 (0-1)

        Returns:
            混合后的RGBColor
        """
        r = int(color[0] * alpha + 255 * (1 - alpha))
        g = int(color[1] * alpha + 255 * (1 - alpha))
        b = int(color[2] * alpha + 255 * (1 - alpha))
        return RGBColor(r, g, b)
