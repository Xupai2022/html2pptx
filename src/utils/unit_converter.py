"""
单位转换工具
用于在px、pt、EMU等单位之间转换
"""

from pptx.util import Inches, Pt, Emu


class UnitConverter:
    """单位转换器"""

    # 常量定义
    DPI = 96  # 默认屏幕DPI
    EMU_PER_INCH = 914400  # 每英寸的EMU数
    EMU_PER_PT = 12700  # 每点的EMU数
    PT_PER_INCH = 72  # 每英寸的点数

    # 幻灯片尺寸 (1920x1080 px)
    SLIDE_WIDTH_PX = 1920
    SLIDE_HEIGHT_PX = 1080
    SLIDE_WIDTH_EMU = None  # 将在类初始化后计算
    SLIDE_HEIGHT_EMU = None

    @classmethod
    def px_to_emu(cls, px: float) -> int:
        """
        像素转EMU

        Args:
            px: 像素值

        Returns:
            EMU值
        """
        return int(px * cls.EMU_PER_INCH / cls.DPI)

    @classmethod
    def pt_to_emu(cls, pt: float) -> int:
        """
        点转EMU

        Args:
            pt: 点值

        Returns:
            EMU值
        """
        return int(pt * cls.EMU_PER_PT)

    @classmethod
    def emu_to_px(cls, emu: int) -> float:
        """
        EMU转像素

        Args:
            emu: EMU值

        Returns:
            像素值
        """
        return emu * cls.DPI / cls.EMU_PER_INCH

    @classmethod
    def emu_to_pt(cls, emu: int) -> float:
        """
        EMU转点

        Args:
            emu: EMU值

        Returns:
            点值
        """
        return emu / cls.EMU_PER_PT

    @classmethod
    def pt_to_px(cls, pt: float) -> float:
        """
        点转像素 (假设96 DPI)

        Args:
            pt: 点值

        Returns:
            像素值
        """
        return pt * cls.DPI / cls.PT_PER_INCH

    @classmethod
    def px_to_pt(cls, px: float) -> float:
        """
        像素转点 (假设96 DPI)

        Args:
            px: 像素值

        Returns:
            点值
        """
        return px * cls.PT_PER_INCH / cls.DPI

    @classmethod
    def font_size_px_to_pt(cls, px_size: int) -> int:
        """
        字体大小专用转换：像素转点

        专门用于HTML到PPTX的字体大小转换，确保转换精度和一致性

        Args:
            px_size: HTML中的字体大小(像素)

        Returns:
            PPTX中的字体大小(点)，确保最小值为1
        """
        if px_size <= 0:
            return 1

        # 使用标准转换：1px = 0.75pt (96 DPI标准)
        pt_size = px_size * cls.PT_PER_INCH / cls.DPI

        # 四舍五入并确保最小值为1
        return max(1, int(round(pt_size)))

    @classmethod
    def parse_html_font_size(cls, font_size_str: str, parent_size_px: int = None) -> int:
        """
        解析HTML字体大小字符串并转换为pt

        支持多种单位：px, pt, em, rem, %

        Args:
            font_size_str: HTML字体大小字符串(如"16px", "1.2em", "120%")
            parent_size_px: 父元素字体大小(px)，用于相对单位计算

        Returns:
            转换后的pt值
        """
        import re

        if not font_size_str:
            return cls.font_size_px_to_pt(16)  # 默认16px -> 12pt

        # 移除空白字符
        font_size_str = font_size_str.strip()

        # 正则表达式匹配数字和单位
        match = re.match(r'^([\d.]+)\s*(px|pt|em|rem|%)?$', font_size_str.lower())
        if not match:
            # 无法解析时返回默认值
            return cls.font_size_px_to_pt(16)

        value = float(match.group(1))
        unit = match.group(2) or 'px'  # 默认单位为px

        # 默认字体大小
        default_px = 16

        # 根据单位转换为px
        if unit == 'px':
            px_size = int(value)
        elif unit == 'pt':
            # pt已经是目标单位，直接返回
            return max(1, int(round(value)))
        elif unit == 'em':
            # em相对于父元素
            base_px = parent_size_px or default_px
            px_size = int(value * base_px)
        elif unit == 'rem':
            # rem相对于根元素
            px_size = int(value * default_px)
        elif unit == '%':
            # 百分比相对于父元素
            base_px = parent_size_px or default_px
            px_size = int(value * base_px / 100)
        else:
            # 未知单位，当作px处理
            px_size = int(value)

        # 转换为pt
        return cls.font_size_px_to_pt(px_size)

    @classmethod
    def get_slide_dimensions(cls):
        """
        获取幻灯片尺寸 (EMU)

        Returns:
            (width_emu, height_emu)
        """
        if cls.SLIDE_WIDTH_EMU is None:
            cls.SLIDE_WIDTH_EMU = cls.px_to_emu(cls.SLIDE_WIDTH_PX)
        if cls.SLIDE_HEIGHT_EMU is None:
            cls.SLIDE_HEIGHT_EMU = cls.px_to_emu(cls.SLIDE_HEIGHT_PX)

        return cls.SLIDE_WIDTH_EMU, cls.SLIDE_HEIGHT_EMU

    @classmethod
    def normalize_percentage(cls, percentage_str: str) -> float:
        """
        标准化百分比字符串为浮点数

        Args:
            percentage_str: 百分比字符串,如"92.7%"

        Returns:
            0-1之间的浮点数
        """
        if isinstance(percentage_str, (int, float)):
            return float(percentage_str)

        percentage_str = str(percentage_str).strip()
        if percentage_str.endswith('%'):
            return float(percentage_str[:-1]) / 100.0
        return float(percentage_str)


# 初始化幻灯片尺寸
UnitConverter.get_slide_dimensions()
