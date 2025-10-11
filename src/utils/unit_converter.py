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
