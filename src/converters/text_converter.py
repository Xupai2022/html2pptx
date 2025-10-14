"""
文本转换器
处理H1, H2, P等文本元素
"""

from pptx.util import Pt
from src.converters.base_converter import BaseConverter
from src.mapper.style_mapper import StyleMapper
from src.utils.unit_converter import UnitConverter
from src.utils.color_parser import ColorParser
from src.utils.logger import setup_logger
from src.utils.font_manager import get_font_manager
from src.utils.style_computer import get_style_computer

logger = setup_logger(__name__)


class TextConverter(BaseConverter):
    """文本元素转换器"""

    def convert_title(self, title_text: str, subtitle_text: str = None, x: int = 80, y: int = 20) -> int:
        """
        转换标题和副标题

        Tailwind CSS v2.2.19 规则 (1rem = 16px):
        - mt-10 = 10 × 0.25rem = 2.5rem = 40px
        - mt-2 = 2 × 0.25rem = 0.5rem = 8px
        - mb-4 = 4 × 0.25rem = 1rem = 16px
        - h-1 = 1 × 0.25rem = 0.25rem = 4px
        - line-height: 1.5 (Tailwind默认)

        文本高度计算:
        - h1 (font-size: 48px): 48px × 1.5 = 72px
        - h2 (font-size: 36px): 36px × 1.5 = 54px

        布局计算 (从content-section padding-top开始):
        初始y = 20px (content-section padding-top)
        + mt-10: 40px → y = 60px
        + h1: 72px → y = 132px
        + mt-2: 8px → y = 140px
        + h2: 54px → y = 194px
        + line: 4px → y = 198px (装饰线紧接h2)
        + mb-4: 16px → y = 214px (装饰线的下边距)
        标题区域结束: y = 214px

        Args:
            title_text: 标题文本
            subtitle_text: 副标题文本
            x: X坐标(px)
            y: Y坐标(px) - content-section的padding-top起始位置，默认20px

        Returns:
            标题区域结束后的Y坐标(px)
        """
        # 初始y值应该是content-section的padding-top (20px)
        current_y = y

        # mt-10: 2.5rem = 40px
        current_y += 40

        # 获取样式计算器和字体
        style_computer = get_style_computer(self.css_parser)
        font_manager = get_font_manager(self.css_parser)
        font_name = font_manager.get_font('h1')

        # 创建临时h1元素来获取字体大小
        from bs4 import BeautifulSoup
        temp_soup = BeautifulSoup('<h1>' + title_text + '</h1>', 'html.parser')
        h1_element = temp_soup.h1

        # 获取h1的字体大小 (现在get_font_size_pt返回pt值)
        h1_font_size_pt = style_computer.get_font_size_pt(h1_element)
        # 转换回px用于高度计算
        h1_font_size_px = UnitConverter.pt_to_px(h1_font_size_pt)
        h1_height = int(h1_font_size_px * 1.5)  # 行高1.5

        logger.debug(f"H1标题字体大小: {h1_font_size_px}px → {h1_font_size_pt}pt, 高度: {h1_height}px")

        # 添加h1标题
        left = UnitConverter.px_to_emu(x)
        top = UnitConverter.px_to_emu(current_y)
        width = UnitConverter.px_to_emu(1760)
        height = UnitConverter.px_to_emu(h1_height)

        title_box = self.slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.text = title_text
        title_frame.word_wrap = True
        title_frame.margin_top = 0
        title_frame.margin_bottom = 0
        title_frame.margin_left = 0

        for paragraph in title_frame.paragraphs:
            for run in paragraph.runs:
                # 使用转换后的pt值设置字体大小
                run.font.size = Pt(h1_font_size_pt)
                run.font.bold = True
                run.font.name = font_name

        logger.info(f"添加标题: {title_text}")

        current_y += h1_height  # 72px

        # 添加副标题
        if subtitle_text:
            # mt-2: 0.5rem = 8px
            current_y += 8

            # 创建临时h2元素来获取字体大小
            temp_soup_h2 = BeautifulSoup('<h2>' + subtitle_text + '</h2>', 'html.parser')
            h2_element = temp_soup_h2.h2

            # 获取h2的字体大小 (现在get_font_size_pt返回pt值)
            h2_font_size_pt = style_computer.get_font_size_pt(h2_element)
            # 转换回px用于高度计算
            h2_font_size_px = UnitConverter.pt_to_px(h2_font_size_pt)
            h2_height = int(h2_font_size_px * 1.5)  # 行高1.5

            logger.debug(f"H2副标题字体大小: {h2_font_size_px}px → {h2_font_size_pt}pt, 高度: {h2_height}px")

            subtitle_top = UnitConverter.px_to_emu(current_y)
            subtitle_box = self.slide.shapes.add_textbox(
                left, subtitle_top, width, UnitConverter.px_to_emu(h2_height)
            )
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = subtitle_text
            subtitle_frame.word_wrap = True
            subtitle_frame.margin_top = 0
            subtitle_frame.margin_bottom = 0
            subtitle_frame.margin_left = 0

            font_name_h2 = font_manager.get_font('h2')

            for paragraph in subtitle_frame.paragraphs:
                for run in paragraph.runs:
                    # 使用转换后的pt值设置字体大小
                    run.font.size = Pt(h2_font_size_pt)
                    run.font.bold = True
                    run.font.color.rgb = ColorParser.get_primary_color()
                    run.font.name = font_name_h2

            logger.info(f"添加副标题: {subtitle_text}")

            current_y += h2_height  # 54px

            # 添加装饰线 (w-20 h-1) - 紧接h2，无间距
            line_top = UnitConverter.px_to_emu(current_y)
            line_left = UnitConverter.px_to_emu(x)
            line_width = UnitConverter.px_to_emu(80)
            # h-1: 0.25rem = 4px
            line_height = UnitConverter.px_to_emu(4)

            line_shape = self.slide.shapes.add_shape(
                1,  # Rectangle
                line_left, line_top, line_width, line_height
            )
            line_shape.fill.solid()
            line_shape.fill.fore_color.rgb = ColorParser.get_primary_color()
            line_shape.line.fill.background()

            current_y += 4  # 装饰线高度

            # mb-4: 1rem = 16px (装饰线的下边距)
            current_y += 16

        # 标题区域结束位置
        # 正确计算: y(20) + mt-10(40) + h1(72) + mt-2(8) + h2(54) + line(4) + mb-4(16) = 214px
        # 装饰线紧接h2(194px)，装饰线结束(198px)，装饰线mb-4下边距(16px) → 第一个容器起始(214px)
        logger.info(f"标题区域结束位置: y={current_y}px")
        return current_y

    def convert_paragraph(self, p_element, x: int, y: int, width: int = 1760):
        """
        转换段落

        Args:
            p_element: 段落HTML元素
            x: X坐标(px)
            y: Y坐标(px)
            width: 宽度(px)
        """
        text = p_element.get_text(strip=True)
        if not text:
            return

        # 获取样式计算器和字体
        style_computer = get_style_computer(self.css_parser)
        font_manager = get_font_manager(self.css_parser)

        # 获取段落的字体大小 (现在get_font_size_pt返回pt值)
        p_font_size_pt = style_computer.get_font_size_pt(p_element)
        # 转换回px用于高度计算
        p_font_size_px = UnitConverter.pt_to_px(p_font_size_pt)
        p_height = int(p_font_size_px * 1.5)  # 行高1.5

        logger.debug(f"段落字体大小: {p_font_size_px}px → {p_font_size_pt}pt, 高度: {p_height}px")

        left = UnitConverter.px_to_emu(x)
        top = UnitConverter.px_to_emu(y)
        w = UnitConverter.px_to_emu(width)
        h = UnitConverter.px_to_emu(p_height)

        text_box = self.slide.shapes.add_textbox(left, top, w, h)
        text_frame = text_box.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0

        # 应用样式
        p_style = style_computer.compute_computed_style(p_element)
        inline_style = self._extract_inline_style(p_element)
        font_name = font_manager.get_font('p', inline_style)

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                # 使用转换后的pt值设置字体大小
                run.font.size = Pt(p_font_size_pt)
                run.font.name = font_name

                # 颜色
                color_str = p_style.get('color') or inline_style.get('color')
                if color_str:
                    color = ColorParser.parse_color(color_str)
                    if color:
                        run.font.color.rgb = color

                # 字体粗细
                font_weight = p_style.get('font-weight')
                if font_weight:
                    run.font.bold = StyleMapper.parse_font_weight(font_weight)

    def _extract_inline_style(self, element) -> dict:
        """提取内联样式"""
        style_dict = {}
        style_str = element.get('style', '')

        if not style_str:
            # 检查class属性
            classes = element.get('class', [])
            if 'primary-color' in classes:
                style_dict['color'] = 'rgb(10, 66, 117)'

        return style_dict

    def convert(self, element, **kwargs):
        """转换文本元素"""
        # 注意：这个方法在main.py中用于单个元素的转换
        # 但是convert_title会创建新的文本框，可能导致重复文字
        # 为了避免重复，这里应该使用convert_paragraph方法处理所有文本元素

        tag_name = element.name
        if tag_name in ['h1', 'h2', 'h3', 'p']:
            # 统一使用段落转换方法，传入实际元素
            x = kwargs.get('x', 80)
            y = kwargs.get('y', 0)
            width = kwargs.get('width', 1760)
            self.convert_paragraph(element, x, y, width)
