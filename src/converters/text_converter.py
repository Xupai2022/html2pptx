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

        # 添加h1标题
        left = UnitConverter.px_to_emu(x)
        top = UnitConverter.px_to_emu(current_y)
        width = UnitConverter.px_to_emu(1760)
        # h1高度: 48px × 1.5 = 72px
        h1_height = int(48 * 1.5)
        height = UnitConverter.px_to_emu(h1_height)

        title_box = self.slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.text = title_text
        title_frame.word_wrap = True
        title_frame.margin_top = 0
        title_frame.margin_bottom = 0
        title_frame.margin_left = 0

        # 获取字体
        font_manager = get_font_manager(self.css_parser)
        font_name = font_manager.get_font('h1')

        for paragraph in title_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(48)
                run.font.bold = True
                run.font.name = font_name

        logger.info(f"添加标题: {title_text}")

        current_y += h1_height  # 72px

        # 添加副标题
        if subtitle_text:
            # mt-2: 0.5rem = 8px
            current_y += 8

            subtitle_top = UnitConverter.px_to_emu(current_y)
            # h2高度: 36px × 1.5 = 54px
            h2_height = int(36 * 1.5)
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
                    run.font.size = Pt(36)
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

        left = UnitConverter.px_to_emu(x)
        top = UnitConverter.px_to_emu(y)
        w = UnitConverter.px_to_emu(width)
        h = UnitConverter.px_to_emu(30)

        text_box = self.slide.shapes.add_textbox(left, top, w, h)
        text_frame = text_box.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0

        # 应用P样式
        p_style = self.css_parser.get_element_style('p') or {}
        inline_style = self._extract_inline_style(p_element)

        # 获取字体
        font_manager = get_font_manager(self.css_parser)
        font_name = font_manager.get_font('p', inline_style)

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(20)
                run.font.name = font_name

                # 颜色
                color_str = inline_style.get('color') or p_style.get('color')
                if color_str:
                    color = ColorParser.parse_color(color_str)
                    if color:
                        run.font.color.rgb = color

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
        tag_name = element.name

        if tag_name == 'h1':
            self.convert_title(element.get_text(strip=True), **kwargs)
        elif tag_name == 'p':
            self.convert_paragraph(element, **kwargs)
