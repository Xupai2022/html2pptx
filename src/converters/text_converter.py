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

logger = setup_logger(__name__)


class TextConverter(BaseConverter):
    """文本元素转换器"""

    def convert_title(self, title_text: str, subtitle_text: str = None, x: int = 80, y: int = 80):
        """
        转换标题和副标题

        Args:
            title_text: 标题文本
            subtitle_text: 副标题文本
            x: X坐标(px)
            y: Y坐标(px)
        """
        # 添加标题文本框
        left = UnitConverter.px_to_emu(x)
        top = UnitConverter.px_to_emu(y)
        width = UnitConverter.px_to_emu(1760)  # 1920 - 80*2
        height = UnitConverter.px_to_emu(60)

        # H1样式
        h1_style = self.css_parser.get_element_style('h1') or {}

        title_box = self.slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.text = title_text
        title_frame.word_wrap = True
        title_frame.margin_top = 0
        title_frame.margin_bottom = 0
        title_frame.margin_left = 0

        # 应用H1样式
        for paragraph in title_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(48)
                run.font.bold = True
                run.font.name = 'Microsoft YaHei'

        logger.info(f"添加标题: {title_text}")

        # 添加副标题
        if subtitle_text:
            subtitle_top = UnitConverter.px_to_emu(y + 70)
            subtitle_box = self.slide.shapes.add_textbox(
                left, subtitle_top, width, UnitConverter.px_to_emu(50)
            )
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = subtitle_text
            subtitle_frame.word_wrap = True
            subtitle_frame.margin_top = 0
            subtitle_frame.margin_bottom = 0
            subtitle_frame.margin_left = 0

            for paragraph in subtitle_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(36)
                    run.font.bold = True
                    run.font.color.rgb = ColorParser.get_primary_color()
                    run.font.name = 'Microsoft YaHei'

            logger.info(f"添加副标题: {subtitle_text}")

            # 添加装饰线
            line_top = UnitConverter.px_to_emu(y + 130)
            line_left = UnitConverter.px_to_emu(x)
            line_width = UnitConverter.px_to_emu(80)
            line_height = UnitConverter.px_to_emu(4)

            line_shape = self.slide.shapes.add_shape(
                1,  # Rectangle
                line_left, line_top, line_width, line_height
            )
            line_shape.fill.solid()
            line_shape.fill.fore_color.rgb = ColorParser.get_primary_color()
            line_shape.line.fill.background()

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

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(20)
                run.font.name = 'Microsoft YaHei'

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
