"""
HTML转PPTX主程序
"""

import sys
import re
from pathlib import Path

from src.parser.html_parser import HTMLParser
from src.parser.css_parser import CSSParser
from src.renderer.pptx_builder import PPTXBuilder
from src.converters.text_converter import TextConverter
from src.converters.table_converter import TableConverter
from src.converters.shape_converter import ShapeConverter
from src.converters.chart_converter import ChartConverter
from src.converters.timeline_converter import TimelineConverter
from src.utils.logger import setup_logger
from src.utils.unit_converter import UnitConverter
from src.utils.color_parser import ColorParser
from src.utils.chart_capture import ChartCapture
from src.utils.font_manager import get_font_manager
from src.utils.style_computer import get_style_computer
from pptx.util import Pt
from pptx.enum.text import MSO_ANCHOR

logger = setup_logger(__name__)


class HTML2PPTX:
    """HTML转PPTX转换器"""

    def __init__(self, html_path: str):
        """
        初始化转换器

        Args:
            html_path: HTML文件路径
        """
        self.html_path = html_path
        self.html_parser = HTMLParser(html_path)
        self.css_parser = CSSParser(self.html_parser.soup)
        self.pptx_builder = PPTXBuilder()

        # 初始化全局字体管理器和样式计算器
        self.font_manager = get_font_manager(self.css_parser)
        self.style_computer = get_style_computer(self.css_parser)

    def convert(self, output_path: str):
        """
        执行转换

        Args:
            output_path: 输出PPTX路径
        """
        logger.info("=" * 50)
        logger.info("开始HTML转PPTX转换")
        logger.info("=" * 50)

        # 获取所有幻灯片
        slides = self.html_parser.get_slides()

        for slide_html in slides:
            logger.info(f"\n处理幻灯片...")

            # 创建空白幻灯片
            pptx_slide = self.pptx_builder.add_blank_slide()

            # 初始化转换器
            text_converter = TextConverter(pptx_slide, self.css_parser)
            table_converter = TableConverter(pptx_slide, self.css_parser)
            shape_converter = ShapeConverter(pptx_slide, self.css_parser)

            # 1. 添加顶部装饰条
            shape_converter.add_top_bar()

            # 2. 添加标题和副标题
            title = self.html_parser.get_title(slide_html)
            subtitle = self.html_parser.get_subtitle(slide_html)
            if title:
                # content-section的padding-top是20px
                title_end_y = text_converter.convert_title(title, subtitle, x=80, y=20)
                # space-y-10的第一个子元素紧接标题区域（无上间距）
                y_offset = title_end_y
            else:
                # 没有标题时使用默认位置（content-section padding-top）
                y_offset = 20

            # 3. 处理内容区域

            # 统一处理所有容器：查找space-y-10容器，按顺序处理其子元素
            space_y_container = slide_html.find('div', class_='space-y-10')
            if space_y_container:
                # 按顺序遍历直接子元素
                is_first_container = True
                for container in space_y_container.find_all(recursive=False):
                    if not container.name:
                        continue

                    container_classes = container.get('class', [])

                    # space-y-10: 第一个元素无上间距，后续元素有40px间距
                    if not is_first_container:
                        y_offset += 40  # space-y-10间距
                    is_first_container = False

                    # 根据class路由到对应的处理方法
                    if 'stats-container' in container_classes:
                        # 顶层stats-container（不在stat-card内）
                        y_offset = self._convert_stats_container(
                            container, pptx_slide, y_offset
                        )
                    elif 'stat-card' in container_classes:
                        y_offset = self._convert_stat_card(
                            container, pptx_slide, y_offset
                        )
                    elif 'data-card' in container_classes:
                        y_offset = self._convert_data_card(
                            container, pptx_slide, shape_converter, y_offset
                        )
                    elif 'strategy-card' in container_classes:
                        y_offset = self._convert_strategy_card(
                            container, pptx_slide, y_offset
                        )
                    else:
                        # 未知容器类型，记录警告
                        logger.warning(f"遇到未知容器类型: {container_classes}")
            else:
                # 降级处理：如果没有space-y-10，使用旧逻辑
                logger.warning("未找到space-y-10容器，使用降级处理")

                # 处理统计卡片 (.stat-card)
                stat_cards = self.html_parser.get_stat_cards(slide_html)
                for card in stat_cards:
                    y_offset = self._convert_stat_card(card, pptx_slide, y_offset)

                # 处理数据卡片 (.data-card)
                data_cards = self.html_parser.get_data_cards(slide_html)
                for card in data_cards:
                    y_offset = self._convert_data_card(
                        card, pptx_slide, shape_converter, y_offset
                    )

                # 处理策略卡片 (.strategy-card)
                strategy_cards = self.html_parser.get_strategy_cards(slide_html)
                for card in strategy_cards:
                    y_offset = self._convert_strategy_card(
                        card, pptx_slide, y_offset
                    )

            # 4. 添加页码
            page_num = self.html_parser.get_page_number(slide_html)
            if page_num:
                shape_converter.add_page_number(page_num)

        # 保存PPTX
        self.pptx_builder.save(output_path)

        logger.info("=" * 50)
        logger.info(f"转换完成! 输出: {output_path}")
        logger.info("=" * 50)

    def _convert_stats_container(self, container, pptx_slide, y_start: int) -> int:
        """
        转换统计卡片容器 (.stats-container)

        Returns:
            下一个元素的Y坐标
        """
        stat_boxes = container.find_all('div', class_='stat-box')
        num_boxes = len(stat_boxes)

        if num_boxes == 0:
            return y_start

        # 动态获取列数：优先从inline style，其次从CSS规则
        num_columns = 4  # 默认值

        # 1. 检查inline style属性
        inline_style = container.get('style', '')
        if 'grid-template-columns' in inline_style:
            # 解析inline style中的grid-template-columns
            import re
            repeat_match = re.search(r'repeat\((\d+),', inline_style)
            if repeat_match:
                num_columns = int(repeat_match.group(1))
                logger.info(f"从inline style检测到列数: {num_columns}列")
            else:
                fr_count = len(re.findall(r'1fr', inline_style))
                if fr_count > 0:
                    num_columns = fr_count
                    logger.info(f"从inline style检测到列数: {num_columns}列")
        else:
            # 2. 从CSS规则获取
            num_columns = self.css_parser.get_grid_columns('.stats-container')
            logger.info(f"从CSS规则检测到列数: {num_columns}列")

        # 根据列数动态计算box宽度
        # 总宽度 = 1920 - 2*80(左右边距) = 1760
        # box_width = (1760 - (num_columns-1) * gap) / num_columns
        gap = 20
        total_width = 1760
        box_width = int((total_width - (num_columns - 1) * gap) / num_columns)
        box_height = 220
        x_start = 80

        logger.info(f"计算box尺寸: 宽度={box_width}px, 高度={box_height}px, 间距={gap}px")

        for idx, box in enumerate(stat_boxes):
            col = idx % num_columns
            row = idx // num_columns

            x = x_start + col * (box_width + gap)
            y = y_start + row * (box_height + gap)

            # 添加背景
            shape_converter = ShapeConverter(pptx_slide, self.css_parser)
            shape_converter.add_stat_box_background(x, y, box_width, box_height)

            # 提取内容
            icon = box.find('i')
            title_elem = box.find('div', class_='stat-title')
            h2 = box.find('h2')
            p = box.find('p')

            # 添加图标(使用文本替代，顶部居中)
            current_y = y + 20
            if icon:
                icon_classes = icon.get('class', [])
                icon_char = self._get_icon_char(icon_classes)

                # 图标居中
                icon_left = UnitConverter.px_to_emu(x + box_width // 2 - 25)
                icon_top = UnitConverter.px_to_emu(current_y)
                icon_box = pptx_slide.shapes.add_textbox(
                    icon_left, icon_top,
                    UnitConverter.px_to_emu(50), UnitConverter.px_to_emu(40)
                )
                icon_frame = icon_box.text_frame
                icon_frame.text = icon_char
                icon_frame.vertical_anchor = 1  # 居中
                for paragraph in icon_frame.paragraphs:
                    paragraph.alignment = 2  # PP_ALIGN.CENTER
                    for run in paragraph.runs:
                        # 图标字体大小应该是父元素的1.5倍，默认36pt
                        run.font.size = Pt(36)
                        run.font.color.rgb = ColorParser.get_primary_color()

                current_y += 45

            # 添加标题（居中）
            if title_elem:
                title_text = title_elem.get_text(strip=True)
                title_left = UnitConverter.px_to_emu(x + 10)
                title_top = UnitConverter.px_to_emu(current_y)
                title_box = pptx_slide.shapes.add_textbox(
                    title_left, title_top,
                    UnitConverter.px_to_emu(box_width - 20), UnitConverter.px_to_emu(30)
                )
                title_frame = title_box.text_frame
                title_frame.text = title_text
                title_frame.word_wrap = True
                for paragraph in title_frame.paragraphs:
                    paragraph.alignment = 2  # PP_ALIGN.CENTER
                    for run in paragraph.runs:
                        # 使用样式计算器获取正确的字体大小
                        title_font_size_pt = self.style_computer.get_font_size_pt(title_elem)
                        run.font.size = Pt(title_font_size_pt)
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = self.font_manager.get_font('body')

                current_y += 30

            # 添加主数据（居中）
            if h2:
                h2_text = h2.get_text(strip=True)
                h2_left = UnitConverter.px_to_emu(x + 10)
                h2_top = UnitConverter.px_to_emu(current_y)
                h2_box = pptx_slide.shapes.add_textbox(
                    h2_left, h2_top,
                    UnitConverter.px_to_emu(box_width - 20), UnitConverter.px_to_emu(50)
                )
                h2_frame = h2_box.text_frame
                h2_frame.text = h2_text
                for paragraph in h2_frame.paragraphs:
                    paragraph.alignment = 2  # PP_ALIGN.CENTER
                    for run in paragraph.runs:
                        # 使用样式计算器获取正确的字体大小
                        h2_font_size_pt = self.style_computer.get_font_size_pt(h2)
                        run.font.size = Pt(h2_font_size_pt)
                        run.font.bold = True
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = self.font_manager.get_font('body')

                current_y += 50

            # 添加描述（居中）
            if p:
                p_text = p.get_text(strip=True)
                p_left = UnitConverter.px_to_emu(x + 10)
                p_top = UnitConverter.px_to_emu(current_y)
                p_box = pptx_slide.shapes.add_textbox(
                    p_left, p_top,
                    UnitConverter.px_to_emu(box_width - 20), UnitConverter.px_to_emu(30)
                )
                p_frame = p_box.text_frame
                p_frame.text = p_text
                p_frame.word_wrap = True
                for paragraph in p_frame.paragraphs:
                    paragraph.alignment = 2  # PP_ALIGN.CENTER
                    for run in paragraph.runs:
                        # 使用样式计算器获取正确的字体大小
                        p_font_size_pt = self.style_computer.get_font_size_pt(p)
                        run.font.size = Pt(p_font_size_pt)
                        run.font.name = self.font_manager.get_font('body')

        # 计算下一个元素的Y坐标
        # 注意：这里计算的是所有stat-box渲染完毕后的Y坐标
        # 每一行占用：box_height + gap（除了最后一行没有gap）
        # 正确公式：y_start + num_rows * box_height + (num_rows - 1) * gap
        num_rows = (num_boxes + num_columns - 1) // num_columns
        actual_height = num_rows * box_height + (num_rows - 1) * gap

        logger.info(f"stats-container高度计算: 行数={num_rows}, box高度={box_height}px, gap={gap}px, 总高度={actual_height}px")

        return y_start + actual_height

    def _convert_stat_card(self, card, pptx_slide, y_start: int) -> int:
        """转换统计卡片(.stat-card) - 支持多种内部结构"""
        # 1. 检查是否包含stats-container (stat-box容器类型)
        stats_container = card.find('div', class_='stats-container')
        if stats_container:
            logger.info("stat-card包含stats-container,处理嵌套的stat-box结构")

            # 估算stat-card高度用于添加背景
            stat_boxes = stats_container.find_all('div', class_='stat-box')
            num_boxes = len(stat_boxes)

            # 动态获取列数
            num_columns = 4
            inline_style = stats_container.get('style', '')
            if 'grid-template-columns' in inline_style:
                import re
                repeat_match = re.search(r'repeat\((\d+),', inline_style)
                if repeat_match:
                    num_columns = int(repeat_match.group(1))
                else:
                    fr_count = len(re.findall(r'1fr', inline_style))
                    if fr_count > 0:
                        num_columns = fr_count
            else:
                num_columns = self.css_parser.get_grid_columns('.stats-container')

            # 从CSS读取约束
            stat_card_padding_top = 20
            stat_card_padding_bottom = 20
            stats_container_gap = 20
            stat_box_height = 220  # TODO阶段2: 改为动态计算

            # 计算stats-container的实际高度
            num_rows = (num_boxes + num_columns - 1) // num_columns
            stats_container_height = num_rows * stat_box_height + (num_rows - 1) * stats_container_gap

            # 计算stat-card总高度（包括自身padding）
            # stat-card = padding-top + (可选标题35px) + stats-container + padding-bottom
            has_title = card.find('p', class_='primary-color') is not None
            title_height = 35 if has_title else 0

            card_height = stat_card_padding_top + title_height + stats_container_height + stat_card_padding_bottom

            logger.info(f"stat-card高度计算: padding={stat_card_padding_top+stat_card_padding_bottom}px, "
                       f"标题={title_height}px, stats-container={stats_container_height}px, 总高度={card_height}px")

            # 添加stat-card背景
            bg_color_str = self.css_parser.get_background_color('.stat-card')
            if bg_color_str:
                from pptx.enum.shapes import MSO_SHAPE
                bg_shape = pptx_slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    UnitConverter.px_to_emu(80),
                    UnitConverter.px_to_emu(y_start),
                    UnitConverter.px_to_emu(1760),
                    UnitConverter.px_to_emu(card_height)
                )
                bg_shape.fill.solid()
                bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
                if bg_rgb:
                    if alpha < 1.0:
                        bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                    bg_shape.fill.fore_color.rgb = bg_rgb
                bg_shape.line.fill.background()
                bg_shape.shadow.inherit = False
                logger.info(f"添加stat-card背景色: {bg_color_str}, 高度={card_height}px")

            y_start += 15  # 顶部padding

            # 添加标题(如果有)
            p_elem = card.find('p', class_='primary-color', recursive=False)
            if not p_elem:
                # 尝试在第一层查找
                for child in card.children:
                    if hasattr(child, 'get') and 'primary-color' in child.get('class', []):
                        p_elem = child
                        break

            if p_elem and p_elem.name == 'p':
                text = p_elem.get_text(strip=True)
                if text:
                    text_left = UnitConverter.px_to_emu(95)
                    text_top = UnitConverter.px_to_emu(y_start)
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(1730), UnitConverter.px_to_emu(30)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            # 使用样式计算器获取正确的字体大小
                            p_font_size_pt = self.style_computer.get_font_size_pt(p_elem)
                            run.font.size = Pt(p_font_size_pt)
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                    y_start += 35

            # 处理嵌套的stats-container
            next_y = self._convert_stats_container(stats_container, pptx_slide, y_start)
            return next_y + 20  # 底部padding

        # 2. 检查是否包含timeline (时间线类型)
        timeline = card.find('div', class_='timeline')
        if timeline:
            logger.info("stat-card包含timeline,处理时间线结构")

            # 计算stat-card总高度（用于背景）
            timeline_items = timeline.find_all('div', class_='timeline-item')
            num_items = len(timeline_items)
            # 每个timeline-item约85px, 加上标题35px和padding 30px
            card_height = num_items * 85 + 65

            # 添加stat-card背景
            shape_converter = ShapeConverter(pptx_slide, self.css_parser)
            # 从CSS获取背景颜色
            bg_color = self.css_parser.get_background_color('.stat-card')
            if bg_color:
                # 添加带颜色的背景矩形
                from pptx.enum.shapes import MSO_SHAPE
                bg_shape = pptx_slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    UnitConverter.px_to_emu(80),
                    UnitConverter.px_to_emu(y_start),
                    UnitConverter.px_to_emu(1760),
                    UnitConverter.px_to_emu(card_height)
                )
                bg_shape.fill.solid()
                # 解析背景颜色（支持rgba透明度）
                bg_rgb, alpha = ColorParser.parse_rgba(bg_color)
                if bg_rgb:
                    # 如果有透明度，与白色混合
                    if alpha < 1.0:
                        bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                        logger.info(f"混合透明度: alpha={alpha}, 结果=RGB({bg_rgb[0]}, {bg_rgb[1]}, {bg_rgb[2]})")
                    bg_shape.fill.fore_color.rgb = bg_rgb
                bg_shape.line.fill.background()  # 无边框
                # 移除阴影效果
                bg_shape.shadow.inherit = False
                logger.info(f"添加stat-card背景色: {bg_color}, 高度={card_height}px")

            y_start += 15  # 顶部padding

            # 添加标题(如果有)
            p_elem = card.find('p', class_='primary-color')
            if p_elem:
                text = p_elem.get_text(strip=True)
                if text:
                    text_left = UnitConverter.px_to_emu(95)  # 左侧padding
                    text_top = UnitConverter.px_to_emu(y_start)
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(1730), UnitConverter.px_to_emu(30)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            # 使用样式计算器获取正确的字体大小
                            title_font_size_pt = self.style_computer.get_font_size_pt(p_elem)
                            run.font.size = Pt(title_font_size_pt)
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                    y_start += 35

            # 处理timeline
            timeline_converter = TimelineConverter(pptx_slide, self.css_parser)
            next_y = timeline_converter.convert_timeline(timeline, x=95, y=y_start, width=1730)

            return next_y + 35  # 时间线后留一些间距

        # 3. 检查是否包含canvas (图表类型)
        canvas = card.find('canvas')
        if canvas:
            logger.info("stat-card包含canvas,处理图表")

            # 从CSS读取约束
            stat_card_padding_top = 20
            stat_card_padding_bottom = 20

            # 标题高度
            has_title = card.find('p', class_='primary-color') is not None
            title_height = 35 if has_title else 0

            # canvas高度（固定220px，这是convert_chart传入的height）
            canvas_height = 220

            # stat-card总高度
            card_height = stat_card_padding_top + title_height + canvas_height + stat_card_padding_bottom

            logger.info(f"stat-card(canvas)高度计算: padding={stat_card_padding_top+stat_card_padding_bottom}px, "
                       f"标题={title_height}px, canvas={canvas_height}px, 总高度={card_height}px")

            bg_color_str = self.css_parser.get_background_color('.stat-card')
            if bg_color_str:
                from pptx.enum.shapes import MSO_SHAPE
                bg_shape = pptx_slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    UnitConverter.px_to_emu(80),
                    UnitConverter.px_to_emu(y_start),
                    UnitConverter.px_to_emu(1760),
                    UnitConverter.px_to_emu(card_height)
                )
                bg_shape.fill.solid()
                bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
                if bg_rgb:
                    if alpha < 1.0:
                        bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                    bg_shape.fill.fore_color.rgb = bg_rgb
                bg_shape.line.fill.background()
                bg_shape.shadow.inherit = False
                logger.info(f"添加stat-card背景色: {bg_color_str}")

            y_start += 15  # 顶部padding

            # 添加标题文本(如果有)
            p_elem = card.find('p', class_='primary-color')
            if p_elem:
                text = p_elem.get_text(strip=True)
                if text:
                    text_left = UnitConverter.px_to_emu(95)
                    text_top = UnitConverter.px_to_emu(y_start)
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(1730), UnitConverter.px_to_emu(30)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            # 使用样式计算器获取正确的字体大小
                            title_font_size_pt = self.style_computer.get_font_size_pt(p_elem)
                            run.font.size = Pt(title_font_size_pt)
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                    y_start += 35

            # 处理canvas图表
            chart_converter = ChartConverter(pptx_slide, self.css_parser, self.html_path)
            success = chart_converter.convert_chart(
                canvas,
                x=95,
                y=y_start,
                width=1730,
                height=220,
                use_screenshot=ChartCapture.is_available()
            )

            if not success:
                logger.warning("图表转换失败,已显示占位文本")

            return y_start + 240

        # 4. 通用降级处理 - 提取所有文本内容
        logger.info("stat-card不包含已知结构,使用通用文本提取")
        return self._convert_generic_card(card, pptx_slide, y_start, card_type='stat-card')

    def _convert_generic_card(self, card, pptx_slide, y_start: int, card_type: str = 'card') -> int:
        """
        通用卡片内容转换 - 降级处理未知结构

        提取所有文本内容，按段落渲染，保持基本样式

        Args:
            card: 卡片元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标
            card_type: 卡片类型（用于样式区分）

        Returns:
            下一个元素的Y坐标
        """
        logger.info(f"使用通用渲染器处理{card_type}")

        x_base = 80
        current_y = y_start

        # 提取所有段落元素 (p, h1, h2, h3, div等)
        text_elements = []

        # 查找所有文本容器
        for elem in card.descendants:
            if hasattr(elem, 'name') and elem.name in ['p', 'h1', 'h2', 'h3', 'h4', 'div', 'span']:
                # 只提取没有子块级元素的文本节点
                if not elem.find_all(['div', 'p', 'h1', 'h2', 'h3']):
                    text = elem.get_text(strip=True)
                    if text and len(text) > 2:  # 过滤空文本和单字符
                        # 检查是否有特殊样式
                        classes = elem.get('class', [])
                        is_primary = 'primary-color' in classes
                        is_bold = 'font-bold' in classes or elem.name in ['h1', 'h2', 'h3', 'h4']

                        text_elements.append({
                            'text': text,
                            'tag': elem.name,
                            'is_primary': is_primary,
                            'is_bold': is_bold
                        })

        # 去重（避免嵌套元素重复提取）
        seen_texts = set()
        unique_elements = []
        for elem in text_elements:
            if elem['text'] not in seen_texts:
                seen_texts.add(elem['text'])
                unique_elements.append(elem)

        logger.info(f"提取了 {len(unique_elements)} 个文本段落")

        # 添加背景和边框（根据容器类型）
        shape_converter = ShapeConverter(pptx_slide, self.css_parser)

        # 预估内容高度
        estimated_height = min(len(unique_elements) * 40 + 40, 280)

        if 'stat-card' in card_type:
            # stat-card有背景色（圆角矩形）
            bg_color = self.css_parser.get_background_color('.stat-card')
            if bg_color:
                from pptx.enum.shapes import MSO_SHAPE
                bg_shape = pptx_slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    UnitConverter.px_to_emu(x_base),
                    UnitConverter.px_to_emu(current_y),
                    UnitConverter.px_to_emu(1760),
                    UnitConverter.px_to_emu(estimated_height)
                )
                bg_shape.fill.solid()
                bg_rgb, alpha = ColorParser.parse_rgba(bg_color)
                if bg_rgb:
                    if alpha < 1.0:
                        bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                    bg_shape.fill.fore_color.rgb = bg_rgb
                bg_shape.line.fill.background()
                bg_shape.shadow.inherit = False  # 无阴影
            current_y += 15  # 顶部padding

        elif 'data-card' in card_type:
            # data-card有左边框
            shape_converter.add_border_left(x_base, current_y, estimated_height, 4)
            current_y += 10

        elif 'stat-box' in card_type:
            # stat-box有背景色
            bg_color = self.css_parser.get_background_color('.stat-box')
            if bg_color:
                from pptx.enum.shapes import MSO_SHAPE
                bg_shape = pptx_slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    UnitConverter.px_to_emu(x_base),
                    UnitConverter.px_to_emu(current_y),
                    UnitConverter.px_to_emu(1760),
                    UnitConverter.px_to_emu(estimated_height)
                )
                bg_shape.fill.solid()
                bg_rgb, alpha = ColorParser.parse_rgba(bg_color)
                if bg_rgb:
                    if alpha < 1.0:
                        bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                    bg_shape.fill.fore_color.rgb = bg_rgb
                bg_shape.line.fill.background()
                bg_shape.shadow.inherit = False  # 无阴影
            current_y += 15

        elif 'strategy-card' in card_type:
            # strategy-card有背景色和左边框
            bg_color = self.css_parser.get_background_color('.strategy-card')
            if bg_color:
                from pptx.enum.shapes import MSO_SHAPE
                bg_shape = pptx_slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    UnitConverter.px_to_emu(x_base),
                    UnitConverter.px_to_emu(current_y),
                    UnitConverter.px_to_emu(1760),
                    UnitConverter.px_to_emu(estimated_height)
                )
                bg_shape.fill.solid()
                bg_rgb, alpha = ColorParser.parse_rgba(bg_color)
                if bg_rgb:
                    if alpha < 1.0:
                        bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                    bg_shape.fill.fore_color.rgb = bg_rgb
                bg_shape.line.fill.background()
                bg_shape.shadow.inherit = False  # 无阴影
            shape_converter.add_border_left(x_base, current_y, estimated_height, 4)
            current_y += 10

        # 渲染文本（unique_elements已在前面提取）
        for elem in unique_elements[:10]:  # 最多渲染10个段落，避免过长
            text = elem['text']
            is_primary = elem['is_primary']
            is_bold = elem['is_bold']
            tag = elem['tag']

            # 根据标签确定字体大小
            if tag in ['h1', 'h2']:
                font_size = 24
            elif tag == 'h3':
                font_size = 20
            elif is_primary:
                font_size = 20
            else:
                font_size = 16

            # 计算文本高度（粗略估算）
            lines = (len(text) // 80) + 1
            text_height = max(30, lines * 25)

            text_left = UnitConverter.px_to_emu(x_base + 20)
            text_top = UnitConverter.px_to_emu(current_y)
            text_box = pptx_slide.shapes.add_textbox(
                text_left, text_top,
                UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(text_height)
            )
            text_frame = text_box.text_frame
            text_frame.text = text
            text_frame.word_wrap = True

            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
                    if is_bold:
                        run.font.bold = True
                    if is_primary:
                        run.font.color.rgb = ColorParser.get_primary_color()
                    run.font.name = self.font_manager.get_font('body')

            current_y += text_height + 10

        return current_y + 20

    def _convert_strategy_card(self, card, pptx_slide, y_start: int) -> int:
        """
        转换策略卡片(.strategy-card)

        处理action-item结构：圆形数字图标 + 标题 + 描述
        """
        logger.info("处理strategy-card")
        x_base = 80

        action_items = card.find_all('div', class_='action-item')

        # 从CSS读取约束
        strategy_card_padding = 10  # top + bottom = 20
        action_item_margin_bottom = 15  # CSS中的margin-bottom

        # 标题高度
        has_title = card.find('p', class_='primary-color') is not None
        title_height = 40 if has_title else 0

        # 每个action-item的高度组成：
        # - 圆形图标: 28px
        # - 标题(action-title): 18px字体 × 1.5 = 27px
        # - 描述(p): 16px字体 × 1.5 × 行数（估算2行）= 48px
        # - margin-bottom: 15px
        # 总计：28 + 27 + 48 + 15 = 118px

        # 简化估算（TODO阶段2：根据实际文本行数计算）
        single_action_item_height = 118

        # strategy-card总高度
        # = padding-top + title + (action-items × height) + padding-bottom
        card_height = (strategy_card_padding + title_height +
                       len(action_items) * single_action_item_height +
                       strategy_card_padding)

        # 限制在max-height范围内（CSS中max-height为300px）
        max_height = 300
        if card_height > max_height:
            logger.warning(f"strategy-card内容高度({card_height}px)超出max-height({max_height}px)")
            card_height = max_height

        logger.info(f"strategy-card高度计算: padding={strategy_card_padding*2}px, "
                   f"标题={title_height}px, action-items={len(action_items)}个×{single_action_item_height}px, "
                   f"总高度={card_height}px")

        bg_color_str = self.css_parser.get_background_color('.strategy-card')
        if bg_color_str:
            from pptx.enum.shapes import MSO_SHAPE
            bg_shape = pptx_slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                UnitConverter.px_to_emu(x_base),
                UnitConverter.px_to_emu(y_start),
                UnitConverter.px_to_emu(1760),
                UnitConverter.px_to_emu(card_height)
            )
            bg_shape.fill.solid()
            bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
            if bg_rgb:
                if alpha < 1.0:
                    bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                bg_shape.fill.fore_color.rgb = bg_rgb
            bg_shape.line.fill.background()
            bg_shape.shadow.inherit = False

        # 添加左边框
        shape_converter = ShapeConverter(pptx_slide, self.css_parser)
        shape_converter.add_border_left(x_base, y_start, card_height, 4)

        current_y = y_start + 15

        # 添加标题
        p_elem = card.find('p', class_='primary-color')
        if p_elem:
            text = p_elem.get_text(strip=True)
            if text:
                text_left = UnitConverter.px_to_emu(x_base + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用样式计算器获取正确的字体大小
                        title_font_size_pt = self.style_computer.get_font_size_pt(p_elem)
                        run.font.size = Pt(title_font_size_pt)
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = self.font_manager.get_font('body')

                current_y += 40

        # 处理每个action-item
        for item in action_items:
            # 获取数字
            number_elem = item.find('div', class_='action-number')
            number_text = number_elem.get_text(strip=True) if number_elem else "•"

            # 获取action-content
            content_elem = item.find('div', class_='action-content')
            if not content_elem:
                continue

            # 获取标题和描述
            title_elem = content_elem.find('div', class_='action-title')
            title_text = title_elem.get_text(strip=True) if title_elem else ""

            desc_elem = content_elem.find('p')
            desc_text = desc_elem.get_text(strip=True) if desc_elem else ""

            # 渲染圆形数字图标
            from pptx.enum.shapes import MSO_SHAPE
            from pptx.enum.text import MSO_ANCHOR

            circle_size = 28
            circle_left = UnitConverter.px_to_emu(x_base + 20)
            circle_top = UnitConverter.px_to_emu(current_y)
            circle = pptx_slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                circle_left,
                circle_top,
                UnitConverter.px_to_emu(circle_size),
                UnitConverter.px_to_emu(circle_size)
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = ColorParser.get_primary_color()
            circle.line.fill.background()

            # 在圆形内添加数字文本
            circle_text_frame = circle.text_frame
            circle_text_frame.text = number_text
            circle_text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            for paragraph in circle_text_frame.paragraphs:
                paragraph.alignment = 2  # PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(14)
                    run.font.color.rgb = ColorParser.parse_color('#FFFFFF')
                    run.font.name = self.font_manager.get_font('body')
                    run.font.bold = True

            # 渲染标题（右侧）
            title_left = UnitConverter.px_to_emu(x_base + 60)
            title_top = UnitConverter.px_to_emu(current_y)
            if title_text:
                title_box = pptx_slide.shapes.add_textbox(
                    title_left, title_top,
                    UnitConverter.px_to_emu(1680), UnitConverter.px_to_emu(25)
                )
                title_frame = title_box.text_frame
                title_frame.text = title_text
                title_frame.word_wrap = True
                for paragraph in title_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用样式计算器获取正确的字体大小
                        title_font_size_pt = self.style_computer.get_font_size_pt(title_elem)
                        run.font.size = Pt(title_font_size_pt)
                        run.font.bold = True
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = self.font_manager.get_font('body')

                current_y += 28

            # 渲染描述（缩进）
            if desc_text:
                desc_left = UnitConverter.px_to_emu(x_base + 60)
                desc_top = UnitConverter.px_to_emu(current_y)
                desc_box = pptx_slide.shapes.add_textbox(
                    desc_left, desc_top,
                    UnitConverter.px_to_emu(1680), UnitConverter.px_to_emu(40)
                )
                desc_frame = desc_box.text_frame
                desc_frame.text = desc_text
                desc_frame.word_wrap = True
                for paragraph in desc_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用样式计算器获取正确的字体大小
                        desc_font_size_pt = self.style_computer.get_font_size_pt(desc_elem)
                        run.font.size = Pt(desc_font_size_pt)
                        run.font.name = self.font_manager.get_font('body')

                current_y += 50
            else:
                current_y += 35

        return current_y + 20

    def _convert_data_card(self, card, pptx_slide, shape_converter, y_start: int) -> int:
        """转换数据卡片(.data-card)"""
        logger.info("处理data-card")
        x_base = 80

        # 注意：左边框的高度需要在计算完实际内容后再添加
        # 暂时记录起始位置，稍后添加边框

        # 初始化当前Y坐标
        current_y = y_start + 10

        # 标题
        p_elem = card.find('p', class_='primary-color')
        if p_elem:
            text = p_elem.get_text(strip=True)
            text_left = UnitConverter.px_to_emu(x_base + 20)
            text_top = UnitConverter.px_to_emu(current_y)
            text_box = pptx_slide.shapes.add_textbox(
                text_left, text_top,
                UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
            )
            text_frame = text_box.text_frame
            text_frame.text = text
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    # 使用样式计算器获取正确的字体大小
                    font_size_px = self.style_computer.get_font_size_pt(p_elem)
                    run.font.size = Pt(font_size_px)
                    run.font.color.rgb = ColorParser.get_primary_color()
                    run.font.name = self.font_manager.get_font('body')

            current_y += 40  # 标题后间距

        # 进度条
        progress_bars = card.find_all('div', class_='progress-container')
        progress_y = current_y
        for progress in progress_bars:
            label_div = progress.find('div', class_='progress-label')
            if label_div:
                spans = label_div.find_all('span')
                if len(spans) >= 2:
                    label_text = spans[0].get_text(strip=True)
                    percent_text = spans[1].get_text(strip=True)
                    percentage = UnitConverter.normalize_percentage(percent_text)

                    shape_converter.add_progress_bar(
                        label_text, percentage, x_base + 20, progress_y, 1700
                    )
                    progress_y += 60

        # 列表项
        bullet_points = card.find_all('div', class_='bullet-point')
        has_content = len(progress_bars) > 0 or len(bullet_points) > 0

        for bullet in bullet_points:
            # 检查是否有嵌套的div结构
            nested_div = bullet.find('div')
            if nested_div:
                # 处理嵌套结构: <div class="bullet-point"><i>...</i><div><p>...</p><p>...</p></div></div>
                all_p = nested_div.find_all('p')

                for idx, p in enumerate(all_p):
                    text = p.get_text(strip=True)
                    if not text:
                        continue

                    bullet_left = UnitConverter.px_to_emu(x_base + 20)
                    bullet_top = UnitConverter.px_to_emu(progress_y)

                    # 第一个p加图标,后续p缩进
                    if idx == 0:
                        prefix = "• "
                        bullet_width = 1720
                    else:
                        prefix = "  "
                        bullet_width = 1720

                    bullet_box = pptx_slide.shapes.add_textbox(
                        bullet_left, bullet_top,
                        UnitConverter.px_to_emu(bullet_width), UnitConverter.px_to_emu(50)
                    )
                    bullet_frame = bullet_box.text_frame
                    bullet_frame.text = f"{prefix}{text}"
                    bullet_frame.word_wrap = True

                    for paragraph in bullet_frame.paragraphs:
                        for run in paragraph.runs:
                            # 使用样式计算器获取正确的字体大小
                            font_size_px = self.style_computer.get_font_size_pt(p)
                            run.font.size = Pt(font_size_px)
                            # 第一个p加粗
                            if idx == 0:
                                run.font.bold = True
                            run.font.name = self.font_manager.get_font('body')

                    progress_y += 28 if idx == 0 else 50
            else:
                # 处理简单结构: <div class="bullet-point"><p>...</p></div>
                p = bullet.find('p')
                if p:
                    text = p.get_text(strip=True)
                    bullet_left = UnitConverter.px_to_emu(x_base + 20)
                    bullet_top = UnitConverter.px_to_emu(progress_y)
                    bullet_box = pptx_slide.shapes.add_textbox(
                        bullet_left, bullet_top,
                        UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
                    )
                    bullet_frame = bullet_box.text_frame
                    bullet_frame.text = f"• {text}"
                    for paragraph in bullet_frame.paragraphs:
                        for run in paragraph.runs:
                            # 使用样式计算器获取正确的字体大小
                            font_size_px = self.style_computer.get_font_size_pt(p)
                            run.font.size = Pt(font_size_px)
                            run.font.name = self.font_manager.get_font('body')

                    progress_y += 35

        # 如果没有识别到任何已知内容，使用通用降级处理
        if not has_content:
            logger.info("data-card不包含progress-bar或bullet-point,使用通用处理")
            return self._convert_generic_card(card, pptx_slide, y_start, card_type='data-card')

        # 计算实际高度
        final_y = progress_y + 20
        actual_height = final_y - y_start

        # 添加左边框（使用实际计算的高度）
        shape_converter.add_border_left(x_base, y_start, actual_height, 4)

        logger.info(f"data-card高度计算: 实际高度={actual_height}px, "
                   f"进度条数={len(progress_bars)}, 列表项数={len(bullet_points)}")

        return final_y

    def _get_icon_char(self, icon_classes: list) -> str:
        """根据FontAwesome类获取对应emoji/Unicode字符"""
        icon_map = {
            # 常用图标
            'fa-search': '🔍',
            'fa-bug': '🐛',
            'fa-check-circle': '✓',
            'fa-exclamation-triangle': '⚠',
            'fa-exclamation-circle': '⚠',
            # 安全相关
            'fa-shield': '🛡',
            'fa-shield-alt': '🛡',
            'fa-shield-virus': '🛡',  # 病毒防护
            'fa-lock': '🔒',
            'fa-unlock': '🔓',
            'fa-key': '🔑',
            # 电子设备
            'fa-laptop': '💻',
            'fa-server': '🖥',
            'fa-database': '🗄',
            'fa-cloud': '☁',
            'fa-mobile': '📱',
            # 网络
            'fa-wifi': '📶',
            'fa-signal': '📶',
            'fa-globe': '🌐',
            'fa-network-wired': '🔌',
            # 状态
            'fa-check': '✓',
            'fa-times': '✗',
            'fa-bolt': '⚡',  # 闪电
            'fa-fire': '🔥',
            'fa-star': '⭐',
            'fa-heart': '♥',
            # 任务
            'fa-tasks': '☑',  # 任务列表
            'fa-list': '📋',
            'fa-clipboard': '📋',
            'fa-calendar': '📅',
            # 文件
            'fa-file': '📄',
            'fa-folder': '📁',
            'fa-download': '⬇',
            'fa-upload': '⬆',
            # 用户
            'fa-user': '👤',
            'fa-users': '👥',
            'fa-user-shield': '🛡',
            # 设置
            'fa-cog': '⚙',
            'fa-wrench': '🔧',
            'fa-tools': '🛠',
            # 其他
            'fa-info-circle': 'ℹ',
            'fa-question-circle': '❓',
            'fa-plus': '+',
            'fa-minus': '-',
            'fa-arrow-right': '→',
            'fa-arrow-left': '←',
            'fa-home': '🏠',
            'fa-bell': '🔔',
            'fa-envelope': '✉',
            'fa-phone': '📞',
            'fa-chart-bar': '📊',
            'fa-chart-line': '📈',
            'fa-chart-pie': '📊',
        }

        for cls in icon_classes:
            if cls in icon_map:
                return icon_map[cls]

        # 如果找不到匹配，返回默认图标
        return '●'


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("用法: python main.py <html文件路径> [输出pptx路径]")
        sys.exit(1)

    html_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "output/output.pptx"

    # 执行转换
    converter = HTML2PPTX(html_path)
    converter.convert(output_path)


if __name__ == "__main__":
    main()
