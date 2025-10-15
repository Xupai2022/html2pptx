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
from pptx.enum.text import MSO_ANCHOR, PP_PARAGRAPH_ALIGNMENT

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
                    elif 'flex' in container_classes and 'justify-between' in container_classes:
                        # 底部信息容器（包含bullet-point的flex布局）
                        y_offset = self._convert_bottom_info(
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
            # p标签将在下面统一处理

            # 智能判断布局方向：检查CSS的align-items设置
            layout_direction = self._determine_layout_direction(box)

            # 智能判断文字对齐方式
            text_alignment = self._determine_text_alignment(box)

            if layout_direction == 'horizontal':
                # 水平布局：图标在左，文字在右
                # 根据CSS样式计算间距：padding: 20px, icon margin-right: 20px
                icon_x = x + 20  # 左padding
                content_x = icon_x + 36 + 20  # icon_x + icon_width + margin-right
                content_width = box_width - 40 - 36 - 20  # box_width - 左padding - icon_width - margin-right

                # 添加图标（左侧）
                if icon:
                    icon_classes = icon.get('class', [])
                    icon_char = self._get_icon_char(icon_classes)

                    # 图标垂直居中（根据CSS font-size: 36px）
                    icon_height = 36
                    icon_top = y + (box_height - icon_height) // 2  # 垂直居中计算
                    icon_left = UnitConverter.px_to_emu(icon_x)
                    icon_top = UnitConverter.px_to_emu(icon_top)
                    icon_box = pptx_slide.shapes.add_textbox(
                        icon_left, icon_top,
                        UnitConverter.px_to_emu(36), UnitConverter.px_to_emu(icon_height)
                    )
                    icon_frame = icon_box.text_frame
                    icon_frame.text = icon_char
                    icon_frame.vertical_anchor = MSO_ANCHOR.MIDDLE  # 垂直居中
                    for paragraph in icon_frame.paragraphs:
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(36)
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                # 添加文字内容（右侧），也垂直居中
                content_height = 0
                if title_elem:
                    title_text = title_elem.get_text(strip=True)
                    title_font_size_pt = self.style_computer.get_font_size_pt(title_elem)
                    title_height = int(title_font_size_pt * 1.5)  # 估算标题高度
                    content_height += title_height + 5  # margin-bottom: 5px

                if h2:
                    h2_text = h2.get_text(strip=True)
                    h2_font_size_pt = self.style_computer.get_font_size_pt(h2)
                    h2_height = int(h2_font_size_pt * 1.5)  # 估算h2高度
                    content_height += h2_height + 5

                # 计算所有p标签的总高度（包括第一个p标签）
                all_p_tags = box.find_all('p')
                for p_tag in all_p_tags:
                    p_text = p_tag.get_text(strip=True)
                    if p_text:
                        p_font_size_pt = self.style_computer.get_font_size_pt(p_tag)
                        # 计算p标签的行数（估算每行80个字符）
                        p_lines = max(1, (len(p_text) + 79) // 80)
                        p_height = p_lines * int(p_font_size_pt * 1.5)
                        content_height += p_height + 5  # 5px间距

                # 垂直居中文字内容
                content_start_y = y + (box_height - content_height) // 2
                current_y = content_start_y

                # 添加标题
                if title_elem:
                    title_text = title_elem.get_text(strip=True)
                    title_left = UnitConverter.px_to_emu(content_x)
                    title_top = UnitConverter.px_to_emu(current_y)
                    title_font_size_pt = self.style_computer.get_font_size_pt(title_elem)
                    title_height = int(title_font_size_pt * 1.5)
                    title_box = pptx_slide.shapes.add_textbox(
                        title_left, title_top,
                        UnitConverter.px_to_emu(content_width), UnitConverter.px_to_emu(title_height)
                    )
                    title_frame = title_box.text_frame
                    title_frame.text = title_text
                    title_frame.word_wrap = True
                    title_frame.vertical_anchor = MSO_ANCHOR.TOP  # 顶部对齐，确保精确定位
                    for paragraph in title_frame.paragraphs:
                        paragraph.alignment = text_alignment
                        for run in paragraph.runs:
                            run.font.size = Pt(title_font_size_pt)
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                    current_y += int(self.style_computer.get_font_size_pt(title_elem) * 1.5) + 5

                # 添加主数据
                if h2:
                    h2_text = h2.get_text(strip=True)
                    h2_left = UnitConverter.px_to_emu(content_x)
                    h2_top = UnitConverter.px_to_emu(current_y)
                    h2_font_size_pt = self.style_computer.get_font_size_pt(h2)
                    h2_height = int(h2_font_size_pt * 1.5)
                    h2_box = pptx_slide.shapes.add_textbox(
                        h2_left, h2_top,
                        UnitConverter.px_to_emu(content_width), UnitConverter.px_to_emu(h2_height)
                    )
                    h2_frame = h2_box.text_frame
                    h2_frame.text = h2_text
                    h2_frame.vertical_anchor = MSO_ANCHOR.TOP  # 顶部对齐
                    for paragraph in h2_frame.paragraphs:
                        paragraph.alignment = text_alignment
                        for run in paragraph.runs:
                            run.font.size = Pt(h2_font_size_pt)
                            run.font.bold = True
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                    current_y += int(self.style_computer.get_font_size_pt(h2) * 1.5) + 5

                # 添加描述（统一处理所有p标签）
                all_p_tags = box.find_all('p')
                for p_tag in all_p_tags:
                    p_text = p_tag.get_text(strip=True)
                    if p_text:
                        p_font_size_pt = self.style_computer.get_font_size_pt(p_tag)
                        # 更精确的行数计算：每行大约80个字符
                        p_lines = max(1, (len(p_text) + 79) // 80)
                        p_height = p_lines * int(p_font_size_pt * 1.5)

                        p_left = UnitConverter.px_to_emu(content_x)
                        p_top = UnitConverter.px_to_emu(current_y)
                        p_box = pptx_slide.shapes.add_textbox(
                            p_left, p_top,
                            UnitConverter.px_to_emu(content_width), UnitConverter.px_to_emu(p_height)
                        )
                        p_frame = p_box.text_frame
                        p_frame.text = p_text
                        p_frame.word_wrap = True
                        p_frame.vertical_anchor = MSO_ANCHOR.TOP  # 顶部对齐
                        for paragraph in p_frame.paragraphs:
                            paragraph.alignment = text_alignment
                            for run in paragraph.runs:
                                run.font.size = Pt(p_font_size_pt)
                                run.font.name = self.font_manager.get_font('body')

                        current_y += p_height + 5  # 间距

            else:
                # 垂直布局：图标在上，文字在下（原有逻辑，但优化间距）
                current_y = y + 25  # 增加顶部间距，避免重合
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
                            run.font.size = Pt(36)
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                    current_y += 50  # 增加图标与文字间距

                # 添加标题
                if title_elem:
                    title_text = title_elem.get_text(strip=True)
                    title_left = UnitConverter.px_to_emu(x + 15)
                    title_top = UnitConverter.px_to_emu(current_y)
                    title_box = pptx_slide.shapes.add_textbox(
                        title_left, title_top,
                        UnitConverter.px_to_emu(box_width - 30), UnitConverter.px_to_emu(25)
                    )
                    title_frame = title_box.text_frame
                    title_frame.text = title_text
                    title_frame.word_wrap = True
                    for paragraph in title_frame.paragraphs:
                        paragraph.alignment = text_alignment
                        for run in paragraph.runs:
                            title_font_size_pt = self.style_computer.get_font_size_pt(title_elem)
                            run.font.size = Pt(title_font_size_pt)
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                    current_y += 30

                # 添加主数据
                if h2:
                    h2_text = h2.get_text(strip=True)
                    h2_left = UnitConverter.px_to_emu(x + 15)
                    h2_top = UnitConverter.px_to_emu(current_y)
                    h2_box = pptx_slide.shapes.add_textbox(
                        h2_left, h2_top,
                        UnitConverter.px_to_emu(box_width - 30), UnitConverter.px_to_emu(40)
                    )
                    h2_frame = h2_box.text_frame
                    h2_frame.text = h2_text
                    for paragraph in h2_frame.paragraphs:
                        paragraph.alignment = text_alignment
                        for run in paragraph.runs:
                            h2_font_size_pt = self.style_computer.get_font_size_pt(h2)
                            run.font.size = Pt(h2_font_size_pt)
                            run.font.bold = True
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = self.font_manager.get_font('body')

                    current_y += 45

                # 添加描述（统一处理所有p标签）
                all_p_tags = box.find_all('p')
                for p_tag in all_p_tags:
                    p_text = p_tag.get_text(strip=True)
                    if p_text:
                        p_font_size_pt = self.style_computer.get_font_size_pt(p_tag)
                        # 更精确的行数计算：每行大约80个字符
                        p_lines = max(1, (len(p_text) + 79) // 80)
                        p_height = p_lines * int(p_font_size_pt * 1.5)

                        p_left = UnitConverter.px_to_emu(x + 15)
                        p_top = UnitConverter.px_to_emu(current_y)
                        p_box = pptx_slide.shapes.add_textbox(
                            p_left, p_top,
                            UnitConverter.px_to_emu(box_width - 30), UnitConverter.px_to_emu(p_height)
                        )
                        p_frame = p_box.text_frame
                        p_frame.text = p_text
                        p_frame.word_wrap = True
                        p_frame.vertical_anchor = MSO_ANCHOR.TOP  # 顶部对齐
                        for paragraph in p_frame.paragraphs:
                            paragraph.alignment = text_alignment
                            for run in paragraph.runs:
                                run.font.size = Pt(p_font_size_pt)
                                run.font.name = self.font_manager.get_font('body')

                        current_y += p_height + 5  # 间距

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

        # 0. 检查是否包含目录布局 (toc-item)
        toc_items = card.find_all('div', class_='toc-item')
        if toc_items:
            logger.info("stat-card包含toc-item目录结构，处理目录布局")
            return self._convert_toc_layout(card, toc_items, pptx_slide, y_start)

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

    def _convert_toc_layout(self, card, toc_items, pptx_slide, y_start: int) -> int:
        """
        转换目录布局 (toc-item)

        处理左右两栏的目录布局，每项包含数字编号和文本

        Args:
            card: stat-card容器
            toc_items: 目录项列表
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理目录布局(toc-item)")
        x_base = 80

        # 检查grid布局列数
        grid_container = card.find('div', class_='grid')
        if grid_container:
            grid_classes = grid_container.get('class', [])
            if 'grid-cols-2' in grid_classes:
                num_columns = 2
            elif 'grid-cols-1' in grid_classes:
                num_columns = 1
            elif 'grid-cols-3' in grid_classes:
                num_columns = 3
            else:
                num_columns = 2  # 默认2列
        else:
            num_columns = 2  # 默认2列

        logger.info(f"检测到目录布局，{num_columns}列，{len(toc_items)}个目录项")

        # 添加stat-card背景
        card_height = len(toc_items) // num_columns * 60 + 80  # 估算高度
        if len(toc_items) % num_columns > 0:
            card_height += 60

        bg_color_str = self.css_parser.get_background_color('.stat-card')
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
            logger.info(f"添加目录卡片背景，高度={card_height}px")

        current_y = y_start + 20

        # 处理目录项
        for idx, toc_item in enumerate(toc_items):
            # 计算位置（网格布局）
            col = idx % num_columns
            row = idx // num_columns
            item_x = x_base + 20 + col * 880  # 每列宽度880px
            item_y = current_y + row * 60  # 每项高度60px

            # 提取数字和文本
            number_elem = toc_item.find('div', class_='toc-number')
            text_elem = toc_item.find('div', class_='toc-text')

            if number_elem and text_elem:
                number_text = number_elem.get_text(strip=True)
                text_content = text_elem.get_text(strip=True)

                # 获取字体大小
                number_font_size = self.style_computer.get_font_size_pt(number_elem)
                text_font_size = self.style_computer.get_font_size_pt(text_elem)

                # 添加数字
                number_left = UnitConverter.px_to_emu(item_x)
                number_top = UnitConverter.px_to_emu(item_y)
                number_box = pptx_slide.shapes.add_textbox(
                    number_left, number_top,
                    UnitConverter.px_to_emu(40), UnitConverter.px_to_emu(30)
                )
                number_frame = number_box.text_frame
                number_frame.text = number_text
                number_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                for paragraph in number_frame.paragraphs:
                    paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(number_font_size)
                        run.font.bold = True
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = self.font_manager.get_font('body')

                # 添加文本
                text_left = UnitConverter.px_to_emu(item_x + 50)
                text_top = UnitConverter.px_to_emu(item_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(800), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = text_content
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(text_font_size)
                        run.font.name = self.font_manager.get_font('body')

        return y_start + card_height + 10

    def _convert_bottom_info(self, bottom_container, pptx_slide, y_start: int) -> int:
        """
        转换底部信息布局

        处理包含bullet-point的flex容器中的底部信息
        HTML中是水平排列，PPTX中也应该水平排列

        Args:
            bottom_container: 底部信息容器
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理底部信息布局（水平排列）")
        x_base = 80
        current_y = y_start

        # 查找所有bullet-point
        bullet_points = bottom_container.find_all('div', class_='bullet-point')

        # 水平排列：计算每个bullet-point的宽度
        total_width = 1760  # 可用总宽度
        item_width = total_width // len(bullet_points)  # 每项平均分配宽度
        gap = 40  # 项目间距

        for idx, bullet_point in enumerate(bullet_points):
            icon_elem = bullet_point.find('i')
            p_elem = bullet_point.find('p')

            if icon_elem and p_elem:
                # 获取图标字符
                icon_classes = icon_elem.get('class', [])
                icon_char = self._get_icon_char(icon_classes)

                # 获取文本
                text = p_elem.get_text(strip=True)

                # 计算水平位置
                item_x = x_base + idx * (item_width + gap)

                # 添加图标（在文本左侧）
                icon_left = UnitConverter.px_to_emu(item_x)
                icon_top = UnitConverter.px_to_emu(current_y)
                icon_box = pptx_slide.shapes.add_textbox(
                    icon_left, icon_top,
                    UnitConverter.px_to_emu(30), UnitConverter.px_to_emu(25)
                )
                icon_frame = icon_box.text_frame
                icon_frame.text = icon_char
                icon_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                for paragraph in icon_frame.paragraphs:
                    paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(16)
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = self.font_manager.get_font('body')

                # 添加文本（在图标右侧）
                text_left = UnitConverter.px_to_emu(item_x + 40)
                text_top = UnitConverter.px_to_emu(current_y)
                text_width = item_width - 40  # 减去图标占用的宽度
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(text_width), UnitConverter.px_to_emu(25)
                )
                text_frame = text_box.text_frame
                text_frame.text = text
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(16)
                        run.font.name = self.font_manager.get_font('body')

        return current_y + 30

    def _convert_data_card(self, card, pptx_slide, shape_converter, y_start: int) -> int:
        """转换数据卡片(.data-card)"""
        logger.info("处理data-card")
        x_base = 80

        # 注意：左边框的高度需要在计算完实际内容后再添加
        # 暂时记录起始位置，稍后添加边框

        # 初始化当前Y坐标
        current_y = y_start + 10

        # === 修复：简化的标题和内容处理逻辑 ===
        # 1. 首先查找并处理标题（查找第一个primary-color的p标签）
        title_elem = card.find('p', class_='primary-color')
        title_text = None

        if title_elem:
            title_text = title_elem.get_text(strip=True)
            if title_text:
                # 渲染标题
                text_left = UnitConverter.px_to_emu(x_base + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = title_text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用样式计算器获取正确的字体大小
                        font_size_px = self.style_computer.get_font_size_pt(title_elem)
                        run.font.size = Pt(font_size_px)
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = self.font_manager.get_font('body')

                current_y += 40  # 标题后间距
                logger.info(f"渲染data-card标题: {title_text}")

        # 2. 处理普通段落内容（明确排除标题元素和bullet-point内的元素）
        content_paragraphs = []
        all_paragraphs = card.find_all('p')

        for p in all_paragraphs:
            # 新增：检查是否在bullet-point里
            parent = p.parent
            is_in_bullet_point = False
            while parent and parent != card:
                if parent.get('class') and 'bullet-point' in parent.get('class', []):
                    is_in_bullet_point = True
                    break
                parent = parent.parent

            if is_in_bullet_point:
                continue

            # 方法1：检查是否有primary-color类
            if 'primary-color' in p.get('class', []):
                continue

            # 方法2：如果是同一个元素对象，也跳过（防止类检查失败的情况）
            if title_elem and p is title_elem:
                continue

            # 方法3：如果文本内容完全相同，也跳过（最后保险）
            p_text = p.get_text(strip=True)
            if title_text and p_text == title_text:
                continue

            # 通过所有检查的才是真正的内容段落
            if p_text:
                content_paragraphs.append(p)

        logger.info(f"data-card段落过滤: 找到{len(all_paragraphs)}个p标签，排除标题后{len(content_paragraphs)}个普通段落")

        # 3. 渲染内容段落
        for p in content_paragraphs:
            text = p.get_text(strip=True)
            if text:
                text_left = UnitConverter.px_to_emu(x_base + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = text
                text_frame.word_wrap = True
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用样式计算器获取正确的字体大小
                        font_size_px = self.style_computer.get_font_size_pt(p)
                        run.font.size = Pt(font_size_px)
                        run.font.name = self.font_manager.get_font('body')

                current_y += 35  # 段落后间距
                logger.info(f"渲染data-card内容: {text[:30]}...")  # 只记录前30个字符

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

        # === 修复：正确判断是否已有内容 ===
        # 不仅要检查progress-bar和bullet-point，还要检查是否已经处理了标题和段落
        has_title_or_content = title_elem is not None or len(content_paragraphs) > 0
        has_special_content = len(progress_bars) > 0 or len(bullet_points) > 0
        has_content = has_title_or_content or has_special_content

        logger.info(f"data-card内容检查: 标题={'是' if title_elem else '否'}, "
                   f"内容段落数={len(content_paragraphs)}, 进度条数={len(progress_bars)}, "
                   f"列表项数={len(bullet_points)}, 总已有内容={'是' if has_content else '否'}")

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
            # === 网络安全相关 ===
            # 核心安全图标
            'fa-shield': '🛡',
            'fa-shield-alt': '🛡',
            'fa-shield-virus': '🦠',  # 病毒防护
            'fa-virus-slash': '🦠',  # 病毒防护（slide01专用）
            'fa-virus': '🦠',        # 病毒
            'fa-lock': '🔒',
            'fa-unlock': '🔓',
            'fa-key': '🔑',
            'fa-fingerprint': '👆',
            'fa-user-shield': '🛡',
            'fa-user-lock': '🔐',

            # 威胁和警告
            'fa-exclamation-triangle': '⚠',
            'fa-exclamation-circle': '⚠',
            'fa-exclamation': '❗',
            'fa-warning': '⚠️',
            'fa-bell': '🔔',
            'fa-bug': '🐛',
            'fa-radiation': '☢️',
            'fa-biohazard': '☣️',

            # === 计算机和硬件 ===
            # 设备
            'fa-laptop': '💻',
            'fa-desktop': '🖥',
            'fa-server': '🖥',
            'fa-mobile': '📱',
            'fa-tablet': '📱',
            'fa-wifi': '📶',
            'fa-network-wired': '🔌',
            'fa-usb': '🔌',
            'fa-plug': '🔌',

            # 存储
            'fa-database': '🗄',
            'fa-hdd': '💾',
            'fa-sd-card': '💾',
            'fa-save': '💾',

            # === 人工智能和机器学习 ===
            'fa-robot': '🤖',       # 机器人（slide01专用）
            'fa-brain': '🧠',
            'fa-microchip': '💻',
            'fa-memory': '🧠',
            'fa-cpu': '💻',
            'fa-server': '🖥',
            'fa-cloud': '☁',
            'fa-cloud-upload-alt': '☁️',
            'fa-cloud-download-alt': '☁️',

            # === 网络和通信 ===
            'fa-globe': '🌐',
            'fa-globe-americas': '🌎',
            'fa-globe-europe': '🌍',
            'fa-globe-asia': '🌏',
            'fa-wifi': '📶',
            'fa-signal': '📶',
            'fa-satellite': '🛰️',
            'fa-ethernet': '🔌',
            'fa-router': '📡',

            # === 法律法规和合规 ===
            'fa-balance-scale': '⚖️',  # 天平（法律）
            'fa-gavel': '🔨',        # 法槌
            'fa-landmark': '🏛️',      # 立法机构
            'fa-courthouse': '🏛️',     # 法院
            'fa-scroll': '📜',        # 法律文件
            'fa-file-contract': '📄',  # 合同
            'fa-file-alt': '📄',      # 文档
            'fa-file-pdf': '📄',       # PDF文档
            'fa-file-word': '📄',      # Word文档
            'fa-file-excel': '📄',     # Excel文档

            # === 身份和权限管理 ===
            'fa-user': '👤',
            'fa-users': '👥',
            'fa-user-check': '✅',
            'fa-user-times': '❌',
            'fa-user-plus': '➕',
            'fa-user-minus': '➖',
            'fa-user-cog': '⚙️',
            'fa-id-card': '🪪',
            'fa-passport': '🪪',
            'fa-fingerprint': '👆',

            # === 数据和监控 ===
            'fa-chart-bar': '📊',
            'fa-chart-line': '📈',
            'fa-chart-pie': '📊',
            'fa-chart-area': '📈',
            'fa-table': '📊',
            'fa-database': '🗄',
            'fa-search': '🔍',
            'fa-search-plus': '🔍',
            'fa-search-minus': '🔍',

            # === 攻击和防御 ===
            'fa-swords': '⚔️',       # 攻击
            'fa-crosshairs': '🎯',    # 瞄准
            'fa-shield-alt': '🛡',     # 防御
            'fa-bomb': '💣',           # 攻击
            'fa-hammer': '🔨',         # 工具
            'fa-wrench': '🔧',         # 维修
            'fa-tools': '🛠',          # 工具集

            # === 时间和流程 ===
            'fa-clock': '🕐',
            'fa-hourglass': '⏳',
            'fa-hourglass-half': '⏳',
            'fa-calendar': '📅',
            'fa-calendar-alt': '📅',
            'fa-tasks': '☑',          # 任务列表
            'fa-list': '📋',
            'fa-clipboard': '📋',
            'fa-clipboard-check': '✅',
            'fa-clipboard-list': '📋',

            # === 系统和设置 ===
            'fa-cog': '⚙',
            'fa-cogs': '⚙️',
            'fa-settings': '⚙️',
            'fa-adjust': '⚙️',
            'fa-sliders-h': '🎚️',
            'fa-toggle-on': '🔛',
            'fa-toggle-off': '🔴',

            # === 文件和数据 ===
            'fa-file': '📄',
            'fa-file-code': '📄',
            'fa-folder': '📁',
            'fa-folder-open': '📂',
            'fa-download': '⬇',
            'fa-upload': '⬆',
            'fa-archive': '📦',
            'fa-file-archive': '📦',

            # === 通信和消息 ===
            'fa-envelope': '✉',
            'fa-envelope-open': '📧',
            'fa-comments': '💬',
            'fa-comment': '💬',
            'fa-comment-dots': '💬',
            'fa-phone': '📞',
            'fa-video': '📹',

            # === 基础图标 ===
            'fa-check': '✓',
            'fa-check-circle': '✓',
            'fa-times': '✗',
            'fa-times-circle': '❌',
            'fa-plus': '+',
            'fa-minus': '-',
            'fa-arrow-right': '→',
            'fa-arrow-left': '←',
            'fa-arrow-up': '↑',
            'fa-arrow-down': '↓',
            'fa-sync': '🔄',
            'fa-redo': '↻',
            'fa-undo': '↺',
            'fa-play': '▶',
            'fa-pause': '⏸',
            'fa-stop': '⏹',
            'fa-home': '🏠',
            'fa-building': '🏢',

            # === 其他相关图标 ===
            'fa-cloud-showers-heavy': '🌧️',  # 大雨云（slide01专用）
            'fa-info-circle': 'ℹ',
            'fa-question-circle': '❓',
            'fa-bolt': '⚡',           # 闪电（电力/网络）
            'fa-fire': '🔥',           # 火灾/威胁
            'fa-star': '⭐',           # 重要性
            'fa-heart': '♥',           # 关注
            'fa-gift': '🎁',           # 奖励
            'fa-tag': '🏷️',            # 标签
        }

        for cls in icon_classes:
            if cls in icon_map:
                return icon_map[cls]

        # 如果找不到匹配，返回默认图标
        return '●'

    def _determine_layout_direction(self, box) -> str:
        """
        智能判断布局方向：水平或垂直

        Args:
            box: stat-box元素

        Returns:
            'horizontal' 或 'vertical'
        """
        # 检查CSS样式，特别是align-items属性
        # align-items: center 通常表示水平布局
        # align-items: flex-start 或未设置通常表示垂直布局

        # 方法1：检查内联样式
        inline_style = box.get('style', '')
        if 'align-items' in inline_style:
            if 'center' in inline_style:
                logger.info("检测到align-items: center，使用水平布局")
                return 'horizontal'
            elif 'flex-start' in inline_style or 'start' in inline_style:
                logger.info("检测到align-items: flex-start，使用垂直布局")
                return 'vertical'

        # 方法2：从CSS解析器获取样式
        computed_styles = self.css_parser.get_style('.stat-box')
        align_items = computed_styles.get('align-items', '').lower() if computed_styles else ''

        if 'center' in align_items:
            logger.info("从CSS检测到align-items: center，使用水平布局")
            return 'horizontal'
        elif 'flex-start' in align_items or 'start' in align_items:
            logger.info("从CSS检测到align-items: flex-start，使用垂直布局")
            return 'vertical'

        # 方法3：检查具体的HTML结构
        # 如果有text-center类，倾向于垂直布局
        box_classes = box.get('class', [])
        if 'text-center' in box_classes:
            logger.info("检测到text-center类，使用垂直布局")
            return 'vertical'

        # 方法4：检查子元素的对齐方式
        title_elem = box.find('div', class_='stat-title')
        if title_elem:
            title_classes = title_elem.get('class', [])
            if 'text-center' in title_classes:
                logger.info("检测到标题居中，使用垂直布局")
                return 'vertical'

        # 默认策略：根据常见模式判断
        # 如果图标存在且有居中类，很可能是垂直布局
        icon = box.find('i')
        if icon:
            icon_parent_classes = icon.parent.get('class', []) if icon.parent else []
            if 'text-center' in icon_parent_classes:
                logger.info("检测到图标居中，使用垂直布局")
                return 'vertical'

        # 默认使用垂直布局（更常见的模式）
        logger.info("未检测到明确的布局方向，使用默认垂直布局")
        return 'vertical'

    def _determine_text_alignment(self, box) -> int:
        """
        智能判断文字对齐方式

        Args:
            box: stat-box元素

        Returns:
            PPTX对齐常量: PP_PARAGRAPH_ALIGNMENT.LEFT, CENTER, RIGHT
        """
        # 方法1：检查内联样式
        inline_style = box.get('style', '')
        if 'text-align' in inline_style:
            if 'center' in inline_style:
                return PP_PARAGRAPH_ALIGNMENT.CENTER
            elif 'right' in inline_style:
                return PP_PARAGRAPH_ALIGNMENT.RIGHT
            elif 'left' in inline_style:
                return PP_PARAGRAPH_ALIGNMENT.LEFT

        # 方法2：检查CSS类
        box_classes = box.get('class', [])
        if 'text-center' in box_classes:
            logger.info("检测到text-center类，使用居中对齐")
            return PP_PARAGRAPH_ALIGNMENT.CENTER
        elif 'text-right' in box_classes:
            logger.info("检测到text-right类，使用右对齐")
            return PP_PARAGRAPH_ALIGNMENT.RIGHT
        elif 'text-left' in box_classes:
            logger.info("检测到text-left类，使用左对齐")
            return PP_PARAGRAPH_ALIGNMENT.LEFT

        # 方法3：检查子元素的对齐类
        title_elem = box.find('div', class_='stat-title')
        if title_elem:
            title_classes = title_elem.get('class', [])
            if 'text-center' in title_classes:
                logger.info("检测到标题居中类，使用居中对齐")
                return PP_PARAGRAPH_ALIGNMENT.CENTER
            elif 'text-right' in title_classes:
                logger.info("检测到标题右对齐类，使用右对齐")
                return PP_PARAGRAPH_ALIGNMENT.RIGHT
            elif 'text-left' in title_classes:
                logger.info("检测到标题左对齐类，使用左对齐")
                return PP_PARAGRAPH_ALIGNMENT.LEFT

        # 方法4：从CSS解析器获取样式
        computed_styles = self.css_parser.get_style('.stat-box')
        text_align = computed_styles.get('text-align', '').lower() if computed_styles else ''

        if 'center' in text_align:
            return PP_PARAGRAPH_ALIGNMENT.CENTER
        elif 'right' in text_align:
            return PP_PARAGRAPH_ALIGNMENT.RIGHT
        elif 'left' in text_align:
            return PP_PARAGRAPH_ALIGNMENT.LEFT

        # 方法5：根据布局方向智能推断
        layout_direction = self._determine_layout_direction(box)
        if layout_direction == 'horizontal':
            # 水平布局通常左对齐更美观
            logger.info("水平布局，默认使用左对齐")
            return PP_PARAGRAPH_ALIGNMENT.LEFT
        else:
            # 垂直布局通常居中对齐更美观
            logger.info("垂直布局，默认使用居中对齐")
            return PP_PARAGRAPH_ALIGNMENT.CENTER


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
