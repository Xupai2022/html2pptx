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
from src.converters.svg_converter import SvgConverter
from src.utils.logger import setup_logger
from src.utils.unit_converter import UnitConverter
from src.utils.color_parser import ColorParser
from src.utils.chart_capture import ChartCapture
from src.utils.font_manager import get_font_manager
from src.utils.style_computer import get_style_computer
from pptx.util import Pt
from pptx.enum.text import MSO_ANCHOR, PP_PARAGRAPH_ALIGNMENT
from pptx.dml.color import RGBColor

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
        # 使用完整的HTML soup来初始化CSS解析器，以便解析head中的style标签
        self.css_parser = CSSParser(self.html_parser.full_soup)
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

            # 统一处理所有容器：优先查找space-y-10容器，如果没有则处理content-section的直接子元素
            space_y_container = slide_html.find('div', class_='space-y-10')
            content_section = slide_html.find('div', class_='content-section')

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
                    y_offset = self._process_container(container, pptx_slide, y_offset, shape_converter)
            else:
                # 新的HTML结构：直接处理content-section的直接子元素（跳过标题区域）
                logger.info("未找到space-y-10容器，处理content-section的直接子元素")

                # 跳过标题区域（第一个包含h1/h2的mb-6容器）
                containers = []
                skip_first_mb = True  # 默认跳过第一个mb容器
                for child in content_section.children:
                    if hasattr(child, 'get') and child.get('class'):
                        classes = child.get('class', [])

                        # 检查是否是标题区域
                        has_title = child.find('h1') or child.find('h2')

                        # 如果是第一个mb容器且有标题，则跳过
                        # 但要确保它不包含其他重要内容（如grid、stat-card等）
                        is_title_container = False
                        if skip_first_mb and any(cls in ['mb-6', 'mb-4', 'mb-8'] for cls in classes) and has_title:
                            # 检查是否真的是纯标题容器（不包含grid、card等内容）
                            has_content = any(cls in classes for cls in ['grid', 'stat-card', 'data-card', 'risk-card', 'flex'])
                            # 检查是否只包含标题元素和简单装饰
                            has_only_titles = True
                            for elem in child.find_all(['div', 'span', 'i']):
                                elem_classes = elem.get('class', [])
                                # 如果找到非装饰性的类，说明不是纯标题容器
                                if any(cls in elem_classes for cls in ['w-', 'h-', 'primary-bg']) and not any(cls in elem_classes for cls in ['fas', 'fa-']):
                                    # 这是装饰条，允许存在
                                    continue
                                elif any(cls in elem_classes for cls in ['fas', 'fa-']):
                                    # 图标也允许
                                    continue
                                elif elem.get('style') and any(prop in elem.get('style', '') for prop in ['width', 'height', 'background']):
                                    # 内联样式的装饰元素
                                    continue

                            if not has_content:
                                is_title_container = True
                                skip_first_mb = False  # 跳过后设置为false

                        if is_title_container:
                            continue  # 跳过纯标题容器

                        # 其他容器都保留
                        if child.name:
                            containers.append(child)

                # 处理所有容器
                for container in containers:
                    if container.name:
                        # 添加间距（模拟mb-6等间距）
                        if containers.index(container) > 0:
                            y_offset += 40  # 间距
                        y_offset = self._process_container(container, pptx_slide, y_offset, shape_converter)

            # 4. 添加页码
            page_num = self.html_parser.get_page_number(slide_html)
            if page_num:
                shape_converter.add_page_number(page_num)

        # 保存PPTX
        self.pptx_builder.save(output_path)

        logger.info("=" * 50)
        logger.info(f"转换完成! 输出: {output_path}")
        logger.info("=" * 50)

    def _process_container(self, container, pptx_slide, y_offset, shape_converter):
        """
        处理单个容器，根据其类型路由到相应的处理方法

        Args:
            container: 容器元素
            pptx_slide: PPTX幻灯片
            y_offset: 当前Y坐标
            shape_converter: 形状转换器

        Returns:
            下一个元素的Y坐标
        """
        container_classes = container.get('class', [])

        # 根据class路由到对应的处理方法
        # 优先检测grid布局（包含grid类）
        if 'grid' in container_classes:
            # 网格容器（新的Tailwind结构）
            return self._convert_grid_container(container, pptx_slide, y_offset, shape_converter)
        elif 'stats-container' in container_classes:
            # 顶层stats-container（不在stat-card内）
            return self._convert_stats_container(container, pptx_slide, y_offset)
        elif 'stat-card' in container_classes:
            return self._convert_stat_card(container, pptx_slide, y_offset)
        elif 'data-card' in container_classes:
            return self._convert_data_card(container, pptx_slide, shape_converter, y_offset)
        elif 'strategy-card' in container_classes:
            return self._convert_strategy_card(container, pptx_slide, y_offset)
        elif 'risk-card' in container_classes:
            return self._convert_risk_card(container, pptx_slide, shape_converter, y_offset)
        elif 'flex' in container_classes and 'gap-6' in container_classes:
            # 检查是否包含SVG图表的flex容器
            svgs_in_container = container.find_all('svg')
            if svgs_in_container:
                logger.info(f"检测到包含 {len(svgs_in_container)} 个SVG的flex容器")
                return self._convert_flex_charts_container(container, pptx_slide, y_offset, shape_converter)
            else:
                # 底部信息容器（包含bullet-point的flex布局）
                return self._convert_bottom_info(container, pptx_slide, y_offset)
        elif 'flex' in container_classes and 'justify-between' in container_classes:
            # 底部信息容器（包含bullet-point的flex布局）
            return self._convert_bottom_info(container, pptx_slide, y_offset)
        elif 'flex-1' in container_classes and 'overflow-hidden' in container_classes:
            # 内容容器（包含多个子容器）
            return self._convert_content_container(container, pptx_slide, y_offset, shape_converter)
        else:
            # 首先检查是否包含SVG元素
            svgs_in_container = container.find_all('svg')
            if svgs_in_container:
                logger.info(f"检测到容器包含 {len(svgs_in_container)} 个SVG元素")
                # 初始化SVG转换器
                svg_converter = SvgConverter(pptx_slide, self.css_parser, self.html_path)

                # 如果是单个SVG，直接转换
                if len(svgs_in_container) == 1:
                    svg_elem = svgs_in_container[0]

                    # 检查是否有标题
                    h3_elem = container.find('h3')
                    if h3_elem:
                        # 处理标题
                        title_text = h3_elem.get_text(strip=True)
                        if title_text:
                            text_box = pptx_slide.shapes.add_textbox(
                                UnitConverter.px_to_emu(80),
                                UnitConverter.px_to_emu(y_offset),
                                UnitConverter.px_to_emu(1760),
                                UnitConverter.px_to_emu(30)
                            )
                            text_frame = text_box.text_frame
                            text_frame.text = title_text
                            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                            for paragraph in text_frame.paragraphs:
                                paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
                                for run in paragraph.runs:
                                    run.font.size = Pt(20)
                                    run.font.bold = True
                                    run.font.name = self.font_manager.get_font('body')
                                    run.font.color.rgb = ColorParser.parse_color('rgb(10, 66, 117)')

                            y_offset += 40

                    # 转换SVG
                    chart_height = svg_converter.convert_svg(
                        svg_elem,
                        container,
                        80,
                        y_offset,
                        1760,
                        0
                    )

                    return y_offset + chart_height + 20
                else:
                    # 多个SVG，使用水平布局
                    chart_height = svg_converter.convert_multiple_svgs(
                        container,
                        80,
                        y_offset,
                        1760,
                        gap=24
                    )
                    return y_offset + chart_height + 20

            # 检测是否包含多个数字列表项（如多个toc-item）
            toc_items = container.find_all('div', class_='toc-item')
            if len(toc_items) > 1:
                return self._convert_numbered_list_group(container, pptx_slide, y_offset)
            elif 'toc-item' in container_classes or self._has_numbered_list_pattern(container):
                # 单个数字列表项
                return self._convert_numbered_list_container(container, pptx_slide, y_offset)

            # 检测flex容器（放在最后，避免误判）
            if 'flex-1' in container_classes or 'flex' in container_classes:
                # 检查flex容器内是否包含网格布局
                grid_child = container.find('div', class_='grid')
                if grid_child:
                    # 如果flex容器内只有一个grid子容器，直接处理grid
                    direct_children = [child for child in container.children if hasattr(child, 'name') and child.name]
                    if len(direct_children) == 1 and direct_children[0] == grid_child:
                        logger.info("flex容器内只包含一个网格容器，直接处理网格布局")
                        return self._convert_grid_container(grid_child, pptx_slide, y_offset, shape_converter)

                # flex容器 - 增强检测，处理居中布局
                # 检查是否是居中容器
                has_justify_center = 'justify-center' in container_classes
                has_items_center = 'items-center' in container_classes
                has_flex_col = 'flex-col' in container_classes

                # 如果是居中布局的flex容器
                if (has_justify_center and has_items_center) or (has_flex_col and has_justify_center):
                    return self._convert_centered_container(container, pptx_slide, y_offset, shape_converter)

                # 普通flex容器
                return self._convert_flex_container(container, pptx_slide, y_offset, shape_converter)

            # 未知容器类型，记录警告但尝试处理
            logger.warning(f"遇到未知容器类型: {container_classes}，尝试降级处理")
            return self._convert_generic_card(container, pptx_slide, y_offset, card_type='unknown')

    def _convert_grid_container(self, container, pptx_slide, y_start, shape_converter):
        """
        转换网格容器（如grid grid-cols-2 gap-6）

        Args:
            container: 网格容器元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标
            shape_converter: 形状转换器

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理网格容器")
        classes = container.get('class', [])

        # 获取网格列数
        num_columns = 2  # 默认2列
        for cls in classes:
            if cls.startswith('grid-cols-') and hasattr(self.css_parser, 'tailwind_grid_columns'):
                columns = self.css_parser.tailwind_grid_columns.get(cls)
                if columns:
                    num_columns = columns
                    logger.info(f"检测到网格列数: {num_columns}")
                    break

        # 获取间距
        gap = 20  # 默认间距
        for cls in classes:
            if cls.startswith('gap-') and hasattr(self.css_parser, 'tailwind_spacing'):
                gap_value = self.css_parser.tailwind_spacing.get(cls)
                if gap_value:
                    # 处理小数值，如1.5rem
                    gap_num = float(gap_value.replace('rem', ''))
                    gap = int(gap_num * 16)  # 转换rem到px
                    logger.info(f"检测到网格间距: {gap}px")
                    break

        # 获取所有子元素
        children = []
        for child in container.children:
            if hasattr(child, 'name') and child.name:
                children.append(child)

        # 计算布局
        total_width = 1760  # 可用宽度
        item_width = (total_width - (num_columns - 1) * gap) // num_columns
        item_height = 200  # 估算高度

        current_y = y_start
        max_y_in_row = y_start

        for idx, child in enumerate(children):
            col = idx % num_columns
            row = idx // num_columns

            if col == 0 and idx > 0:
                # 新行开始
                current_y = max_y_in_row + gap

            x = 80 + col * (item_width + gap)
            y = current_y

            # 处理子元素
            child_classes = child.get('class', [])

            # 检查是否需要添加左边框（data-card有左边框特性）
            needs_left_border = False
            if 'data-card' in child_classes or 'stat-card' in child_classes:
                # 检查CSS中是否有border-left样式
                css_style = self.css_parser.get_style(f".{child_classes[0]}") if child_classes else {}
                if 'border-left' in css_style or 'data-card' in child_classes:
                    needs_left_border = True

            if 'stat-card' in child_classes:
                # 专门处理网格中的stat-card
                child_y = self._convert_grid_stat_card(child, pptx_slide, shape_converter, x, y, item_width)
            elif 'data-card' in child_classes:
                # data-card需要左边框，但我们不能直接调用_convert_data_card
                # 因为它会自己添加边框。我们需要特殊处理
                child_y = self._convert_grid_data_card(child, pptx_slide, shape_converter, x, y, item_width)
            elif 'risk-card' in child_classes:
                # 专门处理网格中的risk-card
                child_y = self._convert_grid_risk_card(child, pptx_slide, shape_converter, x, y, item_width)
            else:
                # 降级处理
                child_y = self._convert_generic_card(child, pptx_slide, y, card_type='grid-item')

                # 如果需要左边框，在这里添加
                if needs_left_border:
                    # 计算实际高度
                    actual_height = child_y - y
                    if actual_height > 0:
                        shape_converter.add_border_left(x, y, actual_height, 4)

            max_y_in_row = max(max_y_in_row, child_y)

        return max_y_in_row + 20  # 返回下一行的起始位置

    def _convert_grid_data_card(self, card, pptx_slide, shape_converter, x, y, width):
        """
        转换网格中的data-card（带左侧竖线）

        Args:
            card: data-card元素
            pptx_slide: PPTX幻灯片
            shape_converter: 形状转换器
            x: X坐标
            y: Y坐标
            width: 宽度

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理网格中的data-card")

        # 初始化变量
        risk_color = None
        bg_color = None

        # 添加data-card背景色
        bg_color_str = 'rgba(10, 66, 117, 0.03)'  # 从CSS获取的背景色
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.dml.color import RGBColor

        # 精确计算所需高度
        # 基础padding: 15px上 + 15px下 = 30px
        estimated_height = 30  # data-card的上下padding

        # 处理h3标题
        h3_elem = card.find('h3')
        if h3_elem:
            # h3高度: 28px字体 + margin-bottom: 12px = 40px
            estimated_height += 40
            logger.info(f"检测到h3标题: {h3_elem.get_text(strip=True)}")

        # 处理risk-item
        risk_items = card.find_all('div', class_='risk-item')
        if risk_items:
            logger.info(f"检测到{len(risk_items)}个risk-item")
            for i, risk_item in enumerate(risk_items):
                # 每个risk-item的高度计算
                item_height = 0

                # 第一个p标签（包含strong和risk-level）
                first_p = risk_item.find('p')
                if first_p:
                    # strong标签: 22px字体
                    # risk-level: 20px字体 + padding(2px上下) + 8px左右padding
                    # 实际高度由最大元素决定，考虑padding: max(22, 20+4) = 24px
                    item_height += 24

                # 两个p标签之间的间距
                item_height += 4  # 小间距

                # 第二个p标签（描述文本）
                desc_p = risk_item.find('p', class_='text-sm')
                if desc_p:
                    # text-sm字体: 14px (根据CSS)
                    # 行高: 1.6 * 14 = 22.4px，实际需要考虑换行
                    # 但实际渲染时是25px字体（约19pt），加上行高1.6 = 40px
                    item_height += 35  # 给描述文本足够的空间

                # risk-item的margin-bottom: 12px
                if i < len(risk_items) - 1:  # 最后一个不加margin
                    item_height += 12

                estimated_height += item_height
                logger.info(f"risk-item {i+1} 精确高度: {item_height}px (总高度: {estimated_height}px)")

        # 处理其他内容（如bullet-point等）
        # 如果没有risk-item但有其他内容
        if not risk_items:
            # 查找所有直接子元素
            direct_children = []
            for child in card.children:
                if hasattr(child, 'name') and child.name:
                    if child.name != 'h3':  # h3已经计算过
                        text = child.get_text(strip=True)
                        if text and len(text) > 2:
                            direct_children.append(child)

            # 每个元素约35px高度
            estimated_height += len(direct_children) * 35

        # 确保最小高度
        estimated_height = max(estimated_height, 120)

        logger.info(f"data-card精确高度计算: {estimated_height}px")

        bg_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            UnitConverter.px_to_emu(x),
            UnitConverter.px_to_emu(y),
            UnitConverter.px_to_emu(width),
            UnitConverter.px_to_emu(estimated_height)
        )
        bg_shape.fill.solid()
        bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
        if bg_rgb:
            if alpha < 1.0:
                bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
            bg_shape.fill.fore_color.rgb = bg_rgb
        bg_shape.line.fill.background()
        logger.info(f"添加data-card背景色，高度={estimated_height}px")

        current_y = y + 15  # 顶部padding，与CSS中的15px保持一致

        # 1. 首先处理h3标题
        h3_elem = card.find('h3')
        if h3_elem:
            h3_text = h3_elem.get_text(strip=True)
            if h3_text:
                text_left = UnitConverter.px_to_emu(x + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(width - 40), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = h3_text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        font_size_pt = self.style_computer.get_font_size_pt(h3_elem)
                        run.font.size = Pt(font_size_pt)
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.bold = True
                        run.font.name = self.font_manager.get_font('h3')

                current_y += 40  # 28px字体 + 12px margin-bottom
                logger.info(f"渲染h3标题: {h3_text}")

        # 2. 处理risk-item或bullet-point
        risk_items = card.find_all('div', class_='risk-item')
        bullet_points = card.find_all('div', class_='bullet-point')

        if risk_items:
            logger.info(f"找到 {len(risk_items)} 个risk-item")
            self._process_risk_items(risk_items, card, pptx_slide, x, y, width, current_y)
        elif bullet_points:
            logger.info(f"找到 {len(bullet_points)} 个bullet-point")
            self._process_bullet_points(bullet_points, card, pptx_slide, x, y, width, current_y)

        # 3. 如果没有risk-item和bullet-point，使用原来的逻辑处理其他内容
        if not risk_items and not bullet_points:
            # 提取文本内容
            text_elements = []
            for elem in card.descendants:
                if hasattr(elem, 'name') and elem.name in ['p', 'h1', 'h2', 'h3', 'h4', 'div', 'span']:
                    # 只提取没有子块级元素的文本节点
                    if not elem.find_all(['div', 'p', 'h1', 'h2', 'h3']):
                        text = elem.get_text(strip=True)
                        if text and len(text) > 2:
                            text_elements.append(elem)

            # 渲染文本
            for elem in text_elements[:5]:  # 最多5个元素
                text = elem.get_text(strip=True)
                if text:
                    text_left = UnitConverter.px_to_emu(x + 20)
                    text_top = UnitConverter.px_to_emu(current_y)
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(width - 40), UnitConverter.px_to_emu(30)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    text_frame.word_wrap = True

                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            # 使用样式计算器获取字体大小
                            font_size_px = self.style_computer.get_font_size_pt(elem)
                            run.font.size = Pt(font_size_px) if font_size_px else Pt(16)
                            run.font.name = self.font_manager.get_font('body')

                    current_y += 35

        # 添加左边框 - 使用与背景相同的高度，确保竖线不会过长
        border_height = estimated_height  # 使用背景矩形的高度
        shape_converter.add_border_left(x, y, border_height, 4)

        # 确保返回正确的位置，保留底部padding
        # 返回背景矩形的底部位置 + 额外间距
        return y + estimated_height + 5

    def _calculate_text_width(self, text: str, font_size: Pt) -> int:
        """
        计算文本的像素宽度

        Args:
            text: 要计算的文本
            font_size: 字体大小（Pt）

        Returns:
            int: 文本的宽度（像素）
        """
        # 将Pt转换为Px
        font_size_px = int(font_size.pt * 0.75)

        # 计算字符宽度
        total_width = 0
        for char in text:
            # 中文字符宽度约为字体大小的1倍
            if '\u4e00' <= char <= '\u9fff':
                char_width = font_size_px
            # 英文字母和数字宽度约为字体大小的0.6倍
            elif char.isalnum() or char in '.,;:!?\'"()[]{}-+/\\=_@#%&*':
                char_width = int(font_size_px * 0.6)
            # 空格宽度约为字体大小的0.3倍
            elif char == ' ':
                char_width = int(font_size_px * 0.3)
            # 其他符号
            else:
                char_width = font_size_px

            total_width += char_width

        return total_width

    def _process_bullet_points(self, bullet_points, card, pptx_slide, x, y, width, current_y):
        """
        处理bullet-point列表

        Args:
            bullet_points: bullet-point元素列表
            card: 父data-card元素
            pptx_slide: PPTX幻灯片
            x: 起始X坐标
            y: 起始Y坐标
            width: 容器宽度
            current_y: 当前Y坐标偏移
        """
        from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
        from pptx.dml.color import RGBColor

        for bp in bullet_points:
            # 获取图标
            icon_elem = bp.find('i')
            icon_char = None
            icon_color = ColorParser.get_primary_color()

            if icon_elem:
                icon_classes = icon_elem.get('class', [])
                # 根据图标类确定颜色
                if 'risk-high' in icon_classes:
                    icon_color = RGBColor(220, 38, 38)  # 红色
                elif 'risk-medium' in icon_classes:
                    icon_color = RGBColor(234, 88, 12)  # 橙色
                elif 'risk-low' in icon_classes:
                    icon_color = RGBColor(202, 138, 4)  # 黄色

                # 获取图标字符（简化处理）
                if 'fa-cloud' in icon_classes:
                    icon_char = "☁"
                elif 'fa-comments' in icon_classes:
                    icon_char = "💬"
                elif 'fa-code-branch' in icon_classes:
                    icon_char = "⚡"
                elif 'fa-globe' in icon_classes:
                    icon_char = "🌐"
                elif 'fa-building' in icon_classes:
                    icon_char = "🏢"
                elif 'fa-link' in icon_classes:
                    icon_char = "🔗"
                elif 'fa-box' in icon_classes:
                    icon_char = "📦"
                elif 'fa-server' in icon_classes:
                    icon_char = "🖥"
                elif 'fa-exclamation-triangle' in icon_classes:
                    icon_char = "⚠"
                elif 'fa-shield-alt' in icon_classes:
                    icon_char = "🛡"
                elif 'fa-clock' in icon_classes:
                    icon_char = "⏰"
                else:
                    icon_char = "•"  # 默认圆点

            # 获取文本内容
            p_elem = bp.find('p')
            if p_elem:
                text = p_elem.get_text(strip=True)

                # 创建文本框
                text_box = pptx_slide.shapes.add_textbox(
                    UnitConverter.px_to_emu(x + 20),
                    UnitConverter.px_to_emu(current_y),
                    UnitConverter.px_to_emu(width - 40),
                    UnitConverter.px_to_emu(35)
                )
                text_frame = text_box.text_frame
                text_frame.clear()

                # 添加段落
                p = text_frame.paragraphs[0]

                # 添加图标
                if icon_char:
                    icon_run = p.add_run()
                    icon_run.text = icon_char + " "
                    icon_run.font.size = Pt(25)
                    icon_run.font.color.rgb = icon_color
                    icon_run.font.name = self.font_manager.get_font('body')

                # 添加文本
                text_run = p.add_run()
                text_run.text = text

                # 获取字体大小（从CSS解析）
                font_size_pt = self.style_computer.get_font_size_pt(p_elem)
                if font_size_pt:
                    text_run.font.size = Pt(font_size_pt)
                else:
                    text_run.font.size = Pt(25)  # 默认25px = 19pt

                text_run.font.name = self.font_manager.get_font('body')
                text_run.font.color.rgb = RGBColor(51, 51, 51)  # 深灰色

                # 设置段落格式
                p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
                p.space_before = Pt(0)
                p.space_after = Pt(4)

                current_y += 35  # 每个bullet-point占35px

        logger.info(f"处理了 {len(bullet_points)} 个bullet-point")

    def _process_risk_items(self, risk_items, card, pptx_slide, x, y, width, current_y):
        """
        处理risk-item列表（原有的逻辑）

        Args:
            risk_items: risk-item元素列表
            card: 父data-card元素
            pptx_slide: PPTX幻灯片
            x: 起始X坐标
            y: 起始Y坐标
            width: 容器宽度
            current_y: 当前Y坐标偏移
        """
        # 这里保留原有的risk-item处理逻辑
        # 由于重构，暂时使用空实现
        logger.info(f"处理了 {len(risk_items)} 个risk-item")

    def _convert_grid_risk_card(self, card, pptx_slide, shape_converter, x, y, width):
        """
        转换网格中的risk-card（带红色边框和特殊样式）

        Args:
            card: risk-card元素
            pptx_slide: PPTX幻灯片
            shape_converter: 形状转换器
            x: X坐标
            y: Y坐标
            width: 宽度

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理网格中的risk-card")

        # 调用完整的risk-card转换方法，但传入指定的位置和宽度
        # 临时保存原始的x_base和card_width
        original_x_base = 80
        original_width = 1760

        # 临时修改risk-card方法的坐标参数以适应网格布局
        card_height = 180

        # 获取CSS样式
        card_style = self.css_parser.get_class_style('risk-card') or {}

        # 添加背景（渐变效果）
        bg_color_str = card_style.get('background', 'linear-gradient(135deg, rgba(239, 68, 68, 0.08) 0%, rgba(239, 68, 68, 0.02) 100%)')

        # 创建矩形背景
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.dml.color import RGBColor
        from pptx.util import Pt
        from src.utils.unit_converter import UnitConverter
        from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

        bg_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            UnitConverter.px_to_emu(x),
            UnitConverter.px_to_emu(y),
            UnitConverter.px_to_emu(width),
            UnitConverter.px_to_emu(card_height)
        )
        bg_shape.fill.solid()

        # 解析渐变背景色，使用最深的颜色
        if 'rgba(239, 68, 68' in bg_color_str:
            # 红色系风险
            bg_rgb = RGBColor(254, 242, 242)  # 非常浅的红色
        elif 'rgba(251, 146' in bg_color_str:
            # 橙色系风险
            bg_rgb = RGBColor(255, 251, 235)  # 非常浅的橙色
        elif 'rgba(250, 204' in bg_color_str:
            # 黄色系风险
            bg_rgb = RGBColor(254, 252, 232)  # 非常浅的黄色
        else:
            # 默认浅灰色
            bg_rgb = RGBColor(249, 250, 251)

        bg_shape.fill.fore_color.rgb = bg_rgb
        bg_shape.line.fill.background()
        bg_shape.shadow.inherit = False

        # 添加左边框
        border_color_str = card_style.get('border-left-color', '#ef4444')
        from src.utils.color_parser import ColorParser
        border_color = ColorParser.parse_color(border_color_str)
        if not border_color:
            # 根据风险等级确定边框颜色
            border_color = ColorParser.parse_color('#ef4444')  # 默认红色

        border_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            UnitConverter.px_to_emu(x),
            UnitConverter.px_to_emu(y),
            UnitConverter.px_to_emu(4),
            UnitConverter.px_to_emu(card_height)
        )
        border_shape.fill.solid()
        border_shape.fill.fore_color.rgb = border_color
        border_shape.line.fill.background()

        current_y = y + 15

        # 缩小内容区域以适应网格
        content_width = width - 40
        content_x = x + 20

        # 处理flex布局内容
        flex_container = card.find('div', class_='flex')
        if flex_container:
            # 获取左侧内容区域
            left_div = flex_container.find('div', class_='flex-1')
            if left_div:
                # 处理风险标题
                title_div = left_div.find('div', class_='risk-title')
                if title_div:
                    # 获取图标
                    icon_elem = title_div.find('i')
                    icon_text = ""
                    icon_color = None

                    if icon_elem:
                        icon_classes = icon_elem.get('class', [])
                        # 根据图标类确定颜色
                        if 'severity-critical' in icon_classes:
                            icon_color = ColorParser.parse_color('#dc2626')
                            icon_text = "⚠"
                        elif 'severity-high' in icon_classes:
                            icon_color = ColorParser.parse_color('#ea580c')
                            icon_text = "⚠"
                        elif 'severity-medium' in icon_classes:
                            icon_color = ColorParser.parse_color('#d97706')
                            icon_text = "⚠"
                        else:
                            icon_text = "•"

                    # 获取标题文本
                    title_text = title_div.get_text(strip=True)
                    if icon_text:
                        title_text = title_text.replace(icon_text, "").strip()

                    # 添加标题文本（缩小字体以适应网格）
                    text_left = UnitConverter.px_to_emu(content_x)
                    text_top = UnitConverter.px_to_emu(current_y)

                    if icon_text and icon_color:
                        # 如果有图标，创建两段式文本
                        text_box = pptx_slide.shapes.add_textbox(
                            text_left, text_top,
                            UnitConverter.px_to_emu(content_width - 150), UnitConverter.px_to_emu(30)
                        )
                        text_frame = text_box.text_frame
                        p = text_frame.paragraphs[0]

                        # 图标run
                        icon_run = p.add_run()
                        icon_run.text = icon_text + " "
                        icon_run.font.size = Pt(20)
                        icon_run.font.name = self.font_manager.get_font('body')
                        icon_run.font.color.rgb = icon_color
                        icon_run.font.bold = True

                        # 标题run
                        title_run = p.add_run()
                        title_run.text = title_text
                        title_run.font.size = Pt(20)
                        title_run.font.name = self.font_manager.get_font('body')
                        title_run.font.bold = True
                        title_run.font.color.rgb = RGBColor(51, 51, 51)  # 深灰色
                    else:
                        # 没有图标，直接添加标题
                        text_box = pptx_slide.shapes.add_textbox(
                            text_left, text_top,
                            UnitConverter.px_to_emu(content_width - 150), UnitConverter.px_to_emu(30)
                        )
                        text_frame = text_box.text_frame
                        text_frame.text = title_text

                        for paragraph in text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(20)
                                run.font.name = self.font_manager.get_font('body')
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(51, 51, 51)

                    current_y += 30

                # 处理风险描述（缩小字体）
                desc_div = left_div.find('div', class_='risk-desc')
                if desc_div:
                    desc_text = desc_div.get_text(strip=True)
                    if desc_text:
                        text_box = pptx_slide.shapes.add_textbox(
                            UnitConverter.px_to_emu(content_x),
                            UnitConverter.px_to_emu(current_y),
                            UnitConverter.px_to_emu(content_width - 150),
                            UnitConverter.px_to_emu(40)
                        )
                        text_frame = text_box.text_frame
                        text_frame.text = desc_text

                        for paragraph in text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(16)
                                run.font.name = self.font_manager.get_font('body')
                                run.font.color.rgb = RGBColor(102, 102, 102)  # 灰色

                        current_y += 25

                # 处理标签（缩小尺寸）
                tag_div = left_div.find('div', class_='mt-3')
                if tag_div:
                    span_elem = tag_div.find('span')
                    if span_elem:
                        tag_text = span_elem.get_text(strip=True)
                        tag_classes = span_elem.get('class', [])

                        # 确定标签颜色
                        tag_bg_color = RGBColor(254, 226, 226)  # 浅红色背景
                        tag_text_color = RGBColor(153, 27, 27)  # 深红色文字

                        if 'bg-orange-100' in tag_classes:
                            tag_bg_color = RGBColor(255, 237, 213)  # 浅橙色背景
                            tag_text_color = RGBColor(154, 52, 18)  # 深橙色文字
                        elif 'bg-yellow-100' in tag_classes:
                            tag_bg_color = RGBColor(254, 249, 195)  # 浅黄色背景
                            tag_text_color = RGBColor(120, 53, 15)  # 深黄色文字

                        # 创建标签背景
                        tag_box = pptx_slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            UnitConverter.px_to_emu(content_x),
                            UnitConverter.px_to_emu(current_y),
                            UnitConverter.px_to_emu(80),
                            UnitConverter.px_to_emu(24)
                        )
                        tag_box.fill.solid()
                        tag_box.fill.fore_color.rgb = tag_bg_color
                        tag_box.line.fill.background()

                        # 添加标签文本
                        tag_text_box = pptx_slide.shapes.add_textbox(
                            UnitConverter.px_to_emu(content_x),
                            UnitConverter.px_to_emu(current_y + 2),
                            UnitConverter.px_to_emu(80),
                            UnitConverter.px_to_emu(20)
                        )
                        tag_text_frame = tag_text_box.text_frame
                        tag_text_frame.text = tag_text

                        for paragraph in tag_text_frame.paragraphs:
                            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                            for run in paragraph.runs:
                                run.font.size = Pt(12)
                                run.font.name = self.font_manager.get_font('body')
                                run.font.color.rgb = tag_text_color
                                run.font.bold = True

            # 获取右侧CVSS分数区域（缩小字体）
            right_div = flex_container.find('div', class_='text-center')
            if right_div:
                # 获取CVSS分数
                score_div = right_div.find('div', class_='cvss-score')
                if score_div:
                    score_text = score_div.get_text(strip=True)

                    # 添加CVSS分数
                    score_box = pptx_slide.shapes.add_textbox(
                        UnitConverter.px_to_emu(x + width - 120),
                        UnitConverter.px_to_emu(y + 40),
                        UnitConverter.px_to_emu(100),
                        UnitConverter.px_to_emu(50)
                    )
                    score_frame = score_box.text_frame
                    score_frame.text = score_text

                    for paragraph in score_frame.paragraphs:
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(36)
                            run.font.name = self.font_manager.get_font('body')
                            run.font.bold = True

                            # 根据分数确定颜色
                            if '10.0' in score_text or '9.' in score_text:
                                run.font.color.rgb = RGBColor(239, 68, 68)  # 红色
                            elif '8.' in score_text or '7.' in score_text:
                                run.font.color.rgb = RGBColor(234, 88, 12)  # 橙色
                            elif '6.' in score_text or '5.' in score_text:
                                run.font.color.rgb = RGBColor(217, 119, 6)  # 黄色
                            else:
                                run.font.color.rgb = RGBColor(107, 114, 128)  # 灰色

                # 获取CVSS标签
                label_div = right_div.find('div', class_='cvss-label')
                if label_div:
                    label_text = label_div.get_text(strip=True)

                    label_box = pptx_slide.shapes.add_textbox(
                        UnitConverter.px_to_emu(x + width - 120),
                        UnitConverter.px_to_emu(y + 90),
                        UnitConverter.px_to_emu(100),
                        UnitConverter.px_to_emu(25)
                    )
                    label_frame = label_box.text_frame
                    label_frame.text = label_text

                    for paragraph in label_frame.paragraphs:
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(14)
                            run.font.name = self.font_manager.get_font('body')
                            run.font.color.rgb = RGBColor(102, 102, 102)

        return y + card_height + 10

    def _convert_grid_stat_card(self, card, pptx_slide, shape_converter, x, y, width):
        """
        转换网格中的stat-card（带背景色）

        Args:
            card: stat-card元素
            pptx_slide: PPTX幻灯片
            shape_converter: 形状转换器
            x: X坐标
            y: Y坐标
            width: 宽度

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理网格中的stat-card")

        # 添加背景色
        bg_color_str = self.css_parser.get_background_color('.stat-card')
        if bg_color_str:
            from pptx.enum.shapes import MSO_SHAPE
            # 估算高度
            height = 180
            bg_shape = pptx_slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                UnitConverter.px_to_emu(x),
                UnitConverter.px_to_emu(y),
                UnitConverter.px_to_emu(width),
                UnitConverter.px_to_emu(height)
            )
            bg_shape.fill.solid()
            bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
            if bg_rgb:
                if alpha < 1.0:
                    bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                bg_shape.fill.fore_color.rgb = bg_rgb
            bg_shape.line.fill.background()
            logger.info(f"添加stat-card背景色: {bg_color_str}")

        # 添加左边框
        border_left_style = self.css_parser.get_style('.stat-card').get('border-left', '')
        if '4px solid' in border_left_style:
            shape_converter.add_border_left(x, y, 180, 4)

        # 首先检查是否包含risk-level标签（风险分布）
        risk_levels = card.find_all('span', class_='risk-level')
        if risk_levels:
            logger.info(f"stat-card包含{len(risk_levels)}个risk-level标签，处理为风险分布")

            # 处理h3标题
            h3_elem = card.find('h3')
            current_y = y + 20

            if h3_elem:
                h3_text = h3_elem.get_text(strip=True)
                if h3_text:
                    h3_font_size_pt = self.style_computer.get_font_size_pt(h3_elem)
                    h3_color = self._get_element_color(h3_elem) or ColorParser.get_primary_color()

                    text_left = UnitConverter.px_to_emu(x + 20)
                    text_top = UnitConverter.px_to_emu(current_y)
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(width - 40), UnitConverter.px_to_emu(30)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = h3_text
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(h3_font_size_pt)
                            run.font.name = self.font_manager.get_font('h3')
                            # 智能判断是否应该加粗
                            if self._should_be_bold(h3_elem):
                                run.font.bold = True
                            run.font.color.rgb = h3_color

                    current_y += 35

            # 处理risk-level标签
            num_risks = len(risk_levels)
            total_width = width - 40
            risk_width = total_width // num_risks - 20
            current_x = x + 20

            for risk_level in risk_levels:
                risk_text = risk_level.get_text(strip=True)
                risk_classes = risk_level.get('class', [])

                # 获取风险等级的颜色
                risk_color = None
                bg_color = None
                if 'risk-high' in risk_classes:
                    risk_color = ColorParser.parse_color('#dc2626')  # 红色
                    bg_color = RGBColor(252, 231, 229)  # 浅红色背景
                elif 'risk-medium' in risk_classes:
                    risk_color = ColorParser.parse_color('#f59e0b')  # 橙色
                    bg_color = RGBColor(254, 243, 199)  # 浅橙色背景
                elif 'risk-low' in risk_classes:
                    risk_color = ColorParser.parse_color('#3b82f6')  # 蓝色
                    bg_color = RGBColor(239, 246, 255)  # 浅蓝色背景

                # 添加背景形状
                if bg_color:
                    from pptx.enum.shapes import MSO_SHAPE
                    bg_shape = pptx_slide.shapes.add_shape(
                        MSO_SHAPE.ROUNDED_RECTANGLE,
                        UnitConverter.px_to_emu(current_x),
                        UnitConverter.px_to_emu(current_y),
                        UnitConverter.px_to_emu(risk_width),
                        UnitConverter.px_to_emu(35)
                    )
                    bg_shape.fill.solid()
                    bg_shape.fill.fore_color.rgb = bg_color
                    bg_shape.line.fill.background()

                # 创建风险等级文本框
                text_left = UnitConverter.px_to_emu(current_x + 5)
                text_top = UnitConverter.px_to_emu(current_y + 5)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(risk_width - 10), UnitConverter.px_to_emu(25)
                )
                text_frame = text_box.text_frame
                text_frame.text = risk_text
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                for paragraph in text_frame.paragraphs:
                    paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(20)
                        run.font.bold = True
                        run.font.name = self.font_manager.get_font('body')

                        # 应用风险等级颜色
                        if risk_color:
                            run.font.color.rgb = risk_color

                # 移动到下一个位置
                current_x += risk_width + 20

            return y + 180

        # 查找内部的flex容器
        flex_container = card.find('div', class_='flex')
        if flex_container:
            # 收集所有内容（标题和数字）
            content_elements = []

            # 查找左侧内容区域（包含标题和数字的div）
            content_div = flex_container.find('div')
            if content_div:
                # 收集标题
                h3_elem = content_div.find('h3')
                if h3_elem:
                    content_elements.append(('h3', h3_elem))

                # 收集所有p标签（数字）
                for p_elem in content_div.find_all('p'):
                    content_elements.append(('p', p_elem))

            # 计算内容的总高度
            total_content_height = 0
            element_heights = []

            for elem_type, elem in content_elements:
                if elem_type == 'h3':
                    height = 30  # 标题高度
                else:
                    # 检查是否是大数字
                    p_classes = elem.get('class', [])
                    is_large_number = any(cls in p_classes for cls in ['text-4xl', 'text-3xl', 'text-2xl'])
                    height = 50 if is_large_number else 25

                element_heights.append(height)
                total_content_height += height

            # 计算垂直起始位置（垂直居中）
            card_height = 180
            start_y = y + (card_height - total_content_height) // 2
            if start_y < y + 15:
                start_y = y + 15  # 保证最小内边距

            # 渲染内容
            current_y = start_y
            for idx, (elem_type, elem) in enumerate(content_elements):
                text = elem.get_text(strip=True)
                if not text:
                    continue

                text_left = UnitConverter.px_to_emu(x + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                height = element_heights[idx]

                if elem_type == 'h3':
                    # 处理标题
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(width - 80), UnitConverter.px_to_emu(height)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    text_frame.word_wrap = True
                    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                    # 获取字体大小和颜色
                    font_size_pt = self.style_computer.get_font_size_pt(elem)
                    element_color = self._get_element_color(elem)

                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(font_size_pt)
                            run.font.name = self.font_manager.get_font('h3')
                            # 智能判断是否应该加粗
                            if self._should_be_bold(elem):
                                run.font.bold = True

                            # 应用颜色
                            if element_color:
                                run.font.color.rgb = element_color
                            else:
                                # 检查类名
                                classes = elem.get('class', [])
                                if 'text-gray-600' in classes:
                                    run.font.color.rgb = RGBColor(102, 102, 102)  # 灰色

                else:
                    # 处理数字
                    p_classes = elem.get('class', [])
                    is_large_number = any(cls in p_classes for cls in ['text-4xl', 'text-3xl', 'text-2xl'])

                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(width - 80), UnitConverter.px_to_emu(height)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    text_frame.word_wrap = True
                    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                    # 获取字体大小和颜色
                    font_size_pt = self.style_computer.get_font_size_pt(elem)
                    element_color = self._get_element_color(elem)

                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(font_size_pt)

                            if is_large_number:
                                run.font.bold = True

                            run.font.name = self.font_manager.get_font('body')

                            # 应用颜色 - 检查Tailwind CSS颜色类
                            color_found = False
                            for cls in p_classes:
                                # 处理Tailwind CSS颜色类
                                if cls == 'primary-color':
                                    run.font.color.rgb = ColorParser.parse_color('rgb(10, 66, 117)')
                                    color_found = True
                                    break
                                elif cls.startswith('text-') and hasattr(self.css_parser, 'tailwind_colors'):
                                    color_str = self.css_parser.tailwind_colors.get(cls)
                                    if color_str:
                                        color_rgb = ColorParser.parse_color(color_str)
                                        if color_rgb:
                                            run.font.color.rgb = color_rgb
                                            color_found = True
                                            break

                            # 如果element_color存在，优先使用
                            if element_color and not color_found:
                                run.font.color.rgb = element_color
                            elif not color_found:
                                # 默认文本颜色
                                run.font.color.rgb = ColorParser.get_text_color()

                current_y += height

            # 处理右侧图标
            icon_elem = flex_container.find('i')
            if icon_elem:
                icon_classes = icon_elem.get('class', [])
                icon_char = self._get_icon_char(icon_classes)

                if icon_char:
                    # 获取图标字体大小
                    icon_font_size_px = 36  # 默认值
                    icon_font_size_pt = 27  # 默认值

                    # 检查是否有text-4xl等字体大小类
                    icon_classes = icon_elem.get('class', [])
                    for cls in icon_classes:
                        if cls.startswith('text-') and hasattr(self.css_parser, 'tailwind_font_sizes'):
                            font_size_str = self.css_parser.tailwind_font_sizes.get(cls)
                            if font_size_str:
                                icon_font_size_px = int(font_size_str.replace('px', ''))
                                icon_font_size_pt = UnitConverter.font_size_px_to_pt(icon_font_size_px)
                                break

                    # 图标框尺寸基于字体大小
                    icon_box_size = icon_font_size_px + 4  # 稍微留点边距
                    icon_left = UnitConverter.px_to_emu(x + width - icon_box_size - 10)
                    icon_top = UnitConverter.px_to_emu(y + (180 - icon_box_size) // 2)  # 垂直居中

                    icon_box = pptx_slide.shapes.add_textbox(
                        icon_left, icon_top,
                        UnitConverter.px_to_emu(icon_box_size), UnitConverter.px_to_emu(icon_box_size)
                    )
                    icon_frame = icon_box.text_frame
                    icon_frame.text = icon_char
                    icon_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                    for paragraph in icon_frame.paragraphs:
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(icon_font_size_pt)
                            run.font.name = self.font_manager.get_font('body')

                            # 图标颜色
                            icon_color = self._get_element_color(icon_elem)
                            if icon_color:
                                run.font.color.rgb = icon_color
                            elif 'primary-color' in icon_classes:
                                run.font.color.rgb = ColorParser.get_primary_color()
                            elif 'text-orange-600' in icon_classes:
                                run.font.color.rgb = ColorParser.get_color_by_name('orange')
                            else:
                                run.font.color.rgb = RGBColor(200, 200, 200)  # 浅色（opacity-30效果）

        else:
            # 降级处理：查找所有文本内容
            text_elements = []
            for elem in card.descendants:
                if hasattr(elem, 'name') and elem.name in ['p', 'h1', 'h2', 'h3', 'h4']:
                    # 只提取没有子块级元素的文本节点
                    if not elem.find_all(['div', 'p', 'h1', 'h2', 'h3']):
                        text = elem.get_text(strip=True)
                        if text and len(text) > 0:
                            text_elements.append(elem)

            # 初始化current_y
            current_y = y + 20

            # 渲染文本
            for elem in text_elements[:5]:  # 最多5个元素
                text = elem.get_text(strip=True)
                if text:
                    text_left = UnitConverter.px_to_emu(x + 20)
                    text_top = UnitConverter.px_to_emu(current_y)

                    # 根据元素类型和字体大小计算高度
                    font_size_pt = self.style_computer.get_font_size_pt(elem)
                    if elem.name == 'h3':
                        height = 40
                    elif font_size_pt and font_size_pt > 30:  # text-4xl 等大字体
                        height = 50
                    elif font_size_pt and font_size_pt > 20:  # text-lg 等中等字体
                        height = 35
                    else:
                        height = 30

                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(width - 40), UnitConverter.px_to_emu(height)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    text_frame.word_wrap = True
                    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            # 使用样式计算器获取字体大小
                            run.font.size = Pt(font_size_pt) if font_size_pt else Pt(16)

                            # 设置字体
                            if elem.name == 'h3':
                                run.font.name = self.font_manager.get_font('h3')
                                run.font.bold = True
                            else:
                                run.font.name = self.font_manager.get_font('body')

                            # 处理文字颜色
                            elem_classes = elem.get('class', [])
                            element_color = self._get_element_color(elem)

                            if element_color:
                                run.font.color.rgb = element_color
                            else:
                                # 检查特定的颜色类
                                if 'primary-color' in elem_classes:
                                    run.font.color.rgb = ColorParser.get_primary_color()
                                elif 'text-red-600' in elem_classes:
                                    run.font.color.rgb = RGBColor(220, 38, 38)  # 红色
                                elif 'text-green-600' in elem_classes:
                                    run.font.color.rgb = RGBColor(22, 163, 74)  # 绿色
                                elif 'text-gray-800' in elem_classes:
                                    run.font.color.rgb = RGBColor(31, 41, 55)  # 深灰色
                                elif 'text-gray-600' in elem_classes:
                                    run.font.color.rgb = RGBColor(75, 85, 99)  # 中灰色
                                else:
                                    # 默认文本颜色
                                    run.font.color.rgb = ColorParser.get_text_color()

                    current_y += height + 10  # 增加间距

        return y + 180  # 返回固定高度

    def _get_element_relative_position(self, element, container):
        """
        获取元素相对于容器的位置

        Args:
            element: 元素
            container: 容器元素

        Returns:
            (x, y) 相对位置
        """
        # 初始化相对位置
        rel_x = 0
        rel_y = 0

        # 获取元素的样式
        style_str = element.get('style', '')
        classes = element.get('class', [])

        # 解析margin
        if style_str:
            # 解析margin-top
            import re
            margin_match = re.search(r'margin-top:\s*(\d+)px', style_str)
            if margin_match:
                rel_y += int(margin_match.group(1))

            # 解析margin-bottom
            margin_match = re.search(r'margin-bottom:\s*(\d+)px', style_str)
            if margin_match:
                # margin-bottom会在后续处理
                pass

        # 根据class判断位置
        if isinstance(classes, str):
            classes = classes.split()

        # Bootstrap/Tailwind margin classes
        for cls in classes:
            if cls.startswith('mb-') or cls.startswith('margin-bottom-'):
                # 提取margin-bottom值
                try:
                    value = int(cls.replace('mb-', '').replace('margin-bottom-', ''))
                    # Tailwind默认单位是0.25rem (4px)
                    if 'mb-' in cls:
                        rel_y += value * 4
                except:
                    pass
            elif cls.startswith('mt-') or cls.startswith('margin-top-'):
                # 提取margin-top值
                try:
                    value = int(cls.replace('mt-', '').replace('margin-top-', ''))
                    if 'mt-' in cls:
                        rel_y += value * 4
                except:
                    pass

        return rel_x, rel_y

    def _determine_title_text_alignment(self, title_elem):
        """
        智能检测标题的文本对齐方式

        Args:
            title_elem: 标题元素

        Returns:
            PP_PARAGRAPH_ALIGNMENT 枚举值
        """
        from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

        # 1. 检查内联样式
        style_str = title_elem.get('style', '')
        if 'text-align' in style_str:
            import re
            align_match = re.search(r'text-align:\s*(\w+)', style_str)
            if align_match:
                align_value = align_match.group(1).lower()
                if align_value == 'center':
                    return PP_PARAGRAPH_ALIGNMENT.CENTER
                elif align_value == 'right':
                    return PP_PARAGRAPH_ALIGNMENT.RIGHT
                elif align_value == 'justify':
                    return PP_PARAGRAPH_ALIGNMENT.JUSTIFY

        # 2. 检查CSS类
        classes = title_elem.get('class', [])
        if isinstance(classes, str):
            classes = classes.split()

        for cls in classes:
            if cls == 'text-center':
                return PP_PARAGRAPH_ALIGNMENT.CENTER
            elif cls == 'text-right':
                return PP_PARAGRAPH_ALIGNMENT.RIGHT
            elif cls == 'text-justify':
                return PP_PARAGRAPH_ALIGNMENT.JUSTIFY
            elif cls == 'text-left':
                return PP_PARAGRAPH_ALIGNMENT.LEFT

        # 3. 检查父容器的对齐设置
        parent = title_elem.parent
        while parent:
            parent_style = parent.get('style', '')
            if 'text-align' in parent_style:
                import re
                align_match = re.search(r'text-align:\s*(\w+)', parent_style)
                if align_match:
                    align_value = align_match.group(1).lower()
                    if align_value == 'center':
                        return PP_PARAGRAPH_ALIGNMENT.CENTER
                    elif align_value == 'right':
                        return PP_PARAGRAPH_ALIGNMENT.RIGHT
                    elif align_value == 'justify':
                        return PP_PARAGRAPH_ALIGNMENT.JUSTIFY

            parent_classes = parent.get('class', [])
            if isinstance(parent_classes, str):
                parent_classes = parent_classes.split()

            for cls in parent_classes:
                if cls == 'text-center':
                    return PP_PARAGRAPH_ALIGNMENT.CENTER
                elif cls == 'text-right':
                    return PP_PARAGRAPH_ALIGNMENT.RIGHT
                elif cls == 'text-justify':
                    return PP_PARAGRAPH_ALIGNMENT.JUSTIFY
                elif cls == 'text-left':
                    return PP_PARAGRAPH_ALIGNMENT.LEFT

            parent = parent.parent

        # 4. 检查CSS计算样式
        # 尝试从CSS解析器获取样式
        if hasattr(self, 'css_parser'):
            # 获取元素的样式
            selector = title_elem.name
            if classes:
                selector += '.' + '.'.join(classes)

            style = self.css_parser.get_style(selector)
            if style and 'text-align' in style:
                align_value = style['text-align'].lower()
                if align_value == 'center':
                    return PP_PARAGRAPH_ALIGNMENT.CENTER
                elif align_value == 'right':
                    return PP_PARAGRAPH_ALIGNMENT.RIGHT
                elif align_value == 'justify':
                    return PP_PARAGRAPH_ALIGNMENT.JUSTIFY

        # 5. 根据上下文推断对齐方式
        # 检查是否在flex容器中
        flex_container = title_elem.find_parent(class_='flex')
        if flex_container:
            # 检查flex容器的对齐类
            flex_classes = flex_container.get('class', [])
            if isinstance(flex_classes, str):
                flex_classes = flex_classes.split()

            # 如果有justify-between，可能是多列布局，默认左对齐
            if 'justify-between' in flex_classes or 'justify-around' in flex_classes or 'justify-evenly' in flex_classes:
                return PP_PARAGRAPH_ALIGNMENT.LEFT

        # 6. 默认值：左对齐（大多数图表标题的默认选择）
        return PP_PARAGRAPH_ALIGNMENT.LEFT

    def _convert_flex_charts_container(self, container, pptx_slide, y_start, shape_converter):
        """
        转换包含SVG图表的flex容器

        Args:
            container: flex容器元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标
            shape_converter: 形状转换器

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理包含SVG图表的flex容器")

        # 初始化SVG转换器
        svg_converter = SvgConverter(pptx_slide, self.css_parser, self.html_path)

        # 获取所有直接子元素（应该是图表容器）
        chart_containers = []
        for child in container.children:
            if hasattr(child, 'name') and child.name == 'div':
                chart_containers.append(child)

        if not chart_containers:
            logger.warning("flex容器中未找到图表容器")
            return y_start

        logger.info(f"找到 {len(chart_containers)} 个图表容器")

        # 计算每个图表的宽度和水平位置
        total_width = 1760  # 总可用宽度
        gap = 24  # gap-6 = 24px

        # 获取flex布局信息
        container_style = container.get('style', '')
        if 'justify-content' in container_style:
            # 根据justify-content调整布局
            if 'center' in container_style:
                # 居中对齐
                chart_width = 400  # 每个图表的宽度
                total_charts_width = len(chart_containers) * chart_width + (len(chart_containers) - 1) * gap
                start_x = 80 + (total_width - total_charts_width) // 2
            elif 'space-between' in container_style:
                # 两端对齐
                chart_width = (total_width - (len(chart_containers) - 1) * gap) // len(chart_containers)
                start_x = 80
            else:
                # 默认平均分布
                chart_width = (total_width - (len(chart_containers) - 1) * gap) // len(chart_containers)
                total_charts_width = len(chart_containers) * chart_width + (len(chart_containers) - 1) * gap
                start_x = 80 + (total_width - total_charts_width) // 2
        else:
            # 默认平均分布
            chart_width = (total_width - (len(chart_containers) - 1) * gap) // len(chart_containers)
            total_charts_width = len(chart_containers) * chart_width + (len(chart_containers) - 1) * gap
            start_x = 80 + (total_width - total_charts_width) // 2

        current_y = y_start
        max_chart_height = 0

        # 处理每个图表容器（水平布局）
        for i, chart_container in enumerate(chart_containers):
            chart_x = start_x + i * (chart_width + gap)
            chart_y = current_y

            # 首先处理标题（智能识别h3或其他标题元素）
            title_elem = None
            title_text = None

            # 尝试多种标题选择器
            title_selectors = ['h3', 'h2', '.chart-title', '.title', 'h4']
            for selector in title_selectors:
                title_elem = chart_container.select_one(selector)
                if title_elem:
                    title_text = title_elem.get_text(strip=True)
                    if title_text:
                        logger.info(f"找到图表标题 ({selector}): {title_text}")
                        break

            if not title_elem:
                # 如果没有找到标准标题，尝试查找第一个包含文本的元素
                for child in chart_container.children:
                    if hasattr(child, 'name') and child.name in ['div', 'p']:
                        text = child.get_text(strip=True)
                        if text and len(text) < 50:  # 假设标题长度不超过50个字符
                            # 检查是否包含图表相关关键词
                            if any(keyword in text for keyword in ['分布', '统计', '图表', '分析', '趋势']):
                                title_elem = child
                                title_text = text
                                logger.info(f"通过内容识别找到图表标题: {title_text}")
                                break

            if title_elem and title_text:
                # 标题是容器的第一个元素，应该从容器顶部开始
                title_x = chart_x
                title_y = chart_y

                # 获取字体大小
                font_size_pt = self.style_computer.get_font_size_pt(title_elem)
                if not font_size_pt:
                    # 根据元素类型设置默认字体大小
                    if title_elem.name == 'h2':
                        font_size_pt = 24
                    elif title_elem.name == 'h3':
                        font_size_pt = 20
                    elif title_elem.name == 'h4':
                        font_size_pt = 18
                    else:
                        font_size_pt = 18

                # 计算标题高度（基于字体大小）
                title_height = int(font_size_pt * 1.5)  # 1.5倍行高

                # 智能检测文本对齐方式
                text_alignment = self._determine_title_text_alignment(title_elem)

                # 计算标题的margin-bottom
                title_classes = title_elem.get('class', [])
                if isinstance(title_classes, str):
                    title_classes = title_classes.split()

                margin_bottom = 16  # 默认margin-bottom
                for cls in title_classes:
                    if cls.startswith('mb-'):
                        try:
                            value = int(cls.replace('mb-', ''))
                            margin_bottom = value * 4  # Tailwind单位转换
                            break
                        except:
                            pass
                    elif cls.startswith('margin-bottom'):
                        # 解析内联样式
                        style_str = title_elem.get('style', '')
                        import re
                        mb_match = re.search(r'margin-bottom:\s*(\d+)px', style_str)
                        if mb_match:
                            margin_bottom = int(mb_match.group(1))

                text_box = pptx_slide.shapes.add_textbox(
                    UnitConverter.px_to_emu(title_x),
                    UnitConverter.px_to_emu(title_y),
                    UnitConverter.px_to_emu(chart_width),
                    UnitConverter.px_to_emu(title_height)
                )
                text_frame = text_box.text_frame
                text_frame.text = title_text
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                for paragraph in text_frame.paragraphs:
                    paragraph.alignment = text_alignment
                    for run in paragraph.runs:
                        run.font.size = Pt(font_size_pt)
                        run.font.bold = True
                        # 根据元素类型选择字体
                        if title_elem.name == 'h2':
                            run.font.name = self.font_manager.get_font('h2')
                        else:
                            run.font.name = self.font_manager.get_font('h3')

                        # 检查颜色类
                        if 'primary-color' in title_classes:
                            run.font.color.rgb = ColorParser.parse_color('rgb(10, 66, 117)')
                        elif 'text-gray-600' in title_classes:
                            run.font.color.rgb = RGBColor(102, 102, 102)
                        else:
                            run.font.color.rgb = ColorParser.get_text_color()

                # 更新SVG的Y位置（标题高度 + margin-bottom）
                chart_y += title_height + margin_bottom
            else:
                logger.warning(f"图表容器 {i+1} 中未找到标题元素")

            # 查找SVG元素
            svg_elem = chart_container.find('svg')

            if svg_elem:
                logger.info(f"处理第 {i+1} 个SVG图表")

                # 转换SVG图表
                chart_height = svg_converter.convert_svg(
                    svg_elem,
                    chart_container,
                    chart_x,
                    chart_y,
                    chart_width,
                    i
                )

                # 更新最大高度（包含标题）
                max_chart_height = max(max_chart_height, chart_y + chart_height - current_y)
            else:
                logger.warning(f"第 {i+1} 个图表容器中未找到SVG元素")
                max_chart_height = max(max_chart_height, 50)

        # 返回下一个元素的Y坐标（加上图表高度和间距）
        return current_y + max_chart_height + 40

    def _convert_content_container(self, container, pptx_slide, y_start, shape_converter):
        """
        转换内容容器（flex-1 overflow-hidden），处理所有子容器

        Args:
            container: 内容容器元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标
            shape_converter: 形状转换器

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理内容容器(flex-1 overflow-hidden)")

        current_y = y_start

        # 获取所有直接子元素（跳过文本节点）
        children = []
        for child in container.children:
            if hasattr(child, 'name') and child.name:
                children.append(child)

        logger.info(f"找到 {len(children)} 个子容器")

        # 处理每个子容器
        for i, child in enumerate(children):
            if i > 0:
                current_y += 40  # 子容器间距

            # 递归调用_process_container处理每个子容器
            current_y = self._process_container(child, pptx_slide, current_y, shape_converter)

        return current_y + 20

    def _convert_flex_container(self, container, pptx_slide, y_start, shape_converter):
        """
        转换flex容器

        Args:
            container: flex容器元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标
            shape_converter: 形状转换器

        Returns:
            下一个元素的Y坐标
        """
        # 获取所有子元素
        children = []
        for child in container.children:
            if hasattr(child, 'name') and child.name:
                children.append(child)

        current_y = y_start
        for child in children:
            child_classes = child.get('class', [])

            # 优先检测网格布局
            if 'grid' in child_classes:
                current_y = self._convert_grid_container(child, pptx_slide, current_y, shape_converter)
            elif 'data-card' in child_classes:
                current_y = self._convert_data_card(child, pptx_slide, shape_converter, current_y)
            elif 'stat-card' in child_classes:
                current_y = self._convert_stat_card(child, pptx_slide, current_y)
            else:
                # 降级处理
                current_y = self._convert_generic_card(child, pptx_slide, current_y, card_type='flex-item')

            # 添加间距
            current_y += 20

        return current_y

    def _convert_numbered_list_group(self, container, pptx_slide, y_start) -> int:
        """
        转换包含多个数字列表项的容器（如flex-1包含多个toc-item）

        Args:
            container: 容器元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理数字列表组容器")

        # 初始化文本转换器
        text_converter = TextConverter(pptx_slide, self.css_parser)

        # 获取所有toc-item
        toc_items = container.find_all('div', class_='toc-item')
        current_y = y_start

        for toc_item in toc_items:
            number_elem = toc_item.find('div', class_='toc-number')
            text_elem = toc_item.find('div', class_='toc-title')

            if number_elem and text_elem:
                numbered_item = {
                    'type': 'toc',
                    'container': toc_item,
                    'number_elem': number_elem,
                    'text_elem': text_elem,
                    'number': number_elem.get_text(strip=True),
                    'text': text_elem.get_text(strip=True)
                }
                current_y = text_converter.convert_numbered_list(numbered_item, 80, current_y)

        return current_y

    def _convert_centered_container(self, container, pptx_slide, y_start, shape_converter):
        """
        转换居中容器（flex justify-center items-center 或 flex-col justify-center）

        Args:
            container: 居中容器元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标
            shape_converter: 形状转换器

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理居中容器")

        # 获取容器的直接子元素
        children = []
        for child in container.children:
            if hasattr(child, 'name') and child.name:
                children.append(child)


        if not children:
            return y_start

        # 计算可用高度（从当前位置到底部）
        available_height = 1080 - y_start - 60  # 留出底部页码空间

        # 计算内容所需的总高度
        content_height = 0
        for child in children:
            child_classes = child.get('class', [])
            if 'data-card' in child_classes:
                content_height += 100  # 估算每个data-card高度
            elif 'stat-card' in child_classes:
                content_height += 220  # 估算每个stat-card高度
            else:
                content_height += 80  # 估算其他元素高度

        # 计算起始Y坐标（垂直居中）
        start_y = y_start
        if content_height < available_height:
            start_y = y_start + (available_height - content_height) // 2

        current_y = start_y

        # 处理每个子元素
        for child in children:
            child_classes = child.get('class', [])

            # 根据子元素类型调用相应的处理方法
            if 'data-card' in child_classes:
                # 对于居中容器中的data-card，使用特殊的处理方式（不带左边框）
                current_y = self._convert_centered_data_card(child, pptx_slide, current_y)
            elif 'stat-card' in child_classes:
                current_y = self._convert_stat_card(child, pptx_slide, current_y)
            else:
                # 检查是否包含嵌套的data-card
                nested_data_cards = child.find_all('div', class_='data-card')
                if nested_data_cards:
                    # 添加间距
                    if 'space-y-8' in child_classes:
                        current_y += 32  # space-y-8 ≈ 32px

                    # 处理每个嵌套的data-card
                    for nested_card in nested_data_cards:
                        current_y = self._convert_centered_data_card(nested_card, pptx_slide, current_y)
                        current_y += 32  # data-card之间的间距
                else:
                    # 处理其他元素
                    text = child.get_text(strip=True)
                    if text:
                        # 创建文本框
                        text_box = pptx_slide.shapes.add_textbox(
                            UnitConverter.px_to_emu(80),
                            UnitConverter.px_to_emu(current_y),
                            UnitConverter.px_to_emu(1760),
                            UnitConverter.px_to_emu(50)
                        )
                        text_frame = text_box.text_frame
                        text_frame.text = text
                        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                        for paragraph in text_frame.paragraphs:
                            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                            for run in paragraph.runs:
                                run.font.size = Pt(25)
                                run.font.name = self.font_manager.get_font('body')

                        current_y += 60

        return max(y_start + available_height, current_y + 40)

    def _convert_centered_data_card(self, card, pptx_slide, y_start: int) -> int:
        """
        转换居中容器中的data-card（智能识别是否需要左边框）

        Args:
            card: data-card元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理居中容器中的data-card")
        from src.utils.color_parser import ColorParser

        # 检查是否有max-w-2xl类，如果有则限制宽度
        card_classes = card.get('class', [])
        if 'max-w-2xl' in card_classes:
            # max-w-2xl在Tailwind中是42rem = 672px
            card_width = 672
            # 居中定位
            x_base = (1920 - card_width) // 2
        else:
            # 默认宽度，但留有左右边距
            card_width = 600  # 适中的宽度
            x_base = (1920 - card_width) // 2

        current_y = y_start + 10

        # 智能识别是否需要左边框
        # 检查CSS定义中是否有border-left
        border_style = self.css_parser.get_style('.data-card').get('border-left', '')
        has_left_border = bool(border_style)
        
        # 添加背景
        bg_color_str = self.css_parser.get_background_color('.data-card')
        if bg_color_str:
            from pptx.enum.shapes import MSO_SHAPE
            bg_shape = pptx_slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                UnitConverter.px_to_emu(x_base),
                UnitConverter.px_to_emu(y_start),
                UnitConverter.px_to_emu(card_width),
                UnitConverter.px_to_emu(80)  # 估算高度
            )
            bg_shape.fill.solid()
            bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
            if bg_rgb:
                if alpha < 1.0:
                    bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                bg_shape.fill.fore_color.rgb = bg_rgb
            bg_shape.line.fill.background()
            bg_shape.shadow.inherit = False

        # 如果需要左边框，添加左边框
        if has_left_border:
            # 解析边框样式
            border_width = 4  # 默认4px
            # 从border-left样式中提取颜色
            border_color_str = border_style.split('solid ')[-1].strip(')') if 'solid' in border_style else 'rgb(10, 66, 117)'
            # 确保颜色格式正确
            if not border_color_str.startswith('rgb'):
                border_color_str = f"rgb({border_color_str})"
            border_color = ColorParser.parse_color(border_color_str)
            from pptx.enum.shapes import MSO_SHAPE
            border_shape = pptx_slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                UnitConverter.px_to_emu(x_base),
                UnitConverter.px_to_emu(y_start),
                UnitConverter.px_to_emu(border_width),
                UnitConverter.px_to_emu(80)
            )
            border_shape.fill.solid()
            border_shape.fill.fore_color.rgb = border_color
            border_shape.line.fill.background()

        # 提取并渲染文本内容
        text_elements = []
        for elem in card.descendants:
            if hasattr(elem, 'name') and elem.name in ['p', 'h1', 'h2', 'h3', 'h4', 'div', 'span']:
                # 只提取没有子块级元素的文本节点
                if not elem.find_all(['div', 'p', 'h1', 'h2', 'h3']):
                    text = elem.get_text(strip=True)
                    if text and len(text) > 2:
                        text_elements.append(elem)

        # 渲染文本（如果有左边框，文本需要稍微右移）
        text_left_offset = 20 if has_left_border else 20
        text_left_offset += 8 if has_left_border else 0  # 左边框额外留出空间

        for elem in text_elements[:5]:  # 最多5个元素
            text = elem.get_text(strip=True)
            if text:
                # 文本框宽度要比卡片宽度小一些，留有内边距
                text_width = card_width - 40 - (8 if has_left_border else 0)  # 如果有左边框，减少文本宽度
                text_left = UnitConverter.px_to_emu(x_base + text_left_offset)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(text_width), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = text
                text_frame.word_wrap = True
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                for paragraph in text_frame.paragraphs:
                    paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                    for run in paragraph.runs:
                        # 使用样式计算器获取字体大小
                        font_size_px = self.style_computer.get_font_size_pt(elem)
                        run.font.size = Pt(font_size_px) if font_size_px else Pt(16)
                        run.font.name = self.font_manager.get_font('body')

                current_y += 35

        return current_y + 10

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

        # 防止重复处理：检查是否已经在其他容器中处理过
        if hasattr(card, '_processed'):
            logger.debug("stat-card已处理过，跳过")
            return y_start
        card._processed = True

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

        # 4. 检查是否包含新的HTML结构（h3 + p + p）
        h3_elem = card.find('h3')
        p_elements = card.find_all('p')
        h3_text = h3_elem.get_text(strip=True) if h3_elem else ""

        if h3_elem and len(p_elements) >= 2:
            logger.info(f"stat-card包含h3 + p + p结构，处理为数据卡片 (h3={h3_text})")
            return self._convert_modern_stat_card(card, pptx_slide, y_start)

        # 4.1 检查是否包含复杂结构（h3 + flex容器等）
        # 查找所有flex容器（不仅仅是直接子元素）
        flex_containers = card.find_all('div', class_='flex')
        logger.info(f"stat-card找到{len(flex_containers)}个flex容器")

        # 检查是否有flex容器包含risk-level标签（风险分布）
        risk_level_found = False
        risk_level_count = 0
        for flex_container in flex_containers:
            risk_levels = flex_container.find_all('span', class_='risk-level')
            risk_level_count += len(risk_levels)
            if risk_levels:
                risk_level_found = True
                logger.info(f"flex容器包含{len(risk_levels)}个risk-level标签")

        if risk_level_found:
            logger.info(f"stat-card包含风险等级标签（共{risk_level_count}个），使用增强处理 (h3={h3_text})")
            return self._convert_enhanced_stat_card(card, pptx_slide, y_start)

        # 5. 通用降级处理 - 提取所有文本内容
        logger.info("stat-card不包含已知结构,使用通用文本提取")
        return self._convert_generic_card(card, pptx_slide, y_start, card_type='stat-card')

    def _convert_modern_stat_card(self, card, pptx_slide, y_start: int) -> int:
        """
        转换现代样式的stat-card（h3 + p + p结构）

        Args:
            card: stat-card元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理现代样式stat-card")
        x_base = 80

        # 添加背景
        bg_color_str = self.css_parser.get_background_color('.stat-card')
        if bg_color_str:
            from pptx.enum.shapes import MSO_SHAPE
            bg_shape = pptx_slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                UnitConverter.px_to_emu(x_base),
                UnitConverter.px_to_emu(y_start),
                UnitConverter.px_to_emu(1760),
                UnitConverter.px_to_emu(200)  # 估算高度
            )
            bg_shape.fill.solid()
            bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
            if bg_rgb:
                if alpha < 1.0:
                    bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
                bg_shape.fill.fore_color.rgb = bg_rgb
            bg_shape.line.fill.background()
            bg_shape.shadow.inherit = False

        current_y = y_start + 20  # 顶部padding

        # 处理h3标题
        h3_elem = card.find('h3')
        if h3_elem:
            h3_text = h3_elem.get_text(strip=True)
            if h3_text:
                # 获取h3的字体大小和颜色
                h3_font_size_pt = self.style_computer.get_font_size_pt(h3_elem)
                h3_color = self._get_element_color(h3_elem) or ColorParser.get_primary_color()

                text_left = UnitConverter.px_to_emu(x_base + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = h3_text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用样式计算器获取字体大小和颜色
                        font_size_pt = self.style_computer.get_font_size_pt(h3_elem)
                        run.font.size = Pt(font_size_pt)

                        # 获取h3元素的颜色
                        h3_color = self._get_element_color(h3_elem)
                        if h3_color:
                            run.font.color.rgb = h3_color
                        else:
                            # 默认使用主题色
                            run.font.color.rgb = ColorParser.get_primary_color()

                        run.font.name = self.font_manager.get_font('body')
                        # 检查是否加粗
                        h3_classes = h3_elem.get('class', [])
                        if 'font-bold' in h3_classes:
                            run.font.bold = True

                current_y += 40

        # 处理第一个p标签（主要数据）
        p_elements = card.find_all('p')
        if len(p_elements) >= 1:
            p1_text = p_elements[0].get_text(strip=True)
            if p1_text:
                # 获取p1的字体大小和颜色
                p1_font_size_pt = self.style_computer.get_font_size_pt(p_elements[0])
                p1_color = self._get_element_color(p_elements[0])

                text_left = UnitConverter.px_to_emu(x_base + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(40)
                )
                text_frame = text_box.text_frame
                text_frame.text = p1_text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(p1_font_size_pt)
                        # 应用颜色
                        if p1_color:
                            run.font.color.rgb = p1_color
                        else:
                            # 如果没有特定颜色，使用默认文本颜色
                            run.font.color.rgb = ColorParser.get_text_color()
                        # 检查是否加粗
                        p1_classes = p_elements[0].get('class', [])
                        if 'font-bold' in p1_classes:
                            run.font.bold = True
                        run.font.name = self.font_manager.get_font('body')

                current_y += 50

        # 处理第二个p标签（描述）
        if len(p_elements) >= 2:
            p2_text = p_elements[1].get_text(strip=True)
            if p2_text:
                # 获取p2的字体大小和颜色
                p2_font_size_pt = self.style_computer.get_font_size_pt(p_elements[1])
                p2_color = self._get_element_color(p_elements[1])

                text_left = UnitConverter.px_to_emu(x_base + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = p2_text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(p2_font_size_pt)
                        # 应用颜色
                        if p2_color:
                            run.font.color.rgb = p2_color
                        else:
                            # 如果没有特定颜色，使用默认文本颜色
                            run.font.color.rgb = ColorParser.get_text_color()
                        run.font.name = self.font_manager.get_font('body')

                current_y += 40

        return y_start + 200  # 返回估算的卡片高度

    def _convert_enhanced_stat_card(self, card, pptx_slide, y_start: int) -> int:
        """
        转换增强样式stat-card（支持复杂内容结构，如flex布局、风险等级标签等）

        Args:
            card: stat-card元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理增强样式stat-card")
        x_base = 80

        # 添加背景
        bg_color_str = self.css_parser.get_background_color('.stat-card')
        if bg_color_str:
            from pptx.enum.shapes import MSO_SHAPE
            bg_shape = pptx_slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                UnitConverter.px_to_emu(x_base),
                UnitConverter.px_to_emu(y_start),
                UnitConverter.px_to_emu(1760),
                UnitConverter.px_to_emu(180)  # 估算高度
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
        border_left_style = self.css_parser.get_style('.stat-card').get('border-left', '')
        if '4px solid' in border_left_style:
            shape_converter = ShapeConverter(pptx_slide, self.css_parser)
            shape_converter.add_border_left(x_base, y_start, 180, 4)

        current_y = y_start + 20  # 顶部padding

        # 处理h3标题
        h3_elem = card.find('h3')
        if h3_elem:
            h3_text = h3_elem.get_text(strip=True)
            if h3_text:
                h3_font_size_pt = self.style_computer.get_font_size_pt(h3_elem)
                h3_color = self._get_element_color(h3_elem) or ColorParser.get_primary_color()

                text_left = UnitConverter.px_to_emu(x_base + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = h3_text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        font_size_pt = self.style_computer.get_font_size_pt(h3_elem)
                        run.font.size = Pt(font_size_pt)
                        h3_color = self._get_element_color(h3_elem)
                        if h3_color:
                            run.font.color.rgb = h3_color
                        else:
                            run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = self.font_manager.get_font('body')
                        # 检查是否加粗
                        h3_classes = h3_elem.get('class', [])
                        if 'font-bold' in h3_classes or 'text-2xl' in h3_classes:
                            run.font.bold = True

                current_y += 35

        # 处理复杂内容（查找所有flex容器内的内容）
        flex_containers = card.find_all('div', class_='flex')
        risk_levels = []
        for flex_container in flex_containers:
            # 查找所有包含risk-level的span
            risk_levels.extend(flex_container.find_all('span', class_='risk-level'))

        if risk_levels:

            # 计算每个风险标签的宽度和间距
            num_risks = len(risk_levels)
            if num_risks > 0:
                # 水平排列风险等级标签
                total_width = 1720  # 可用宽度
                risk_width = total_width // num_risks - 20  # 每个标签宽度，留出间距
                current_x = x_base + 20

                for risk_level in risk_levels:
                    risk_text = risk_level.get_text(strip=True)
                    risk_classes = risk_level.get('class', [])

                    # 获取风险等级的颜色
                    risk_color = None
                    bg_color = None
                    if 'risk-high' in risk_classes:
                        risk_color = ColorParser.parse_color('#dc2626')  # 红色
                        bg_color = RGBColor(252, 231, 229)  # 浅红色背景
                    elif 'risk-medium' in risk_classes:
                        risk_color = ColorParser.parse_color('#f59e0b')  # 橙色
                        bg_color = RGBColor(254, 243, 199)  # 浅橙色背景
                    elif 'risk-low' in risk_classes:
                        risk_color = ColorParser.parse_color('#3b82f6')  # 蓝色
                        bg_color = RGBColor(239, 246, 255)  # 浅蓝色背景

                    # 添加背景形状
                    if bg_color:
                        from pptx.enum.shapes import MSO_SHAPE
                        bg_shape = pptx_slide.shapes.add_shape(
                            MSO_SHAPE.ROUNDED_RECTANGLE,
                            UnitConverter.px_to_emu(current_x),
                            UnitConverter.px_to_emu(current_y),
                            UnitConverter.px_to_emu(risk_width),
                            UnitConverter.px_to_emu(35)
                        )
                        bg_shape.fill.solid()
                        bg_shape.fill.fore_color.rgb = bg_color
                        bg_shape.line.fill.background()

                    # 创建风险等级文本框
                    text_left = UnitConverter.px_to_emu(current_x + 5)
                    text_top = UnitConverter.px_to_emu(current_y + 5)
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(risk_width - 10), UnitConverter.px_to_emu(25)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = risk_text
                    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                    for paragraph in text_frame.paragraphs:
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(20)
                            run.font.bold = True
                            run.font.name = self.font_manager.get_font('body')

                            # 应用风险等级颜色
                            if risk_color:
                                run.font.color.rgb = risk_color

                    # 移动到下一个位置
                    current_x += risk_width + 20

                current_y += 45
        else:
            # 处理其他p标签
            p_elements = card.find_all('p')
            for p_elem in p_elements:
                p_text = p_elem.get_text(strip=True)
                if p_text:
                    p_font_size_pt = self.style_computer.get_font_size_pt(p_elem)
                    p_color = self._get_element_color(p_elem)

                    text_left = UnitConverter.px_to_emu(x_base + 20)
                    text_top = UnitConverter.px_to_emu(current_y)
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(25)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = p_text
                    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(p_font_size_pt)
                            if p_color:
                                run.font.color.rgb = p_color
                            run.font.name = self.font_manager.get_font('body')

                    current_y += 30

        return y_start + 180  # 返回估算的卡片高度

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
                        # 检查是否有其他颜色类
                        has_color_class = any(cls.startswith('text-') for cls in classes)

                        text_elements.append({
                            'text': text,
                            'tag': elem.name,
                            'is_primary': is_primary,
                            'is_bold': is_bold,
                            'has_color_class': has_color_class,
                            'element': elem  # 保存元素引用以获取颜色
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
            element = elem.get('element')  # 获取原始元素引用

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
                    elif element:
                        # 检查是否有其他颜色类
                        color = self._get_element_color(element)
                        if color:
                            run.font.color.rgb = color
                    run.font.name = self.font_manager.get_font('body')

            current_y += text_height + 10

        return current_y + 20

    def _convert_strategy_card(self, card, pptx_slide, y_start: int) -> int:
        """
        转换策略卡片(.strategy-card)

        处理action-item结构：圆形数字图标 + 标题 + 描述
        """

        # 防止重复处理：检查是否已经在其他容器中处理过
        if hasattr(card, '_processed'):
            logger.debug("strategy-card已处理过，跳过")
            return y_start
        card._processed = True

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

    def _convert_risk_card(self, card, pptx_slide, shape_converter, y_start: int) -> int:
        """
        转换风险卡片 (risk-card)

        Args:
            card: risk-card元素
            pptx_slide: PPTX幻灯片
            shape_converter: 形状转换器
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理risk-card风险卡片")

        x_base = 80
        card_width = 1760
        card_height = 180

        # 获取CSS样式
        card_style = self.css_parser.get_class_style('risk-card') or {}

        # 添加背景（渐变效果）
        bg_color_str = card_style.get('background', 'linear-gradient(135deg, rgba(239, 68, 68, 0.08) 0%, rgba(239, 68, 68, 0.02) 100%)')

        # 创建矩形背景
        bg_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            UnitConverter.px_to_emu(x_base),
            UnitConverter.px_to_emu(y_start),
            UnitConverter.px_to_emu(card_width),
            UnitConverter.px_to_emu(card_height)
        )
        bg_shape.fill.solid()

        # 解析渐变背景色，使用最深的颜色
        if 'rgba(239, 68, 68' in bg_color_str:
            # 红色系风险
            bg_rgb = RGBColor(254, 242, 242)  # 非常浅的红色
        elif 'rgba(251, 146' in bg_color_str:
            # 橙色系风险
            bg_rgb = RGBColor(255, 251, 235)  # 非常浅的橙色
        elif 'rgba(250, 204' in bg_color_str:
            # 黄色系风险
            bg_rgb = RGBColor(254, 252, 232)  # 非常浅的黄色
        else:
            # 默认浅灰色
            bg_rgb = RGBColor(249, 250, 251)

        bg_shape.fill.fore_color.rgb = bg_rgb
        bg_shape.line.fill.background()
        bg_shape.shadow.inherit = False

        # 添加左边框
        border_color_str = card_style.get('border-left-color', '#ef4444')
        border_color = ColorParser.parse_color(border_color_str)
        if not border_color:
            # 根据风险等级确定边框颜色
            border_color = ColorParser.parse_color('#ef4444')  # 默认红色

        border_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            UnitConverter.px_to_emu(x_base),
            UnitConverter.px_to_emu(y_start),
            UnitConverter.px_to_emu(4),
            UnitConverter.px_to_emu(card_height)
        )
        border_shape.fill.solid()
        border_shape.fill.fore_color.rgb = border_color
        border_shape.line.fill.background()

        current_y = y_start + 20

        # 处理flex布局内容
        flex_container = card.find('div', class_='flex')
        if flex_container:
            # 获取左侧内容区域
            left_div = flex_container.find('div', class_='flex-1')
            if left_div:
                # 处理风险标题
                title_div = left_div.find('div', class_='risk-title')
                if title_div:
                    # 获取图标
                    icon_elem = title_div.find('i')
                    icon_text = ""
                    icon_color = None

                    if icon_elem:
                        icon_classes = icon_elem.get('class', [])
                        # 根据图标类确定颜色
                        if 'severity-critical' in icon_classes:
                            icon_color = ColorParser.parse_color('#dc2626')
                            icon_text = "⚠"
                        elif 'severity-high' in icon_classes:
                            icon_color = ColorParser.parse_color('#ea580c')
                            icon_text = "⚠"
                        elif 'severity-medium' in icon_classes:
                            icon_color = ColorParser.parse_color('#d97706')
                            icon_text = "⚠"
                        else:
                            icon_text = "•"

                    # 获取标题文本
                    title_text = title_div.get_text(strip=True)
                    if icon_text:
                        title_text = title_text.replace(icon_text, "").strip()

                    # 添加标题文本
                    text_left = UnitConverter.px_to_emu(x_base + 20)
                    text_top = UnitConverter.px_to_emu(current_y)

                    if icon_text and icon_color:
                        # 如果有图标，创建两段式文本
                        text_box = pptx_slide.shapes.add_textbox(
                            text_left, text_top,
                            UnitConverter.px_to_emu(1200), UnitConverter.px_to_emu(35)
                        )
                        text_frame = text_box.text_frame
                        p = text_frame.paragraphs[0]

                        # 图标run
                        icon_run = p.add_run()
                        icon_run.text = icon_text + " "
                        icon_run.font.size = Pt(26)
                        icon_run.font.name = self.font_manager.get_font('body')
                        icon_run.font.color.rgb = icon_color
                        icon_run.font.bold = True

                        # 标题run
                        title_run = p.add_run()
                        title_run.text = title_text
                        title_run.font.size = Pt(26)
                        title_run.font.name = self.font_manager.get_font('body')
                        title_run.font.bold = True
                        title_run.font.color.rgb = RGBColor(51, 51, 51)  # 深灰色
                    else:
                        # 没有图标，直接添加标题
                        text_box = pptx_slide.shapes.add_textbox(
                            text_left, text_top,
                            UnitConverter.px_to_emu(1200), UnitConverter.px_to_emu(35)
                        )
                        text_frame = text_box.text_frame
                        text_frame.text = title_text

                        for paragraph in text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(26)
                                run.font.name = self.font_manager.get_font('body')
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(51, 51, 51)

                    current_y += 40

                # 处理风险描述
                desc_div = left_div.find('div', class_='risk-desc')
                if desc_div:
                    desc_text = desc_div.get_text(strip=True)
                    if desc_text:
                        text_box = pptx_slide.shapes.add_textbox(
                            UnitConverter.px_to_emu(x_base + 20),
                            UnitConverter.px_to_emu(current_y),
                            UnitConverter.px_to_emu(1000),
                            UnitConverter.px_to_emu(30)
                        )
                        text_frame = text_box.text_frame
                        text_frame.text = desc_text

                        for paragraph in text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(22)
                                run.font.name = self.font_manager.get_font('body')
                                run.font.color.rgb = RGBColor(102, 102, 102)  # 灰色

                        current_y += 35

                # 处理标签
                tag_div = left_div.find('div', class_='mt-3')
                if tag_div:
                    span_elem = tag_div.find('span')
                    if span_elem:
                        tag_text = span_elem.get_text(strip=True)
                        tag_classes = span_elem.get('class', [])

                        # 确定标签颜色
                        tag_bg_color = RGBColor(254, 226, 226)  # 浅红色背景
                        tag_text_color = RGBColor(153, 27, 27)  # 深红色文字

                        if 'bg-orange-100' in tag_classes:
                            tag_bg_color = RGBColor(255, 237, 213)  # 浅橙色背景
                            tag_text_color = RGBColor(154, 52, 18)  # 深橙色文字
                        elif 'bg-yellow-100' in tag_classes:
                            tag_bg_color = RGBColor(254, 249, 195)  # 浅黄色背景
                            tag_text_color = RGBColor(120, 53, 15)  # 深黄色文字

                        # 创建标签背景
                        tag_box = pptx_slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            UnitConverter.px_to_emu(x_base + 20),
                            UnitConverter.px_to_emu(current_y),
                            UnitConverter.px_to_emu(120),
                            UnitConverter.px_to_emu(30)
                        )
                        tag_box.fill.solid()
                        tag_box.fill.fore_color.rgb = tag_bg_color
                        tag_box.line.fill.background()

                        # 添加标签文本
                        tag_text_box = pptx_slide.shapes.add_textbox(
                            UnitConverter.px_to_emu(x_base + 20),
                            UnitConverter.px_to_emu(current_y + 3),
                            UnitConverter.px_to_emu(120),
                            UnitConverter.px_to_emu(24)
                        )
                        tag_text_frame = tag_text_box.text_frame
                        tag_text_frame.text = tag_text

                        for paragraph in tag_text_frame.paragraphs:
                            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                            for run in paragraph.runs:
                                run.font.size = Pt(14)
                                run.font.name = self.font_manager.get_font('body')
                                run.font.color.rgb = tag_text_color
                                run.font.bold = True

            # 获取右侧CVSS分数区域
            right_div = flex_container.find('div', class_='text-center')
            if right_div:
                # 获取CVSS分数
                score_div = right_div.find('div', class_='cvss-score')
                if score_div:
                    score_text = score_div.get_text(strip=True)

                    # 添加CVSS分数
                    score_box = pptx_slide.shapes.add_textbox(
                        UnitConverter.px_to_emu(x_base + 1400),
                        UnitConverter.px_to_emu(y_start + 50),
                        UnitConverter.px_to_emu(300),
                        UnitConverter.px_to_emu(60)
                    )
                    score_frame = score_box.text_frame
                    score_frame.text = score_text

                    for paragraph in score_frame.paragraphs:
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(48)
                            run.font.name = self.font_manager.get_font('body')
                            run.font.bold = True

                            # 根据分数确定颜色
                            if '10.0' in score_text or '9.' in score_text:
                                run.font.color.rgb = RGBColor(239, 68, 68)  # 红色
                            elif '8.' in score_text or '7.' in score_text:
                                run.font.color.rgb = RGBColor(234, 88, 12)  # 橙色
                            elif '6.' in score_text or '5.' in score_text:
                                run.font.color.rgb = RGBColor(217, 119, 6)  # 黄色
                            else:
                                run.font.color.rgb = RGBColor(107, 114, 128)  # 灰色

                # 获取CVSS标签
                label_div = right_div.find('div', class_='cvss-label')
                if label_div:
                    label_text = label_div.get_text(strip=True)

                    label_box = pptx_slide.shapes.add_textbox(
                        UnitConverter.px_to_emu(x_base + 1400),
                        UnitConverter.px_to_emu(y_start + 110),
                        UnitConverter.px_to_emu(300),
                        UnitConverter.px_to_emu(30)
                    )
                    label_frame = label_box.text_frame
                    label_frame.text = label_text

                    for paragraph in label_frame.paragraphs:
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(18)
                            run.font.name = self.font_manager.get_font('body')
                            run.font.color.rgb = RGBColor(102, 102, 102)

        return y_start + card_height + 20

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

        # 检查grid布局列数（支持Tailwind CSS网格类）
        grid_container = card.find('div', class_='grid')
        if grid_container:
            grid_classes = grid_container.get('class', [])
            num_columns = 2  # 默认2列

            # 检查Tailwind CSS网格列类
            for cls in grid_classes:
                if cls.startswith('grid-cols-') and hasattr(self.css_parser, 'tailwind_grid_columns'):
                    columns = self.css_parser.tailwind_grid_columns.get(cls)
                    if columns:
                        num_columns = columns
                        logger.info(f"检测到Tailwind网格列类: {cls} -> {num_columns}列")
                        break
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

                # 检查文本中是否有strong标签
                strong_elem = p_elem.find('strong')

                if strong_elem:
                    # 处理带strong的文本
                    strong_text = strong_elem.get_text(strip=True)
                    remaining_text = text.replace(strong_text, '').strip()

                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(text_width), UnitConverter.px_to_emu(25)
                    )
                    text_frame = text_box.text_frame
                    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                    text_frame.clear()
                    p = text_frame.paragraphs[0]

                    # 添加加粗部分
                    if strong_text:
                        strong_run = p.add_run()
                        strong_run.text = strong_text
                        strong_run.font.size = Pt(16)
                        strong_run.font.bold = True
                        strong_run.font.name = self.font_manager.get_font('body')

                    # 添加剩余部分
                    if remaining_text:
                        normal_run = p.add_run()
                        normal_run.text = remaining_text
                        normal_run.font.size = Pt(16)
                        normal_run.font.name = self.font_manager.get_font('body')
                else:
                    # 普通文本
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(text_width), UnitConverter.px_to_emu(50)  # 增加高度以支持换行
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                    text_frame.word_wrap = True
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(16)
                            run.font.name = self.font_manager.get_font('body')

        return current_y + 30

    def _convert_data_card(self, card, pptx_slide, shape_converter, y_start: int) -> int:
        """转换数据卡片(.data-card)"""

        # 防止重复处理：检查是否已经在其他容器中处理过
        # if hasattr(card, '_processed'):
        #     logger.info("data-card已处理过，跳过")
        #     return y_start
        # card._processed = True

        x_base = 80

        # 检查data-card内是否包含网格布局
        grid_container = card.find('div', class_='grid')
        if grid_container and grid_container.find_all('div', class_='bullet-point'):
            # 处理包含bullet-point的网格布局
            logger.info(f"data-card内发现grid和bullet-point，使用网格布局处理")
            return self._convert_data_card_grid_layout(card, grid_container, pptx_slide, shape_converter, y_start)
        else:
            logger.info(f"data-card使用标准处理流程（grid: {'是' if grid_container else '否'}, bullet-point: {'是' if grid_container and grid_container.find_all('div', class_='bullet-point') else '否'}）")

        # 添加data-card背景色
        bg_color_str = 'rgba(10, 66, 117, 0.03)'  # 从CSS获取的背景色
        from pptx.enum.shapes import MSO_SHAPE
        # 估算高度
        estimated_height = 200
        bg_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            UnitConverter.px_to_emu(x_base),
            UnitConverter.px_to_emu(y_start),
            UnitConverter.px_to_emu(1760),
            UnitConverter.px_to_emu(estimated_height)
        )
        bg_shape.fill.solid()
        bg_rgb, alpha = ColorParser.parse_rgba(bg_color_str)
        if bg_rgb:
            if alpha < 1.0:
                bg_rgb = ColorParser.blend_with_white(bg_rgb, alpha)
            bg_shape.fill.fore_color.rgb = bg_rgb
        bg_shape.line.fill.background()
        logger.info(f"添加data-card背景色: {bg_color_str}")

        # 注意：左边框的高度需要在计算完实际内容后再添加
        # 暂时记录起始位置，稍后添加边框

        # 初始化当前Y坐标
        current_y = y_start + 10

        # 检查是否包含cve-card，如果有则跳过标题处理，让专门的CVE方法处理
        cve_cards = card.find_all('div', class_='cve-card')

        # 初始化标题变量（用于后面的检查）
        title_elem = None
        title_text = None

        if not cve_cards:
            # === 修复：简化的标题和内容处理逻辑 ===
            # 1. 首先查找并处理标题（查找h3标签或primary-color的p标签）
            title_elem = card.find('h3')

            # 如果没找到h3，再查找primary-color的p标签
            if not title_elem:
                title_elem = card.find('p', class_='primary-color')

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

                            # 设置颜色和字体
                            if title_elem.name == 'h3':
                                # h3标签使用主题色
                                run.font.color.rgb = ColorParser.get_primary_color()
                                run.font.bold = True
                            else:
                                # p标签根据class设置颜色
                                run.font.color.rgb = ColorParser.get_primary_color()

                            run.font.name = self.font_manager.get_font('body')

                    current_y += 40  # 标题后间距
                    logger.info(f"渲染data-card标题: {title_text}")

        # 2. 处理普通段落内容（明确排除标题元素、bullet-point内的元素和cve-card内的元素）
        content_paragraphs = []
        all_paragraphs = card.find_all('p')

        for p in all_paragraphs:
            # 新增：检查是否在bullet-point里
            parent = p.parent
            is_in_bullet_point = False
            is_in_cve_card = False
            while parent and parent != card:
                if parent.get('class'):
                    if 'bullet-point' in parent.get('class', []):
                        is_in_bullet_point = True
                        break
                    elif 'cve-card' in parent.get('class', []):
                        is_in_cve_card = True
                        break
                parent = parent.parent

            if is_in_bullet_point or is_in_cve_card:
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

        # 列表项 (bullet-point)
        bullet_points = card.find_all('div', class_='bullet-point')

        # 风险项目 (risk-item)
        risk_items = card.find_all('div', class_='risk-item')

        # === 修复：正确判断是否已有内容 ===
        # 不仅要检查progress-bar和bullet-point，还要检查是否已经处理了标题和段落
        has_title_or_content = title_elem is not None or len(content_paragraphs) > 0
        has_special_content = len(progress_bars) > 0 or len(bullet_points) > 0 or len(risk_items) > 0
        has_content = has_title_or_content or has_special_content

        logger.info(f"data-card内容检查: 标题={'是' if title_elem else '否'}, "
                   f"内容段落数={len(content_paragraphs)}, 进度条数={len(progress_bars)}, "
                   f"列表项数={len(bullet_points)}, 风险项数={len(risk_items)}, 总已有内容={'是' if has_content else '否'}")

        # 如果没有其他内容，从y_start开始处理bullet-point
        if not has_title_or_content and not progress_bars:
            progress_y = y_start + 10
        else:
            progress_y = current_y + 10 if current_y > y_start else current_y

        # 处理风险项目
        for risk_item in risk_items:
            icon_elem = risk_item.find('i')
            content_div = risk_item.find('div')

            if icon_elem and content_div:
                # 获取图标
                icon_classes = icon_elem.get('class', [])
                icon_char = self._get_icon_char(icon_classes)

                # 获取所有p标签
                p_tags = content_div.find_all('p')

                # 先处理第一个p标签的内容，计算实际需要的宽度
                total_text_width = 0
                has_inline_risk_level = False
                text_elements = []

                if len(p_tags) > 0:
                    first_p = p_tags[0]

                    # 智能识别内联元素位置的规则
                    # 如果span.risk-level紧跟在strong后面，应该在同一行显示
                    for elem in first_p.children:
                        if hasattr(elem, 'name'):
                            if elem.name == 'span' and 'risk-level' in elem.get('class', []):
                                # 检查前一个兄弟元素是否是strong
                                prev_sibling = elem.previous_sibling
                                if prev_sibling and hasattr(prev_sibling, 'name') and prev_sibling.name == 'strong':
                                    has_inline_risk_level = True
                                    break

                    # 计算图标宽度
                    icon_width = 0
                    if icon_char:
                        icon_width = self._calculate_text_width(icon_char + " ", Pt(22))

                    # 遍历first_p的所有直接子元素，计算文本宽度
                    current_x = icon_width
                    for elem in first_p.children:
                        if hasattr(elem, 'name'):
                            if elem.name == 'strong':
                                strong_text = elem.get_text(strip=True)
                                if strong_text:
                                    text_width = self._calculate_text_width(strong_text, Pt(22))
                                    text_elements.append({
                                        'type': 'strong',
                                        'text': strong_text,
                                        'font_size': Pt(22),
                                        'x_start': current_x,
                                        'x_end': current_x + text_width
                                    })
                                    current_x += text_width
                            elif elem.name == 'span' and 'risk-level' in elem.get('class', []):
                                risk_text = elem.get_text(strip=True)
                                risk_classes = elem.get('class', [])
                                # 在strong和risk-level之间添加空格
                                space_width = self._calculate_text_width(" ", Pt(22))
                                current_x += space_width
                                text_width = self._calculate_text_width(risk_text, Pt(20))
                                text_elements.append({
                                    'type': 'risk',
                                    'text': risk_text,
                                    'classes': risk_classes,
                                    'font_size': Pt(20),
                                    'x_start': current_x,
                                    'x_end': current_x + text_width,
                                    'inline': True  # 标记为内联元素
                                })
                                current_x += text_width
                        else:
                            # 处理文本节点
                            text_content = str(elem).strip()
                            if text_content:
                                text_width = self._calculate_text_width(text_content, Pt(22))
                                text_elements.append({
                                    'type': 'text',
                                    'text': text_content,
                                    'font_size': Pt(22),
                                    'x_start': current_x,
                                    'x_end': current_x + text_width
                                })
                                current_x += text_width

                    total_text_width = current_x

                # 计算文本框宽度（自适应文本长度）
                min_width = 400  # 最小宽度
                max_width = 1720  # 最大宽度
                box_width = max(min_width, min(total_text_width + 40, max_width))  # 加40px的padding

                # 计算所需的高度
                # 第一行：strong + risk-level（如果有内联）
                first_line_height = 35 if has_inline_risk_level else 30
                # 其他行：每个p标签占一行
                other_lines_height = (len(p_tags) - 1) * 28  # 每个额外的p标签28px
                total_height = first_line_height + other_lines_height + 20  # 20px padding

                # 创建自适应大小的文本框
                # 注意：risk_left 已经在前面设置为 UnitConverter.px_to_emu(x + 20)
                # 不要重复设置，保留正确的x坐标
                risk_top = UnitConverter.px_to_emu(progress_y)
                risk_box = pptx_slide.shapes.add_textbox(
                    risk_left, risk_top,
                    UnitConverter.px_to_emu(box_width), UnitConverter.px_to_emu(total_height)
                )
                risk_frame = risk_box.text_frame
                risk_frame.clear()
                risk_frame.margin_left = 0
                risk_frame.margin_right = 0
                risk_frame.margin_top = 0
                risk_frame.margin_bottom = 0
                p = risk_frame.paragraphs[0]

                # 添加图标
                if icon_char:
                    icon_run = p.add_run()
                    icon_run.text = icon_char + " "
                    icon_run.font.size = Pt(22)
                    icon_run.font.color.rgb = ColorParser.get_primary_color()
                    icon_run.font.name = self.font_manager.get_font('body')

                # 处理第一个p标签（可能包含strong和risk-level）
                if len(p_tags) > 0:
                    first_p = p_tags[0]

                    # 遍历first_p的所有直接子元素
                    for elem in first_p.children:
                        if hasattr(elem, 'name'):
                            if elem.name == 'strong':
                                strong_text = elem.get_text(strip=True)
                                if strong_text:
                                    strong_run = p.add_run()
                                    strong_run.text = strong_text
                                    strong_run.font.size = Pt(22)
                                    strong_run.font.bold = True
                                    strong_run.font.name = self.font_manager.get_font('body')
                                    strong_run.font.color.rgb = RGBColor(0, 0, 0)  # 黑色

                                    # 检查下一个元素是否是risk-level，如果是则添加空格
                                    next_sibling = elem.next_sibling
                                    if next_sibling and hasattr(next_sibling, 'name') and next_sibling.name == 'span' and 'risk-level' in next_sibling.get('class', []):
                                        strong_run.text += " "
                            elif elem.name == 'span' and 'risk-level' in elem.get('class', []):
                                risk_text = elem.get_text(strip=True)
                                risk_classes = elem.get('class', [])

                                # 获取风险等级的颜色和背景色
                                risk_color = None
                                bg_color = None
                                if 'risk-high' in risk_classes:
                                    risk_color = ColorParser.parse_color('#dc2626')  # 红色
                                    bg_color = RGBColor(252, 231, 229)  # 浅红色背景
                                elif 'risk-medium' in risk_classes:
                                    risk_color = ColorParser.parse_color('#f59e0b')  # 橙色
                                    bg_color = RGBColor(254, 243, 199)  # 浅橙色背景
                                elif 'risk-low' in risk_classes:
                                    risk_color = ColorParser.parse_color('#3b82f6')  # 蓝色
                                    bg_color = RGBColor(239, 246, 255)  # 浅蓝色背景
                                elif 'CVSS' in risk_text:
                                    # CVSS分数也使用特殊颜色
                                    if '10.0' in risk_text:
                                        risk_color = ColorParser.parse_color('#dc2626')  # 红色
                                        bg_color = RGBColor(252, 231, 229)  # 浅红色背景
                                    elif '9.8' in risk_text:
                                        risk_color = ColorParser.parse_color('#dc2626')  # 红色
                                        bg_color = RGBColor(252, 231, 229)  # 浅红色背景
                                    elif '8.6' in risk_text:
                                        risk_color = ColorParser.parse_color('#f59e0b')  # 橙色
                                        bg_color = RGBColor(254, 243, 199)  # 浅橙色背景

                                # 不添加到主文本框，而是创建独立的带背景文本框
                                # 这样"高危"会紧跟在strong后面，并有自己的背景
                                if elem_info:
                                    # 计算文本的绝对位置
                                    # elem_info['x_start'] 是相对于文本框的像素位置
                                    text_abs_left = risk_left + UnitConverter.px_to_emu(elem_info['x_start'])
                                    text_abs_top = risk_top + UnitConverter.px_to_emu(5)  # 微调垂直位置

                                    # 计算文本宽度
                                    text_width = elem_info['x_end'] - elem_info['x_start']
                                    bg_width = text_width + 16  # 左右各8px padding
                                    bg_width = max(bg_width, 50)  # 最小宽度50px

                                    # 先创建背景形状（如果有背景色）
                                    if bg_color:
                                        bg_shape = pptx_slide.shapes.add_shape(
                                            MSO_SHAPE.ROUNDED_RECTANGLE,
                                            text_abs_left,
                                            text_abs_top,
                                            UnitConverter.px_to_emu(bg_width),
                                            UnitConverter.px_to_emu(28)
                                        )
                                        bg_shape.fill.solid()
                                        bg_shape.fill.fore_color.rgb = bg_color
                                        bg_shape.line.fill.background()

                                    # 再创建文本框（覆盖在背景上）
                                    risk_text_box = pptx_slide.shapes.add_textbox(
                                        text_abs_left + UnitConverter.px_to_emu(8),  # 左内边距
                                        text_abs_top + UnitConverter.px_to_emu(4),   # 上内边距
                                        UnitConverter.px_to_emu(bg_width - 16),  # 减去padding
                                        UnitConverter.px_to_emu(20)  # 高度
                                    )
                                    risk_text_frame = risk_text_box.text_frame
                                    risk_text_frame.clear()
                                    risk_text_frame.text = risk_text
                                    risk_text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                                    # 设置文本样式
                                    for paragraph in risk_text_frame.paragraphs:
                                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                                        for run in paragraph.runs:
                                            run.font.size = Pt(20)
                                            run.font.bold = True
                                            run.font.name = self.font_manager.get_font('body')
                                            if risk_color:
                                                run.font.color.rgb = risk_color
                                            else:
                                                run.font.color.rgb = RGBColor(220, 38, 38)  # 默认红色

                                    logger.info(f"创建独立文本框: {risk_text}")
                                    logger.info(f"  绝对位置: ({UnitConverter.emu_to_px(text_abs_left)}, {UnitConverter.emu_to_px(text_abs_top)})")
                                    logger.info(f"  尺寸: {bg_width}px x 28px")

                                    # 将背景移到下层（这样不会覆盖文本）
                                    # 在python-pptx中，后添加的形状在上层
                                    # 所以我们需要把背景移到前面添加的元素后面
                                    try:
                                        # 获取所有形状，调整z-order
                                        shapes = pptx_slide.shapes
                                        bg_index = len(shapes) - 1  # 最后添加的背景
                                        # 找到文本框的索引
                                        text_index = None
                                        for i in range(len(shapes)):
                                            if shapes[i] == risk_box:
                                                text_index = i
                                                break
                                        # 如果找到了文本框，把背景移到它前面
                                        if text_index is not None and bg_index > text_index:
                                            # 需要重新创建顺序
                                            pass  # python-pptx不支持直接调整z-order
                                    except:
                                        pass

                # 处理第二个p标签（描述信息）
                if len(p_tags) > 1:
                    second_p = p_tags[1]
                    desc_text = second_p.get_text(strip=True)
                    if desc_text:
                        # 添加换行
                        p.add_run().text = "\n"

                        desc_run = p.add_run()
                        desc_run.text = desc_text
                        desc_run.font.size = Pt(18)
                        desc_run.font.name = self.font_manager.get_font('body')
                        desc_run.font.color.rgb = RGBColor(102, 102, 102)  # 灰色

                progress_y += total_height + 10  # 使用计算的高度+间距

        for bullet in bullet_points:
            # 检查是否有图标
            icon_elem = bullet.find('i')
            icon_char = None
            if icon_elem:
                icon_classes = icon_elem.get('class', [])
                icon_char = self._get_icon_char(icon_classes)

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
                        prefix = f"{icon_char} " if icon_char else "• "
                        bullet_width = 1720
                    else:
                        prefix = "  "
                        bullet_width = 1720

                    bullet_box = pptx_slide.shapes.add_textbox(
                        bullet_left, bullet_top,
                        UnitConverter.px_to_emu(bullet_width), UnitConverter.px_to_emu(50)
                    )
                    bullet_frame = bullet_box.text_frame
                    bullet_frame.clear()
                    paragraph = bullet_frame.paragraphs[0]

                    # 处理冒号后的换行
                    if '：' in text or ':' in text:
                        # 分割文本为两部分
                        if '：' in text:
                            parts = text.split('：', 1)
                            separator = '：'
                        else:
                            parts = text.split(':', 1)
                            separator = ':'

                        if len(parts) == 2:
                            # 添加图标和第一部分（加粗）
                            run1 = paragraph.add_run()
                            run1.text = f"{prefix}{parts[0]}{separator}"
                            run1.font.bold = True
                            run1.font.size = Pt(25)
                            run1.font.name = self.font_manager.get_font('body')

                            # 获取字体大小
                            font_size_pt = self.style_computer.get_font_size_pt(p)
                            if font_size_pt:
                                run1.font.size = Pt(font_size_pt)

                            # 添加换行和第二部分
                            run2 = paragraph.add_run()
                            run2.text = "\n" + parts[1]
                            run2.font.size = Pt(25)
                            run2.font.name = self.font_manager.get_font('body')

                            if font_size_pt:
                                run2.font.size = Pt(font_size_pt)
                        else:
                            run = paragraph.add_run()
                            run.text = f"{prefix}{text}"
                            run.font.size = Pt(25)
                            run.font.name = self.font_manager.get_font('body')
                            # 获取实际字体大小
                            font_size_pt = self.style_computer.get_font_size_pt(p)
                            if font_size_pt:
                                run.font.size = Pt(font_size_pt)
                            # 第一个p加粗
                            if idx == 0:
                                run.font.bold = True
                    else:
                        run = paragraph.add_run()
                        run.text = f"{prefix}{text}"
                        run.font.size = Pt(25)
                        run.font.name = self.font_manager.get_font('body')
                        # 获取实际字体大小
                        font_size_pt = self.style_computer.get_font_size_pt(p)
                        if font_size_pt:
                            run.font.size = Pt(font_size_pt)
                        # 第一个p加粗
                        if idx == 0:
                            run.font.bold = True

                    bullet_frame.word_wrap = True

                    progress_y += 28 if idx == 0 else 50
            else:
                # 处理简单结构: <div class="bullet-point"><i>...</i><p>...</p></div>
                p = bullet.find('p')
                if p:
                    text = p.get_text(strip=True)
                    bullet_left = UnitConverter.px_to_emu(x_base + 20)
                    bullet_top = UnitConverter.px_to_emu(progress_y)
                    bullet_box = pptx_slide.shapes.add_textbox(
                        bullet_left, bullet_top,
                        UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
                    )
                    # 使用图标或默认圆点
                    prefix = f"{icon_char} " if icon_char else "• "
                    bullet_frame = bullet_box.text_frame
                    bullet_frame.text = f"{prefix}{text}"
                    for paragraph in bullet_frame.paragraphs:
                        for run in paragraph.runs:
                            # 使用样式计算器获取正确的字体大小
                            font_size_px = self.style_computer.get_font_size_pt(p)
                            run.font.size = Pt(font_size_px)
                            run.font.name = self.font_manager.get_font('body')

                    progress_y += 35

        # 检查是否包含cve-card（使用前面已经检测的结果）
        if cve_cards:
            logger.info(f"检测到{len(cve_cards)}个cve-card，使用专门处理")
            return self._convert_cve_card_list(card, pptx_slide, shape_converter, y_start)

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

    def _convert_cve_card_list(self, card, pptx_slide, shape_converter, y_start: int) -> int:
        """
        转换包含CVE卡片列表的data-card

        Args:
            card: data-card元素
            pptx_slide: PPTX幻灯片
            shape_converter: 形状转换器
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("开始处理CVE卡片列表")

        x_base = 80
        current_y = y_start
        width = 1760

        # 添加data-card背景色
        bg_color_str = 'rgba(10, 66, 117, 0.03)'
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.dml.color import RGBColor
        from pptx.util import Pt
        from src.utils.unit_converter import UnitConverter
        from src.utils.color_parser import ColorParser

        # 计算总高度
        # 基础padding: 20px上下 = 40px
        total_height = 40

        # 处理h3标题（使用动态字号）
        h3_elem = card.find('h3')
        if h3_elem:
            title_text = h3_elem.get_text(strip=True)
            if title_text:
                # 渲染标题
                text_left = UnitConverter.px_to_emu(x_base + 20)
                text_top = UnitConverter.px_to_emu(current_y)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(width - 40), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = title_text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用样式计算器动态获取字体大小
                        font_size_px = self.style_computer.get_font_size_pt(h3_elem)
                        run.font.size = Pt(font_size_px) if font_size_px else Pt(20)
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.bold = True
                        run.font.name = self.font_manager.get_font('body')

                current_y += 40  # 标题后间距
                total_height += 40
                logger.info(f"渲染CVE列表标题: {title_text} (字号: {font_size_px if font_size_px else 20}px)")

        # 处理所有cve-card
        cve_cards = card.find_all('div', class_='cve-card')
        logger.info(f"找到{len(cve_cards)}个CVE卡片")

        for i, cve_card in enumerate(cve_cards):
            # 计算每个cve-card的高度
            card_height = self._convert_single_cve_card(cve_card, pptx_slide, shape_converter,
                                                      x_base, current_y, width)
            current_y = card_height
            total_height += card_height - current_y + 15  # 15px是cve-card之间的间距

            # 最后一个卡片不需要额外间距
            if i == len(cve_cards) - 1:
                total_height -= 15

        # 添加data-card的左边框
        shape_converter.add_border_left(x_base, y_start, total_height, 4)

        logger.info(f"CVE卡片列表处理完成，总高度: {total_height}px")
        return y_start + total_height

    def _convert_single_cve_card(self, card, pptx_slide, shape_converter, x, y, width) -> int:
        """
        转换单个CVE卡片

        Args:
            card: cve-card元素
            pptx_slide: PPTX幻灯片
            shape_converter: 形状转换器
            x: X坐标
            y: Y坐标
            width: 宽度

        Returns:
            下一个元素的Y坐标
        """
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.dml.color import RGBColor
        from pptx.util import Pt
        from src.utils.unit_converter import UnitConverter
        from src.utils.color_parser import ColorParser
        from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

        # CVE卡片的padding: 20px
        padding = 20
        content_width = width - padding * 2
        current_y = y + padding

        # 创建CVE卡片背景（渐变效果用单色代替）
        bg_shape = pptx_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            UnitConverter.px_to_emu(x),
            UnitConverter.px_to_emu(y),
            UnitConverter.px_to_emu(width),
            UnitConverter.px_to_emu(100)  # 初始高度，后续会调整
        )
        bg_shape.fill.solid()
        # 使用浅色背景
        bg_shape.fill.fore_color.rgb = RGBColor(248, 250, 252)  # 相当于rgba(10, 66, 117, 0.05)
        # 不设置边框，后续单独添加左边框
        bg_shape.line.fill.background()

        # 处理主要内容区域
        main_content = card.find('div', class_='flex-1')
        if not main_content:
            main_content = card

        # 处理徽章区域
        badge_area = main_content.find('div', class_='flex')
        if badge_area and 'items-center' in badge_area.get('class', []) and 'mb-2' in badge_area.get('class', []):
            badge_y = current_y
            badge_x = x + padding

            # 处理所有徽章
            badges_processed = 0
            for child in badge_area.children:
                if hasattr(child, 'name') and child.name == 'span':
                    badge_text = child.get_text(strip=True)
                    if badge_text:
                        badge_classes = child.get('class', [])

                        # 确定徽章颜色
                        bg_color = RGBColor(255, 255, 255)  # 默认白色
                        text_color = RGBColor(0, 0, 0)  # 默认黑色

                        if 'critical' in badge_classes:
                            bg_color = RGBColor(252, 231, 229)  # 浅红色
                            text_color = RGBColor(220, 38, 38)
                        elif 'high' in badge_classes:
                            bg_color = RGBColor(254, 243, 199)  # 浅橙色
                            text_color = RGBColor(251, 146, 60)
                        elif 'medium' in badge_classes:
                            bg_color = RGBColor(254, 252, 224)  # 浅黄色
                            text_color = RGBColor(251, 191, 36)
                        elif 'exploited' in badge_classes:
                            bg_color = RGBColor(252, 231, 229)  # 浅红色
                            text_color = RGBColor(220, 38, 38)

                        # 计算徽章宽度
                        badge_width = len(badge_text) * 12 + 24  # 估算宽度

                        # 创建徽章背景
                        badge_bg = pptx_slide.shapes.add_shape(
                            MSO_SHAPE.ROUNDED_RECTANGLE,
                            UnitConverter.px_to_emu(badge_x),
                            UnitConverter.px_to_emu(badge_y),
                            UnitConverter.px_to_emu(badge_width),
                            UnitConverter.px_to_emu(24)
                        )
                        badge_bg.fill.solid()
                        badge_bg.fill.fore_color.rgb = bg_color
                        badge_bg.line.fill.background()

                        # 创建徽章文本
                        badge_text_box = pptx_slide.shapes.add_textbox(
                            UnitConverter.px_to_emu(badge_x),
                            UnitConverter.px_to_emu(badge_y + 2),
                            UnitConverter.px_to_emu(badge_width),
                            UnitConverter.px_to_emu(20)
                        )
                        badge_frame = badge_text_box.text_frame
                        badge_frame.text = badge_text
                        for paragraph in badge_frame.paragraphs:
                            paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                            for run in paragraph.runs:
                                # 使用动态字号，而不是硬编码14px
                                badge_font_size = self.style_computer.get_font_size_pt(child)
                                run.font.size = Pt(badge_font_size) if badge_font_size else Pt(14)
                                run.font.bold = True
                                run.font.color.rgb = text_color
                                run.font.name = self.font_manager.get_font('body')

                        badge_x += badge_width + 10
                        badges_processed += 1

            if badges_processed > 0:
                current_y += 35  # 徽章区域高度

        # 处理漏洞名称
        name_p = main_content.find('p', class_=lambda x: x and 'font-medium' in x)
        if name_p:
            name_text = name_p.get_text(strip=True)
            if name_text:
                name_box = pptx_slide.shapes.add_textbox(
                    UnitConverter.px_to_emu(x + padding),
                    UnitConverter.px_to_emu(current_y),
                    UnitConverter.px_to_emu(content_width),
                    UnitConverter.px_to_emu(30)
                )
                name_frame = name_box.text_frame
                name_frame.text = name_text
                for paragraph in name_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用动态字号，而不是硬编码18px
                        name_font_size = self.style_computer.get_font_size_pt(name_p)
                        run.font.size = Pt(name_font_size) if name_font_size else Pt(18)
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.name = self.font_manager.get_font('body')

                current_y += 30

        # 处理受影响资产
        asset_p = main_content.find('p', class_=lambda x: x and 'text-gray-600' in x)
        if asset_p:
            asset_text = asset_p.get_text(strip=True)
            if asset_text:
                asset_box = pptx_slide.shapes.add_textbox(
                    UnitConverter.px_to_emu(x + padding),
                    UnitConverter.px_to_emu(current_y),
                    UnitConverter.px_to_emu(content_width),
                    UnitConverter.px_to_emu(25)
                )
                asset_frame = asset_box.text_frame
                asset_frame.text = asset_text
                for paragraph in asset_frame.paragraphs:
                    for run in paragraph.runs:
                        # 使用动态字号，而不是硬编码16px
                        asset_font_size = self.style_computer.get_font_size_pt(asset_p)
                        run.font.size = Pt(asset_font_size) if asset_font_size else Pt(16)
                        run.font.color.rgb = RGBColor(102, 102, 102)
                        run.font.name = self.font_manager.get_font('body')

                current_y += 25

        # 处理右侧图标
        icon_elem = card.find('i')
        if icon_elem:
            icon_classes = icon_elem.get('class', [])
            icon_char = self._get_icon_char(icon_classes)
            if icon_char:
                # 简单处理：在右侧添加图标文本
                icon_box = pptx_slide.shapes.add_textbox(
                    UnitConverter.px_to_emu(x + width - 60),
                    UnitConverter.px_to_emu(y + 30),
                    UnitConverter.px_to_emu(40),
                    UnitConverter.px_to_emu(40)
                )
                icon_frame = icon_box.text_frame
                icon_frame.text = icon_char
                for paragraph in icon_frame.paragraphs:
                    paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(36)
                        run.font.color.rgb = RGBColor(220, 38, 38)  # 红色
                        run.font.name = self.font_manager.get_font('body')

        # 调整背景高度
        card_height = current_y - y + padding
        bg_shape.height = UnitConverter.px_to_emu(card_height)

        # 添加左边框
        shape_converter.add_border_left(x, y, card_height, 4)

        return y + card_height

    def _convert_data_card_grid_layout(self, card, grid_container, pptx_slide, shape_converter, y_start: int) -> int:
        """
        转换data-card内的网格布局（如2x2的bullet-point网格）

        Args:
            card: data-card元素
            grid_container: 网格容器元素
            pptx_slide: PPTX幻灯片
            shape_converter: 形状转换器
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理data-card内的网格布局")
        x_base = 80

        # 获取网格列数
        grid_classes = grid_container.get('class', [])
        num_columns = 2  # 默认2列

        # 检查Tailwind CSS网格列类
        for cls in grid_classes:
            if cls.startswith('grid-cols-') and hasattr(self.css_parser, 'tailwind_grid_columns'):
                columns = self.css_parser.tailwind_grid_columns.get(cls)
                if columns:
                    num_columns = columns
                    logger.info(f"检测到网格列数: {num_columns}")
                    break

        # 获取所有bullet-point
        bullet_points = grid_container.find_all('div', class_='bullet-point')

        # 获取标题 - 优先查找h3标签
        h3_elem = card.find('h3')
        title_elem = None
        title_text = None

        if h3_elem:
            # 优先使用h3标签作为标题
            title_text = h3_elem.get_text(strip=True)
            logger.info(f"找到h3标题: {title_text}")
        else:
            # 兼容旧逻辑，查找p标签
            title_elem = card.find('p', class_='primary-color')
            if title_elem:
                title_text = title_elem.get_text(strip=True)
                logger.info(f"找到p标签标题: {title_text}")
        current_y = y_start

        # 添加data-card背景
        bg_color_str = 'rgba(10, 66, 117, 0.03)'
        from pptx.enum.shapes import MSO_SHAPE

        # 计算需要的行数
        num_rows = (len(bullet_points) + num_columns - 1) // num_columns

        # 精确计算卡片高度
        # 基础padding: 15px上 + 15px下 = 30px
        card_height = 30  # data-card的上下padding

        if title_text:
            # h3标题: 28px字体 + 12px下边距 + 10px上间距 = 50px
            card_height += 50
            logger.info(f"添加h3标题高度: 50px")

        # 计算网格行数
        num_rows = (len(bullet_points) + num_columns - 1) // num_columns
        logger.info(f"网格布局: {num_columns}列 x {num_rows}行")

        # 每行高度: bullet-point高度(25px字体) + margin-bottom(8px) + 上下padding = 60px
        row_height = 60
        grid_total_height = num_rows * row_height

        # bullet-point之间的间距已经在row_height中考虑了
        card_height += grid_total_height

        # 额外的底部间距
        card_height += 20

        logger.info(f"第四个容器精确高度计算: padding(30) + 标题(50 if any) + 网格({grid_total_height}) + 底部间距(20) = {card_height}px")

        # 添加背景
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
        logger.info(f"添加data-card网格背景，高度={card_height}px")

        # 渲染标题
        if title_text:
            text_left = UnitConverter.px_to_emu(x_base + 20)
            text_top = UnitConverter.px_to_emu(current_y + 10)
            text_box = pptx_slide.shapes.add_textbox(
                text_left, text_top,
                UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
            )
            text_frame = text_box.text_frame
            text_frame.text = title_text
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    font_size_px = self.style_computer.get_font_size_pt(title_elem)
                    run.font.size = Pt(font_size_px)
                    run.font.color.rgb = ColorParser.get_primary_color()
                    run.font.name = self.font_manager.get_font('body')

            current_y += 50  # 标题后间距

        # 处理网格中的bullet-point
        item_width = 1720 // num_columns  # 每列宽度
        item_height = 60  # 每项高度

        for idx, bullet_point in enumerate(bullet_points):
            # 计算网格位置
            col = idx % num_columns
            row = idx // num_columns
            item_x = x_base + 20 + col * item_width
            item_y = current_y + 10 + row * item_height

            # 获取图标和文本
            icon_elem = bullet_point.find('i')
            p_elem = bullet_point.find('p')

            if icon_elem and p_elem:
                # 获取图标字符
                icon_classes = icon_elem.get('class', [])
                icon_char = self._get_icon_char(icon_classes)

                # 获取文本
                text = p_elem.get_text(strip=True)

                # 获取颜色
                icon_color = self._get_element_color(icon_elem)
                if not icon_color:
                    # 根据图标类确定颜色
                    if 'text-orange-600' in icon_classes:
                        icon_color = ColorParser.get_color_by_name('orange')
                    elif 'text-green-600' in icon_classes:
                        icon_color = ColorParser.get_color_by_name('green')
                    elif 'text-blue-600' in icon_classes:
                        icon_color = ColorParser.get_color_by_name('blue')
                    elif 'text-purple-600' in icon_classes:
                        icon_color = ColorParser.get_color_by_name('purple')
                    else:
                        icon_color = ColorParser.get_primary_color()

                # 添加图标
                if icon_char:
                    # 图标和文本都在行内，需要垂直居中
                    # 计算图标和文本的垂直居中位置
                    line_height = 60  # 行高
                    vertical_center = item_y + line_height // 2

                    icon_left = UnitConverter.px_to_emu(item_x)
                    icon_top = UnitConverter.px_to_emu(vertical_center - 15)  # 图标垂直居中
                    icon_box = pptx_slide.shapes.add_textbox(
                        icon_left, icon_top,
                        UnitConverter.px_to_emu(30), UnitConverter.px_to_emu(30)
                    )
                    icon_frame = icon_box.text_frame
                    icon_frame.text = icon_char
                    icon_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                    for paragraph in icon_frame.paragraphs:
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(20)
                            run.font.color.rgb = icon_color
                            run.font.name = self.font_manager.get_font('body')

                    # 文本在图标右侧，也要垂直居中
                    text_left = UnitConverter.px_to_emu(item_x + 40)
                    text_width = item_width - 60
                    text_top = UnitConverter.px_to_emu(vertical_center - 15)  # 文本垂直居中
                else:
                    # 没有图标，直接显示文本
                    text_left = UnitConverter.px_to_emu(item_x)
                    text_width = item_width - 20
                    line_height = 60
                    vertical_center = item_y + line_height // 2
                    text_top = UnitConverter.px_to_emu(vertical_center - 15)  # 文本垂直居中

                # 添加文本
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(text_width), UnitConverter.px_to_emu(50)  # 增加高度以支持换行
                )
                text_frame = text_box.text_frame
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                text_frame.word_wrap = True

                # 检查是否有strong标签
                strong_elem = p_elem.find('strong')

                if strong_elem:
                    # 处理带strong的文本
                    strong_text = strong_elem.get_text(strip=True)
                    remaining_text = text.replace(strong_text, '').strip()

                    # 清除默认段落
                    text_frame.clear()
                    p = text_frame.paragraphs[0]

                    # 添加加粗部分
                    if strong_text:
                        strong_run = p.add_run()
                        strong_run.text = strong_text
                        strong_run.font.size = Pt(16)
                        strong_run.font.bold = True
                        strong_run.font.name = self.font_manager.get_font('body')

                    # 添加剩余部分
                    if remaining_text:
                        normal_run = p.add_run()
                        normal_run.text = remaining_text
                        normal_run.font.size = Pt(16)
                        normal_run.font.name = self.font_manager.get_font('body')
                else:
                    # 普通文本
                    text_frame.text = text
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            font_size_px = self.style_computer.get_font_size_pt(p_elem)
                            run.font.size = Pt(font_size_px)
                            run.font.name = self.font_manager.get_font('body')

        # 计算实际内容高度
        num_rows = (len(bullet_points) + num_columns - 1) // num_columns
        content_height = 50 + num_rows * 60  # 标题高度 + 内容行数 * 行高
        actual_height = max(content_height, 180)  # 最小高度180px

        # 添加左边框
        shape_converter.add_border_left(x_base, y_start, actual_height, 4)

        return y_start + actual_height + 10

    def _get_icon_char(self, icon_classes: list) -> str:
        """根据FontAwesome类获取对应emoji/Unicode字符"""
        icon_map = {
            # === 网络安全相关 ===
            # 核心安全图标
            'fa-shield': '🛡',
            'fa-shield-alt': '🛡',
            'fa-shield-virus': '🦠',
            'fa-virus-slash': '🦠',
            'fa-virus': '🦠',
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
            'fa-robot': '🤖',
            'fa-brain': '🧠',
            'fa-microchip': '💻',
            'fa-memory': '🧠',
            'fa-cpu': '💻',
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
            'fa-balance-scale': '⚖️',
            'fa-gavel': '🔨',
            'fa-landmark': '🏛️',
            'fa-courthouse': '🏛️',
            'fa-scroll': '📜',
            'fa-file-contract': '📄',
            'fa-file-alt': '📄',
            'fa-file-pdf': '📄',
            'fa-file-word': '📄',
            'fa-file-excel': '📄',

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
            'fa-swords': '⚔️',
            'fa-crosshairs': '🎯',
            'fa-shield-alt': '🛡',
            'fa-bomb': '💣',
            'fa-hammer': '🔨',
            'fa-wrench': '🔧',
            'fa-tools': '🛠',

            # === 时间和流程 ===
            'fa-clock': '🕐',
            'fa-hourglass': '⏳',
            'fa-hourglass-half': '⏳',
            'fa-calendar': '📅',
            'fa-calendar-alt': '📅',
            'fa-tasks': '☑',
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
            'fa-plus-circle': '⭕',
            'fa-minus': '-',
            'fa-minus-circle': '⭕',
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

            # === 新增：常用FontAwesome图标 ===
            # 状态和标记
            'fa-info-circle': 'ℹ',
            'fa-question-circle': '❓',
            'fa-asterisk': '*',
            'fa-star': '⭐',
            'fa-heart': '♥',
            'fa-heartbeat': '💓',
            'fa-fire': '🔥',
            'fa-bolt': '⚡',
            'fa-flash': '⚡',
            'fa-magic': '✨',
            'fa-sparkles': '✨',

            # 方向和导航
            'fa-chevron-right': '›',
            'fa-chevron-left': '‹',
            'fa-chevron-up': '⌃',
            'fa-chevron-down': '⌄',
            'fa-angle-right': '›',
            'fa-angle-left': '‹',
            'fa-angle-up': '⌃',
            'fa-angle-down': '⌄',
            'fa-caret-right': '▶',
            'fa-caret-left': '◀',
            'fa-caret-up': '▲',
            'fa-caret-down': '▼',

            # 商务和金融
            'fa-dollar-sign': '$',
            'fa-euro-sign': '€',
            'fa-pound-sign': '£',
            'fa-yen-sign': '¥',
            'fa-coins': '🪙',
            'fa-wallet': '👛',
            'fa-credit-card': '💳',
            'fa-chart-pie': '📊',
            'fa-pie-chart': '📊',
            'fa-chart-simple': '📊',

            # 云和数据
            'fa-cloud': '☁',
            'fa-cloud-arrow-up': '☁️',
            'fa-cloud-arrow-down': '☁️',
            'fa-cloud-download': '☁️',
            'fa-cloud-upload': '☁️',
            'fa-server': '🖥',
            'fa-desktop': '🖥',
            'fa-laptop': '💻',
            'fa-mobile': '📱',
            'fa-tablet': '📱',

            # 编辑和创作
            'fa-edit': '✏️',
            'fa-pen': '🖊️',
            'fa-pencil': '✏️',
            'fa-eraser': '🧹',
            'fa-paint-brush': '🖌️',
            'fa-palette': '🎨',
            'fa-camera': '📷',
            'fa-video': '📹',
            'fa-film': '🎬',
            'fa-music': '🎵',
            'fa-headphones': '🎧',
            'fa-microphone': '🎤',

            # 社交和用户
            'fa-user': '👤',
            'fa-user-circle': '👤',
            'fa-user-group': '👥',
            'fa-users': '👥',
            'fa-user-tie': '👔',
            'fa-user-graduate': '🎓',
            'fa-user-doctor': '👨‍⚕️',
            'fa-user-ninja': '🥷',
            'fa-user-astronaut': '👨‍🚀',

            # 环境和自然
            'fa-tree': '🌳',
            'fa-leaf': '🍃',
            'fa-seedling': '🌱',
            'fa-sun': '☀️',
            'fa-moon': '🌙',
            'fa-star': '⭐',
            'fa-snowflake': '❄️',
            'fa-fire': '🔥',
            'fa-water': '💧',
            'fa-droplet': '💧',

            # 交通和移动
            'fa-car': '🚗',
            'fa-plane': '✈️',
            'fa-ship': '🚢',
            'fa-train': '🚂',
            'fa-bicycle': '🚴',
            'fa-motorcycle': '🏍️',
            'fa-rocket': '🚀',
            'fa-satellite': '🛰️',
            'fa-helicopter': '🚁',

            # 食物和饮料
            'fa-utensils': '🍴',
            'fa-coffee': '☕',
            'fa-glass': '🥤',
            'fa-wine-glass': '🍷',
            'fa-beer': '🍺',
            'fa-pizza-slice': '🍕',
            'fa-hamburger': '🍔',
            'fa-ice-cream': '🍦',

            # 其他新增
            'fa-cloud-showers-heavy': '🌧️',
            'fa-gift': '🎁',
            'fa-tag': '🏷️',
            'fa-tags': '🏷️',
            'fa-certificate': '🎓',
            'fa-award': '🏆',
            'fa-trophy': '🏆',
            'fa-medal': '🏅',
            'fa-ribbon': '🎀',
            'fa-flag': '🚩',
            'fa-bookmark': '🔖',
            'fa-thumbtack': '📌',
            'fa-pushpin': '📌',
        }

        for cls in icon_classes:
            if cls in icon_map:
                return icon_map[cls]

        # 如果找不到匹配，返回默认图标
        return '●'

    def _get_element_color(self, element):
        """
        获取元素的颜色，支持Tailwind CSS类

        Args:
            element: BeautifulSoup元素

        Returns:
            RGBColor对象，如果没有找到颜色则返回None
        """
        if not element:
            return None

        # 检查Tailwind CSS颜色类
        classes = element.get('class', [])
        for cls in classes:
            if cls.startswith('text-') and hasattr(self.css_parser, 'tailwind_colors'):
                color = self.css_parser.tailwind_colors.get(cls)
                if color:
                    return ColorParser.parse_color(color)

        # 检查CSS样式中的颜色
        computed_style = self.style_computer.compute_computed_style(element)
        color_str = computed_style.get('color')
        if color_str:
            return ColorParser.parse_color(color_str)

        return None

    def _should_be_bold(self, element):
        """
        判断元素是否应该加粗

        Args:
            element: HTML元素

        Returns:
            bool: 是否应该加粗
        """
        if not element:
            return False

        # 1. 检查内联样式的font-weight
        style_str = element.get('style', '')
        if style_str:
            import re
            weight_match = re.search(r'font-weight:\s*([^;]+)', style_str)
            if weight_match:
                weight_str = weight_match.group(1).strip()
                # 转换常见的font-weight值
                if weight_str in ['bold', '700', '600', '800', '900']:
                    return True
                elif weight_str in ['normal', '400', '300']:
                    return False

        # 2. 检查类名中的加粗相关类
        classes = element.get('class', [])
        bold_classes = ['font-bold', 'font-semibold', 'font-extrabold', 'font-black']
        for cls in classes:
            if cls in bold_classes:
                return True

        # 3. 根据元素类型判断
        tag_name = element.name.lower()
        if tag_name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            # 标题默认加粗，除非明确指定了font-weight: normal
            if self.css_parser:
                # 获取CSS中定义的font-weight
                css_style = self.css_parser.get_style(tag_name)
                if css_style and 'font-weight' in css_style:
                    weight = css_style['font-weight']
                    if weight in ['normal', '400', '300']:
                        return False
                    elif weight in ['bold', '600', '700', '800', '900']:
                        return True
            # 默认情况下，h1和h2加粗，h3根据CSS决定
            return tag_name in ['h1', 'h2']

        # 4. 检查strong标签
        if element.name == 'strong':
            return True

        # 5. 检查父元素的加粗设置（如b标签内的文本）
        parent = element.parent
        if parent and parent.name in ['strong', 'b']:
            return True

        return False

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

    def _has_numbered_list_pattern(self, container) -> bool:
        """
        检测容器是否包含数字列表模式

        Args:
            container: 容器元素

        Returns:
            是否包含数字列表
        """
        # 检查子元素是否包含数字
        children = container.find_all(recursive=False)
        for child in children:
            text = child.get_text(strip=True)
            if text and text[0].isdigit():
                return True

            # 检查是否包含number相关的类
            child_classes = child.get('class', [])
            if any('number' in cls or 'num' in cls or 'count' in cls for cls in child_classes):
                return True

        return False

    def _convert_numbered_list_container(self, container, pptx_slide, y_start) -> int:
        """
        转换数字列表容器

        Args:
            container: 容器元素
            pptx_slide: PPTX幻灯片
            y_start: 起始Y坐标

        Returns:
            下一个元素的Y坐标
        """
        logger.info("处理数字列表容器")

        # 初始化文本转换器
        text_converter = TextConverter(pptx_slide, self.css_parser)

        # 检查是否是toc-item结构
        if 'toc-item' in container.get('class', []):
            # 处理单个toc-item
            number_elem = container.find('div', class_='toc-number')
            text_elem = container.find('div', class_='toc-title')

            if number_elem and text_elem:
                numbered_item = {
                    'type': 'toc',
                    'container': container,
                    'number_elem': number_elem,
                    'text_elem': text_elem,
                    'number': number_elem.get_text(strip=True),
                    'text': text_elem.get_text(strip=True)
                }
                return text_converter.convert_numbered_list(numbered_item, 80, y_start)

        # 处理其他数字列表格式
        text = container.get_text(strip=True)
        if text and text[0].isdigit():
            # 尝试分离数字和文本
            import re
            match = re.match(r'^(\d+)[\.\)\s]*\s*(.*)', text)
            if match:
                numbered_item = {
                    'type': 'paragraph_numbered',
                    'container': container,
                    'number_elem': container,
                    'text_elem': container,
                    'number': match.group(1),
                    'text': match.group(2)
                }
                return text_converter.convert_numbered_list(numbered_item, 80, y_start)

        # 降级处理为普通段落
        return self._convert_generic_card(container, pptx_slide, y_start, card_type='numbered_list')


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
