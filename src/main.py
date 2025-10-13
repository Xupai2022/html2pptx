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
from src.utils.logger import setup_logger
from src.utils.unit_converter import UnitConverter
from src.utils.color_parser import ColorParser
from src.utils.chart_capture import ChartCapture
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
                text_converter.convert_title(title, subtitle, x=80, y=80)

            # 3. 处理内容区域
            y_offset = 180 if subtitle else 150

            # 处理统计卡片容器 (.stats-container)
            stats_container = slide_html.find('div', class_='stats-container')
            if stats_container:
                y_offset = self._convert_stats_container(
                    stats_container, pptx_slide, y_offset
                )

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

        # 计算布局(4列网格)
        box_width = 400
        box_height = 220
        gap = 20
        x_start = 80

        for idx, box in enumerate(stat_boxes):
            col = idx % 4
            row = idx // 4

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

            # 添加图标(使用文本替代)
            if icon:
                icon_classes = icon.get('class', [])
                icon_char = self._get_icon_char(icon_classes)

                icon_left = UnitConverter.px_to_emu(x + 20)
                icon_top = UnitConverter.px_to_emu(y + 80)
                icon_box = pptx_slide.shapes.add_textbox(
                    icon_left, icon_top,
                    UnitConverter.px_to_emu(50), UnitConverter.px_to_emu(50)
                )
                icon_frame = icon_box.text_frame
                icon_frame.text = icon_char
                for paragraph in icon_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(36)
                        run.font.color.rgb = ColorParser.get_primary_color()

            # 添加标题
            if title_elem:
                title_text = title_elem.get_text(strip=True)
                title_left = UnitConverter.px_to_emu(x + 80)
                title_top = UnitConverter.px_to_emu(y + 20)
                title_box = pptx_slide.shapes.add_textbox(
                    title_left, title_top,
                    UnitConverter.px_to_emu(300), UnitConverter.px_to_emu(30)
                )
                title_frame = title_box.text_frame
                title_frame.text = title_text
                title_frame.word_wrap = True
                for paragraph in title_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(18)
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = 'Microsoft YaHei'

            # 添加主数据
            if h2:
                h2_text = h2.get_text(strip=True)
                h2_left = UnitConverter.px_to_emu(x + 80)
                h2_top = UnitConverter.px_to_emu(y + 60)
                h2_box = pptx_slide.shapes.add_textbox(
                    h2_left, h2_top,
                    UnitConverter.px_to_emu(300), UnitConverter.px_to_emu(50)
                )
                h2_frame = h2_box.text_frame
                h2_frame.text = h2_text
                for paragraph in h2_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(36)
                        run.font.bold = True
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = 'Microsoft YaHei'

            # 添加描述
            if p:
                p_text = p.get_text(strip=True)
                p_left = UnitConverter.px_to_emu(x + 80)
                p_top = UnitConverter.px_to_emu(y + 120)
                p_box = pptx_slide.shapes.add_textbox(
                    p_left, p_top,
                    UnitConverter.px_to_emu(300), UnitConverter.px_to_emu(30)
                )
                p_frame = p_box.text_frame
                p_frame.text = p_text
                for paragraph in p_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(20)
                        run.font.name = 'Microsoft YaHei'

        # 计算下一个元素的Y坐标
        num_rows = (num_boxes + 3) // 4
        return y_start + num_rows * (box_height + gap) + 30

    def _convert_stat_card(self, card, pptx_slide, y_start: int) -> int:
        """转换统计卡片(.stat-card) - 只处理包含canvas图表的卡片"""
        # 检查是否包含canvas图表
        canvas = card.find('canvas')

        # 如果stat-card内包含stats-container,说明这不是图表卡片,而是stat-box容器
        # 需要处理嵌套的stats-container结构
        stats_container = card.find('div', class_='stats-container')
        if stats_container:
            logger.info("stat-card包含stats-container,处理嵌套的stat-box结构")

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
                    text_left = UnitConverter.px_to_emu(80)
                    text_top = UnitConverter.px_to_emu(y_start)
                    text_box = pptx_slide.shapes.add_textbox(
                        text_left, text_top,
                        UnitConverter.px_to_emu(1760), UnitConverter.px_to_emu(30)
                    )
                    text_frame = text_box.text_frame
                    text_frame.text = text
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(20)
                            run.font.color.rgb = ColorParser.get_primary_color()
                            run.font.name = 'Microsoft YaHei'

                    y_start += 40

            # 处理嵌套的stats-container
            return self._convert_stats_container(stats_container, pptx_slide, y_start)

        # 如果没有canvas,说明这个stat-card不是图表类型,跳过
        if not canvas:
            logger.info("stat-card不包含canvas,跳过")
            return y_start

        # 添加标题文本(如果有)
        p_elem = card.find('p', class_='primary-color')
        if p_elem:
            text = p_elem.get_text(strip=True)
            if text:  # 确保文本不为空
                text_left = UnitConverter.px_to_emu(80)
                text_top = UnitConverter.px_to_emu(y_start)
                text_box = pptx_slide.shapes.add_textbox(
                    text_left, text_top,
                    UnitConverter.px_to_emu(1760), UnitConverter.px_to_emu(30)
                )
                text_frame = text_box.text_frame
                text_frame.text = text
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(20)
                        run.font.color.rgb = ColorParser.get_primary_color()
                        run.font.name = 'Microsoft YaHei'

        # 处理canvas图表
        chart_converter = ChartConverter(pptx_slide, self.css_parser, self.html_path)
        success = chart_converter.convert_chart(
            canvas,
            x=80,
            y=y_start + 40,
            width=1760,
            height=220,
            use_screenshot=ChartCapture.is_available()
        )

        if not success:
            logger.warning("图表转换失败,已显示占位文本")

        return y_start + 280

    def _convert_data_card(self, card, pptx_slide, shape_converter, y_start: int) -> int:
        """转换数据卡片(.data-card)"""
        x_base = 80

        # 添加左边框
        shape_converter.add_border_left(x_base, y_start, 280, 4)

        # 标题
        p_elem = card.find('p', class_='primary-color')
        if p_elem:
            text = p_elem.get_text(strip=True)
            text_left = UnitConverter.px_to_emu(x_base + 20)
            text_top = UnitConverter.px_to_emu(y_start + 10)
            text_box = pptx_slide.shapes.add_textbox(
                text_left, text_top,
                UnitConverter.px_to_emu(1720), UnitConverter.px_to_emu(30)
            )
            text_frame = text_box.text_frame
            text_frame.text = text
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(20)
                    run.font.color.rgb = ColorParser.get_primary_color()
                    run.font.name = 'Microsoft YaHei'

        # 进度条
        progress_bars = card.find_all('div', class_='progress-container')
        progress_y = y_start + 50
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
        for bullet in bullet_points:
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
                        run.font.size = Pt(20)
                        run.font.name = 'Microsoft YaHei'

                progress_y += 35

        return progress_y + 20

    def _get_icon_char(self, icon_classes: list) -> str:
        """根据FontAwesome类获取对应字符"""
        icon_map = {
            'fa-search': '🔍',
            'fa-bug': '🐛',
            'fa-check-circle': '✓',
            'fa-exclamation-triangle': '⚠',
            'fa-exclamation-circle': '!',
        }

        for cls in icon_classes:
            if cls in icon_map:
                return icon_map[cls]

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
