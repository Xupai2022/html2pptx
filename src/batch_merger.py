"""
批量HTML转PPTX合并器
将多个HTML文件合并到一个PPTX文件中
"""

import sys
import os
from pathlib import Path
from typing import List
from pptx import Presentation

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
from src.main import HTML2PPTX
from pptx.util import Pt
from pptx.enum.text import MSO_ANCHOR, PP_PARAGRAPH_ALIGNMENT

logger = setup_logger(__name__)


class BatchHTML2PPTXMerger:
    """批量HTML转PPTX合并器 - 使用convert.py的逻辑"""

    def __init__(self, html_dir: str):
        """
        初始化批量合并器

        Args:
            html_dir: HTML文件目录路径
        """
        self.html_dir = Path(html_dir)
        self.html_files = self._get_html_files()

        if not self.html_files:
            raise ValueError(f"在目录 {html_dir} 中未找到HTML文件")

        logger.info(f"找到 {len(self.html_files)} 个HTML文件:")
        for i, html_file in enumerate(self.html_files, 1):
            logger.info(f"  {i}. {html_file.name}")

        # 初始化PPTX构建器（只创建一个Presentation实例）
        self.pptx_builder = PPTXBuilder()

        # 存储每个HTML文件的解析器，避免重复解析
        self.parsers_cache = {}

    def _get_html_files(self) -> List[Path]:
        """获取目录下所有HTML文件，按文件名排序"""
        html_files = []

        # 查找所有HTML文件
        for ext in ['*.html', '*.htm']:
            html_files.extend(self.html_dir.glob(ext))

        # 按文件名排序
        html_files.sort(key=lambda x: x.name.lower())

        return html_files

    def _get_parser_for_file(self, html_file: Path):
        """获取指定HTML文件的解析器（带缓存）"""
        if html_file not in self.parsers_cache:
            self.parsers_cache[html_file] = HTMLParser(str(html_file))
        return self.parsers_cache[html_file]

    def convert(self, output_path: str):
        """
        执行批量转换 - 使用HTML2PPTX类来处理每个HTML文件

        Args:
            output_path: 输出PPTX路径
        """
        logger.info("=" * 60)
        logger.info("开始批量HTML转PPTX合并转换 (使用convert.py逻辑)")
        logger.info("=" * 60)

        total_slides = 0

        for html_file in self.html_files:
            logger.info(f"\n{'='*40}")
            logger.info(f"处理文件: {html_file.name}")
            logger.info(f"{'='*40}")

            try:
                # 为每个HTML文件创建HTML2PPTX实例
                # 这样就能使用convert.py中的最新转换逻辑
                converter = HTML2PPTX(str(html_file))

                # 转换所有幻灯片
                slides = converter.convert_all_slides()

                # 将幻灯片添加到当前的PPTX构建器
                for slide_data in slides:
                    # 使用当前的PPTX构建器添加幻灯片
                    self.pptx_builder.add_slide_from_data(slide_data)
                    total_slides += 1

                logger.info(f"  ✓ 成功处理 {html_file.name}, 添加 {len(slides)} 张幻灯片")

            except Exception as e:
                logger.error(f"  ✗ 处理 {html_file.name} 时出错: {e}")
                import traceback
                traceback.print_exc()
                # 继续处理其他文件
                continue

        # 保存PPTX
        self.pptx_builder.save(output_path)

        logger.info("=" * 60)
        logger.info(f"批量转换完成!")
        logger.info(f"处理文件数: {len(self.html_files)}")
        logger.info(f"总幻灯片数: {total_slides}")
        logger.info(f"输出文件: {output_path}")
        logger.info("=" * 60)

    def _process_slide_content(self, slide_html, pptx_slide, css_parser, html_parser, y_offset: int, html_file: Path) -> int:
        """处理幻灯片内容区域"""

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
                        container, pptx_slide, css_parser, html_parser, y_offset, html_file
                    )
                elif 'stat-card' in container_classes:
                    y_offset = self._convert_stat_card(
                        container, pptx_slide, css_parser, html_parser, y_offset, html_file
                    )
                elif 'data-card' in container_classes:
                    shape_converter = ShapeConverter(pptx_slide, css_parser)
                    y_offset = self._convert_data_card(
                        container, pptx_slide, shape_converter, css_parser, html_parser, y_offset, html_file
                    )
                elif 'strategy-card' in container_classes:
                    y_offset = self._convert_strategy_card(
                        container, pptx_slide, css_parser, html_parser, y_offset, html_file
                    )
                elif 'flex' in container_classes and 'justify-between' in container_classes:
                    # 底部信息容器（包含bullet-point的flex布局）
                    y_offset = self._convert_bottom_info(
                        container, pptx_slide, css_parser, html_parser, y_offset, html_file
                    )
                else:
                    # 未知容器类型，记录警告
                    logger.warning(f"遇到未知容器类型: {container_classes}")
        else:
            # 降级处理：如果没有space-y-10，使用旧逻辑
            logger.warning("未找到space-y-10容器，使用降级处理")

            # 处理统计卡片 (.stat-card)
            stat_cards = html_parser.get_stat_cards(slide_html)
            for card in stat_cards:
                y_offset = self._convert_stat_card(card, pptx_slide, css_parser, html_parser, y_offset, html_file)

            # 处理数据卡片 (.data-card)
            data_cards = html_parser.get_data_cards(slide_html)
            for card in data_cards:
                shape_converter = ShapeConverter(pptx_slide, css_parser)
                y_offset = self._convert_data_card(card, pptx_slide, shape_converter, css_parser, html_parser, y_offset, html_file)

            # 处理策略卡片 (.strategy-card)
            strategy_cards = html_parser.get_strategy_cards(slide_html)
            for card in strategy_cards:
                y_offset = self._convert_strategy_card(card, pptx_slide, css_parser, html_parser, y_offset, html_file)

        return y_offset

    def _convert_stats_container(self, container, pptx_slide, css_parser, html_parser, y_start: int, html_file: Path) -> int:
        """转换统计卡片容器 - 复制自main.py的逻辑"""
        # 这里复制main.py中的_convert_stats_container方法
        # 为了简化，我们直接调用main.py中的HTML2PPTX类的静态方法

        # 创建一个临时的HTML2PPTX实例来复用现有方法
        temp_converter = HTML2PPTX(str(html_file))
        temp_converter.css_parser = css_parser
        temp_converter.font_manager = get_font_manager(css_parser)
        temp_converter.style_computer = get_style_computer(css_parser)
        # 清理样式计算器缓存，避免不同HTML文件间的样式污染
        temp_converter.style_computer.clear_cache()
        # 设置HTML文件ID
        temp_converter.style_computer.set_html_file_id(str(html_file))
        temp_converter.font_manager.css_parser = css_parser
        temp_converter.font_size_extractor.set_html_file_id(str(html_file))

        return temp_converter._convert_stats_container(container, pptx_slide, y_start)

    def _convert_stat_card(self, card, pptx_slide, css_parser, html_parser, y_start: int, html_file: Path) -> int:
        """转换统计卡片"""
        temp_converter = HTML2PPTX(str(html_file))
        temp_converter.css_parser = css_parser
        temp_converter.font_manager = get_font_manager(css_parser)
        temp_converter.style_computer = get_style_computer(css_parser)
        temp_converter.html_path = str(html_file)
        # 清理样式计算器缓存，避免不同HTML文件间的样式污染
        temp_converter.style_computer.clear_cache()
        # 设置HTML文件ID
        temp_converter.style_computer.set_html_file_id(str(html_file))
        temp_converter.font_manager.css_parser = css_parser
        temp_converter.font_size_extractor.set_html_file_id(str(html_file))

        return temp_converter._convert_stat_card(card, pptx_slide, y_start)

    def _convert_data_card(self, card, pptx_slide, shape_converter, css_parser, html_parser, y_start: int, html_file: Path) -> int:
        """转换数据卡片"""
        temp_converter = HTML2PPTX(str(html_file))
        temp_converter.css_parser = css_parser
        temp_converter.font_manager = get_font_manager(css_parser)
        temp_converter.style_computer = get_style_computer(css_parser)
        temp_converter.html_path = str(html_file)
        # 清理样式计算器缓存，避免不同HTML文件间的样式污染
        temp_converter.style_computer.clear_cache()
        # 设置HTML文件ID
        temp_converter.style_computer.set_html_file_id(str(html_file))
        temp_converter.font_manager.css_parser = css_parser
        temp_converter.font_size_extractor.set_html_file_id(str(html_file))

        return temp_converter._convert_data_card(card, pptx_slide, shape_converter, y_start)

    def _convert_strategy_card(self, card, pptx_slide, css_parser, html_parser, y_start: int, html_file: Path) -> int:
        """转换策略卡片"""
        temp_converter = HTML2PPTX(str(html_file))
        temp_converter.css_parser = css_parser
        temp_converter.font_manager = get_font_manager(css_parser)
        temp_converter.style_computer = get_style_computer(css_parser)
        temp_converter.html_path = str(html_file)
        # 清理样式计算器缓存，避免不同HTML文件间的样式污染
        temp_converter.style_computer.clear_cache()
        # 设置HTML文件ID
        temp_converter.style_computer.set_html_file_id(str(html_file))
        temp_converter.font_manager.css_parser = css_parser
        temp_converter.font_size_extractor.set_html_file_id(str(html_file))

        return temp_converter._convert_strategy_card(card, pptx_slide, y_start)

    def _convert_bottom_info(self, bottom_container, pptx_slide, css_parser, html_parser, y_start: int, html_file: Path) -> int:
        """转换底部信息布局"""
        temp_converter = HTML2PPTX(str(html_file))
        temp_converter.css_parser = css_parser
        temp_converter.font_manager = get_font_manager(css_parser)
        temp_converter.style_computer = get_style_computer(css_parser)
        temp_converter.html_path = str(html_file)
        # 清理样式计算器缓存，避免不同HTML文件间的样式污染
        temp_converter.style_computer.clear_cache()
        # 设置HTML文件ID
        temp_converter.style_computer.set_html_file_id(str(html_file))
        temp_converter.font_manager.css_parser = css_parser
        temp_converter.font_size_extractor.set_html_file_id(str(html_file))

        return temp_converter._convert_bottom_info(bottom_container, pptx_slide, y_start)


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("用法: python batch_merger.py <HTML目录> [输出pptx路径]")
        print("\n示例:")
        print("  python batch_merger.py ./input output/merged.pptx")
        print("  python batch_merger.py ./input")
        sys.exit(1)

    html_dir = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "output/merged.pptx"

    try:
        # 创建输出目录
        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)

        # 执行批量转换
        merger = BatchHTML2PPTXMerger(html_dir)
        merger.convert(output_path)

        print(f"\n✓ 批量转换完成！输出文件: {output_path}")

    except Exception as e:
        logger.error(f"批量转换失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()