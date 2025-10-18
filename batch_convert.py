"""
HTML批量转PPTX工具（完整样式保留版）
使用修改后的HTML2PPTX类实现完美的样式保留
"""

import os
import sys
import glob
from pathlib import Path
from datetime import datetime
import shutil

# 添加src目录到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from src.main import HTML2PPTX
from pptx import Presentation
import logging

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class BatchConverter:
    """批量转换器 - 完整样式保留版"""

    def __init__(self):
        self.output_dir = Path("output")

    def find_slide_files(self):
        """
        查找input目录下所有以slide开头的HTML文件

        Returns:
            排序后的HTML文件路径列表
        """
        # 检查input目录是否存在
        input_dir = Path("input")
        if not input_dir.exists():
            logger.error("input目录不存在！")
            return []

        # 查找input目录下所有以slide开头的HTML文件
        pattern = "input/slide*.html"
        html_files = sorted(glob.glob(pattern))

        logger.info(f"找到 {len(html_files)} 个slide文件")
        return html_files

    def create_output_directory(self):
        """创建输出目录"""
        self.output_dir.mkdir(exist_ok=True)
        logger.info(f"输出目录: {self.output_dir}")

    def convert_all_to_single_pptx(self):
        """
        将所有HTML文件转换为单个PPTX文件
        使用共享的Presentation对象确保样式一致性
        """
        try:
            # 查找HTML文件
            html_files = self.find_slide_files()

            if not html_files:
                logger.warning("未找到任何slide开头的HTML文件")
                return

            # 创建输出目录
            self.create_output_directory()

            logger.info(f"\n开始批量转换 {len(html_files)} 个HTML文件...")

            # 创建一个共享的Presentation对象
            # 这确保了所有幻灯片使用相同的样式和格式
            shared_presentation = Presentation()

            # 设置幻灯片尺寸
            from src.utils.unit_converter import UnitConverter
            shared_presentation.slide_width = UnitConverter.px_to_emu(1920)
            shared_presentation.slide_height = UnitConverter.px_to_emu(1080)

            # 移除默认的标题幻灯片（如果有）
            if len(shared_presentation.slides) > 0:
                xml_slides = shared_presentation.slides._sldIdLst
                xml_slides.remove(xml_slides[0])

            total_slides_processed = 0

            # 处理每个HTML文件
            for i, html_file in enumerate(html_files):
                logger.info(f"\n处理第 {i+1} 个文件: {os.path.basename(html_file)}")

                try:
                    # 清理全局缓存，避免不同HTML文件间的样式污染
                    from src.utils.style_computer import _style_computer_instance
                    from src.utils.font_manager import _font_manager_instance
                    from src.utils.font_size_extractor import _font_size_extractor_instance

                    if _style_computer_instance is not None:
                        _style_computer_instance.clear_cache()
                    if _font_manager_instance is not None:
                        _font_manager_instance._cached_fonts.clear()
                    if _font_size_extractor_instance is not None:
                        _font_size_extractor_instance.clear_cache()

                    logger.debug(f"已清理缓存，准备处理: {html_file}")

                    # 创建转换器，传入共享的Presentation对象
                    converter = HTML2PPTX(
                        html_path=html_file,
                        existing_presentation=shared_presentation
                    )

                    # 使用特殊的批量转换方法
                    # 这里我们需要修改convert方法以支持批量模式
                    self._convert_html_to_shared_presentation(converter, html_file)

                    total_slides_processed += 1
                    logger.info(f"  [OK] 已处理")

                except Exception as e:
                    logger.error(f"  [ERROR] 处理失败: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    continue

            # 保存最终文件
            if total_slides_processed > 0:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = self.output_dir / f"merged_presentation_{timestamp}.pptx"
                shared_presentation.save(output_file)

                logger.info(f"\n[SUCCESS] 批量转换完成！")
                logger.info(f"输出文件: {output_file}")
                logger.info(f"总处理幻灯片数: {total_slides_processed}")
                logger.info(f"最终PPTX幻灯片数: {len(shared_presentation.slides)}")
            else:
                logger.error("没有成功处理任何HTML文件")

        except Exception as e:
            logger.error(f"批量转换失败: {str(e)}")
            import traceback
            traceback.print_exc()
            raise

    def _convert_html_to_shared_presentation(self, converter, html_file):
        """
        将HTML内容转换到共享的Presentation对象
        这是核心方法，确保样式完全保留
        """
        # 获取HTML解析器
        html_parser = converter.html_parser

        # 获取所有幻灯片
        slides = html_parser.get_slides()

        for slide_html in slides:
            logger.info(f"  处理幻灯片...")

            # 使用共享的pptx_builder添加幻灯片
            # 这里会使用完全相同的转换逻辑
            pptx_slide = converter.pptx_builder.add_blank_slide()

            # 初始化转换器
            from src.converters.text_converter import TextConverter
            from src.converters.table_converter import TableConverter
            from src.converters.shape_converter import ShapeConverter
            from src.utils.font_manager import get_font_manager
            from src.utils.style_computer import get_style_computer

            text_converter = TextConverter(pptx_slide, converter.css_parser)
            table_converter = TableConverter(pptx_slide, converter.css_parser)
            shape_converter = ShapeConverter(pptx_slide, converter.css_parser)

            # 1. 添加顶部装饰条
            shape_converter.add_top_bar()

            # 2. 添加标题和副标题
            title_info = html_parser.get_title_info(slide_html)
            if title_info:
                # content-section的padding-top是20px
                title_end_y = text_converter.convert_title(
                    title_info['text'],
                    title_info['subtitle'],
                    x=80,
                    y=20,
                    is_cover=title_info.get('is_cover', False),
                    title_classes=title_info.get('classes', []),
                    h1_element=title_info.get('h1_element')
                )
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
                    y_offset = converter._process_container(container, pptx_slide, y_offset, shape_converter)
            else:
                # 处理content-section的直接子元素
                logger.info("  未找到space-y-10容器，处理content-section的直接子元素")

                # 跳过标题区域
                containers = []
                skip_first_mb = True
                for child in content_section.children:
                    if hasattr(child, 'get') and child.get('class'):
                        classes = child.get('class', [])
                        has_title = child.find('h1') or child.find('h2')

                        is_title_container = False
                        if skip_first_mb and any(cls in ['mb-6', 'mb-4', 'mb-8'] for cls in classes) and has_title:
                            has_content = any(cls in classes for cls in ['grid', 'stat-card', 'data-card', 'risk-card', 'flex'])
                            if not has_content:
                                is_title_container = True
                                skip_first_mb = False

                        if is_title_container:
                            continue

                        if child.name:
                            containers.append(child)

                # 处理所有容器
                for container in containers:
                    if container.name:
                        if containers.index(container) > 0:
                            y_offset += 40
                        y_offset = converter._process_container(container, pptx_slide, y_offset, shape_converter)

            # 4. 添加页码
            page_num = html_parser.get_page_number(slide_html)
            if page_num:
                shape_converter.add_page_number(page_num)


def main():
    """主函数"""
    print("=" * 60)
    print("HTML批量转PPTX工具（完整样式保留版）")
    print("=" * 60)

    # 检查input目录
    input_dir = Path("input")
    if not input_dir.exists():
        print("\n错误：input目录不存在！")
        print("请创建input目录并放入slide开头的HTML文件")
        return

    # 查找HTML文件
    html_files = sorted(glob.glob("input/slide*.html"))
    if not html_files:
        print("\n未找到任何slide开头的HTML文件！")
        print(f"请检查input目录中的文件")
        return

    print(f"\n找到 {len(html_files)} 个HTML文件:")
    for f in html_files:
        print(f"  - {os.path.basename(f)}")

    # 创建转换器并执行转换
    converter = BatchConverter()

    try:
        converter.convert_all_to_single_pptx()
    except KeyboardInterrupt:
        print("\n\n用户中断转换")
    except Exception as e:
        print(f"\n错误：{str(e)}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("转换结束")
    print("=" * 60)


if __name__ == "__main__":
    main()