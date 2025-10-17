"""
SVG转换器
将SVG图表转换为PPTX内容，支持截图和内容提取两种方式
"""

from typing import Optional, Dict, List, Tuple
from pathlib import Path
import re

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Pt

from src.converters.base_converter import BaseConverter
from src.utils.chart_capture import ChartCapture
from src.utils.unit_converter import UnitConverter
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class SvgConverter(BaseConverter):
    """SVG图表转换器"""

    def __init__(self, slide, css_parser, html_path: str = None):
        """
        初始化SVG转换器

        Args:
            slide: PPTX幻灯片对象
            css_parser: CSS解析器
            html_path: HTML文件路径（用于截图）
        """
        super().__init__(slide, css_parser)
        self.html_path = html_path
        self.chart_capturer = ChartCapture()

    def convert(self, element, **kwargs):
        """
        转换SVG元素（实现BaseConverter的抽象方法）

        Args:
            element: SVG元素
            **kwargs: 其他参数

        Returns:
            转换结果
        """
        # 获取参数
        x = kwargs.get('x', 80)
        y = kwargs.get('y', 100)
        width = kwargs.get('width', 1760)
        chart_index = kwargs.get('chart_index', 0)

        # 获取容器
        container = kwargs.get('container', element.parent)

        # 调用SVG转换方法
        return self.convert_svg(element, container, x, y, width, chart_index)

    def convert_svg(
        self,
        svg_element,
        container,
        x: int,
        y: int,
        width: int,
        chart_index: int = 0
    ) -> int:
        """
        转换SVG图表

        Args:
            svg_element: SVG元素
            container: 容器元素
            x: X坐标
            y: Y坐标
            width: 宽度
            chart_index: 图表索引（用于多个图表时的截图）

        Returns:
            实际高度
        """
        logger.info(f"开始转换SVG图表 {chart_index}")

        # 获取SVG的原始尺寸
        svg_width, svg_height = self._get_svg_dimensions(svg_element)

        # 计算目标尺寸，保持宽高比
        target_width = width
        target_height = int(target_width * svg_height / svg_width) if svg_width > 0 else 250

        logger.info(f"SVG原始尺寸: {svg_width}x{svg_height}, 目标尺寸: {target_width}x{target_height}")

        # 尝试截图（优先方案）
        actual_height = target_height
        screenshot_success = False

        if self.html_path:
            screenshot_path = self._capture_svg_screenshot(svg_element, chart_index, container)
            if screenshot_path:
                actual_height = self._insert_svg_screenshot(
                    self.slide, screenshot_path, x, y, target_width, target_height
                )
                screenshot_success = True
                logger.info(f"SVG图表 {chart_index} 截图成功，尺寸: {target_width}x{actual_height}")

        # 优雅降级到内容提取
        if not screenshot_success:
            logger.warning(f"SVG图表 {chart_index} 截图失败，降级到内容提取")
            actual_height = self._render_svg_fallback(svg_element, x, y, target_width, target_height)

        return actual_height

    def _get_svg_dimensions(self, svg_element) -> Tuple[int, int]:
        """
        获取SVG的尺寸

        Args:
            svg_element: SVG元素

        Returns:
            (宽度, 高度)
        """
        # 尝试从viewBox获取
        viewbox = svg_element.get('viewBox')
        if viewbox:
            try:
                values = viewbox.split()
                if len(values) >= 4:
                    width = int(float(values[2]))
                    height = int(float(values[3]))
                    return width, height
            except (ValueError, IndexError):
                pass

        # 尝试从width和height属性获取
        width_str = svg_element.get('width', '400')
        height_str = svg_element.get('height', '250')

        # 清理单位
        width = self._parse_dimension(width_str)
        height = self._parse_dimension(height_str)

        # 使用默认值
        if width <= 0:
            width = 400
        if height <= 0:
            height = 250

        return width, height

    def _parse_dimension(self, dimension_str: str) -> int:
        """
        解析尺寸字符串（处理px、em等单位）

        Args:
            dimension_str: 尺寸字符串

        Returns:
            像素值
        """
        if not dimension_str:
            return 0

        # 移除空格并转为小写
        dimension_str = dimension_str.strip().lower()

        # 提取数字部分
        match = re.search(r'([\d.]+)', dimension_str)
        if match:
            value = float(match.group(1))

            # 处理单位
            if dimension_str.endswith('px'):
                return int(value)
            elif dimension_str.endswith('em'):
                return int(value * 16)  # 假设1em=16px
            elif dimension_str.endswith('%'):
                return int(value * 4)  # 相对于某个基准
            else:
                return int(value)

        return 0

    def _capture_svg_screenshot(self, svg_element, chart_index: int, container) -> Optional[str]:
        """
        截取SVG截图

        Args:
            svg_element: SVG元素
            chart_index: 图表索引
            container: 容器元素

        Returns:
            截图路径，失败返回None
        """
        try:
            # 方案1：尝试基于父容器和SVG索引的组合选择器
            # 获取容器在父容器中的位置
            parent_containers = container.parent.find_all('div', class_='flex-1') if container.parent else []
            container_index = parent_containers.index(container) if container in parent_containers else 0

            logger.info(f"容器索引: {container_index}, SVG索引: {chart_index}")

            # 生成唯一的选择器，基于容器中的SVG内容
            # 检查SVG的特征来创建唯一选择器
            viewbox = svg_element.get('viewBox', '')
            svg_content = str(svg_element)[:200]  # 取前200个字符作为内容标识

            # 尝试使用XPath选择器
            # 基于viewBox或内容特征创建XPath
            if viewbox:
                # 使用viewBox作为特征
                xpath_selector = f"//div[@class='flex-1'][{container_index + 1}]//svg[@viewBox='{viewbox}']"
                logger.info(f"使用XPath选择器: {xpath_selector}")
                result = self.chart_capturer.capture_svg(
                    self.html_path,
                    xpath_selector,
                    wait_time=2000
                )
                if result:
                    return result

            # 方案2：基于容器内的所有SVG数量，使用更精确的索引
            # 找到父级flex容器（包含多个图表的容器）
            parent_flex = None
            current = container
            while current:
                if 'flex' in current.get('class', []) and 'gap-6' in current.get('class', []):
                    parent_flex = current
                    break
                current = current.parent

            if parent_flex:
                # 获取父flex容器中的所有SVG
                all_svg_in_flex = parent_flex.find_all('svg')
                logger.info(f"父flex容器中有 {len(all_svg_in_flex)} 个SVG")

                # 找到当前SVG在整个flex容器中的索引
                global_svg_index = all_svg_in_flex.index(svg_element)
                logger.info(f"SVG在flex容器中的全局索引: {global_svg_index}")

                return self.chart_capturer.capture_svg_by_index(
                    self.html_path,
                    global_svg_index,
                    wait_time=2000
                )

            # 方案3：降级到普通选择器
            svg_classes = svg_element.get('class', [])
            if svg_classes:
                selector = '.'.join(['svg'] + svg_classes)
                logger.info(f"使用CSS选择器截取SVG: {selector}")
                return self.chart_capturer.capture_svg(
                    self.html_path,
                    selector,
                    wait_time=2000
                )

            # 最后尝试直接截取第一个SVG
            logger.info("使用默认选择器截取SVG")
            return self.chart_capturer.capture_svg(
                self.html_path,
                "svg",
                wait_time=2000
            )

        except Exception as e:
            logger.error(f"SVG截图失败: {e}")
            return None

    def _insert_svg_screenshot(
        self,
        pptx_slide,
        screenshot_path: str,
        x: int,
        y: int,
        width: int,
        height: int
    ) -> int:
        """
        插入SVG截图到PPTX，保持宽高比

        Args:
            pptx_slide: PPTX幻灯片
            screenshot_path: 截图路径
            x: X坐标
            y: Y坐标
            width: 期望宽度
            height: 期望高度

        Returns:
            实际高度
        """
        try:
            # 使用PIL库获取截图的实际尺寸
            if PIL_AVAILABLE:
                with Image.open(screenshot_path) as img:
                    actual_width, actual_height = img.size

                logger.info(f"截图实际尺寸: {actual_width}x{actual_height}px")
                logger.info(f"期望插入尺寸: {width}x{height}px")

                # 计算保持宽高比的尺寸
                scaled_height = int(width * actual_height / actual_width)
                logger.info(f"实际插入尺寸: {width}x{scaled_height}px")
            else:
                logger.warning("PIL库不可用，使用原始尺寸")
                scaled_height = height

            # 添加图片，使用保持宽高比的尺寸
            pic = pptx_slide.shapes.add_picture(
                screenshot_path,
                UnitConverter.px_to_emu(x),
                UnitConverter.px_to_emu(y),
                UnitConverter.px_to_emu(width),
                UnitConverter.px_to_emu(scaled_height)
            )

            return scaled_height

        except Exception as e:
            logger.error(f"插入SVG截图失败: {e}")
            return height

    def _render_svg_fallback(
        self,
        svg_element,
        x: int,
        y: int,
        width: int,
        height: int
    ) -> int:
        """
        SVG截图失败时的降级渲染

        Args:
            svg_element: SVG元素
            x: X坐标
            y: Y坐标
            width: 宽度
            height: 高度

        Returns:
            实际高度
        """
        logger.info("使用降级方案渲染SVG内容")

        try:
            # 添加占位符矩形
            placeholder = self.slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                UnitConverter.px_to_emu(x),
                UnitConverter.px_to_emu(y),
                UnitConverter.px_to_emu(width),
                UnitConverter.px_to_emu(height)
            )

            # 设置样式
            placeholder.fill.solid()
            placeholder.fill.fore_color.rgb = RGBColor(240, 240, 240)
            placeholder.line.color.rgb = RGBColor(200, 200, 200)
            placeholder.line.width = Pt(1)

            # 添加文本
            text_frame = placeholder.text_frame
            text_frame.text = "SVG图表\n（无法显示）"
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

            for paragraph in text_frame.paragraphs:
                paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(16)
                    run.font.color.rgb = RGBColor(100, 100, 100)

            return height

        except Exception as e:
            logger.error(f"SVG降级渲染失败: {e}")
            return height

    def convert_multiple_svgs(
        self,
        container,
        x: int,
        y: int,
        total_width: int,
        gap: int = 24
    ) -> int:
        """
        转换容器中的多个SVG图表（水平布局）

        Args:
            container: 容器元素
            x: 起始X坐标
            y: Y坐标
            total_width: 总宽度
            gap: 图表间距

        Returns:
            实际高度
        """
        svg_elements = container.find_all('svg')
        num_svgs = len(svg_elements)

        if num_svgs == 0:
            logger.warning("容器中未找到SVG元素")
            return 0

        logger.info(f"找到 {num_svgs} 个SVG元素")

        # 计算每个SVG的宽度
        chart_width = (total_width - (num_svgs - 1) * gap) // num_svgs

        # 计算起始X位置（水平居中）
        total_charts_width = num_svgs * chart_width + (num_svgs - 1) * gap
        start_x = x + (total_width - total_charts_width) // 2

        current_y = y
        max_height = 0

        # 处理每个SVG图表
        for i, svg_elem in enumerate(svg_elements):
            chart_x = start_x + i * (chart_width + gap)

            # 转换SVG
            chart_height = self.convert_svg(
                svg_elem,
                container,
                chart_x,
                current_y,
                chart_width,
                i
            )

            # 更新最大高度
            max_height = max(max_height, chart_height)

        return max_height


# 导入必要的RGBColor和其他常量
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_PARAGRAPH_ALIGNMENT