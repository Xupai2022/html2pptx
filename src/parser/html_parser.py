"""
HTML解析器
使用BeautifulSoup解析HTML结构
"""

from bs4 import BeautifulSoup
from pathlib import Path
from typing import List, Optional
import re

from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class HTMLParser:
    """HTML解析器"""

    def __init__(self, html_path: str):
        """
        初始化解析器

        Args:
            html_path: HTML文件路径
        """
        self.html_path = Path(html_path)
        self.soup = None
        self._parse()

    def _parse(self):
        """解析HTML文件"""
        if not self.html_path.exists():
            raise FileNotFoundError(f"HTML文件不存在: {self.html_path}")

        with open(self.html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        self.soup = BeautifulSoup(html_content, 'lxml')
        logger.info(f"成功解析HTML: {self.html_path}")

    def get_slides(self) -> List:
        """
        获取所有幻灯片容器

        Returns:
            幻灯片元素列表
        """
        slides = self.soup.find_all('div', class_='slide-container')
        logger.info(f"找到 {len(slides)} 个幻灯片")
        return slides

    def get_title(self, slide) -> Optional[str]:
        """
        获取幻灯片标题 (H1)

        Args:
            slide: 幻灯片元素

        Returns:
            标题文本
        """
        h1 = slide.find('h1')
        return h1.get_text(strip=True) if h1 else None

    def get_subtitle(self, slide) -> Optional[str]:
        """
        获取幻灯片副标题 (H2)

        Args:
            slide: 幻灯片元素

        Returns:
            副标题文本
        """
        h2 = slide.find('h2')
        return h2.get_text(strip=True) if h2 else None

    def get_page_number(self, slide) -> Optional[str]:
        """
        获取页码

        Args:
            slide: 幻灯片元素

        Returns:
            页码文本
        """
        page_num = slide.find('div', class_='page-number')
        return page_num.get_text(strip=True) if page_num else None

    def get_paragraphs(self, slide) -> List:
        """
        获取段落元素

        Args:
            slide: 幻灯片元素

        Returns:
            段落元素列表
        """
        return slide.find_all('p')

    def get_tables(self, slide) -> List:
        """
        获取表格元素

        Args:
            slide: 幻灯片元素

        Returns:
            表格元素列表
        """
        return slide.find_all('table')

    def get_stat_boxes(self, slide) -> List:
        """
        获取统计卡片 (.stat-box)

        Args:
            slide: 幻灯片元素

        Returns:
            统计卡片列表
        """
        return slide.find_all('div', class_='stat-box')

    def get_stat_cards(self, slide) -> List:
        """
        获取统计卡片容器 (.stat-card)

        Args:
            slide: 幻灯片元素

        Returns:
            统计卡片列表
        """
        return slide.find_all('div', class_='stat-card')

    def get_data_cards(self, slide) -> List:
        """
        获取数据卡片 (.data-card)

        Args:
            slide: 幻灯片元素

        Returns:
            数据卡片列表
        """
        return slide.find_all('div', class_='data-card')

    def get_progress_bars(self, element) -> List:
        """
        获取进度条元素

        Args:
            element: 父元素

        Returns:
            进度条容器列表
        """
        return element.find_all('div', class_='progress-container')

    def get_bullet_points(self, element) -> List:
        """
        获取列表项

        Args:
            element: 父元素

        Returns:
            列表项元素列表
        """
        return element.find_all('div', class_='bullet-point')

    def extract_inline_style(self, element) -> dict:
        """
        提取元素的内联样式

        Args:
            element: HTML元素

        Returns:
            样式字典
        """
        style_dict = {}
        style_str = element.get('style', '')

        if not style_str:
            return style_dict

        # 解析style属性
        for item in style_str.split(';'):
            if ':' in item:
                key, value = item.split(':', 1)
                style_dict[key.strip()] = value.strip()

        return style_dict

    def get_canvas_elements(self, slide) -> List:
        """
        获取canvas图表元素

        Args:
            slide: 幻灯片元素

        Returns:
            canvas元素列表
        """
        return slide.find_all('canvas')

    def extract_chart_data(self, slide) -> List[dict]:
        """
        从script标签中提取图表数据

        Args:
            slide: 幻灯片元素

        Returns:
            图表配置列表
        """
        charts = []
        scripts = slide.find_all('script')

        for script in scripts:
            script_text = script.string
            if not script_text or 'Chart' not in script_text:
                continue

            # 提取canvas ID
            canvas_id_match = re.search(r"getElementById\('([^']+)'\)", script_text)
            if not canvas_id_match:
                continue

            canvas_id = canvas_id_match.group(1)

            # 提取图表类型
            chart_type_match = re.search(r"type:\s*'([^']+)'", script_text)
            chart_type = chart_type_match.group(1) if chart_type_match else 'bar'

            # 提取标签
            labels_match = re.search(r"labels:\s*\[(.*?)\]", script_text)
            labels = []
            if labels_match:
                labels_str = labels_match.group(1)
                labels = [l.strip().strip("'\"") for l in labels_str.split(',')]

            charts.append({
                'canvas_id': canvas_id,
                'type': chart_type,
                'labels': labels,
                'element': slide.find('canvas', id=canvas_id)
            })

        return charts
