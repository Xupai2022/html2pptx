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
        self.full_soup = None  # 保存完整的HTML soup
        self._parse()

    def _parse(self):
        """解析HTML文件"""
        if not self.html_path.exists():
            raise FileNotFoundError(f"HTML文件不存在: {self.html_path}")

        # 尝试多种编码
        encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'latin-1', 'cp1252']
        html_content = None

        for encoding in encodings:
            try:
                with open(self.html_path, 'r', encoding=encoding) as f:
                    html_content = f.read()
                logger.info(f"使用编码 {encoding} 成功读取文件: {self.html_path}")
                break
            except UnicodeDecodeError:
                continue

        if html_content is None:
            # 如果所有编码都失败，使用错误处理模式
            with open(self.html_path, 'r', encoding='utf-8', errors='replace') as f:
                html_content = f.read()
            logger.warning(f"使用替换模式读取文件（部分字符可能丢失）: {self.html_path}")

        # 保存完整的HTML soup
        self.full_soup = BeautifulSoup(html_content, 'lxml')
        # 创建slide-container的副本（用于向后兼容）
        slide_container = self.full_soup.find('div', class_='slide-container')
        self.soup = slide_container if slide_container else self.full_soup
        logger.info(f"成功解析HTML: {self.html_path}")

    def get_slides(self) -> List:
        """
        获取所有幻灯片容器

        Returns:
            幻灯片元素列表
        """
        # 如果soup本身就是slide-container，直接返回
        if self.soup.name == 'div' and self.soup.get('class') and 'slide-container' in self.soup.get('class', []):
            logger.info("找到 1 个幻灯片 (soup本身就是slide-container)")
            return [self.soup]

        # 否则查找slide-container
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

    def get_strategy_cards(self, slide) -> List:
        """
        获取策略卡片 (.strategy-card)

        Args:
            slide: 幻灯片元素

        Returns:
            策略卡片列表
        """
        return slide.find_all('div', class_='strategy-card')

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

    def get_toc_items(self, slide) -> List:
        """
        获取目录项 (.toc-item)

        Args:
            slide: 幻灯片元素

        Returns:
            目录项列表
        """
        return slide.find_all('div', class_='toc-item')

    def detect_numbered_lists(self, slide) -> List[dict]:
        """
        智能检测数字列表

        Args:
            slide: 幻灯片元素

        Returns:
            数字列表信息列表，每个元素包含类型、元素、数字和文本
        """
        numbered_lists = []

        # 1. 检测toc-item结构
        toc_items = self.get_toc_items(slide)
        for toc_item in toc_items:
            number_elem = toc_item.find('div', class_='toc-number')
            text_elem = toc_item.find('div', class_='toc-title')

            if number_elem and text_elem:
                numbered_lists.append({
                    'type': 'toc',
                    'container': toc_item,
                    'number_elem': number_elem,
                    'text_elem': text_elem,
                    'number': number_elem.get_text(strip=True),
                    'text': text_elem.get_text(strip=True)
                })

        # 2. 检测带数字类名的元素
        number_patterns = ['number', 'num', 'count', 'index', 'step', 'order']
        for pattern in number_patterns:
            elems = slide.find_all('div', class_=lambda x: x and pattern in str(x))
            for elem in elems:
                text = elem.get_text(strip=True)
                if text and text[0].isdigit():
                    # 分离数字和文本
                    match = re.match(r'^(\d+)[\.\)\s]*\s*(.*)', text)
                    if match:
                        numbered_lists.append({
                            'type': 'numbered_class',
                            'container': elem,
                            'number_elem': elem,
                            'text_elem': elem,
                            'number': match.group(1),
                            'text': match.group(2)
                        })

        # 3. 检测flex布局中的数字结构
        flex_containers = slide.find_all('div', style=lambda x: x and 'display' in x and 'flex' in x)
        for flex_container in flex_containers:
            children = [child for child in flex_container.children if hasattr(child, 'get')]
            if len(children) >= 2:
                # 检查第一个子元素是否为数字
                first_text = children[0].get_text(strip=True)
                if re.match(r'^\d+(\.\d+)?$', first_text) or re.match(r'^\d{2}$', first_text):
                    numbered_lists.append({
                        'type': 'flex_numbered',
                        'container': flex_container,
                        'number_elem': children[0],
                        'text_elem': children[1] if len(children) > 1 else None,
                        'number': first_text,
                        'text': children[1].get_text(strip=True) if len(children) > 1 else ''
                    })

        # 4. 检测有序列表
        ol_lists = slide.find_all('ol')
        for ol in ol_lists:
            li_items = ol.find_all('li')
            for idx, li in enumerate(li_items, 1):
                numbered_lists.append({
                    'type': 'ordered_list',
                    'container': li,
                    'number_elem': li,
                    'text_elem': li,
                    'number': str(idx),
                    'text': li.get_text(strip=True)
                })

        # 5. 检测段落开头的数字模式
        paragraphs = slide.find_all('p')
        for p in paragraphs:
            text = p.get_text(strip=True)
            # 匹配各种数字开头格式
            patterns = [
                r'^(\d+)\.\s*(.*)',  # 1. 文本
                r'^(\d+)\)\s*(.*)',  # 1) 文本
                r'^(\d+)、\s*(.*)',  # 1、文本
                r'^([①②③④⑤⑥⑦⑧⑨⑩])\s*(.*)',  # 圆圈数字
                r'^([⓵⓶⓷⓸⓹⓺⓻⓼⓽⓾])\s*(.*)',  # 带圈数字
            ]

            for pattern in patterns:
                match = re.match(pattern, text)
                if match:
                    numbered_lists.append({
                        'type': 'paragraph_numbered',
                        'container': p,
                        'number_elem': p,
                        'text_elem': p,
                        'number': match.group(1),
                        'text': match.group(2)
                    })
                    break

        return numbered_lists

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
