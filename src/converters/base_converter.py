"""
转换器基类
"""

from abc import ABC, abstractmethod


class BaseConverter(ABC):
    """转换器基类"""

    def __init__(self, slide, css_parser):
        """
        初始化转换器

        Args:
            slide: python-pptx幻灯片对象
            css_parser: CSS解析器
        """
        self.slide = slide
        self.css_parser = css_parser

    @abstractmethod
    def convert(self, element, **kwargs):
        """
        转换元素

        Args:
            element: HTML元素
            **kwargs: 其他参数
        """
        pass
