"""
图表截图工具
使用Playwright无头浏览器截取Chart.js图表
"""

import asyncio
import hashlib
from pathlib import Path
from typing import Optional, List, Dict
import tempfile

try:
    from playwright.async_api import async_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ChartCapture:
    """Chart.js图表截图工具"""

    def __init__(self, cache_dir: str = None, use_system_chrome: bool = True):
        """
        初始化图表截图工具

        Args:
            cache_dir: 缓存目录路径
            use_system_chrome: 是否优先使用系统已安装的Chrome(推荐True)
        """
        if not PLAYWRIGHT_AVAILABLE:
            logger.warning("Playwright未安装,图表截图功能不可用")
            logger.warning("请运行: pip install playwright")

        self.cache_dir = Path(cache_dir) if cache_dir else Path(tempfile.gettempdir()) / "html2pptx_charts"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.use_system_chrome = use_system_chrome
        logger.info(f"图表缓存目录: {self.cache_dir}")

        if use_system_chrome:
            logger.info("优先使用系统Chrome浏览器")

    async def capture_svg_async(
        self,
        html_path: str,
        svg_selector: str = "svg",
        output_path: str = None,
        wait_time: int = 2000
    ) -> Optional[str]:
        """
        异步截取SVG图表

        Args:
            html_path: HTML文件路径
            svg_selector: SVG元素选择器，支持CSS选择器或XPath
            output_path: 输出路径,不指定则自动生成
            wait_time: 等待时间(毫秒),确保SVG渲染完成

        Returns:
            截图文件路径,失败返回None
        """
        if not PLAYWRIGHT_AVAILABLE:
            logger.error("Playwright未安装,无法截图")
            return None

        html_path = Path(html_path).absolute()
        if not html_path.exists():
            logger.error(f"HTML文件不存在: {html_path}")
            return None

        # 生成输出路径
        if output_path is None:
            cache_key = self._get_cache_key(str(html_path), svg_selector)
            output_path = self.cache_dir / f"svg_{cache_key}.png"
        else:
            output_path = Path(output_path)

        # 检查缓存
        if output_path.exists():
            logger.info(f"使用缓存的SVG截图: {output_path}")
            return str(output_path)

        try:
            async with async_playwright() as p:
                # 启动浏览器（复用现有逻辑）
                browser = None

                if self.use_system_chrome:
                    try:
                        browser = await p.chromium.launch(
                            headless=True,
                            channel="chrome"
                        )
                        logger.info("使用系统Chrome浏览器")
                    except Exception as chrome_error:
                        logger.warning(f"系统Chrome不可用: {chrome_error}")
                        logger.info("尝试使用Playwright Chromium...")

                if browser is None:
                    try:
                        browser = await p.chromium.launch(headless=True)
                        logger.info("使用Playwright Chromium浏览器")
                    except Exception as chromium_error:
                        error_msg = str(chromium_error)
                        if "Executable doesn't exist" in error_msg or "playwright install" in error_msg:
                            logger.error("浏览器不可用!")
                            logger.error("解决方案:")
                            logger.error("  1. 确保Chrome已安装(推荐)")
                            logger.error("  2. 或运行: playwright install chromium")
                            return None
                        else:
                            raise chromium_error

                # 创建页面
                page = await browser.new_page(
                    viewport={'width': 1920, 'height': 1080}
                )

                # 加载HTML
                file_url = html_path.as_uri()
                await page.goto(file_url, wait_until='networkidle')
                logger.info(f"加载HTML: {html_path}")

                # 判断选择器类型
                is_xpath = svg_selector.startswith('//') or svg_selector.startswith('(')

                # 等待SVG元素
                if is_xpath:
                    await page.wait_for_selector(f'xpath={svg_selector}', timeout=10000)
                    logger.info(f"找到SVG元素(XPath): {svg_selector}")
                    svg_element = await page.query_selector(f'xpath={svg_selector}')
                else:
                    await page.wait_for_selector(svg_selector, timeout=10000)
                    logger.info(f"找到SVG元素(CSS): {svg_selector}")
                    svg_element = await page.query_selector(svg_selector)

                if svg_element:
                    # 等待SVG渲染完成
                    await page.wait_for_timeout(wait_time)
                    logger.info(f"等待SVG渲染 {wait_time}ms")

                    # 截取SVG元素，保持原始比例
                    screenshot_bytes = await svg_element.screenshot(
                        path=str(output_path),
                        type='png'
                    )
                    logger.info(f"SVG截图成功: {output_path}")
                else:
                    logger.error(f"未找到SVG元素: {svg_selector}")
                    await browser.close()
                    return None

                # 关闭浏览器
                await browser.close()

                return str(output_path)

        except Exception as e:
            logger.error(f"SVG截图失败: {e}")
            return None

    async def capture_svg_by_index_async(
        self,
        html_path: str,
        svg_index: int = 0,
        output_path: str = None,
        wait_time: int = 2000
    ) -> Optional[str]:
        """
        按索引截取SVG图表

        Args:
            html_path: HTML文件路径
            svg_index: SVG元素索引(从0开始)
            output_path: 输出路径,不指定则自动生成
            wait_time: 等待时间(毫秒)

        Returns:
            截图文件路径,失败返回None
        """
        if not PLAYWRIGHT_AVAILABLE:
            logger.error("Playwright未安装,无法截图")
            return None

        html_path = Path(html_path).absolute()
        if not html_path.exists():
            logger.error(f"HTML文件不存在: {html_path}")
            return None

        # 生成输出路径
        if output_path is None:
            cache_key = self._get_cache_key(str(html_path), f"svg_index_{svg_index}")
            output_path = self.cache_dir / f"svg_{cache_key}.png"
        else:
            output_path = Path(output_path)

        # 检查缓存
        if output_path.exists():
            logger.info(f"使用缓存的SVG截图: {output_path}")
            return str(output_path)

        try:
            async with async_playwright() as p:
                # 启动浏览器（复用现有逻辑）
                browser = None

                if self.use_system_chrome:
                    try:
                        browser = await p.chromium.launch(
                            headless=True,
                            channel="chrome"
                        )
                    except Exception:
                        browser = None

                if browser is None:
                    try:
                        browser = await p.chromium.launch(headless=True)
                    except Exception as chromium_error:
                        error_msg = str(chromium_error)
                        if "Executable doesn't exist" in error_msg or "playwright install" in error_msg:
                            logger.error("浏览器不可用!")
                            return None
                        else:
                            raise chromium_error

                # 创建页面
                page = await browser.new_page(
                    viewport={'width': 1920, 'height': 1080}
                )

                # 加载HTML
                file_url = html_path.as_uri()
                await page.goto(file_url, wait_until='networkidle')
                logger.info(f"加载HTML: {html_path}")

                # 获取所有SVG元素
                svg_elements = await page.query_selector_all("svg")
                logger.info(f"找到 {len(svg_elements)} 个SVG元素")

                if svg_index < len(svg_elements):
                    svg_element = svg_elements[svg_index]
                    logger.info(f"截取第 {svg_index} 个SVG元素")

                    # 等待SVG渲染完成
                    await page.wait_for_timeout(wait_time)
                    logger.info(f"等待SVG渲染 {wait_time}ms")

                    # 截取SVG元素
                    screenshot_bytes = await svg_element.screenshot(
                        path=str(output_path),
                        type='png'
                    )
                    logger.info(f"SVG截图成功: {output_path}")
                else:
                    logger.error(f"SVG索引 {svg_index} 超出范围(总共 {len(svg_elements)} 个)")
                    await browser.close()
                    return None

                # 关闭浏览器
                await browser.close()

                return str(output_path)

        except Exception as e:
            logger.error(f"SVG截图失败: {e}")
            return None

    def capture_svg(
        self,
        html_path: str,
        svg_selector: str = "svg",
        output_path: str = None,
        wait_time: int = 2000
    ) -> Optional[str]:
        """
        同步截取SVG(内部调用异步方法)

        Args:
            html_path: HTML文件路径
            svg_selector: SVG元素选择器
            output_path: 输出路径
            wait_time: 等待时间(毫秒)

        Returns:
            截图文件路径,失败返回None
        """
        return asyncio.run(
            self.capture_svg_async(html_path, svg_selector, output_path, wait_time)
        )

    def capture_svg_by_index(
        self,
        html_path: str,
        svg_index: int = 0,
        output_path: str = None,
        wait_time: int = 2000
    ) -> Optional[str]:
        """
        同步按索引截取SVG(内部调用异步方法)

        Args:
            html_path: HTML文件路径
            svg_index: SVG元素索引
            output_path: 输出路径
            wait_time: 等待时间(毫秒)

        Returns:
            截图文件路径,失败返回None
        """
        return asyncio.run(
            self.capture_svg_by_index_async(html_path, svg_index, output_path, wait_time)
        )

    async def capture_chart_async(
        self,
        html_path: str,
        canvas_selector: str = "canvas",
        output_path: str = None,
        wait_time: int = 2000
    ) -> Optional[str]:
        """
        异步截取图表

        Args:
            html_path: HTML文件路径
            canvas_selector: Canvas元素选择器
            output_path: 输出路径,不指定则自动生成
            wait_time: 等待时间(毫秒),确保图表渲染完成

        Returns:
            截图文件路径,失败返回None
        """
        if not PLAYWRIGHT_AVAILABLE:
            logger.error("Playwright未安装,无法截图")
            return None

        html_path = Path(html_path).absolute()
        if not html_path.exists():
            logger.error(f"HTML文件不存在: {html_path}")
            return None

        # 生成输出路径
        if output_path is None:
            # 使用HTML内容hash作为缓存键
            cache_key = self._get_cache_key(str(html_path), canvas_selector)
            output_path = self.cache_dir / f"chart_{cache_key}.png"
        else:
            output_path = Path(output_path)

        # 检查缓存
        if output_path.exists():
            logger.info(f"使用缓存的图表截图: {output_path}")
            return str(output_path)

        try:
            async with async_playwright() as p:
                # 启动浏览器
                browser = None

                # 方案1: 尝试使用系统Chrome
                if self.use_system_chrome:
                    try:
                        browser = await p.chromium.launch(
                            headless=True,
                            channel="chrome"  # 使用系统安装的Chrome
                        )
                        logger.info("使用系统Chrome浏览器")
                    except Exception as chrome_error:
                        logger.warning(f"系统Chrome不可用: {chrome_error}")
                        logger.info("尝试使用Playwright Chromium...")

                # 方案2: 使用Playwright的Chromium
                if browser is None:
                    try:
                        browser = await p.chromium.launch(headless=True)
                        logger.info("使用Playwright Chromium浏览器")
                    except Exception as chromium_error:
                        error_msg = str(chromium_error)
                        if "Executable doesn't exist" in error_msg or "playwright install" in error_msg:
                            logger.error("浏览器不可用!")
                            logger.error("解决方案:")
                            logger.error("  1. 确保Chrome已安装(推荐)")
                            logger.error("  2. 或运行: playwright install chromium")
                            return None
                        else:
                            raise chromium_error

                # 创建页面
                page = await browser.new_page(
                    viewport={'width': 1920, 'height': 1080}
                )

                # 加载HTML
                file_url = html_path.as_uri()
                await page.goto(file_url, wait_until='networkidle')
                logger.info(f"加载HTML: {html_path}")

                # 等待Canvas元素
                await page.wait_for_selector(canvas_selector, timeout=10000)
                logger.info(f"找到Canvas元素: {canvas_selector}")

                # 等待图表渲染完成
                await page.wait_for_timeout(wait_time)
                logger.info(f"等待图表渲染 {wait_time}ms")

                # 截取Canvas元素
                canvas_element = await page.query_selector(canvas_selector)
                if canvas_element:
                    screenshot_bytes = await canvas_element.screenshot(
                        path=str(output_path),
                        type='png'
                    )
                    logger.info(f"图表截图成功: {output_path}")
                else:
                    logger.error(f"未找到Canvas元素: {canvas_selector}")
                    await browser.close()
                    return None

                # 关闭浏览器
                await browser.close()

                return str(output_path)

        except Exception as e:
            logger.error(f"图表截图失败: {e}")
            return None

    def capture_chart(
        self,
        html_path: str,
        canvas_selector: str = "canvas",
        output_path: str = None,
        wait_time: int = 2000
    ) -> Optional[str]:
        """
        同步截取图表(内部调用异步方法)

        Args:
            html_path: HTML文件路径
            canvas_selector: Canvas元素选择器
            output_path: 输出路径
            wait_time: 等待时间(毫秒)

        Returns:
            截图文件路径,失败返回None
        """
        return asyncio.run(
            self.capture_chart_async(html_path, canvas_selector, output_path, wait_time)
        )

    async def capture_multiple_charts_async(
        self,
        html_path: str,
        chart_configs: List[Dict]
    ) -> Dict[str, str]:
        """
        异步截取多个图表

        Args:
            html_path: HTML文件路径
            chart_configs: 图表配置列表,每项包含:
                - canvas_id: Canvas元素ID
                - selector: CSS选择器(可选,优先级高于canvas_id)
                - output_path: 输出路径(可选)

        Returns:
            图表ID到截图路径的映射
        """
        if not PLAYWRIGHT_AVAILABLE:
            logger.error("Playwright未安装,无法截图")
            return {}

        results = {}

        for config in chart_configs:
            canvas_id = config.get('canvas_id')
            selector = config.get('selector', f'#{canvas_id}' if canvas_id else 'canvas')
            output_path = config.get('output_path')

            screenshot_path = await self.capture_chart_async(
                html_path, selector, output_path
            )

            if screenshot_path:
                results[canvas_id or selector] = screenshot_path

        return results

    def capture_multiple_charts(
        self,
        html_path: str,
        chart_configs: List[Dict]
    ) -> Dict[str, str]:
        """
        同步截取多个图表

        Args:
            html_path: HTML文件路径
            chart_configs: 图表配置列表

        Returns:
            图表ID到截图路径的映射
        """
        return asyncio.run(
            self.capture_multiple_charts_async(html_path, chart_configs)
        )

    def _get_cache_key(self, html_path: str, selector: str) -> str:
        """
        生成缓存键

        Args:
            html_path: HTML文件路径
            selector: 选择器

        Returns:
            缓存键(hash)
        """
        # 读取HTML内容生成hash
        try:
            with open(html_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 组合HTML内容和选择器生成唯一键
            key_str = f"{content}:{selector}"
            cache_key = hashlib.md5(key_str.encode()).hexdigest()[:16]
            return cache_key
        except Exception as e:
            logger.error(f"生成缓存键失败: {e}")
            # 降级为文件名+选择器
            return hashlib.md5(f"{html_path}:{selector}".encode()).hexdigest()[:16]

    def clear_cache(self):
        """清除所有缓存的图表截图"""
        if self.cache_dir.exists():
            # 清除chart缓存
            for file in self.cache_dir.glob("chart_*.png"):
                try:
                    file.unlink()
                    logger.info(f"删除缓存: {file}")
                except Exception as e:
                    logger.error(f"删除缓存失败: {e}")

            # 清除svg缓存
            for file in self.cache_dir.glob("svg_*.png"):
                try:
                    file.unlink()
                    logger.info(f"删除SVG缓存: {file}")
                except Exception as e:
                    logger.error(f"删除SVG缓存失败: {e}")

    @staticmethod
    def is_available() -> bool:
        """检查Playwright是否可用"""
        return PLAYWRIGHT_AVAILABLE


# 便捷函数
def capture_chart(html_path: str, canvas_selector: str = "canvas", output_path: str = None) -> Optional[str]:
    """
    快速截取图表(单例模式)

    Args:
        html_path: HTML文件路径
        canvas_selector: Canvas选择器
        output_path: 输出路径

    Returns:
        截图路径
    """
    capturer = ChartCapture()
    return capturer.capture_chart(html_path, canvas_selector, output_path)


def capture_svg(html_path: str, svg_selector: str = "svg", output_path: str = None) -> Optional[str]:
    """
    快速截取SVG(单例模式)

    Args:
        html_path: HTML文件路径
        svg_selector: SVG选择器
        output_path: 输出路径

    Returns:
        截图路径
    """
    capturer = ChartCapture()
    return capturer.capture_svg(html_path, svg_selector, output_path)


def capture_svg_by_index(html_path: str, svg_index: int = 0, output_path: str = None) -> Optional[str]:
    """
    快速按索引截取SVG(单例模式)

    Args:
        html_path: HTML文件路径
        svg_index: SVG索引
        output_path: 输出路径

    Returns:
        截图路径
    """
    capturer = ChartCapture()
    return capturer.capture_svg_by_index(html_path, svg_index, output_path)
