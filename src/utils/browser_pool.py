"""
浏览器池管理
提供浏览器实例复用，提高截图效率
"""

import asyncio
from typing import Optional, List
from pathlib import Path
import logging

try:
    from playwright.async_api import async_playwright, Browser, BrowserContext, Page
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

logger = logging.getLogger(__name__)


class BrowserPool:
    """浏览器池管理器"""

    def __init__(self, max_pages: int = 3, use_system_chrome: bool = True):
        """
        初始化浏览器池

        Args:
            max_pages: 最大页面数
            use_system_chrome: 是否使用系统Chrome
        """
        if not PLAYWRIGHT_AVAILABLE:
            raise ImportError("Playwright未安装，请运行: pip install playwright")

        self.max_pages = max_pages
        self.use_system_chrome = use_system_chrome
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.available_pages: asyncio.Queue = asyncio.Queue(maxsize=max_pages)
        self.used_pages: List[Page] = []
        self._lock = asyncio.Lock()
        self._initialized = False

    async def _initialize(self):
        """初始化浏览器"""
        if self._initialized:
            return

        async with self._lock:
            if self._initialized:
                return

            logger.info("初始化浏览器池...")

            try:
                playwright = await async_playwright().start()

                # 尝试启动浏览器
                if self.use_system_chrome:
                    try:
                        self.browser = await playwright.chromium.launch(
                            headless=True,
                            channel="chrome",
                            timeout=5000  # 5秒启动超时
                        )
                        logger.info("使用系统Chrome浏览器")
                    except Exception as e:
                        logger.warning(f"系统Chrome不可用: {e}")
                        logger.info("使用Playwright Chromium...")

                if not self.browser:
                    self.browser = await playwright.chromium.launch(
                        headless=True,
                        timeout=5000
                    )
                    logger.info("使用Playwright Chromium浏览器")

                # 创建上下文
                self.context = await self.browser.new_context(
                    viewport={'width': 1920, 'height': 1080},
                    ignore_https_errors=True
                )

                # 预创建页面
                for _ in range(self.max_pages):
                    page = await self.context.new_page()
                    # 设置更短的超时
                    page.set_default_timeout(3000)
                    page.set_default_navigation_timeout(5000)
                    await self.available_pages.put(page)

                self._initialized = True
                logger.info(f"浏览器池初始化完成，创建了 {self.max_pages} 个页面")

            except Exception as e:
                logger.error(f"浏览器池初始化失败: {e}")
                raise

    async def get_page(self) -> Page:
        """
        获取一个可用页面（简化版，避免队列问题）

        Returns:
            页面对象
        """
        await self._initialize()

        # 简化：总是创建新页面，避免队列在不同事件循环中的问题
        if self.context:
            page = await self.context.new_page()
            page.set_default_timeout(3000)
            page.set_default_navigation_timeout(5000)
            logger.debug("创建新页面")
            return page
        else:
            raise RuntimeError("浏览器未初始化")

    async def close_page(self, page: Page):
        """
        关闭或回收页面

        Args:
            page: 要关闭的页面
        """
        if page in self.used_pages:
            self.used_pages.remove(page)

            try:
                # 清理页面内容
                await page.evaluate("""
                    // 清除所有内容
                """)

                # 如果页面池未满，回收页面
                if self.available_pages.qsize() < self.max_pages:
                    await self.available_pages.put(page)
                    logger.debug(f"回收页面，可用: {self.available_pages.qsize()}")
                else:
                    # 页面池已满，关闭页面
                    await page.close()
                    logger.debug("关闭多余页面")
            except Exception as e:
                logger.warning(f"回收页面失败: {e}")
                try:
                    await page.close()
                except:
                    pass

    async def close_all(self):
        """关闭所有页面和浏览器"""
        logger.info("关闭浏览器池...")

        # 关闭所有使用的页面
        for page in self.used_pages[:]:
            try:
                await page.close()
            except:
                pass
        self.used_pages.clear()

        # 关闭池中的页面
        while not self.available_pages.empty():
            try:
                page = self.available_pages.get_nowait()
                await page.close()
            except:
                pass

        # 关闭上下文
        if self.context:
            try:
                await self.context.close()
            except:
                pass
            self.context = None

        # 关闭浏览器
        if self.browser:
            try:
                await self.browser.close()
            except:
                pass
            self.browser = None

        self._initialized = False
        logger.info("浏览器池已关闭")


# 全局浏览器池实例
_browser_pool: Optional[BrowserPool] = None


async def get_browser_pool() -> BrowserPool:
    """获取全局浏览器池实例"""
    global _browser_pool
    if _browser_pool is None:
        _browser_pool = BrowserPool()
    return _browser_pool


# 便捷函数
async def get_page() -> Page:
    """获取一个页面"""
    pool = await get_browser_pool()
    return await pool.get_page()


async def close_page(page: Page):
    """关闭页面"""
    pool = await get_browser_pool()
    await pool.close_page(page)


async def cleanup_browser_pool():
    """清理浏览器池"""
    global _browser_pool
    if _browser_pool:
        await _browser_pool.close_all()
        _browser_pool = None


# 确保程序退出时清理
import atexit
atexit.register(lambda: asyncio.create_task(cleanup_browser_pool()))