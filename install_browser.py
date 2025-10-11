"""
Playwright浏览器安装辅助脚本
"""

import subprocess
import sys
from pathlib import Path

# 添加src到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from src.utils.logger import setup_logger

logger = setup_logger(__name__)


def check_playwright():
    """检查Playwright是否已安装"""
    try:
        import playwright
        logger.info("Playwright已安装")
        return True
    except ImportError:
        logger.error("Playwright未安装!")
        logger.error("请先运行: pip install playwright")
        return False


def install_chromium():
    """安装Chromium浏览器"""
    logger.info("=" * 50)
    logger.info("开始安装Chromium浏览器...")
    logger.info("=" * 50)

    try:
        # 运行playwright install chromium
        result = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            timeout=300  # 5分钟超时
        )

        if result.returncode == 0:
            logger.info("=" * 50)
            logger.info("Chromium浏览器安装成功!")
            logger.info("=" * 50)
            return True
        else:
            logger.error("安装失败!")
            logger.error(f"错误信息: {result.stderr}")
            return False

    except subprocess.TimeoutExpired:
        logger.error("安装超时(5分钟),请检查网络连接")
        return False
    except Exception as e:
        logger.error(f"安装异常: {e}")
        return False


def verify_installation():
    """验证安装"""
    logger.info("\n验证安装...")

    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            logger.info("Chromium浏览器启动成功!")
            browser.close()
            return True

    except Exception as e:
        logger.error(f"验证失败: {e}")
        return False


def main():
    """主函数"""
    logger.info("Playwright浏览器安装工具\n")

    # 检查Playwright
    if not check_playwright():
        return 1

    # 安装Chromium
    logger.info("\n即将下载Chromium浏览器(约300MB)")
    logger.info("这可能需要几分钟时间,请耐心等待...\n")

    if not install_chromium():
        return 1

    # 验证安装
    if verify_installation():
        logger.info("\n" + "=" * 50)
        logger.info("安装完成!现在可以使用图表截图功能了")
        logger.info("=" * 50)
        logger.info("\n运行转换:")
        logger.info("  python convert.py slidewithtable.html output/result.pptx")
        return 0
    else:
        logger.error("\n安装验证失败,请手动运行:")
        logger.error("  playwright install chromium")
        return 1


if __name__ == "__main__":
    sys.exit(main())
