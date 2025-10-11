"""
测试图表截图功能
"""

import sys
from pathlib import Path

# 添加src到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from src.utils.chart_capture import ChartCapture
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


def test_chart_capture():
    """测试图表截图"""

    # 检查Playwright是否可用
    if not ChartCapture.is_available():
        logger.warning("Playwright未安装!")
        logger.info("请运行以下命令安装:")
        logger.info("  pip install playwright")
        logger.info("  playwright install chromium")
        return False

    logger.info("Playwright已安装")

    # 测试文件
    html_file = Path(__file__).parent / "slidewithtable.html"

    if not html_file.exists():
        print(f"❌ 测试文件不存在: {html_file}")
        return False

    logger.info(f"找到测试文件: {html_file}")

    # 创建截图工具
    capturer = ChartCapture()
    logger.info(f"图表缓存目录: {capturer.cache_dir}")

    # 测试截图
    logger.info("开始截取图表...")
    logger.info("-" * 50)

    screenshot_path = capturer.capture_chart(
        str(html_file),
        canvas_selector="#vulnerabilityChart",
        wait_time=3000  # 等待3秒确保图表渲染完成
    )

    if screenshot_path:
        screenshot_file = Path(screenshot_path)
        if screenshot_file.exists():
            size_kb = screenshot_file.stat().st_size / 1024
            logger.info("图表截图成功!")
            logger.info(f"  路径: {screenshot_path}")
            logger.info(f"  大小: {size_kb:.2f} KB")
            return True
        else:
            logger.error(f"截图文件不存在: {screenshot_path}")
            return False
    else:
        logger.error("图表截图失败")
        return False


def main():
    """主函数"""
    logger.info("=" * 50)
    logger.info("图表截图功能测试")
    logger.info("=" * 50)

    success = test_chart_capture()

    logger.info("=" * 50)
    if success:
        logger.info("测试通过!")
    else:
        logger.error("测试失败!")
    logger.info("=" * 50)

    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
