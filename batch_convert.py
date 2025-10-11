"""
批量转换工具
支持一次转换多个HTML文件
"""

import sys
from pathlib import Path
from src.main import HTML2PPTX
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


def batch_convert(input_dir: str, output_dir: str, pattern: str = "*.html"):
    """
    批量转换HTML文件

    Args:
        input_dir: 输入目录
        output_dir: 输出目录
        pattern: 文件匹配模式
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists():
        logger.error(f"输入目录不存在: {input_dir}")
        return

    # 创建输出目录
    output_path.mkdir(parents=True, exist_ok=True)

    # 查找所有HTML文件
    html_files = list(input_path.glob(pattern))

    if not html_files:
        logger.warning(f"未找到匹配的HTML文件: {pattern}")
        return

    logger.info(f"找到 {len(html_files)} 个HTML文件")

    # 批量转换
    success_count = 0
    fail_count = 0

    for html_file in html_files:
        try:
            logger.info(f"\n{'='*60}")
            logger.info(f"转换: {html_file.name}")
            logger.info(f"{'='*60}")

            # 生成输出文件名
            output_file = output_path / f"{html_file.stem}.pptx"

            # 执行转换
            converter = HTML2PPTX(str(html_file))
            converter.convert(str(output_file))

            success_count += 1
            logger.info(f"✓ 成功: {html_file.name} → {output_file.name}")

        except Exception as e:
            fail_count += 1
            logger.error(f"✗ 失败: {html_file.name} - {e}")

    # 输出统计
    logger.info(f"\n{'='*60}")
    logger.info(f"批量转换完成!")
    logger.info(f"成功: {success_count} 个")
    logger.info(f"失败: {fail_count} 个")
    logger.info(f"总计: {len(html_files)} 个")
    logger.info(f"{'='*60}")


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("用法: python batch_convert.py <输入目录> [输出目录] [文件模式]")
        print("\n示例:")
        print("  python batch_convert.py ./slides ./output")
        print("  python batch_convert.py ./slides ./output '*.html'")
        sys.exit(1)

    input_dir = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "./output"
    pattern = sys.argv[3] if len(sys.argv) > 3 else "*.html"

    batch_convert(input_dir, output_dir, pattern)


if __name__ == "__main__":
    main()
