"""
批量合并HTML转PPTX启动脚本
将input目录下的所有HTML文件合并为一个PPTX文件
"""

import sys
from pathlib import Path

# 添加src目录到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from src.batch_merger import BatchHTML2PPTXMerger

def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(
        description='批量合并HTML文件为单个PPTX文件',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python batch_merge.py input                    # 将input目录下的HTML合并为output/merged.pptx
  python batch_merge.py input output/report.pptx # 指定输出路径
  python batch_merge.py ./slides ./final.pptx    # 使用相对路径
        """
    )

    parser.add_argument('input_dir', help='HTML文件目录路径')
    parser.add_argument('output_path', nargs='?',
                       default='output/merged.pptx',
                       help='输出PPTX文件路径 (默认: output/merged.pptx)')

    args = parser.parse_args()

    try:
        print("=" * 60)
        print("批量HTML转PPTX合并工具")
        print("=" * 60)
        print(f"输入目录: {args.input_dir}")
        print(f"输出文件: {args.output_path}")
        print("-" * 60)

        # 执行批量转换
        merger = BatchHTML2PPTXMerger(args.input_dir)
        merger.convert(args.output_path)

        print("=" * 60)
        print("✓ 批量转换完成！")
        print(f"输出文件: {args.output_path}")
        print("=" * 60)

    except Exception as e:
        print(f"\n✗ 错误: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()