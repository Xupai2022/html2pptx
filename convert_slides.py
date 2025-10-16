#!/usr/bin/env python3
"""
批量转换slide开头的HTML文件为PPTX

使用方法:
python convert_slides.py

功能:
- 自动查找input目录下所有以"slide"开头的.html文件
- 批量转换为PPTX文件
- 输出到output目录
- 显示转换进度和结果统计
"""

import os
import sys
import glob
from pathlib import Path
import subprocess
import time

def find_slide_html_files():
    """查找input目录下所有以slide开头的HTML文件"""
    input_dir = "input"
    pattern = os.path.join(input_dir, "slide*.html")
    files = glob.glob(pattern)
    files.sort()  # 按文件名排序
    return files

def convert_single_file(html_file, output_dir):
    """转换单个HTML文件为PPTX"""
    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        # 生成输出文件名
        html_name = Path(html_file).stem
        output_file = os.path.join(output_dir, f"{html_name}.pptx")

        print(f"  转换: {html_file} -> {output_file}")

        # 调用转换程序
        result = subprocess.run([
            sys.executable, "convert.py", html_file, output_file
        ], capture_output=True, text=True, encoding='utf-8')

        if result.returncode == 0:
            print(f"  [SUCCESS] 成功: {output_file}")
            return True, output_file
        else:
            print(f"  [FAILED] 失败: {html_file}")
            print(f"     错误信息: {result.stderr}")
            return False, None

    except Exception as e:
        print(f"  [ERROR] 异常: {html_file} - {str(e)}")
        return False, None

def main():
    """主函数"""
    print("=" * 60)
    print("批量转换slide开头的HTML文件为PPTX")
    print("=" * 60)

    # 查找所有slide开头的HTML文件
    html_files = find_slide_html_files()

    if not html_files:
        print("[ERROR] 未找到任何以'slide'开头的HTML文件")
        print("   请确保input目录下有slide*.html文件")
        return

    print(f"[INFO] 找到 {len(html_files)} 个HTML文件:")
    for i, file in enumerate(html_files, 1):
        print(f"   {i}. {file}")

    print("\n[INFO] 开始批量转换...")
    print("-" * 60)

    # 输出目录
    output_dir = "output"

    # 统计结果
    success_count = 0
    failed_count = 0
    success_files = []
    failed_files = []

    # 开始转换
    start_time = time.time()

    for i, html_file in enumerate(html_files, 1):
        print(f"\n[{i}/{len(html_files)}] 正在处理: {html_file}")

        success, output_file = convert_single_file(html_file, output_dir)

        if success:
            success_count += 1
            success_files.append((html_file, output_file))
        else:
            failed_count += 1
            failed_files.append(html_file)

    # 显示转换结果统计
    end_time = time.time()
    duration = end_time - start_time

    print("\n" + "=" * 60)
    print("[STATS] 转换结果统计")
    print("=" * 60)
    print(f"总文件数: {len(html_files)}")
    print(f"成功转换: {success_count}")
    print(f"转换失败: {failed_count}")
    print(f"总耗时: {duration:.2f} 秒")

    if success_count > 0:
        print(f"\n[SUCCESS] 成功转换的文件:")
        for html_file, output_file in success_files:
            file_size = os.path.getsize(output_file) / 1024  # KB
            print(f"   {html_file} -> {output_file} ({file_size:.1f} KB)")

    if failed_count > 0:
        print(f"\n[FAILED] 转换失败的文件:")
        for html_file in failed_files:
            print(f"   {html_file}")

    print(f"\n[OUTPUT] 输出目录: {os.path.abspath(output_dir)}")

    if success_count == len(html_files):
        print("\n[COMPLETE] 所有文件转换成功！")
    elif success_count > 0:
        print(f"\n[PARTIAL] 部分文件转换成功，请检查失败的文件")
    else:
        print(f"\n[FAILED] 所有文件转换失败，请检查错误信息")

if __name__ == "__main__":
    main()