"""批量更新所有converter文件的字体硬编码"""
import os
import re

files_to_update = [
    'src/converters/shape_converter.py',
    'src/converters/chart_converter.py',
    'src/converters/timeline_converter.py',
    'src/converters/table_converter.py',
]

def update_file(file_path):
    """更新单个文件"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 统计原有的硬编码字体数量
    original_count = content.count("run.font.name = 'Microsoft YaHei'")
    
    if original_count == 0:
        print(f"跳过 {file_path} (没有硬编码字体)")
        return 0
    
    # 1. 添加import (如果还没有)
    if 'from src.utils.font_manager import get_font_manager' not in content:
        # 在logger导入后添加
        content = content.replace(
            'from src.utils.logger import setup_logger',
            'from src.utils.logger import setup_logger\nfrom src.utils.font_manager import get_font_manager'
        )
    
    # 2. 替换所有硬编码字体为动态获取
    # 使用 get_font_manager(self.css_parser).get_font('body')
    content = content.replace(
        "run.font.name = 'Microsoft YaHei'",
        "run.font.name = get_font_manager(self.css_parser).get_font('body')"
    )
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"已更新 {file_path} ({original_count} 处)")
    return original_count

if __name__ == '__main__':
    total = 0
    for file_path in files_to_update:
        if os.path.exists(file_path):
            total += update_file(file_path)
        else:
            print(f"文件不存在: {file_path}")
    
    print(f"\n总计更新: {total} 处")
