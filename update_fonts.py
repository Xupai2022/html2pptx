"""批量更新字体硬编码为动态获取"""
import re

def update_main_py():
    with open('src/main.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 简单替换所有 'Microsoft YaHei' 为 self.font_manager.get_font('body')
    # 这是一个通用方案，使用body作为默认选择器
    
    original = content
    # 替换所有硬编码字体
    content = content.replace(
        "run.font.name = 'Microsoft YaHei'",
        "run.font.name = self.font_manager.get_font('body')"
    )
    
    changes = original.count("run.font.name = 'Microsoft YaHei'")
    
    with open('src/main.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✅ 更新了 src/main.py 中的 {changes} 处字体设置")
    return changes

if __name__ == '__main__':
    total = update_main_py()
    print(f"\n总计更新: {total} 处")
