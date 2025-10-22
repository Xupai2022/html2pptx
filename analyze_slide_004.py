from bs4 import BeautifulSoup

html_path = "input/slide_004.html"
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

soup = BeautifulSoup(html, 'lxml')
slide = soup.find('div', class_='slide-container')
content_section = slide.find('div', class_='content-section')

print("=== Slide 004 HTML结构分析 ===\n")
print("标题区域:")
print("- .mb-6 容器（标题区域）")
h1 = slide.find('h1')
h2 = slide.find('h2')
print(f"  - h1: {h1.get_text(strip=True)} (font-size: 48px)")
print(f"  - h2: {h2.get_text(strip=True)} (font-size: 36px)")
print(f"  - 装饰线 (h-1 = 4px)")

print("\n内容区域 (.flex-1.overflow-hidden):")
flex_overflow = content_section.find('div', class_='flex-1')
if flex_overflow:
    # 第一个容器 (grid grid-cols-2 gap-6 mb-6)
    grid1 = flex_overflow.find('div', class_='grid')
    if grid1:
        print(f"\n1. 第一个网格容器: {grid1.get('class')}")
        stat_cards = grid1.find_all('div', class_='stat-card', recursive=False)
        for idx, card in enumerate(stat_cards):
            h3 = card.find('h3')
            p = card.find('p', class_='text-4xl')
            print(f"   - stat-card #{idx+1}:")
            print(f"     - h3: {h3.get_text(strip=True)} (font-size: 28px)")
            print(f"     - p.text-4xl: {p.get_text(strip=True)}")
            print(f"     - 内部结构: stat-card")
        
        # 分析间距
        print(f"   - mb-6 = 24px margin-bottom")
    
    # 获取所有data-card
    data_cards = flex_overflow.find_all('div', class_='data-card', recursive=False)
    print(f"\n2-4. Data-card容器: {len(data_cards)}个")
    for idx, card in enumerate(data_cards):
        h3 = card.find('h3')
        p = card.find('p')
        print(f"   - data-card #{idx+1}:")
        if h3:
            print(f"     - h3: {h3.get_text(strip=True)[:30]}... (font-size: 28px)")
        if p:
            print(f"     - p: {p.get_text(strip=True)[:50]}... (font-size: 25px)")
        print(f"     - CSS: border-left: 4px, padding: 15px 20px")
        print(f"     - CSS: margin-bottom: 20px")

print("\n=== 计算精确的Y坐标 ===")
print("假设content-section padding-top = 40px")
print("从代码看：content-section { flex: 1; padding: 40px 80px 60px 80px; }")
print("\nY坐标计算:")
y = 40  # content-section padding-top
print(f"起始: y = {y}px (content-section padding-top)")

# 标题区域 (.mb-6)
print(f"\n标题区域 (.mb-6):")
y += 0  # mb-6容器没有mt
print(f"  - h1 (48px * 1.5 line-height) = 72px")
print(f"    - mb-20 = 20px")
y_after_h1 = y + 72 + 20
print(f"  - y_after_h1 = {y_after_h1}px")
print(f"  - h2 (36px * 1.5 line-height) = 54px")
y_after_h2 = y_after_h1 + 54
print(f"  - y_after_h2 = {y_after_h2}px")
print(f"  - 装饰线 (h-1 = 4px)")
y_after_line = y_after_h2 + 4
print(f"  - y_after_line = {y_after_line}px")
print(f"  - mb-6容器的margin-bottom = 24px")
y = y_after_line + 24
print(f"  - 标题区域结束: y = {y}px")

# 第一个grid容器
print(f"\n第一个grid容器 (grid grid-cols-2 gap-6 mb-6):")
print(f"  - 起始: y = {y}px")
print(f"  - stat-card内部:")
print(f"    - h3 (28px) + mb-10px = 38px")
print(f"    - p.text-4xl (36px) = 36px")
print(f"    - 总高度约: padding (15+15) + h3 (38) + p (36) = 104px")
print(f"  - mb-6 = 24px margin-bottom")
y_after_grid1 = y + 104 + 24
print(f"  - 第一个grid结束: y = {y_after_grid1}px")

# 三个data-card
print(f"\n三个data-card:")
for i in range(3):
    print(f"  - data-card #{i+1}:")
    print(f"    - 起始: y = {y_after_grid1}px")
    print(f"    - h3 (28px) + mb-10px = 38px")
    print(f"    - p (25px * 1.6 line-height * 2行) = 80px")
    print(f"    - padding (15+15) = 30px")
    print(f"    - 总高度约: 30 + 38 + 80 = 148px")
    print(f"    - margin-bottom = 20px")
    y_after_grid1 += 148 + 20
    print(f"    - 结束: y = {y_after_grid1}px")

