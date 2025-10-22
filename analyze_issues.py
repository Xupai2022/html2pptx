from bs4 import BeautifulSoup

# 分析slide_005
print("=== Slide 005 分析 ===")
with open("input/slide_005.html", 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'lxml')
    
flex_container = soup.find('div', class_='flex-1')
grid = flex_container.find('div', class_='grid') if flex_container else None
if grid:
    stat_cards = grid.find_all('div', class_='stat-card', recursive=False)
    for idx, card in enumerate(stat_cards):
        print(f"\nstat-card #{idx+1}:")
        print(f"  classes: {card.get('class')}")
        # 检查是否有text-center和flex
        if 'text-center' in card.get('class', []):
            print("  -> 有text-center类")
        if 'flex' in card.get('class', []):
            print("  -> 有flex类")
            flex_classes = card.get('class', [])
            if 'flex-col' in flex_classes:
                print("  -> 有flex-col类")
            if 'justify-center' in flex_classes:
                print("  -> 有justify-center类（100%高度）")
            if 'items-center' in flex_classes:
                print("  -> 有items-center类")
        
        # 查找内部元素
        text_7xl = card.find('div', class_='text-7xl')
        h3 = card.find('h3')
        if text_7xl:
            print(f"  text-7xl: {text_7xl.get_text(strip=True)}")
        if h3:
            print(f"  h3 (mt-2): {h3.get_text(strip=True)}")

# 分析slide_006
print("\n\n=== Slide 006 分析 ===")
with open("input/slide_006.html", 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'lxml')
    
flex_container = soup.find('div', class_='flex-1')
# 先查找grid
grid = flex_container.find('div', class_='grid') if flex_container else None
if grid:
    print("\n第一个容器是grid:")
    print(f"  grid classes: {grid.get('class')}")
    data_cards = grid.find_all('div', class_='data-card', recursive=False)
    print(f"  包含{len(data_cards)}个data-card")
    for idx, card in enumerate(data_cards):
        h3 = card.find('h3')
        p = card.find('p')
        if h3:
            print(f"    data-card #{idx+1} h3: {h3.get_text(strip=True)}")
            # 检查h3是否会被重复渲染
            print(f"      h3.parent: {h3.parent.name}, classes: {h3.parent.get('class', [])}")

# 分析后面的data-card
second_container = None
for child in flex_container.children if flex_container else []:
    if hasattr(child, 'get') and child.get('class'):
        classes = child.get('class', [])
        if 'data-card' in classes and child != grid:
            second_container = child
            break

if second_container:
    print(f"\n第二个容器是data-card:")
    p = second_container.find('p')
    if p:
        print(f"  p内容: {p.get_text(strip=True)[:80]}...")

# 分析slide_008
print("\n\n=== Slide 008 分析 ===")
with open("input/slide_008.html", 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'lxml')
    
flex_container = soup.find('div', class_='flex-1')
if flex_container:
    # 第一个grid容器
    grid1 = flex_container.find('div', class_='grid')
    if grid1:
        print("第一个grid容器:")
        stat_cards = grid1.find_all('div', class_='stat-card', recursive=False)
        for idx, card in enumerate(stat_cards):
            h3 = card.find('h3')
            p = card.find('p')
            if h3:
                print(f"  stat-card #{idx+1}: h3={h3.get_text(strip=True)}")
            if p:
                print(f"    p={p.get_text(strip=True)}")
    
    # 查找第二个容器 (包含"平均响应时间"的data-card)
    data_cards = flex_container.find_all('div', class_='data-card', recursive=False)
    print(f"\n找到{len(data_cards)}个data-card:")
    for idx, card in enumerate(data_cards):
        h3 = card.find('h3')
        p = card.find('p')
        if h3:
            h3_text = h3.get_text(strip=True)
            print(f"  data-card #{idx+1}: h3={h3_text}")
            if "平均响应时间" in h3_text:
                print(f"    -> 这是包含'平均响应时间'的data-card")
                if p:
                    p_text = p.get_text(strip=True)
                    print(f"    p={p_text}")
                    if "41.0分钟" not in p_text:
                        print(f"    ❌ 警告：p标签没有包含'41.0分钟'")
                        print(f"    检查HTML结构...")
                        # 检查是否p标签在h3外面
                        for elem in card.children:
                            if hasattr(elem, 'name') and elem.name == 'p':
                                print(f"      找到p标签: {elem.get_text(strip=True)}")
                else:
                    print(f"    ❌ 警告：没有找到p标签")

