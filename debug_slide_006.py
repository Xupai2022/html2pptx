from bs4 import BeautifulSoup

with open("input/slide_006.html", 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'lxml')

flex_container = soup.find('div', class_='flex-1')
grid = flex_container.find('div', class_='grid')

print("=== Grid容器内的data-card ===")
data_cards = grid.find_all('div', class_='data-card', recursive=False)
for idx, card in enumerate(data_cards):
    print(f"\ndata-card #{idx+1}:")
    # 查找h3 (recursive=False vs recursive=True)
    h3_direct = card.find('h3', recursive=False)
    h3_all = card.find_all('h3')
    
    print(f"  find('h3', recursive=False): {h3_direct.get_text(strip=True) if h3_direct else 'None'}")
    print(f"  find_all('h3'): {[h.get_text(strip=True) for h in h3_all]}")
    
    # 查找p标签
    p_direct = card.find('p', recursive=False)
    p_all = card.find_all('p')
    print(f"  find('p', recursive=False): {p_direct.get_text(strip=True)[:50] if p_direct else 'None'}...")
    print(f"  find_all('p'): 共{len(p_all)}个")

