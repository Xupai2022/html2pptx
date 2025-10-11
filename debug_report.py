"""
调试报告生成器
生成HTML元素到PPTX元素的映射报告,便于调试
"""

import sys
from pathlib import Path
from src.parser.html_parser import HTMLParser
from src.parser.css_parser import CSSParser
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class DebugReporter:
    """调试报告生成器"""

    def __init__(self, html_path: str):
        """初始化"""
        self.html_path = html_path
        self.html_parser = HTMLParser(html_path)
        self.css_parser = CSSParser(self.html_parser.soup)

    def generate_report(self, output_path: str = None):
        """
        生成调试报告

        Args:
            output_path: 输出路径
        """
        if output_path is None:
            output_path = "output/debug_report.md"

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        report_lines = []
        report_lines.append("# HTML转PPTX调试报告\n")
        report_lines.append(f"**源文件**: {self.html_path}\n")
        report_lines.append(f"**生成时间**: {Path(self.html_path).stat().st_mtime}\n\n")

        # CSS规则统计
        report_lines.append("## CSS样式规则\n")
        report_lines.append(f"共解析 {len(self.css_parser.style_rules)} 条规则\n\n")
        report_lines.append("| 选择器 | 属性数量 | 关键属性 |\n")
        report_lines.append("|--------|---------|----------|\n")

        for selector, props in self.css_parser.style_rules.items():
            key_props = []
            if 'font-size' in props:
                key_props.append(f"font-size: {props['font-size']}")
            if 'color' in props:
                key_props.append(f"color: {props['color']}")
            if 'background-color' in props:
                key_props.append(f"bg: {props['background-color']}")

            report_lines.append(f"| `{selector}` | {len(props)} | {', '.join(key_props) if key_props else '-'} |\n")

        # 幻灯片结构分析
        slides = self.html_parser.get_slides()
        report_lines.append(f"\n## 幻灯片结构\n")
        report_lines.append(f"共 {len(slides)} 个幻灯片\n\n")

        for idx, slide in enumerate(slides, 1):
            report_lines.append(f"### 幻灯片 {idx}\n\n")

            # 标题
            title = self.html_parser.get_title(slide)
            subtitle = self.html_parser.get_subtitle(slide)
            if title:
                report_lines.append(f"**标题**: {title}\n")
            if subtitle:
                report_lines.append(f"**副标题**: {subtitle}\n")

            # 元素统计
            stat_boxes = self.html_parser.get_stat_boxes(slide)
            stat_cards = self.html_parser.get_stat_cards(slide)
            data_cards = self.html_parser.get_data_cards(slide)
            tables = self.html_parser.get_tables(slide)
            paragraphs = self.html_parser.get_paragraphs(slide)

            report_lines.append("\n**元素统计**:\n")
            report_lines.append(f"- 统计卡片(stat-box): {len(stat_boxes)}\n")
            report_lines.append(f"- 图表卡片(stat-card): {len(stat_cards)}\n")
            report_lines.append(f"- 数据卡片(data-card): {len(data_cards)}\n")
            report_lines.append(f"- 表格(table): {len(tables)}\n")
            report_lines.append(f"- 段落(p): {len(paragraphs)}\n")

            # stat-box详情
            if stat_boxes:
                report_lines.append("\n**统计卡片内容**:\n")
                for i, box in enumerate(stat_boxes, 1):
                    icon = box.find('i')
                    title_elem = box.find('div', class_='stat-title')
                    h2 = box.find('h2')
                    p = box.find('p')

                    icon_text = f"[{icon.get('class', [])}]" if icon else "-"
                    title_text = title_elem.get_text(strip=True) if title_elem else "-"
                    h2_text = h2.get_text(strip=True) if h2 else "-"
                    p_text = p.get_text(strip=True) if p else "-"

                    report_lines.append(f"{i}. 图标: {icon_text}, 标题: {title_text}, 数据: {h2_text}, 描述: {p_text}\n")

            # data-card详情
            if data_cards:
                report_lines.append("\n**数据卡片内容**:\n")
                for i, card in enumerate(data_cards, 1):
                    title_p = card.find('p', class_='primary-color')
                    progress_bars = self.html_parser.get_progress_bars(card)
                    bullet_points = self.html_parser.get_bullet_points(card)

                    title_text = title_p.get_text(strip=True) if title_p else "-"
                    report_lines.append(f"{i}. 标题: {title_text}\n")
                    report_lines.append(f"   - 进度条: {len(progress_bars)} 个\n")
                    report_lines.append(f"   - 列表项: {len(bullet_points)} 个\n")

            # 页码
            page_num = self.html_parser.get_page_number(slide)
            if page_num:
                report_lines.append(f"\n**页码**: {page_num}\n")

            report_lines.append("\n---\n\n")

        # 写入文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.writelines(report_lines)

        logger.info(f"调试报告已生成: {output_path}")
        print(f"\n✓ 调试报告已生成: {output_path}")


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("用法: python debug_report.py <HTML文件路径> [输出报告路径]")
        print("\n示例:")
        print("  python debug_report.py slidewithtable.html")
        print("  python debug_report.py slidewithtable.html output/report.md")
        sys.exit(1)

    html_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None

    reporter = DebugReporter(html_path)
    reporter.generate_report(output_path)


if __name__ == "__main__":
    main()
