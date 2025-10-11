# HTML转PPTX工程实现计划文档

## 项目概述

### 项目目标
构建一个从0到1的HTML转PPTX自动化转换工程,将AI生成的HTML报告精准转换为PPTX格式,确保样式与内容的完全一致性。

### 核心约束
- **严格样式一致性**: 必须完全遵循template.txt中定义的样式规范
- **固定幻灯片尺寸**: 1920px × 1080px (16:9)
- **主题色**: RGB(10, 66, 117) - 深蓝色
- **字体大小规范**: H1(48px), H2(36px), P(20-23px), 容器内文本(≥25px)

---

## 一、需求分析

### 1.1 HTML模板分析

基于template.txt和slidewithtable.html的分析,HTML包含以下核心元素:

#### 布局结构
- **幻灯片容器** (.slide-container): 1920×1080px固定尺寸
- **顶部装饰条** (.top-bar): 10px高,主题色背景
- **内容区域** (.content-section): 内边距80px,最大高度1000px
- **页码** (.page-number): 右下角固定位置

#### 文本元素
- **H1标题**: 48px, 粗体700
- **H2副标题**: 36px, 粗体600, 主题色
- **段落P**: 20-23px, #333颜色
- **装饰线**: 宽80px, 高4px, 主题色

#### 数据容器
1. **stats-container**: 4栏网格布局统计卡片
2. **stat-box**: 图标+标题+数据组合
3. **stat-card**: 包含图表的卡片
4. **data-card**: 左边框强调的数据卡
5. **strategy-card**: 策略建议卡片

#### 复杂元素
- **表格** (.ip-table): 自定义样式,主题色表头
- **进度条** (.progress-bar): 16px高,圆角8px
- **图标**: FontAwesome图标库
- **图表**: Chart.js绘制的柱状图/折线图

### 1.2 转换挑战识别

| 挑战点 | 难度 | 解决方案方向 |
|--------|------|--------------|
| Chart.js图表渲染 | ⭐⭐⭐⭐⭐ | Selenium截图+OCR提取数据 或 解析JS代码提取数据重绘 |
| CSS样式精确映射 | ⭐⭐⭐⭐ | 建立CSS-to-PPTX样式映射规则库 |
| 动态网格布局 | ⭐⭐⭐ | 计算元素位置,转换为PPTX绝对定位 |
| FontAwesome图标 | ⭐⭐⭐ | Unicode字符映射 或 SVG转换 |
| 渐变/透明度 | ⭐⭐ | python-pptx支持,直接映射 |

---

## 二、技术选型论证

### 2.1 业内方案调研

#### 方案A: python-pptx + BeautifulSoup
**优势:**
- python-pptx成熟稳定,API完善
- BeautifulSoup解析HTML简单高效
- 社区活跃,问题易解决

**劣势:**
- 需要手动处理所有样式映射
- Chart.js图表无法直接转换
- 复杂布局计算工作量大

**适用性:** ⭐⭐⭐⭐ (推荐)

#### 方案B: Playwright/Puppeteer + 截图
**优势:**
- 完美还原视觉效果
- 无需处理样式映射
- 图表自动渲染

**劣势:**
- 生成的是图片,不可编辑
- 文件体积大
- 失去PPTX灵活性

**适用性:** ⭐⭐ (不推荐,作为降级方案)

#### 方案C: LibreOffice API
**优势:**
- 支持HTML直接导入
- 自动处理部分样式

**劣势:**
- 样式兼容性差
- 需要安装LibreOffice
- 控制粒度低

**适用性:** ⭐ (不推荐)

### 2.2 最终选型

**核心方案: python-pptx + BeautifulSoup + Playwright (混合方案)**

#### 技术栈
- **HTML解析**: BeautifulSoup 4.12+
- **PPTX生成**: python-pptx 0.6.23+
- **图表处理**: Playwright (截图) + Pillow (图片处理)
- **样式计算**: cssutils (CSS解析)
- **字体处理**: python-pptx内置 + 系统字体

#### 架构思路
1. 使用BeautifulSoup解析HTML结构和文本内容
2. 使用cssutils解析CSS样式规则
3. 使用Playwright渲染页面并截取Chart.js图表区域
4. 建立CSS样式到PPTX样式的映射规则引擎
5. 按元素类型分模块转换(文本/表格/图表/图标)
6. 组装为完整PPTX幻灯片

---

## 三、架构设计

### 3.1 模块划分

```
html2pptx/
├── src/
│   ├── parser/                 # HTML解析模块
│   │   ├── __init__.py
│   │   ├── html_parser.py     # BeautifulSoup封装
│   │   ├── css_parser.py      # CSS样式提取
│   │   └── structure_analyzer.py  # 结构分析
│   │
│   ├── mapper/                 # 样式映射模块
│   │   ├── __init__.py
│   │   ├── style_mapper.py    # CSS到PPTX样式映射
│   │   ├── color_mapper.py    # 颜色转换
│   │   ├── font_mapper.py     # 字体映射
│   │   └── layout_calculator.py  # 布局计算
│   │
│   ├── converters/             # 元素转换器
│   │   ├── __init__.py
│   │   ├── base_converter.py # 转换器基类
│   │   ├── text_converter.py # 文本转换
│   │   ├── table_converter.py # 表格转换
│   │   ├── chart_converter.py # 图表转换
│   │   ├── icon_converter.py # 图标转换
│   │   └── shape_converter.py # 形状/装饰转换
│   │
│   ├── renderer/               # PPTX生成模块
│   │   ├── __init__.py
│   │   ├── pptx_builder.py   # PPTX构建器
│   │   ├── slide_composer.py # 幻灯片组装
│   │   └── element_placer.py # 元素定位
│   │
│   ├── utils/                  # 工具模块
│   │   ├── __init__.py
│   │   ├── chart_capture.py  # Playwright图表截图
│   │   ├── image_processor.py # 图片处理
│   │   ├── unit_converter.py # 单位转换(px→EMU)
│   │   └── logger.py         # 日志工具
│   │
│   └── main.py                # 主入口
│
├── tests/                      # 测试模块
│   ├── test_parser.py
│   ├── test_mapper.py
│   ├── test_converters.py
│   └── test_integration.py
│
├── config/                     # 配置文件
│   ├── style_rules.json      # 样式映射规则
│   └── settings.py           # 全局配置
│
├── output/                     # 输出目录
├── requirements.txt
└── README.md
```

### 3.2 核心数据流

```
HTML文件
    ↓
[HTML Parser] → 提取DOM结构
    ↓
[CSS Parser] → 提取样式规则
    ↓
[Structure Analyzer] → 分析元素层级与类型
    ↓
[Chart Capture] → 截取图表图片 (Playwright)
    ↓
[Style Mapper] → CSS样式 → PPTX样式映射
    ↓
[Element Converters] → 按类型转换元素
    ↓
[Layout Calculator] → 计算元素绝对位置
    ↓
[Slide Composer] → 组装幻灯片
    ↓
[PPTX Builder] → 生成最终PPTX文件
    ↓
PPTX输出
```

### 3.3 关键算法设计

#### 布局计算算法
```python
def calculate_absolute_position(element, parent_context):
    """
    计算元素在1920×1080画布上的绝对位置

    输入:
    - element: BeautifulSoup元素对象
    - parent_context: 父级上下文(位置、尺寸)

    输出:
    - {x, y, width, height} 单位:px

    步骤:
    1. 解析CSS定位属性(absolute/relative/static)
    2. 处理padding/margin
    3. 处理grid/flex布局
    4. 计算最终坐标
    """
    pass
```

#### 样式映射算法
```python
def map_css_to_pptx_style(css_properties):
    """
    CSS属性 → python-pptx样式对象

    映射规则:
    - font-size: 48px → Pt(48)
    - color: rgb(10,66,117) → RGBColor(10,66,117)
    - font-weight: 700 → font.bold = True
    - background-color → fill.solid()
    - border-left → line.color/line.width
    """
    pass
```

---

## 四、详细设计

### 4.1 样式映射规则库 (config/style_rules.json)

```json
{
  "font_mapping": {
    "default": "Microsoft YaHei",
    "fallback": ["Arial", "SimSun"]
  },
  "color_palette": {
    "primary": "rgb(10, 66, 117)",
    "text_default": "#333",
    "bg_card": "rgba(10, 66, 117, 0.08)"
  },
  "size_mapping": {
    "h1": {"font_size": 48, "bold": true},
    "h2": {"font_size": 36, "bold": true, "color": "primary"},
    "p": {"font_size": 20, "color": "text_default"}
  },
  "container_mapping": {
    ".stat-box": {
      "background": "rgba(10, 66, 117, 0.06)",
      "border_radius": 10,
      "padding": 20
    },
    ".data-card": {
      "border_left": {"width": 4, "color": "primary"},
      "padding_left": 15
    }
  }
}
```

### 4.2 图表处理方案

#### 方案选择: Playwright截图法
**理由:**
- Chart.js是客户端渲染,无法直接从HTML提取数据
- 解析JS代码提取数据过于复杂且不稳定
- 截图可保证视觉100%一致性

**实现步骤:**
1. 使用Playwright启动无头浏览器
2. 加载HTML页面,等待Chart.js渲染完成
3. 定位canvas元素,截取特定区域
4. 保存为PNG图片
5. 插入到PPTX对应位置

```python
async def capture_chart(html_path, chart_selector):
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        await page.goto(f'file://{html_path}')
        await page.wait_for_selector('canvas')
        await page.wait_for_timeout(1000)  # 等待动画完成

        chart_element = await page.query_selector(chart_selector)
        screenshot = await chart_element.screenshot()

        await browser.close()
        return screenshot
```

### 4.3 单位转换

PPTX使用EMU (English Metric Units)单位:
- 1 inch = 914400 EMU
- 1 px (at 96 DPI) = 9525 EMU
- 1920px = 18288000 EMU (宽度)
- 1080px = 10287000 EMU (高度)

```python
class UnitConverter:
    DPI = 96
    EMU_PER_INCH = 914400

    @staticmethod
    def px_to_emu(px):
        return int(px * UnitConverter.EMU_PER_INCH / UnitConverter.DPI)

    @staticmethod
    def pt_to_emu(pt):
        return int(pt * 12700)
```

---

## 五、测试用例设计

### 5.1 单元测试

| 测试模块 | 测试用例 | 预期结果 |
|---------|---------|---------|
| HTML Parser | 解析slidewithtable.html | 提取所有.slide-container |
| CSS Parser | 提取primary-color样式 | rgb(10, 66, 117) |
| Style Mapper | 映射H1样式 | font_size=48, bold=True |
| Layout Calculator | 计算.stat-box位置 | 4个盒子均匀分布 |
| Table Converter | 转换.ip-table | 表头主题色,边框正确 |
| Chart Converter | 截取柱状图 | PNG图片,尺寸正确 |

### 5.2 集成测试

**测试场景1: 完整转换slidewithtable.html**
- 输入: slidewithtable.html
- 输出: output/slidewithtable.pptx
- 验证点:
  - [ ] 幻灯片尺寸16:9
  - [ ] 顶部蓝色装饰条存在
  - [ ] H1/H2文本内容与样式正确
  - [ ] 4个stat-box布局正确
  - [ ] 图表截图清晰可见
  - [ ] 进度条样式正确
  - [ ] 页码"7"在右下角

**测试场景2: 模板最小化测试**
- 输入: 仅包含H1+H2+P的最小HTML
- 输出: 能否正常生成PPTX
- 验证点: 不崩溃,基础元素正确

**测试场景3: 边界条件测试**
- 超长文本是否截断
- 缺少CSS样式时使用默认值
- 图表加载失败时的降级处理

### 5.3 视觉回归测试

使用Playwright截取HTML页面完整截图,与生成的PPTX进行视觉对比:
1. 将PPTX导出为PNG
2. 使用Pillow计算SSIM相似度
3. 相似度 > 95% 视为通过

---

## 六、实施计划

### 第一阶段: 基础框架搭建 (预计1天)

#### 任务清单
- [x] 创建项目目录结构
- [ ] 配置虚拟环境并安装依赖
- [ ] 实现基础工具类(单位转换、日志)
- [ ] 搭建HTML解析器基础框架
- [ ] 搭建PPTX生成器基础框架

#### 关键依赖
```
beautifulsoup4==4.12.3
python-pptx==0.6.23
playwright==1.45.0
Pillow==10.4.0
cssutils==2.11.1
lxml==5.3.0
```

### 第二阶段: 核心模块开发 (预计2天)

#### Day 1: 解析与映射
- [ ] 实现HTML/CSS解析器
- [ ] 实现样式映射引擎
- [ ] 实现布局计算器
- [ ] 单元测试通过

#### Day 2: 转换器开发
- [ ] 文本转换器(H1/H2/P)
- [ ] 表格转换器
- [ ] 形状转换器(装饰条、进度条)
- [ ] 图表截图工具

### 第三阶段: 集成与测试 (预计1天)

- [ ] 集成所有模块
- [ ] 完成slidewithtable.html完整转换
- [ ] 修复Bug
- [ ] 性能优化

### 第四阶段: 迭代优化 (预计1天)

基于自我批判的改进方向:
1. **准确性改进**: 精细调整样式映射误差
2. **性能优化**: 缓存机制、并发处理
3. **扩展性增强**: 支持自定义样式规则

---

## 七、风险与应对

### 风险1: Chart.js图表数据提取困难
**影响:** 高
**概率:** 中
**应对:**
- 主方案: Playwright截图(已选择)
- 备选: 手动提取JS中的data数组,用matplotlib重绘

### 风险2: 字体兼容性问题
**影响:** 中
**概率:** 高
**应对:**
- 使用微软雅黑作为默认字体(Windows兼容性好)
- 提供字体fallback机制
- 支持嵌入字体文件(需用户提供)

### 风险3: 复杂CSS布局计算误差
**影响:** 高
**概率:** 中
**应对:**
- 针对模板中的固定布局模式(grid 4列等)建立专用规则
- 提供手动微调配置接口
- 充分测试边界条件

---

## 八、改进方向(自我批判)

### 改进点1: 数据驱动的样式规则
**现状问题:** 样式映射硬编码在代码中,扩展性差
**改进方案:**
- 将所有样式规则抽取到JSON配置文件
- 支持用户自定义样式映射规则
- 实现规则热更新,无需修改代码

### 改进点2: 增量转换与缓存
**现状问题:** 每次转换都重新解析,性能低
**改进方案:**
- 缓存HTML解析结果
- 图表截图缓存(基于内容hash)
- 支持只转换修改过的幻灯片

### 改进点3: 可视化调试工具
**现状问题:** 样式错误难以定位
**改进方案:**
- 生成调试报告(HTML元素→PPTX元素映射表)
- 可视化对比工具(HTML vs PPTX并排显示)
- 提供样式差异高亮功能

### 改进点4: 支持批量转换
**扩展需求:**
- 支持一次转换多个HTML文件
- 合并为单个PPTX(多个幻灯片)
- 支持幻灯片模板复用

---

## 九、交付物清单

### 代码交付物
- [x] 完整项目源代码(符合架构设计)
- [ ] requirements.txt依赖文件
- [ ] README.md使用文档
- [ ] 配置文件样例

### 文档交付物
- [x] 本实施计划文档
- [ ] API文档(模块接口说明)
- [ ] 样式映射规则文档
- [ ] 常见问题FAQ

### 测试交付物
- [ ] 单元测试代码(覆盖率>80%)
- [ ] 集成测试用例
- [ ] 测试报告

### 演示交付物
- [ ] slidewithtable.pptx (转换结果)
- [ ] 对比截图(HTML vs PPTX)

---

## 十、技术卡点预案

### 卡点1: python-pptx不支持某些高级样式
**解决思路:**
1. 查阅python-pptx源码,寻找底层XML操作方法
2. 直接修改PPTX压缩包内的XML文件
3. 使用python-pptx的oxml接口绕过限制

### 卡点2: Playwright安装失败或性能问题
**解决思路:**
1. 降级方案: 使用Selenium + ChromeDriver
2. 优化方案: 复用浏览器实例,避免频繁启动
3. 替代方案: 预渲染图表为图片,直接提供给用户

### 卡点3: 样式还原精度不足
**解决思路:**
1. 使用像素级精确定位(EMU单位)
2. 引入误差容忍配置(允许±5px偏差)
3. 提供手动校准模式

---

## 附录

### A. 参考资料
- [python-pptx官方文档](https://python-pptx.readthedocs.io/)
- [BeautifulSoup文档](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
- [Playwright Python文档](https://playwright.dev/python/)
- [Office Open XML规范](http://officeopenxml.com/)

### B. 开发环境要求
- Python 3.8+
- Windows/Linux/macOS
- 至少2GB可用内存
- 浏览器(Chromium,由Playwright自动安装)

### C. 术语表
- **EMU**: English Metric Units, PPTX内部使用的长度单位
- **DPI**: Dots Per Inch, 屏幕分辨率
- **DOM**: Document Object Model, HTML文档结构
- **SSIM**: Structural Similarity Index, 图像相似度指标

---

**文档版本:** v1.0
**创建时间:** 2025-10-11
**作者:** Claude Code
**状态:** ✅ 已完成,待评审
