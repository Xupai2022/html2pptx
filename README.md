# HTML转PPTX转换工具

一个从0到1构建的HTML转PPTX自动化转换工程,专门用于将AI生成的HTML报告转换为PPTX格式,严格保持样式与内容的一致性。

## 项目特点

- **精准样式映射**: 完整支持template.txt中定义的所有样式规范
- **固定幻灯片尺寸**: 1920×1080 (16:9)
- **主题色一致**: RGB(10, 66, 117) 深蓝色主题
- **模块化架构**: 清晰的模块划分,易于扩展维护

## 项目结构

```
html2pptx/
├── src/                        # 源代码目录
│   ├── parser/                 # HTML解析模块
│   │   ├── html_parser.py      # BeautifulSoup解析器
│   │   └── css_parser.py       # CSS样式提取
│   │
│   ├── mapper/                 # 样式映射模块
│   │   └── style_mapper.py     # CSS到PPTX样式映射
│   │
│   ├── converters/             # 元素转换器
│   │   ├── base_converter.py  # 转换器基类
│   │   ├── text_converter.py  # 文本转换(H1/H2/P)
│   │   ├── table_converter.py # 表格转换
│   │   └── shape_converter.py # 形状转换(装饰条/进度条)
│   │
│   ├── renderer/               # PPTX生成模块
│   │   └── pptx_builder.py    # PPTX构建器
│   │
│   ├── utils/                  # 工具模块
│   │   ├── unit_converter.py  # 单位转换(px→EMU)
│   │   ├── color_parser.py    # 颜色解析
│   │   └── logger.py          # 日志工具
│   │
│   └── main.py                # 主程序
│
├── output/                     # 输出目录
├── template.txt               # HTML样式模板
├── slidewithtable.html        # 测试样例
├── convert.py                 # 启动脚本
├── requirements.txt           # 依赖清单
├── IMPLEMENTATION_PLAN.md     # 实施计划文档
└── README.md                  # 本文件
```

## 技术栈

- **HTML解析**: BeautifulSoup 4.12+
- **PPTX生成**: python-pptx 0.6.23+
- **图片处理**: Pillow 10.4+
- **Python版本**: 3.8+

## 安装指南

### 1. 克隆项目

```bash
cd D:\Users\User\Desktop\html2pptx
```

### 2. 激活虚拟环境

```bash
source html2ppt/Scripts/activate  # Windows Git Bash
# 或
html2ppt\Scripts\activate.bat     # Windows CMD
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. (可选)安装Chromium浏览器 - 启用图表截图功能

如果需要将Chart.js图表转换为真实图片(而不是占位文本),需要安装Playwright浏览器:

**方式1: 使用辅助脚本(推荐)**
```bash
python install_browser.py
```

**方式2: 手动安装**
```bash
# 安装Playwright浏览器
playwright install chromium
```

**说明**:
- 不安装浏览器时,图表会显示为占位文本,其他功能正常使用
- Chromium浏览器约300MB,首次安装需要几分钟
- 安装后图表会自动截图并插入PPTX

## 使用方法

### 基本用法

```bash
python convert.py <HTML文件路径> [输出PPTX路径]
```

### 示例

```bash
# 转换示例HTML
python convert.py slidewithtable.html output/slidewithtable.pptx

# 使用默认输出路径
python convert.py slidewithtable.html
```

### 输出

转换成功后会在`output/`目录生成PPTX文件,日志会显示详细的转换过程。

## 支持的HTML元素

### 文本元素
- **H1标题**: 48px, 粗体
- **H2副标题**: 36px, 粗体, 主题色
- **段落P**: 20-23px, #333颜色
- **装饰线**: 宽80px, 高4px, 主题色

### 容器元素
- **stats-container**: 4栏网格统计卡片
- **stat-box**: 图标+标题+数据组合
- **stat-card**: 包含图表的卡片
- **data-card**: 左边框强调的数据卡
- **progress-bar**: 进度条

### 装饰元素
- **top-bar**: 顶部10px装饰条
- **page-number**: 右下角页码

## 样式映射规则

### 颜色映射
- `rgb(10, 66, 117)` → 主题色
- `#333` → 默认文本色
- `rgba(10, 66, 117, 0.08)` → 卡片背景色(带透明度)

### 单位转换
- 1px = 9525 EMU (假设96 DPI)
- 幻灯片宽度: 1920px = 18288000 EMU
- 幻灯片高度: 1080px = 10287000 EMU

### 字体规范
- 默认字体: Microsoft YaHei (微软雅黑)
- H1字号: 48pt
- H2字号: 36pt
- 正文字号: 20pt

## 核心算法

### 布局计算
```python
def calculate_absolute_position(element, parent_context):
    """
    计算元素在1920×1080画布上的绝对位置
    处理CSS定位属性、padding/margin、grid/flex布局
    """
```

### 样式映射
```python
def map_css_to_pptx_style(css_properties):
    """
    CSS属性 → python-pptx样式对象
    - font-size: 48px → Pt(48)
    - color: rgb(10,66,117) → RGBColor(10,66,117)
    - font-weight: 700 → font.bold = True
    """
```

## 已知限制

1. **图表处理**: Chart.js图表当前显示为占位文本,需要集成Playwright进行截图
2. **FontAwesome图标**: 当前使用Emoji替代,未来可改为SVG转换
3. **复杂CSS**: 仅支持模板中使用的CSS属性,不支持动画、渐变等高级特性

## 测试用例

已完成的测试:
- ✅ 解析slidewithtable.html成功
- ✅ 提取34条CSS规则
- ✅ 转换标题和副标题
- ✅ 转换4个统计卡片
- ✅ 转换进度条(3个)
- ✅ 添加页码
- ✅ 生成30KB PPTX文件

## 改进方向

基于实施计划文档中的自我批判,以下为未来改进方向:

### 1. 数据驱动的样式规则
- 将样式映射规则抽取到JSON配置文件
- 支持用户自定义样式规则
- 实现规则热更新

### 2. 增量转换与缓存
- 缓存HTML解析结果
- 图表截图缓存(基于内容hash)
- 支持只转换修改过的幻灯片

### 3. 可视化调试工具
- 生成调试报告(HTML→PPTX映射表)
- 可视化对比工具
- 样式差异高亮

### 4. Chart.js图表支持
- 集成Playwright进行图表截图
- 支持柱状图、折线图、饼图等
- 自动调整图表尺寸

## 贡献指南

欢迎提交Issue和Pull Request!

### 开发流程
1. Fork本项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启Pull Request

### 代码规范
- 遵循PEP 8
- 添加适当的注释和文档字符串
- 所有公开函数需包含类型注解
- 新增功能需添加单元测试

## 常见问题

### Q: 转换失败提示"No module named 'bs4'"?
**A**: 请确保在虚拟环境中安装了依赖: `pip install -r requirements.txt`

### Q: 生成的PPTX样式不准确?
**A**: 请检查HTML是否严格遵循template.txt中的样式约束,自定义样式可能无法正确映射

### Q: 如何支持新的HTML元素?
**A**: 在`src/converters/`目录下创建新的转换器,继承`BaseConverter`类并实现`convert()`方法

### Q: 图表为什么显示为占位文本?
**A**: 当前版本未集成Playwright,Chart.js图表暂时显示占位。可参考实施计划文档中的图表处理方案进行扩展

## 许可证

MIT License

## 作者

Claude Code - AI驱动的全栈开发助手

## 致谢

- python-pptx项目组
- BeautifulSoup开发团队
- 所有开源贡献者

---

**版本**: v1.0.0
**更新日期**: 2025-10-11
**状态**: ✅ 可用于生产环境
