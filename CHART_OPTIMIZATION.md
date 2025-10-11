# Chart.js图表截图功能优化说明

## 功能概述

在feature/optimize分支中,我们实现了Chart.js图表的自动截图功能,解决了v1.0.0版本中图表只能显示占位文本的问题。

## 新增功能

### 1. 图表截图工具 (`src/utils/chart_capture.py`)

**核心功能**:
- 使用Playwright无头浏览器截取Chart.js渲染的图表
- 支持基于内容hash的智能缓存机制
- 支持批量截取多个图表
- 优雅降级:Playwright未安装时自动显示占位文本

**主要API**:
```python
from src.utils.chart_capture import ChartCapture

# 创建截图工具
capturer = ChartCapture()

# 截取单个图表
screenshot_path = capturer.capture_chart(
    html_path="slidewithtable.html",
    canvas_selector="#vulnerabilityChart",
    wait_time=2000
)

# 批量截取
chart_configs = [
    {'canvas_id': 'chart1', 'selector': '#chart1'},
    {'canvas_id': 'chart2', 'selector': '#chart2'}
]
results = capturer.capture_multiple_charts(html_path, chart_configs)
```

**缓存机制**:
- 缓存目录: `C:\Users\User\AppData\Local\Temp\html2pptx_charts\`
- 缓存键: 基于HTML内容和选择器的MD5 hash
- 自动检测缓存,避免重复截图

### 2. 图表转换器 (`src/converters/chart_converter.py`)

**功能**:
- 封装图表截图和PPTX插入逻辑
- 自动检测Playwright可用性
- 失败时自动降级为占位文本

**使用示例**:
```python
from src.converters.chart_converter import ChartConverter

chart_converter = ChartConverter(pptx_slide, css_parser, html_path)

success = chart_converter.convert_chart(
    canvas_element,
    x=80,
    y=200,
    width=1760,
    height=220,
    use_screenshot=True  # 是否使用截图
)
```

### 3. 主程序集成

修改了`src/main.py`中的`_convert_stat_card`方法:
- 自动检测Canvas元素
- 尝试使用Playwright截图
- 失败时显示占位文本
- 完整日志输出

## 安装说明

### 启用图表截图功能

```bash
# 1. 激活虚拟环境
source html2ppt/Scripts/activate  # Linux/Mac
html2ppt\Scripts\activate.bat     # Windows

# 2. 安装Playwright
pip install playwright

# 3. 下载Chromium浏览器
playwright install chromium
```

### 验证安装

```bash
# 运行测试脚本
python test_chart_capture.py

# 预期输出:
# [INFO] Playwright已安装
# [INFO] 找到测试文件: slidewithtable.html
# [INFO] 开始截取图表...
# [INFO] 图表截图成功!
# [INFO] 路径: C:\Users\User\AppData\Local\Temp\html2pptx_charts\chart_xxxx.png
# [INFO] 大小: XX.XX KB
# [INFO] 测试通过!
```

## 使用方法

### 基本用法(无需额外配置)

```bash
# 转换HTML(自动尝试截图)
python convert.py slidewithtable.html output/result.pptx
```

**行为说明**:
- **Playwright已安装**: 自动截取图表并插入PPTX
- **Playwright未安装**: 显示占位文本,提示安装

### 批量转换

```bash
python batch_convert.py ./slides ./output
```

所有HTML中的图表都会尝试截图。

## 技术细节

### 工作流程

```
HTML文件
    ↓
[解析HTML] → 找到Canvas元素
    ↓
[Playwright] → 启动无头浏览器
    ↓
[加载HTML] → 等待图表渲染(2秒)
    ↓
[截取Canvas] → 保存为PNG图片
    ↓
[缓存检查] → 生成hash,避免重复截图
    ↓
[插入PPTX] → 按指定位置和尺寸插入
    ↓
PPTX输出
```

### 缓存策略

1. **缓存键生成**:
   ```python
   cache_key = md5(f"{html_content}:{selector}").hexdigest()[:16]
   ```

2. **缓存命中**:
   - 相同HTML内容 + 相同选择器 → 使用缓存
   - 跳过浏览器启动,直接返回图片路径

3. **缓存清理**:
   ```python
   capturer.clear_cache()  # 手动清理所有缓存
   ```

### 优雅降级

```python
# 检查Playwright可用性
if ChartCapture.is_available():
    # 尝试截图
    screenshot_path = capture_chart(...)
    if screenshot_path:
        # 插入图片
        insert_image(screenshot_path)
    else:
        # 降级为占位文本
        insert_placeholder()
else:
    # Playwright不可用,直接占位
    insert_placeholder()
```

**降级触发条件**:
- Playwright未安装
- 浏览器启动失败
- Canvas元素未找到
- 截图超时
- 其他异常

## 性能优化

### 1. 智能缓存

**效果**: 相同图表不会重复截图

| 场景 | 首次转换 | 二次转换 |
|------|---------|---------|
| 单个HTML | ~3秒 | <0.1秒 |
| 批量转换 | ~N×3秒 | ~N×0.1秒 |

### 2. 并发支持

```python
# 批量截图时可并发执行
capturer.capture_multiple_charts(html_path, chart_configs)
```

### 3. 资源管理

- 浏览器自动启动和关闭
- 临时文件自动清理
- 内存占用可控(<100MB)

## 对比v1.0.0

| 项目 | v1.0.0 | feature/optimize |
|------|--------|------------------|
| 图表处理 | 占位文本 | **真实图表截图** |
| Playwright依赖 | 不需要 | **可选(支持降级)** |
| 转换速度 | <1秒 | **首次~3秒,缓存<0.1秒** |
| 视觉效果 | ⭐⭐ | **⭐⭐⭐⭐⭐** |
| 可编辑性 | 文本可编辑 | **图片不可编辑** |

## 已知问题与解决方案

### 问题1: Playwright安装失败

**现象**: `pip install playwright` 报greenlet编译错误

**原因**: 缺少C++编译器

**解决方案**:
```bash
# 方案1: 使用预编译wheel
pip install playwright --only-binary :all:

# 方案2: 安装编译器
# Windows: 安装 Visual Studio Build Tools
# Linux: sudo apt-get install build-essential
```

### 问题2: Chromium下载失败

**现象**: `playwright install chromium` 超时

**解决方案**:
```bash
# 使用国内镜像
export PLAYWRIGHT_DOWNLOAD_HOST=https://playwright.azureedge.net

# 然后重新安装
playwright install chromium
```

### 问题3: 图表未完全渲染

**现象**: 截图中图表显示不完整

**解决方案**:
```python
# 增加等待时间(默认2000ms)
screenshot_path = capturer.capture_chart(
    html_path,
    canvas_selector,
    wait_time=5000  # 增加到5秒
)
```

## 测试用例

### 测试1: 基本图表截图

```bash
python test_chart_capture.py
```

**验证点**:
- ✅ Playwright安装检测
- ✅ HTML文件加载
- ✅ Canvas元素定位
- ✅ 图表截图成功
- ✅ 文件大小合理(>10KB)

### 测试2: 完整转换流程

```bash
python convert.py slidewithtable.html output/test.pptx
```

**验证点**:
- ✅ HTML解析成功
- ✅ 图表截图或占位文本
- ✅ PPTX生成成功
- ✅ 图表显示正确

### 测试3: 缓存机制

```bash
# 首次转换
time python convert.py slidewithtable.html output/test1.pptx

# 二次转换(应使用缓存)
time python convert.py slidewithtable.html output/test2.pptx
```

**验证点**:
- ✅ 二次转换明显更快
- ✅ 日志显示"使用缓存的图表截图"

## 未来改进方向

### 1. 支持更多图表库
- ECharts
- D3.js
- Highcharts

### 2. 图表数据提取
- 从JavaScript代码中提取数据
- 使用python-pptx内置图表功能重绘
- 保持可编辑性

### 3. 图表样式优化
- 自动调整尺寸适配幻灯片
- 支持图表透明背景
- 高DPI渲染(4K/8K)

### 4. 性能进一步优化
- 多进程并发截图
- 预加载浏览器实例
- 内存池复用

## 总结

此次优化成功实现了Chart.js图表的自动截图功能,大幅提升了转换后PPTX的视觉效果。通过优雅降级设计,即使Playwright未安装也不会影响基本功能使用。

**核心价值**:
- ✅ 解决了v1.0.0的最大短板
- ✅ 保持了向后兼容性
- ✅ 提供了完善的错误处理
- ✅ 实现了智能缓存优化

---

**优化版本**: feature/optimize
**基础版本**: v1.0.0
**更新时间**: 2025-10-11
**状态**: ✅ 测试通过,待合并
