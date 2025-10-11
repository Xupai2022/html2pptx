"""
配置加载器
支持从JSON文件加载样式规则,实现数据驱动的样式映射
"""

import json
from pathlib import Path
from typing import Dict, Any, Optional

from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ConfigLoader:
    """配置加载器"""

    _instance = None
    _config = None

    def __new__(cls):
        """单例模式"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        """初始化配置加载器"""
        if self._config is None:
            self.load_config()

    def load_config(self, config_path: str = None):
        """
        加载配置文件

        Args:
            config_path: 配置文件路径,默认为config/style_rules.json
        """
        if config_path is None:
            # 默认配置路径
            base_path = Path(__file__).parent.parent.parent
            config_path = base_path / 'config' / 'style_rules.json'

        config_path = Path(config_path)

        if not config_path.exists():
            logger.warning(f"配置文件不存在: {config_path},使用默认配置")
            self._config = self._get_default_config()
            return

        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                self._config = json.load(f)
            logger.info(f"成功加载配置: {config_path}")
        except Exception as e:
            logger.error(f"加载配置失败: {e},使用默认配置")
            self._config = self._get_default_config()

    def _get_default_config(self) -> Dict[str, Any]:
        """获取默认配置"""
        return {
            "font_mapping": {
                "default": "Microsoft YaHei",
                "fallback": ["Arial"]
            },
            "color_palette": {
                "primary": "rgb(10, 66, 117)",
                "text_default": "#333"
            },
            "layout": {
                "slide_width": 1920,
                "slide_height": 1080
            }
        }

    def get(self, key_path: str, default: Any = None) -> Any:
        """
        获取配置值(支持点号分隔的路径)

        Args:
            key_path: 配置键路径,如'color_palette.primary'
            default: 默认值

        Returns:
            配置值
        """
        keys = key_path.split('.')
        value = self._config

        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default

        return value

    def get_color(self, color_key: str) -> Optional[str]:
        """
        获取颜色值

        Args:
            color_key: 颜色键名

        Returns:
            颜色字符串
        """
        return self.get(f'color_palette.{color_key}')

    def get_font(self, font_type: str = 'default') -> str:
        """
        获取字体名称

        Args:
            font_type: 字体类型

        Returns:
            字体名称
        """
        return self.get(f'font_mapping.{font_type}', 'Microsoft YaHei')

    def get_layout(self, layout_key: str) -> Any:
        """
        获取布局配置

        Args:
            layout_key: 布局键名

        Returns:
            布局值
        """
        return self.get(f'layout.{layout_key}')

    def reload(self, config_path: str = None):
        """
        重新加载配置(支持热更新)

        Args:
            config_path: 配置文件路径
        """
        logger.info("重新加载配置...")
        self.load_config(config_path)


# 全局配置实例
config = ConfigLoader()
