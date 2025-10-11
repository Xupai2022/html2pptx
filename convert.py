"""
HTML转PPTX启动脚本
"""

import sys
from pathlib import Path

# 添加src目录到路径
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from src.main import main

if __name__ == "__main__":
    main()
