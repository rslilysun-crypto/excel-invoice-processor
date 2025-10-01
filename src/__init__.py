# -*- coding: utf-8 -*-
"""
Excel发票数据处理软件
主包初始化文件
"""

__version__ = "1.0.0"
__author__ = "AI Assistant"
__description__ = "Excel发票数据处理软件 - 专业的Excel数据列删除和重排版工具"

# 导入主要模块
from . import core
from . import ui
from . import utils

__all__ = [
    'core',
    'ui', 
    'utils',
    '__version__',
    '__author__',
    '__description__'
]