# -*- coding: utf-8 -*-
"""
配置管理工具模块
处理用户设置、模板保存和加载等功能
"""

import json
import os
import sys
from typing import Dict, List, Any
from src.utils.logger import get_logger

logger = get_logger("Config")

class ConfigManager:
    """
    配置管理器
    负责用户设置和模板的保存、加载
    """
    
    def __init__(self):
        # 在EXE环境下，使用用户目录保存配置，避免临时目录问题
        if getattr(sys, 'frozen', False):
            # 运行在PyInstaller打包的EXE中
            user_dir = os.path.expanduser("~")
            self.config_dir = os.path.join(user_dir, "AppData", "Local", "Excel发票数据处理软件")
        else:
            # 开发环境，使用项目目录
            self.config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config")
        
        self.templates_file = os.path.join(self.config_dir, "templates.json")
        self.settings_file = os.path.join(self.config_dir, "settings.json")
        
        # 确保配置目录存在
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
            logger.info(f"创建配置目录: {self.config_dir}")
        
        # 初始化默认配置
        self._init_default_config()
    
    def _init_default_config(self):
        """
        初始化默认配置
        """
        # 默认模板配置
        default_templates = {
            "发票数据标准模板": {
                "description": "发票数据处理的标准模板，删除常见的冗余列",
                "columns_to_delete": [
                    "发票代码",
                    "发票号码", 
                    "销方识别号",
                    "销方名称",
                    "购方识别号",
                    "发票来源",
                    "是否正数",
                    "发票风险等级",
                    "开票人"
                ],
                "created_time": "2024-01-01 00:00:00",
                "is_default": True
            }
        }
        
        # 默认设置
        default_settings = {
            "last_input_directory": "",
            "last_output_directory": "",
            "default_output_format": ".xlsx",
            "auto_open_result": True,
            "create_backup": True,
            "default_template": "发票数据标准模板",
            "ui_theme": "default",
            "language": "zh_CN"
        }
        
        # 如果模板文件不存在，创建默认模板
        if not os.path.exists(self.templates_file):
            self.save_templates(default_templates)
        
        # 如果设置文件不存在，创建默认设置
        if not os.path.exists(self.settings_file):
            self.save_settings(default_settings)
    
    def load_templates(self) -> Dict[str, Any]:
        """
        加载用户模板
        
        Returns:
            Dict[str, Any]: 模板字典
        """
        try:
            with open(self.templates_file, 'r', encoding='utf-8') as f:
                templates = json.load(f)
            logger.info(f"成功加载 {len(templates)} 个模板")
            return templates
        except Exception as e:
            logger.error(f"加载模板失败: {e}")
            return {}
    
    def save_templates(self, templates: Dict[str, Any]):
        """
        保存用户模板
        
        Args:
            templates: 模板字典
        """
        try:
            with open(self.templates_file, 'w', encoding='utf-8') as f:
                json.dump(templates, f, ensure_ascii=False, indent=2)
            logger.info(f"成功保存 {len(templates)} 个模板")
        except Exception as e:
            logger.error(f"保存模板失败: {e}")
    
    def add_template(self, name: str, columns_to_delete: List[str], description: str = ""):
        """
        添加新模板
        
        Args:
            name: 模板名称
            columns_to_delete: 要删除的列名列表
            description: 模板描述
        """
        templates = self.load_templates()
        
        from datetime import datetime
        templates[name] = {
            "description": description,
            "columns_to_delete": columns_to_delete,
            "created_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "is_default": False
        }
        
        self.save_templates(templates)
        logger.info(f"添加新模板: {name}")
    
    def delete_template(self, name: str):
        """
        删除模板
        
        Args:
            name: 模板名称
        """
        templates = self.load_templates()
        
        if name in templates:
            # 不允许删除默认模板
            if templates[name].get("is_default", False):
                logger.warning(f"不能删除默认模板: {name}")
                return False
            
            del templates[name]
            self.save_templates(templates)
            logger.info(f"删除模板: {name}")
            return True
        else:
            logger.warning(f"模板不存在: {name}")
            return False
    
    def load_settings(self) -> Dict[str, Any]:
        """
        加载用户设置
        
        Returns:
            Dict[str, Any]: 设置字典
        """
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            logger.info("成功加载用户设置")
            return settings
        except Exception as e:
            logger.error(f"加载设置失败: {e}")
            return {}
    
    def save_settings(self, settings: Dict[str, Any]):
        """
        保存用户设置
        
        Args:
            settings: 设置字典
        """
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            logger.info("成功保存用户设置")
        except Exception as e:
            logger.error(f"保存设置失败: {e}")
    
    def update_setting(self, key: str, value: Any):
        """
        更新单个设置项
        
        Args:
            key: 设置键
            value: 设置值
        """
        settings = self.load_settings()
        settings[key] = value
        self.save_settings(settings)
        logger.info(f"更新设置: {key} = {value}")
    
    def get_setting(self, key: str, default=None):
        """
        获取设置值
        
        Args:
            key: 设置键
            default: 默认值
        
        Returns:
            设置值或默认值
        """
        settings = self.load_settings()
        return settings.get(key, default)