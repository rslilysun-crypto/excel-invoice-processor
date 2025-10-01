# -*- coding: utf-8 -*-
"""
Excel文件读取器
负责Excel文件的读取、工作表管理和数据提取
"""

import pandas as pd
import os
from typing import Optional, Dict, List, Any
from src.utils.logger import get_logger

logger = get_logger("ExcelReader")

class ExcelReader:
    """
    Excel文件读取器类
    提供Excel文件读取、工作表选择、数据提取等功能
    """
    
    def __init__(self):
        self.file_path = None
        self.excel_file = None
        self.worksheets_info = {}
        self.current_worksheet = None
        
    def load_file(self, file_path: str) -> bool:
        """
        加载Excel文件
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            bool: 加载是否成功
        """
        try:
            if not os.path.exists(file_path):
                logger.error(f"文件不存在: {file_path}")
                return False
                
            if not file_path.lower().endswith(('.xlsx', '.xls')):
                logger.error(f"不支持的文件格式: {file_path}")
                return False
                
            # 读取Excel文件
            self.excel_file = pd.ExcelFile(file_path)
            self.file_path = file_path
            
            # 获取工作表信息
            self._load_worksheets_info()
            
            logger.info(f"Excel文件加载成功: {file_path}")
            return True
            
        except Exception as e:
            logger.error(f"加载Excel文件失败: {e}")
            return False
    
    def _load_worksheets_info(self):
        """
        加载工作表信息
        """
        try:
            self.worksheets_info = {}
            
            for sheet_name in self.excel_file.sheet_names:
                try:
                    # 读取工作表基本信息
                    df = pd.read_excel(self.excel_file, sheet_name=sheet_name, nrows=0)
                    
                    # 获取实际数据行数
                    full_df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                    
                    self.worksheets_info[sheet_name] = {
                        'max_row': len(full_df),
                        'max_column': len(full_df.columns),
                        'columns': list(full_df.columns),
                        'has_data': len(full_df) > 0
                    }
                    
                except Exception as e:
                    logger.warning(f"读取工作表 {sheet_name} 信息失败: {e}")
                    self.worksheets_info[sheet_name] = {
                        'max_row': 0,
                        'max_column': 0,
                        'columns': [],
                        'has_data': False
                    }
                    
        except Exception as e:
            logger.error(f"加载工作表信息失败: {e}")
    
    def get_worksheets_list(self) -> List[Dict[str, Any]]:
        """
        获取工作表列表
        
        Returns:
            List[Dict]: 工作表信息列表
        """
        worksheets = []
        
        for name, info in self.worksheets_info.items():
            worksheets.append({
                'name': name,
                'rows': info['max_row'],
                'columns': info['max_column'],
                'has_data': info['has_data']
            })
            
        return worksheets
    
    def get_target_worksheet(self) -> Optional[str]:
        """
        获取目标工作表名称（智能识别发票基础信息工作表）
        
        Returns:
            Optional[str]: 目标工作表名称
        """
        # 优先查找包含"发票基础信息"的工作表
        for sheet_name in self.worksheets_info.keys():
            if "发票基础信息" in sheet_name:
                return sheet_name
        
        # 查找包含"发票信息"的工作表
        for sheet_name in self.worksheets_info.keys():
            if "发票信息" in sheet_name:
                return sheet_name
        
        # 查找包含"基础信息"的工作表
        for sheet_name in self.worksheets_info.keys():
            if "基础信息" in sheet_name:
                return sheet_name
        
        # 返回第一个有数据的工作表
        for sheet_name, info in self.worksheets_info.items():
            if info['has_data']:
                return sheet_name
        
        # 返回第一个工作表
        if self.worksheets_info:
            return list(self.worksheets_info.keys())[0]
        
        return None
    
    def select_worksheet(self, worksheet_name: str) -> bool:
        """
        选择工作表
        
        Args:
            worksheet_name: 工作表名称
            
        Returns:
            bool: 选择是否成功
        """
        try:
            if worksheet_name not in self.worksheets_info:
                logger.error(f"工作表不存在: {worksheet_name}")
                return False
                
            self.current_worksheet = worksheet_name
            logger.info(f"选择工作表: {worksheet_name}")
            return True
            
        except Exception as e:
            logger.error(f"选择工作表失败: {e}")
            return False
    
    def read_headers(self, worksheet_name: str) -> Optional[List[str]]:
        """
        读取工作表表头
        
        Args:
            worksheet_name: 工作表名称
            
        Returns:
            Optional[List[str]]: 表头列表
        """
        try:
            if worksheet_name not in self.worksheets_info:
                logger.error(f"工作表不存在: {worksheet_name}")
                return None
                
            df = pd.read_excel(self.excel_file, sheet_name=worksheet_name, nrows=0)
            headers = list(df.columns)
            
            logger.info(f"读取表头成功，共 {len(headers)} 列")
            return headers
            
        except Exception as e:
            logger.error(f"读取表头失败: {e}")
            return None
    
    def read_data_preview(self, worksheet_name: str, preview_rows: int = 10) -> pd.DataFrame:
        """
        读取数据预览
        
        Args:
            worksheet_name: 工作表名称
            preview_rows: 预览行数
            
        Returns:
            pd.DataFrame: 预览数据
        """
        try:
            if worksheet_name not in self.worksheets_info:
                logger.error(f"工作表不存在: {worksheet_name}")
                return pd.DataFrame()
                
            df = pd.read_excel(self.excel_file, sheet_name=worksheet_name, nrows=preview_rows)
            logger.info(f"读取预览数据成功，{len(df)} 行 x {len(df.columns)} 列")
            return df
            
        except Exception as e:
            logger.error(f"读取预览数据失败: {e}")
            return pd.DataFrame()
    
    def read_full_data(self, worksheet_name: str) -> pd.DataFrame:
        """
        读取完整数据
        
        Args:
            worksheet_name: 工作表名称
            
        Returns:
            pd.DataFrame: 完整数据
        """
        try:
            if worksheet_name not in self.worksheets_info:
                logger.error(f"工作表不存在: {worksheet_name}")
                return pd.DataFrame()
                
            df = pd.read_excel(self.excel_file, sheet_name=worksheet_name)
            logger.info(f"读取完整数据成功，{len(df)} 行 x {len(df.columns)} 列")
            return df
            
        except Exception as e:
            logger.error(f"读取完整数据失败: {e}")
            return pd.DataFrame()
    
    def get_file_info(self) -> Optional[Dict[str, Any]]:
        """
        获取文件信息
        
        Returns:
            Optional[Dict]: 文件信息
        """
        try:
            if not self.file_path:
                return None
                
            file_stat = os.stat(self.file_path)
            file_size_mb = round(file_stat.st_size / (1024 * 1024), 2)
            
            return {
                'file_path': self.file_path,
                'file_name': os.path.basename(self.file_path),
                'file_size_mb': file_size_mb,
                'worksheet_count': len(self.worksheets_info),
                'worksheets': list(self.worksheets_info.keys())
            }
            
        except Exception as e:
            logger.error(f"获取文件信息失败: {e}")
            return None
    
    def get_all_worksheets_data(self) -> Optional[Dict[str, pd.DataFrame]]:
        """
        获取所有工作表的数据
        
        Returns:
            Optional[Dict[str, pd.DataFrame]]: 所有工作表数据的字典
        """
        try:
            if not self.excel_file or not self.worksheets_info:
                logger.error("Excel文件未加载")
                return None
                
            all_sheets_data = {}
            for sheet_name in self.worksheets_info.keys():
                try:
                    df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                    all_sheets_data[sheet_name] = df
                    logger.debug(f"读取工作表 {sheet_name}: {len(df)} 行 x {len(df.columns)} 列")
                except Exception as e:
                    logger.warning(f"读取工作表 {sheet_name} 失败: {e}")
                    continue
                    
            logger.info(f"成功读取 {len(all_sheets_data)} 个工作表的数据")
            return all_sheets_data
            
        except Exception as e:
            logger.error(f"获取所有工作表数据失败: {e}")
            return None
    
    def close(self):
        """
        关闭Excel文件
        """
        try:
            if self.excel_file:
                self.excel_file.close()
                
            self.file_path = None
            self.excel_file = None
            self.worksheets_info = {}
            self.current_worksheet = None
            
            logger.info("Excel文件已关闭")
            
        except Exception as e:
            logger.error(f"关闭Excel文件失败: {e}")