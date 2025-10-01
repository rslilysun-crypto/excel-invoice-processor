# -*- coding: utf-8 -*-
"""
数据处理模块
负责数据的列删除、预览生成和处理逻辑
"""

import pandas as pd
from typing import List, Dict, Any, Optional, Tuple
import copy
from src.utils.logger import get_logger

logger = get_logger("DataProcessor")

class DataProcessor:
    """
    数据处理器
    提供列删除、数据预览等功能
    """
    
    def __init__(self):
        self.original_data = None
        self.processed_data = None
        self.columns_to_delete = []
        self.columns_to_recalculate = []  # 需要重新计算合计的列
        self.processing_history = []
        self.cross_sheet_data = {}  # 存储跨工作表关联数据
    
    def load_data(self, data: pd.DataFrame) -> bool:
        """
        加载原始数据
        
        Args:
            data: 原始数据DataFrame
        
        Returns:
            bool: 加载是否成功
        """
        try:
            if data.empty:
                logger.error("数据为空")
                return False
            
            self.original_data = data.copy()
            self.processed_data = data.copy()
            
            logger.info(f"加载数据成功: {len(data)} 行 x {len(data.columns)} 列")
            return True
            
        except Exception as e:
            logger.error(f"加载数据失败: {e}")
            return False
    
    def set_columns_to_delete(self, columns: List[str]) -> bool:
        """
        设置要删除的列
        
        Args:
            columns: 要删除的列名列表
        
        Returns:
            bool: 设置是否成功
        """
        try:
            if not self.original_data is not None:
                logger.error("没有加载数据")
                return False
            
            # 验证列名是否存在
            available_columns = self.original_data.columns.tolist()
            invalid_columns = [col for col in columns if col not in available_columns]
            
            if invalid_columns:
                logger.warning(f"以下列名不存在: {invalid_columns}")
            
            # 只保留存在的列名
            valid_columns = [col for col in columns if col in available_columns]
            self.columns_to_delete = valid_columns
            
            logger.info(f"设置要删除的列: {len(valid_columns)} 个")
            return True
            
        except Exception as e:
            logger.error(f"设置删除列失败: {e}")
            return False
    
    def validate_deletion(self) -> Tuple[bool, str]:
        """
        验证删除操作的有效性
        
        Returns:
            Tuple[bool, str]: (是否有效, 错误信息)
        """
        try:
            if self.original_data is None:
                return False, "没有加载数据"
            
            # 如果没有要删除的列，也是有效的（表示不删除任何列）
            if not self.columns_to_delete:
                return True, "无需删除列"
            
            # 检查是否会删除所有列
            remaining_columns = [col for col in self.original_data.columns if col not in self.columns_to_delete]
            
            if not remaining_columns:
                return False, "不能删除所有列，请至少保留一列数据"
            
            # 检查删除的列是否存在
            available_columns = self.original_data.columns.tolist()
            invalid_columns = [col for col in self.columns_to_delete if col not in available_columns]
            
            if invalid_columns:
                return False, f"以下列不存在: {', '.join(invalid_columns)}"
            
            return True, "验证通过"
            
        except Exception as e:
            logger.error(f"验证删除操作失败: {e}")
            return False, f"验证失败: {str(e)}"
    
    def generate_preview(self, preview_rows: int = 10) -> Dict[str, pd.DataFrame]:
        """
        生成删除前后的数据预览
        
        Args:
            preview_rows: 预览行数
        
        Returns:
            Dict[str, pd.DataFrame]: 包含'before'和'after'的预览数据
        """
        try:
            if self.original_data is None:
                logger.error("没有加载数据")
                return {'before': pd.DataFrame(), 'after': pd.DataFrame()}
            
            # 删除前的预览
            before_preview = self.original_data.head(preview_rows).copy()
            
            # 删除后的预览
            remaining_columns = [col for col in self.original_data.columns if col not in self.columns_to_delete]
            after_preview = self.original_data[remaining_columns].head(preview_rows).copy()
            
            logger.info(f"生成预览数据: 删除前 {len(before_preview.columns)} 列, 删除后 {len(after_preview.columns)} 列")
            
            return {
                'before': before_preview,
                'after': after_preview
            }
            
        except Exception as e:
            logger.error(f"生成预览失败: {e}")
            return {'before': pd.DataFrame(), 'after': pd.DataFrame()}
    
    def process_data(self) -> bool:
        """
        执行数据处理（删除指定列）
        
        Returns:
            bool: 处理是否成功
        """
        try:
            if self.original_data is None:
                logger.error("没有加载数据")
                return False
            
            # 如果有要删除的列，验证删除操作
            if self.columns_to_delete:
                is_valid, error_msg = self.validate_deletion()
                if not is_valid:
                    logger.error(f"删除操作验证失败: {error_msg}")
                    return False
                
                # 执行删除操作
                remaining_columns = [col for col in self.original_data.columns if col not in self.columns_to_delete]
                self.processed_data = self.original_data[remaining_columns].copy()
                
                # 记录处理历史
                processing_record = {
                    'action': 'delete_columns',
                    'deleted_columns': self.columns_to_delete.copy(),
                    'original_columns_count': len(self.original_data.columns),
                    'remaining_columns_count': len(self.processed_data.columns),
                    'data_rows': len(self.processed_data)
                }
                self.processing_history.append(processing_record)
                
                logger.info(f"数据处理完成: 删除了 {len(self.columns_to_delete)} 列, 保留 {len(remaining_columns)} 列")
                logger.info(f"删除的列: {', '.join(self.columns_to_delete)}")
            else:
                # 没有要删除的列，直接复制原始数据
                self.processed_data = self.original_data.copy()
                
                # 记录处理历史
                processing_record = {
                    'action': 'no_deletion',
                    'deleted_columns': [],
                    'original_columns_count': len(self.original_data.columns),
                    'remaining_columns_count': len(self.processed_data.columns),
                    'data_rows': len(self.processed_data)
                }
                self.processing_history.append(processing_record)
                
                logger.info("数据处理完成: 未删除任何列，保持原始数据结构")
            
            # 如果有需要重新计算的列，更新合计行
            if self.columns_to_recalculate:
                # 获取保留的数值列（排除已删除的列）
                remaining_columns = self.processed_data.columns.tolist()
                remaining_recalc_columns = [col for col in self.columns_to_recalculate if col in remaining_columns]
                if remaining_recalc_columns:
                    self.update_summary_row(remaining_recalc_columns)
                    logger.info(f"重新计算了合计行中的 {len(remaining_recalc_columns)} 列")
            
            return True
            
        except Exception as e:
            logger.error(f"数据处理失败: {e}")
            return False
    
    def get_processing_summary(self) -> Dict[str, Any]:
        """
        获取处理结果摘要
        
        Returns:
            Dict[str, Any]: 处理摘要信息
        """
        if self.original_data is None:
            return {}
        
        summary = {
            'original_columns_count': len(self.original_data.columns),
            'original_rows_count': len(self.original_data),
            'deleted_columns_count': len(self.columns_to_delete),
            'deleted_columns': self.columns_to_delete.copy(),
            'remaining_columns_count': len(self.processed_data.columns) if self.processed_data is not None else 0,
            'remaining_columns': self.processed_data.columns.tolist() if self.processed_data is not None else [],
            'data_rows_count': len(self.processed_data) if self.processed_data is not None else 0,
            'processing_history': self.processing_history.copy()
        }
        
        return summary
    
    def get_processed_data(self) -> Optional[pd.DataFrame]:
        """
        获取处理后的数据
        
        Returns:
            Optional[pd.DataFrame]: 处理后的数据，如果没有处理则返回None
        """
        return self.processed_data.copy() if self.processed_data is not None else None
    
    def get_original_data(self) -> Optional[pd.DataFrame]:
        """
        获取原始数据
        
        Returns:
            Optional[pd.DataFrame]: 原始数据
        """
        return self.original_data.copy() if self.original_data is not None else None
    
    def reset(self):
        """
        重置处理器状态
        """
        self.processed_data = self.original_data.copy() if self.original_data is not None else None
        self.columns_to_delete = []
        self.processing_history = []
        
        logger.info("数据处理器已重置")
    
    def get_column_info(self) -> Dict[str, Any]:
        """
        获取列信息统计
        
        Returns:
            Dict[str, Any]: 列信息
        """
        if self.original_data is None:
            return {}
        
        columns_info = []
        for i, col in enumerate(self.original_data.columns):
            col_info = {
                'index': i,
                'excel_column': self._get_excel_column_name(i),
                'name': col,
                'data_type': str(self.original_data[col].dtype),
                'non_null_count': self.original_data[col].count(),
                'null_count': self.original_data[col].isnull().sum(),
                'is_selected_for_deletion': col in self.columns_to_delete
            }
            columns_info.append(col_info)
        
        return {
            'total_columns': len(self.original_data.columns),
            'columns_info': columns_info
        }
    
    def _get_excel_column_name(self, index: int) -> str:
        """
        将列索引转换为Excel列名（A, B, C, ...）
        
        Args:
            index: 列索引（从0开始）
        
        Returns:
            str: Excel列名
        """
        result = ""
        while index >= 0:
            result = chr(65 + (index % 26)) + result
            index = index // 26 - 1
        return result
    
    def apply_template(self, template_columns: List[str]) -> bool:
        """
        应用模板（设置要删除的列）
        
        Args:
            template_columns: 模板中要删除的列名列表
        
        Returns:
            bool: 应用是否成功
        """
        try:
            if self.original_data is None:
                logger.error("没有加载数据")
                return False
            
            # 使用模糊匹配找到对应的列
            available_columns = self.original_data.columns.tolist()
            matched_columns = []
            
            for template_col in template_columns:
                # 精确匹配
                if template_col in available_columns:
                    matched_columns.append(template_col)
                else:
                    # 模糊匹配（包含关键词）
                    for available_col in available_columns:
                        if template_col in available_col or available_col in template_col:
                            if available_col not in matched_columns:
                                matched_columns.append(available_col)
                            break
            
            self.columns_to_delete = matched_columns
            
            logger.info(f"应用模板成功: 匹配到 {len(matched_columns)} 个列")
            logger.info(f"匹配的列: {', '.join(matched_columns)}")
            
            return True
            
        except Exception as e:
            logger.error(f"应用模板失败: {e}")
            return False
    
    def set_columns_to_recalculate(self, columns: List[str]) -> bool:
        """
        设置需要重新计算合计的列
        
        Args:
            columns: 需要重新计算的列名列表
        
        Returns:
            bool: 设置是否成功
        """
        try:
            if self.original_data is None:
                logger.error("没有加载数据")
                return False
            
            # 验证列是否存在且为数值类型
            valid_columns = []
            available_columns = self.original_data.columns.tolist()
            
            for col in columns:
                if col in available_columns:
                    if pd.api.types.is_numeric_dtype(self.original_data[col]):
                        valid_columns.append(col)
                    else:
                        logger.warning(f"列 {col} 不是数值类型，跳过")
                else:
                    logger.warning(f"列 {col} 不存在，跳过")
            
            self.columns_to_recalculate = valid_columns
            
            logger.info(f"设置需要重新计算的列: {len(valid_columns)} 个")
            logger.info(f"重新计算的列: {', '.join(valid_columns)}")
            
            return True
            
        except Exception as e:
            logger.error(f"设置重新计算列失败: {e}")
            return False
    
    def calculate_all_numeric_sums(self, data: pd.DataFrame = None) -> Dict[str, Any]:
        """
        计算所有数值列的求和统计
        
        Args:
            data: 要计算的数据，如果为None则使用processed_data
        
        Returns:
            Dict: 包含统计结果的字典
        """
        try:
            # 确定要计算的数据
            if data is None:
                if self.processed_data is None:
                    logger.error("没有可用的数据进行统计")
                    return {'success': False, 'error': '没有可用的数据'}
                target_data = self.processed_data
            else:
                target_data = data
            
            if target_data.empty:
                logger.error("数据为空")
                return {'success': False, 'error': '数据为空'}
            
            # 识别数值列
            numeric_columns = target_data.select_dtypes(include=['number']).columns.tolist()
            
            if not numeric_columns:
                logger.info("未找到数值列")
                return {'success': True, 'sums': {}, 'total_numeric_columns': 0}
            
            # 计算每列的统计信息
            column_stats = {}
            
            for col in numeric_columns:
                try:
                    # 获取列数据
                    col_data = target_data[col]
                    
                    # 计算统计信息
                    total_count = len(col_data)
                    null_count = col_data.isnull().sum()
                    valid_count = total_count - null_count
                    
                    if valid_count > 0:
                        # 计算求和
                        col_sum = col_data.sum()
                        
                        # 格式化显示（不使用千分位分隔符）
                        formatted_sum = f"{col_sum:.2f}"
                        
                        column_stats[col] = {
                            'sum': col_sum,
                            'formatted_sum': formatted_sum,
                            'total_count': total_count,
                            'valid_count': valid_count,
                            'null_count': null_count,
                            'average': col_sum / valid_count if valid_count > 0 else 0
                        }
                        
                        logger.debug(f"列 {col} 统计: 合计={formatted_sum}, 有效数据={valid_count}/{total_count}")
                    
                except Exception as e:
                    logger.warning(f"计算列 {col} 统计失败: {e}")
                    continue
            
            result = {
                'success': True,
                'sums': column_stats,
                'total_numeric_columns': len(numeric_columns),
                'processed_columns': len(column_stats)
            }
            
            logger.info(f"数值列统计完成: 共 {len(numeric_columns)} 个数值列，成功处理 {len(column_stats)} 个")
            return result
            
        except Exception as e:
            logger.error(f"计算数值列统计失败: {e}")
            return {'success': False, 'error': str(e)}
    
    def get_column_sum(self, column_name: str, data: pd.DataFrame = None) -> Dict[str, Any]:
        """
        获取指定列的求和信息
        
        Args:
            column_name: 列名
            data: 要计算的数据，如果为None则使用processed_data
        
        Returns:
            Dict: 包含求和结果的字典
        """
        try:
            # 确定要计算的数据
            if data is None:
                if self.processed_data is None:
                    return {'success': False, 'error': '没有可用的数据'}
                target_data = self.processed_data
            else:
                target_data = data
            
            if column_name not in target_data.columns:
                return {'success': False, 'error': f'列 {column_name} 不存在'}
            
            col_data = target_data[column_name]
            
            # 检查是否为数值列
            if not pd.api.types.is_numeric_dtype(col_data):
                return {'success': False, 'error': f'列 {column_name} 不是数值类型'}
            
            # 计算统计信息
            total_count = len(col_data)
            null_count = col_data.isnull().sum()
            valid_count = total_count - null_count
            
            if valid_count == 0:
                return {'success': False, 'error': f'列 {column_name} 没有有效的数值数据'}
            
            col_sum = col_data.sum()
            
            # 格式化显示（不使用千分位分隔符）
            formatted_sum = f"{col_sum:.2f}"
            
            result = {
                'success': True,
                'column_name': column_name,
                'sum': col_sum,
                'formatted_sum': formatted_sum,
                'total_count': total_count,
                'valid_count': valid_count,
                'null_count': null_count,
                'average': col_sum / valid_count
            }
            
            logger.info(f"列 {column_name} 求和: {formatted_sum}")
            return result
            
        except Exception as e:
            logger.error(f"计算列 {column_name} 求和失败: {e}")
            return {'success': False, 'error': str(e)}
    
    def identify_summary_row(self, data: pd.DataFrame = None) -> Dict[str, Any]:
        """
        识别合计行
        
        Args:
            data: 要分析的数据，如果为None则使用processed_data
        
        Returns:
            Dict: 包含合计行信息的字典
        """
        try:
            # 确定要分析的数据
            if data is None:
                if self.processed_data is None:
                    return {'success': False, 'error': '没有可用的数据'}
                target_data = self.processed_data
            else:
                target_data = data
            
            if target_data.empty:
                return {'success': False, 'error': '数据为空'}
            
            # 查找包含"合计"关键词的行
            summary_row_index = None
            summary_keywords = ['合计', '总计', '小计', 'Total', 'Sum']
            
            for index, row in target_data.iterrows():
                # 检查每一行的所有列是否包含合计关键词
                for col in target_data.columns:
                    cell_value = str(row[col]).strip()
                    if any(keyword in cell_value for keyword in summary_keywords):
                        summary_row_index = index
                        break
                if summary_row_index is not None:
                    break
            
            if summary_row_index is not None:
                result = {
                    'success': True,
                    'summary_row_index': summary_row_index,
                    'summary_row_data': target_data.iloc[summary_row_index].to_dict(),
                    'has_summary_row': True
                }
                logger.info(f"找到合计行，位置: 第{summary_row_index + 1}行")
            else:
                result = {
                    'success': True,
                    'summary_row_index': None,
                    'summary_row_data': None,
                    'has_summary_row': False
                }
                logger.info("未找到合计行")
            
            return result
            
        except Exception as e:
            logger.error(f"识别合计行失败: {e}")
            return {'success': False, 'error': str(e)}
    
    def update_summary_row(self, columns_to_recalculate: List[str], data: pd.DataFrame = None) -> bool:
        """
        更新合计行中指定列的求和值
        
        Args:
            columns_to_recalculate: 需要重新计算的列名列表
            data: 要处理的数据，如果为None则使用processed_data
        
        Returns:
            bool: 更新是否成功
        """
        try:
            # 确定要处理的数据
            if data is None:
                if self.processed_data is None:
                    logger.error("没有可用的数据")
                    return False
                target_data = self.processed_data.copy()
            else:
                target_data = data.copy()
            
            # 识别合计行
            summary_info = self.identify_summary_row(target_data)
            if not summary_info['success'] or not summary_info['has_summary_row']:
                logger.warning("未找到合计行，跳过更新")
                return True  # 没有合计行不算错误
            
            summary_row_index = summary_info['summary_row_index']
            
            # 获取数据行（排除合计行）
            data_rows = target_data.drop(summary_row_index)
            
            # 重新计算指定列的合计
            updated_values = {}
            for col_name in columns_to_recalculate:
                if col_name in target_data.columns:
                    # 检查是否为数值列
                    if pd.api.types.is_numeric_dtype(data_rows[col_name]):
                        # 计算合计（排除空值）
                        col_sum = data_rows[col_name].sum()
                        updated_values[col_name] = col_sum
                        logger.info(f"重新计算列 {col_name} 合计: {col_sum}")
                    else:
                        logger.warning(f"列 {col_name} 不是数值类型，跳过计算")
                else:
                    logger.warning(f"列 {col_name} 不存在，跳过计算")
            
            # 更新合计行的值
            for col_name, new_value in updated_values.items():
                target_data.at[summary_row_index, col_name] = new_value
            
            # 更新processed_data
            if data is None:
                self.processed_data = target_data
            
            logger.info(f"成功更新合计行，重新计算了 {len(updated_values)} 列")
            return True
            
        except Exception as e:
            logger.error(f"更新合计行失败: {e}")
            return False
    
    def get_numeric_columns_for_summary(self, data: pd.DataFrame = None) -> List[str]:
        """
        获取可用于合计计算的数值列
        
        Args:
            data: 要分析的数据，如果为None则使用processed_data
        
        Returns:
            List[str]: 数值列名列表
        """
        try:
            # 确定要分析的数据
            if data is None:
                if self.processed_data is None:
                    return []
                target_data = self.processed_data
            else:
                target_data = data
            
            if target_data.empty:
                return []
            
            # 识别合计行
            summary_info = self.identify_summary_row(target_data)
            
            # 获取数据行（如果有合计行则排除）
            if summary_info['success'] and summary_info['has_summary_row']:
                data_rows = target_data.drop(summary_info['summary_row_index'])
            else:
                data_rows = target_data
            
            # 获取数值列
            numeric_columns = data_rows.select_dtypes(include=['number']).columns.tolist()
            
            logger.info(f"找到 {len(numeric_columns)} 个可用于合计的数值列")
            return numeric_columns
            
        except Exception as e:
            logger.error(f"获取数值列失败: {e}")
            return []
    
    def load_cross_sheet_data(self, sheet_data: Dict[str, pd.DataFrame]) -> bool:
        """
        加载跨工作表数据用于关联
        
        Args:
            sheet_data: 包含所有工作表数据的字典，键为工作表名，值为DataFrame
            
        Returns:
            bool: 加载是否成功
        """
        try:
            self.cross_sheet_data = sheet_data.copy()
            logger.info(f"加载跨工作表数据成功，共 {len(sheet_data)} 个工作表")
            return True
        except Exception as e:
            logger.error(f"加载跨工作表数据失败: {e}")
            return False
    
    def extract_goods_names_by_invoice(self, invoice_sheet_name: str, detail_sheet_name: str, 
                                     invoice_column: str = None, 
                                     goods_column: str = '货物或应税劳务名称') -> Dict[str, str]:
        """
        根据发票号码从明细表中提取货物名称（智能字段选择）
        
        Args:
            invoice_sheet_name: 发票基础信息表名称
            detail_sheet_name: 发票明细表名称（信息汇总表）
            invoice_column: 发票号码列名（如果为None则自动选择）
            goods_column: 货物名称列名
            
        Returns:
            Dict[str, str]: 发票号码到货物名称的映射（只取第一条）
        """
        try:
            if not self.cross_sheet_data:
                logger.error("未加载跨工作表数据")
                return {}
            
            # 获取发票基础信息表和信息汇总表
            if invoice_sheet_name not in self.cross_sheet_data:
                logger.error(f"未找到发票基础信息表: {invoice_sheet_name}")
                return {}
            
            if detail_sheet_name not in self.cross_sheet_data:
                logger.error(f"未找到信息汇总表: {detail_sheet_name}")
                return {}
            
            invoice_df = self.cross_sheet_data[invoice_sheet_name]
            detail_df = self.cross_sheet_data[detail_sheet_name]
            
            # 智能选择关联字段
            if invoice_column is None:
                # 检查发票基础信息表中的可用字段
                available_invoice_columns = []
                if '发票号码' in invoice_df.columns:
                    available_invoice_columns.append('发票号码')
                if '数电发票号码' in invoice_df.columns:
                    available_invoice_columns.append('数电发票号码')
                
                if not available_invoice_columns:
                    logger.error("发票基础信息表中未找到发票号码相关字段")
                    return {}
                
                # 优先使用数电发票号码，如果不存在则使用发票号码
                invoice_column = '数电发票号码' if '数电发票号码' in available_invoice_columns else '发票号码'
                logger.info(f"自动选择关联字段: {invoice_column}")
            
            # 检查必要的列是否存在
            if invoice_column not in invoice_df.columns:
                logger.error(f"发票基础信息表中未找到列: {invoice_column}")
                return {}
            
            if invoice_column not in detail_df.columns:
                logger.error(f"信息汇总表中未找到列: {invoice_column}")
                return {}
            
            if goods_column not in detail_df.columns:
                logger.error(f"信息汇总表中未找到列: {goods_column}")
                return {}
            
            # 提取发票基础信息表中的发票号码
            invoice_numbers = invoice_df[invoice_column].dropna().unique()
            
            # 构建发票号码到货物名称的映射（只取第一条）
            invoice_goods_map = {}
            
            for invoice_num in invoice_numbers:
                # 在信息汇总表中查找对应的货物名称
                matching_rows = detail_df[detail_df[invoice_column] == invoice_num]
                
                if not matching_rows.empty:
                    # 只取第一条匹配记录的货物名称
                    first_goods_name = matching_rows[goods_column].iloc[0]
                    if pd.notna(first_goods_name) and str(first_goods_name).strip():
                        invoice_goods_map[str(invoice_num)] = str(first_goods_name).strip()
            
            logger.info(f"成功提取 {len(invoice_goods_map)} 个发票的货物名称信息（使用字段: {invoice_column}）")
            return invoice_goods_map
            
        except Exception as e:
            logger.error(f"提取货物名称失败: {e}")
            return {}
    
    def add_goods_names_to_invoice_sheet(self, invoice_sheet_name: str, 
                                       invoice_goods_map: Dict[str, str], 
                                       invoice_column: str = None,
                                       new_column_name: str = '货物或应税劳务名称') -> bool:
        """
        将货物名称添加到发票基础信息表的最后一列（智能字段选择）
        
        Args:
            invoice_sheet_name: 发票基础信息表名称
            invoice_goods_map: 发票号码到货物名称的映射
            invoice_column: 发票号码列名（如果为None则自动选择）
            new_column_name: 新增列的名称
            
        Returns:
            bool: 操作是否成功
        """
        try:
            if not self.cross_sheet_data or invoice_sheet_name not in self.cross_sheet_data:
                logger.error(f"未找到发票基础信息表: {invoice_sheet_name}")
                return False
            
            invoice_df = self.cross_sheet_data[invoice_sheet_name].copy()
            
            # 智能选择关联字段（与extract_goods_names_by_invoice保持一致）
            if invoice_column is None:
                # 检查发票基础信息表中的可用字段
                available_invoice_columns = []
                if '发票号码' in invoice_df.columns:
                    available_invoice_columns.append('发票号码')
                if '数电发票号码' in invoice_df.columns:
                    available_invoice_columns.append('数电发票号码')
                
                if not available_invoice_columns:
                    logger.error("发票基础信息表中未找到发票号码相关字段")
                    return False
                
                # 优先使用数电发票号码，如果不存在则使用发票号码
                invoice_column = '数电发票号码' if '数电发票号码' in available_invoice_columns else '发票号码'
                logger.info(f"自动选择关联字段: {invoice_column}")
            
            # 检查发票号码列是否存在
            if invoice_column not in invoice_df.columns:
                logger.error(f"发票基础信息表中未找到列: {invoice_column}")
                return False
            
            # 创建新的货物名称列
            goods_names_list = []
            
            for _, row in invoice_df.iterrows():
                invoice_num = str(row[invoice_column]) if pd.notna(row[invoice_column]) else ''
                
                if invoice_num in invoice_goods_map:
                    # 直接使用映射中的货物名称（已经是第一条）
                    goods_names_list.append(invoice_goods_map[invoice_num])
                else:
                    goods_names_list.append('')  # 未找到匹配的货物名称
            
            # 添加新列到DataFrame的最后
            invoice_df[new_column_name] = goods_names_list
            
            # 更新存储的数据
            self.cross_sheet_data[invoice_sheet_name] = invoice_df
            
            # 更新主数据源，确保后续处理能使用包含货物名称的数据
            if self.original_data is not None and len(self.original_data) == len(invoice_df):
                self.original_data = invoice_df.copy()
                logger.info("已更新原始数据源，包含货物名称列")
            
            # 如果当前处理的数据是发票基础信息表，也更新processed_data
            if self.processed_data is not None and len(self.processed_data) == len(invoice_df):
                self.processed_data = invoice_df.copy()
                logger.info("已更新处理数据，包含货物名称列")
            
            logger.info(f"成功添加货物名称列到 {invoice_sheet_name}，共 {len(goods_names_list)} 行数据（使用字段: {invoice_column}）")
            return True
            
        except Exception as e:
            logger.error(f"添加货物名称列失败: {e}")
            return False
    
    def process_cross_sheet_association(self, invoice_sheet_name: str, detail_sheet_name: str,
                                      invoice_column: str = None,
                                      goods_column: str = '货物或应税劳务名称',
                                      new_column_name: str = '货物或应税劳务名称') -> bool:
        """
        执行跨工作表数据关联的完整流程（智能字段选择）
        
        Args:
            invoice_sheet_name: 发票基础信息表名称
            detail_sheet_name: 信息汇总表名称
            invoice_column: 发票号码列名（如果为None则自动选择）
            goods_column: 货物名称列名
            new_column_name: 新增列的名称
            
        Returns:
            bool: 操作是否成功
        """
        try:
            # 1. 提取货物名称映射（使用智能字段选择）
            invoice_goods_map = self.extract_goods_names_by_invoice(
                invoice_sheet_name, detail_sheet_name, invoice_column, goods_column
            )
            
            if not invoice_goods_map:
                logger.warning("未找到任何发票号码匹配的货物名称")
                return False
            
            # 2. 将货物名称添加到发票基础信息表（使用智能字段选择）
            success = self.add_goods_names_to_invoice_sheet(
                invoice_sheet_name, invoice_goods_map, invoice_column, new_column_name
            )
            
            if success:
                logger.info("跨工作表数据关联处理完成")
                # 记录处理历史
                self.processing_history.append({
                    'operation': 'cross_sheet_association',
                    'invoice_sheet': invoice_sheet_name,
                    'detail_sheet': detail_sheet_name,
                    'matched_invoices': len(invoice_goods_map)
                })
            
            return success
            
        except Exception as e:
            logger.error(f"跨工作表数据关联处理失败: {e}")
            return False
