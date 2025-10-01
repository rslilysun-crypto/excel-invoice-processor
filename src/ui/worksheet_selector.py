# -*- coding: utf-8 -*-
"""
工作表选择器
提供工作表选择和预览功能
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Any, Optional, Callable
from src.utils.logger import get_logger

logger = get_logger("WorksheetSelector")

class WorksheetSelector:
    """
    工作表选择器类
    提供工作表选择、预览和确认功能
    """
    
    def __init__(self, parent: tk.Tk, worksheets: List[Dict[str, Any]], 
                 current_selection: str = None, on_select: Optional[Callable] = None):
        """
        初始化工作表选择器
        
        Args:
            parent: 父窗口
            worksheets: 工作表信息列表
            current_selection: 当前选中的工作表
            on_select: 选择回调函数
        """
        self.parent = parent
        self.worksheets = worksheets
        self.current_selection = current_selection
        self.on_select = on_select
        self.selected_worksheet = current_selection
        
        # 界面组件
        self.dialog = None
        self.worksheet_listbox = None
        self.info_text = None
        self.preview_text = None
        self.confirm_button = None
        
        self._create_dialog()
        
    def _create_dialog(self):
        """
        创建选择器对话框
        """
        try:
            # 创建顶级窗口
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title("选择工作表")
            self.dialog.geometry("600x500")
            self.dialog.resizable(True, True)
            
            # 设置为模态对话框
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            
            # 居中显示
            self._center_dialog()
            
            # 创建主框架
            main_frame = ttk.Frame(self.dialog, padding="10")
            main_frame.pack(fill="both", expand=True)
            
            # 标题标签
            title_label = ttk.Label(
                main_frame,
                text="请选择要处理的工作表",
                font=("Arial", 12, "bold")
            )
            title_label.pack(pady=(0, 10))
            
            # 创建左右分栏
            paned_window = ttk.PanedWindow(main_frame, orient="horizontal")
            paned_window.pack(fill="both", expand=True, pady=(0, 10))
            
            # 左侧：工作表列表
            left_frame = ttk.LabelFrame(paned_window, text="工作表列表", padding="5")
            paned_window.add(left_frame, weight=1)
            
            # 工作表列表框
            list_frame = ttk.Frame(left_frame)
            list_frame.pack(fill="both", expand=True)
            
            self.worksheet_listbox = tk.Listbox(
                list_frame,
                selectmode="single",
                font=("Arial", 10)
            )
            
            # 滚动条
            list_scrollbar = ttk.Scrollbar(
                list_frame,
                orient="vertical",
                command=self.worksheet_listbox.yview
            )
            self.worksheet_listbox.configure(yscrollcommand=list_scrollbar.set)
            
            self.worksheet_listbox.pack(side="left", fill="both", expand=True)
            list_scrollbar.pack(side="right", fill="y")
            
            # 绑定选择事件
            self.worksheet_listbox.bind("<<ListboxSelect>>", self._on_worksheet_selected)
            
            # 右侧：工作表信息和预览
            right_frame = ttk.Frame(paned_window)
            paned_window.add(right_frame, weight=1)
            
            # 工作表信息
            info_frame = ttk.LabelFrame(right_frame, text="工作表信息", padding="5")
            info_frame.pack(fill="x", pady=(0, 5))
            
            self.info_text = tk.Text(
                info_frame,
                height=6,
                wrap=tk.WORD,
                state="disabled",
                font=("Arial", 9)
            )
            self.info_text.pack(fill="x")
            
            # 数据预览
            preview_frame = ttk.LabelFrame(right_frame, text="数据预览", padding="5")
            preview_frame.pack(fill="both", expand=True)
            
            preview_text_frame = ttk.Frame(preview_frame)
            preview_text_frame.pack(fill="both", expand=True)
            
            self.preview_text = tk.Text(
                preview_text_frame,
                wrap=tk.NONE,
                state="disabled",
                font=("Consolas", 8)
            )
            
            # 预览滚动条
            preview_v_scrollbar = ttk.Scrollbar(
                preview_text_frame,
                orient="vertical",
                command=self.preview_text.yview
            )
            preview_h_scrollbar = ttk.Scrollbar(
                preview_text_frame,
                orient="horizontal",
                command=self.preview_text.xview
            )
            
            self.preview_text.configure(
                yscrollcommand=preview_v_scrollbar.set,
                xscrollcommand=preview_h_scrollbar.set
            )
            
            self.preview_text.grid(row=0, column=0, sticky="nsew")
            preview_v_scrollbar.grid(row=0, column=1, sticky="ns")
            preview_h_scrollbar.grid(row=1, column=0, sticky="ew")
            
            preview_text_frame.grid_rowconfigure(0, weight=1)
            preview_text_frame.grid_columnconfigure(0, weight=1)
            
            # 按钮框架
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x")
            
            # 刷新按钮
            refresh_button = ttk.Button(
                button_frame,
                text="刷新",
                command=self._refresh_worksheets
            )
            refresh_button.pack(side="left")
            
            # 取消按钮
            cancel_button = ttk.Button(
                button_frame,
                text="取消",
                command=self._on_cancel
            )
            cancel_button.pack(side="right", padx=(5, 0))
            
            # 确认按钮
            self.confirm_button = ttk.Button(
                button_frame,
                text="确认选择",
                command=self._on_confirm,
                state="disabled"
            )
            self.confirm_button.pack(side="right")
            
            # 加载工作表列表
            self._load_worksheets()
            
            # 绑定关闭事件
            self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
            
            logger.info("工作表选择器创建完成")
            
        except Exception as e:
            logger.error(f"创建工作表选择器失败: {e}")
    
    def _center_dialog(self):
        """
        将对话框居中显示
        """
        try:
            self.dialog.update_idletasks()
            
            # 获取对话框尺寸
            dialog_width = self.dialog.winfo_width()
            dialog_height = self.dialog.winfo_height()
            
            # 获取父窗口位置和尺寸
            parent_x = self.parent.winfo_x()
            parent_y = self.parent.winfo_y()
            parent_width = self.parent.winfo_width()
            parent_height = self.parent.winfo_height()
            
            # 计算居中位置
            x = parent_x + (parent_width - dialog_width) // 2
            y = parent_y + (parent_height - dialog_height) // 2
            
            self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
            
        except Exception as e:
            logger.error(f"居中对话框失败: {e}")
    
    def _load_worksheets(self):
        """
        加载工作表列表
        """
        try:
            # 清空列表
            self.worksheet_listbox.delete(0, tk.END)
            
            # 添加工作表
            for i, worksheet in enumerate(self.worksheets):
                name = worksheet['name']
                rows = worksheet.get('rows', 0)
                columns = worksheet.get('columns', 0)
                has_data = worksheet.get('has_data', False)
                
                # 格式化显示文本
                status = "有数据" if has_data else "无数据"
                display_text = f"{name} ({rows}行 x {columns}列) - {status}"
                
                self.worksheet_listbox.insert(tk.END, display_text)
                
                # 如果是当前选中的工作表，设置选中状态
                if name == self.current_selection:
                    self.worksheet_listbox.selection_set(i)
                    self.worksheet_listbox.activate(i)
                    self._show_worksheet_info(worksheet)
            
            # 如果没有当前选择，选择第一个有数据的工作表
            if not self.current_selection and self.worksheets:
                for i, worksheet in enumerate(self.worksheets):
                    if worksheet.get('has_data', False):
                        self.worksheet_listbox.selection_set(i)
                        self.worksheet_listbox.activate(i)
                        self._show_worksheet_info(worksheet)
                        break
                else:
                    # 如果没有有数据的工作表，选择第一个
                    self.worksheet_listbox.selection_set(0)
                    self.worksheet_listbox.activate(0)
                    self._show_worksheet_info(self.worksheets[0])
            
            logger.info(f"加载了 {len(self.worksheets)} 个工作表")
            
        except Exception as e:
            logger.error(f"加载工作表列表失败: {e}")
    
    def _on_worksheet_selected(self, event):
        """
        工作表选择事件处理
        
        Args:
            event: 选择事件
        """
        try:
            selection = self.worksheet_listbox.curselection()
            if selection:
                index = selection[0]
                worksheet = self.worksheets[index]
                self.selected_worksheet = worksheet['name']
                
                # 显示工作表信息
                self._show_worksheet_info(worksheet)
                
                # 启用确认按钮
                self.confirm_button.config(state="normal")
                
                logger.info(f"选择工作表: {self.selected_worksheet}")
            
        except Exception as e:
            logger.error(f"处理工作表选择事件失败: {e}")
    
    def _show_worksheet_info(self, worksheet: Dict[str, Any]):
        """
        显示工作表信息
        
        Args:
            worksheet: 工作表信息
        """
        try:
            # 更新信息文本
            self.info_text.config(state="normal")
            self.info_text.delete(1.0, tk.END)
            
            info_lines = [
                f"工作表名称: {worksheet['name']}",
                f"数据行数: {worksheet.get('rows', 0)}",
                f"数据列数: {worksheet.get('columns', 0)}",
                f"数据状态: {'有数据' if worksheet.get('has_data', False) else '无数据'}"
            ]
            
            # 如果有列信息，显示列名
            columns = worksheet.get('columns_list', [])
            if columns:
                info_lines.append(f"\n列名列表 (共{len(columns)}列):")
                for i, col in enumerate(columns[:10]):  # 最多显示前10列
                    info_lines.append(f"  {i+1}. {col}")
                if len(columns) > 10:
                    info_lines.append(f"  ... 还有 {len(columns) - 10} 列")
            
            self.info_text.insert(1.0, "\n".join(info_lines))
            self.info_text.config(state="disabled")
            
            # 更新预览（这里可以添加实际的数据预览功能）
            self._show_data_preview(worksheet)
            
        except Exception as e:
            logger.error(f"显示工作表信息失败: {e}")
    
    def _show_data_preview(self, worksheet: Dict[str, Any]):
        """
        显示数据预览
        
        Args:
            worksheet: 工作表信息
        """
        try:
            # 更新预览文本
            self.preview_text.config(state="normal")
            self.preview_text.delete(1.0, tk.END)
            
            # 这里应该调用实际的数据读取功能来获取预览数据
            # 目前只显示基本信息
            preview_lines = [
                "数据预览功能",
                "=" * 50,
                f"工作表: {worksheet['name']}",
                f"行数: {worksheet.get('rows', 0)}",
                f"列数: {worksheet.get('columns', 0)}",
                "",
                "注意: 实际数据预览需要连接到Excel读取器"
            ]
            
            self.preview_text.insert(1.0, "\n".join(preview_lines))
            self.preview_text.config(state="disabled")
            
        except Exception as e:
            logger.error(f"显示数据预览失败: {e}")
    
    def _refresh_worksheets(self):
        """
        刷新工作表列表
        """
        try:
            # 重新加载工作表列表
            self._load_worksheets()
            logger.info("工作表列表已刷新")
            
        except Exception as e:
            logger.error(f"刷新工作表列表失败: {e}")
    
    def _on_confirm(self):
        """
        确认选择
        """
        try:
            if self.selected_worksheet:
                if self.on_select:
                    self.on_select(self.selected_worksheet)
                
                self._close_dialog()
                logger.info(f"确认选择工作表: {self.selected_worksheet}")
            
        except Exception as e:
            logger.error(f"确认选择失败: {e}")
    
    def _on_cancel(self):
        """
        取消选择
        """
        try:
            self._close_dialog()
            logger.info("取消工作表选择")
            
        except Exception as e:
            logger.error(f"取消选择失败: {e}")
    
    def _close_dialog(self):
        """
        关闭对话框
        """
        try:
            if self.dialog:
                self.dialog.grab_release()
                self.dialog.destroy()
                self.dialog = None
            
        except Exception as e:
            logger.error(f"关闭对话框失败: {e}")
    
    def show(self):
        """
        显示选择器
        """
        try:
            if self.dialog:
                self.dialog.deiconify()
                self.dialog.lift()
                self.dialog.focus_set()
            
        except Exception as e:
            logger.error(f"显示工作表选择器失败: {e}")

# 便捷函数
def select_worksheet(parent: tk.Tk, worksheets: List[Dict[str, Any]], 
                    current_selection: str = None, on_select: Optional[Callable] = None) -> WorksheetSelector:
    """
    显示工作表选择器
    
    Args:
        parent: 父窗口
        worksheets: 工作表信息列表
        current_selection: 当前选中的工作表
        on_select: 选择回调函数
        
    Returns:
        WorksheetSelector: 工作表选择器实例
    """
    return WorksheetSelector(parent, worksheets, current_selection, on_select)