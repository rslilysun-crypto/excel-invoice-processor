# -*- coding: utf-8 -*-
"""
求和列选择器组件
专门用于选择需要重新计算合计的数值列
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Callable

from src.utils.logger import get_logger

logger = get_logger("SumColumnSelector")

class SumColumnSelector:
    """
    求和列选择器
    专门用于选择需要重新计算合计的数值列
    """
    
    def __init__(self, parent: tk.Widget, numeric_columns: List[str], 
                 selected_columns: List[str], callback: Callable[[List[str]], None]):
        """
        初始化求和列选择器
        
        Args:
            parent: 父窗口
            numeric_columns: 可用的数值列列表
            selected_columns: 已选择的列
            callback: 选择完成回调函数
        """
        self.parent = parent
        self.numeric_columns = numeric_columns
        self.selected_columns = selected_columns.copy()
        self.callback = callback
        
        self.window = None
        self.column_vars = {}  # 列复选框变量
        
        self._create_window()
        self._create_widgets()
        self._setup_layout()
        self._populate_columns()
        
        logger.info("求和列选择器初始化完成")
    
    def _create_window(self):
        """
        创建窗口
        """
        self.window = tk.Toplevel(self.parent)
        self.window.title("选择需要重新计算合计的列")
        self.window.geometry("500x400")
        self.window.resizable(True, True)
        
        # 设置窗口图标（如果有的话）
        try:
            self.window.iconbitmap(default="")
        except:
            pass
        
        # 设置为模态窗口
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # 居中显示
        self._center_window()
    
    def _center_window(self):
        """
        窗口居中显示
        """
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
    
    def _create_widgets(self):
        """
        创建界面组件
        """
        # 主框架
        self.main_frame = ttk.Frame(self.window, padding="10")
        
        # 说明标签
        self.description_label = ttk.Label(
            self.main_frame,
            text="请选择需要重新计算合计值的数值列：\n选中的列将在数据处理后重新计算合计行中的求和值。",
            font=("Arial", 10),
            justify="left"
        )
        
        # 列选择区域
        self.selection_frame = ttk.LabelFrame(self.main_frame, text="可用的数值列", padding="10")
        
        # 滚动框架
        self.canvas = tk.Canvas(self.selection_frame, height=200)
        self.scrollbar = ttk.Scrollbar(self.selection_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 统计信息
        self.stats_label = ttk.Label(self.main_frame, text="", font=("Arial", 9))
        
        # 按钮区域
        self.button_frame = ttk.Frame(self.main_frame)
        
        self.select_all_btn = ttk.Button(
            self.button_frame,
            text="全选",
            command=self._select_all,
            width=8
        )
        
        self.select_none_btn = ttk.Button(
            self.button_frame,
            text="全不选",
            command=self._select_none,
            width=8
        )
        
        self.confirm_btn = ttk.Button(
            self.button_frame,
            text="确认选择",
            command=self._confirm_selection,
            width=10
        )
        
        self.cancel_btn = ttk.Button(
            self.button_frame,
            text="取消",
            command=self._cancel_selection,
            width=8
        )
    
    def _setup_layout(self):
        """
        设置布局
        """
        # 主框架
        self.main_frame.pack(fill="both", expand=True)
        
        # 说明标签
        self.description_label.pack(pady=(0, 10), anchor="w")
        
        # 列选择区域
        self.selection_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # 滚动区域
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # 统计信息
        self.stats_label.pack(pady=(0, 10))
        
        # 按钮区域
        self.button_frame.pack(fill="x")
        
        self.select_all_btn.pack(side="left", padx=(0, 5))
        self.select_none_btn.pack(side="left", padx=(0, 20))
        self.cancel_btn.pack(side="right")
        self.confirm_btn.pack(side="right", padx=(0, 10))
    
    def _populate_columns(self):
        """
        填充列选择项
        """
        # 清除现有内容
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.column_vars.clear()
        
        # 创建列选择项
        for i, column in enumerate(self.numeric_columns):
            var = tk.BooleanVar()
            var.set(column in self.selected_columns)
            self.column_vars[column] = var
            
            # 创建复选框
            checkbox = ttk.Checkbutton(
                self.scrollable_frame,
                text=f"{column}",
                variable=var,
                command=self._update_stats
            )
            checkbox.grid(row=i, column=0, sticky="w", padx=5, pady=2)
        
        # 更新统计信息
        self._update_stats()
        
        # 绑定鼠标滚轮
        self._bind_mousewheel()
    
    def _bind_mousewheel(self):
        """
        绑定鼠标滚轮事件
        """
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        self.canvas.bind("<MouseWheel>", _on_mousewheel)
    
    def _update_stats(self):
        """
        更新统计信息
        """
        selected_count = sum(1 for var in self.column_vars.values() if var.get())
        total_count = len(self.column_vars)
        
        self.stats_label.config(
            text=f"已选择 {selected_count} / {total_count} 列用于重新计算合计"
        )
    
    def _select_all(self):
        """
        全选
        """
        for var in self.column_vars.values():
            var.set(True)
        self._update_stats()
    
    def _select_none(self):
        """
        全不选
        """
        for var in self.column_vars.values():
            var.set(False)
        self._update_stats()
    
    def _confirm_selection(self):
        """
        确认选择
        """
        try:
            # 获取选中的列
            selected = [col for col, var in self.column_vars.items() if var.get()]
            
            logger.info(f"确认选择 {len(selected)} 列用于重新计算合计")
            
            # 调用回调函数
            if self.callback:
                self.callback(selected)
            
            # 关闭窗口
            self.window.destroy()
            
        except Exception as e:
            logger.error(f"确认选择失败: {e}")
            messagebox.showerror("错误", f"确认选择失败: {str(e)}")
    
    def _cancel_selection(self):
        """
        取消选择
        """
        logger.info("取消求和列选择")
        self.window.destroy()
    
    def show(self):
        """
        显示窗口
        """
        if self.window:
            self.window.deiconify()
            self.window.lift()
            self.window.focus_force()