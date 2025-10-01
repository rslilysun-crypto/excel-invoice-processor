# -*- coding: utf-8 -*-
"""
列选择器组件
提供动态列选择、模板应用等功能
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import List, Dict, Any, Callable, Optional
import re

from src.utils.logger import get_logger

logger = get_logger("ColumnSelector")

class ColumnSelector:
    """
    列选择器
    提供动态列选择界面和模板管理功能
    """
    
    def __init__(self, parent: tk.Widget, headers: List[str], 
                 selected_columns: List[str], config_manager, 
                 callback: Callable[[List[str]], None]):
        """
        初始化列选择器
        
        Args:
            parent: 父窗口
            headers: 列标题列表
            selected_columns: 已选择的列
            config_manager: 配置管理器
            callback: 选择完成回调函数
        """
        self.parent = parent
        self.headers = headers
        self.selected_columns = selected_columns.copy()
        self.config_manager = config_manager
        self.callback = callback
        
        self.window = None
        self.column_vars = {}  # 列复选框变量
        self.search_var = tk.StringVar()
        self.template_var = tk.StringVar()
        
        self._create_window()
        self._create_widgets()
        self._setup_layout()
        self._load_templates()
        self._populate_columns()
        self._update_selection_count()
        
        logger.info("列选择器初始化完成")
    
    def _create_window(self):
        """
        创建选择器窗口
        """
        self.window = tk.Toplevel(self.parent)
        self.window.title("选择要删除的列")
        self.window.geometry("900x750")
        self.window.resizable(True, True)
        
        # 设置为模态窗口
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # 居中显示
        self._center_window()
    
    def _center_window(self):
        """
        将窗口居中显示
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
        self.main_frame = ttk.Frame(self.window, padding="15")
        
        # 标题和说明
        self.title_label = ttk.Label(
            self.main_frame, 
            text="选择要删除的列", 
            font=("Arial", 14, "bold")
        )
        
        self.desc_label = ttk.Label(
            self.main_frame,
            text="请选择需要从Excel文件中删除的列。选中的列将在处理后被移除。",
            foreground="gray"
        )
        
        # 模板选择区域
        self._create_template_area()
        
        # 搜索区域
        self._create_search_area()
        
        # 列选择区域
        self._create_column_selection_area()
        
        # 选择统计区域
        self._create_selection_stats_area()
        
        # 按钮区域
        self._create_button_area()
    
    def _create_template_area(self):
        """
        创建模板选择区域
        """
        self.template_frame = ttk.LabelFrame(self.main_frame, text="预设模板", padding="10")
        
        # 模板选择下拉框
        self.template_combo = ttk.Combobox(
            self.template_frame,
            textvariable=self.template_var,
            state="readonly",
            width=30
        )
        
        # 模板操作按钮
        self.template_buttons_frame = ttk.Frame(self.template_frame)
        
        self.apply_template_btn = ttk.Button(
            self.template_buttons_frame,
            text="应用模板",
            command=self._apply_template,
            width=10
        )
        
        self.save_template_btn = ttk.Button(
            self.template_buttons_frame,
            text="保存为模板",
            command=self._save_template,
            width=12
        )
        
        self.delete_template_btn = ttk.Button(
            self.template_buttons_frame,
            text="删除模板",
            command=self._delete_template,
            width=10
        )
    
    def _create_search_area(self):
        """
        创建搜索区域
        """
        self.search_frame = ttk.LabelFrame(self.main_frame, text="搜索和筛选", padding="10")
        
        # 搜索输入框
        self.search_entry = ttk.Entry(
            self.search_frame,
            textvariable=self.search_var,
            width=30
        )
        self.search_var.trace("w", self._on_search_changed)
        
        # 搜索按钮
        self.search_btn = ttk.Button(
            self.search_frame,
            text="搜索",
            command=self._filter_columns,
            width=8
        )
        
        # 清除搜索按钮
        self.clear_search_btn = ttk.Button(
            self.search_frame,
            text="清除",
            command=self._clear_search,
            width=8
        )
    
    def _create_column_selection_area(self):
        """
        创建列选择区域
        """
        self.selection_frame = ttk.LabelFrame(self.main_frame, text="列选择", padding="10")
        
        # 操作按钮行
        self.selection_buttons_frame = ttk.Frame(self.selection_frame)
        
        self.select_all_btn = ttk.Button(
            self.selection_buttons_frame,
            text="全选",
            command=self._select_all,
            width=8
        )
        
        self.select_none_btn = ttk.Button(
            self.selection_buttons_frame,
            text="全不选",
            command=self._select_none,
            width=8
        )
        
        self.invert_selection_btn = ttk.Button(
            self.selection_buttons_frame,
            text="反选",
            command=self._invert_selection,
            width=8
        )
        
        # 列列表框架
        self.columns_list_frame = ttk.Frame(self.selection_frame)
        
        # 创建滚动的列选择区域
        self.columns_canvas = tk.Canvas(self.columns_list_frame, height=180)
        self.columns_scrollbar = ttk.Scrollbar(
            self.columns_list_frame, 
            orient="vertical", 
            command=self.columns_canvas.yview
        )
        self.columns_canvas.configure(yscrollcommand=self.columns_scrollbar.set)
        
        # 创建内部框架
        self.columns_inner_frame = ttk.Frame(self.columns_canvas)
        self.columns_canvas.create_window((0, 0), window=self.columns_inner_frame, anchor="nw")
    
    def _create_selection_stats_area(self):
        """
        创建选择统计区域
        """
        self.stats_frame = ttk.LabelFrame(self.main_frame, text="选择统计", padding="10")
        
        self.stats_var = tk.StringVar()
        self.stats_label = ttk.Label(
            self.stats_frame,
            textvariable=self.stats_var,
            font=("Arial", 10)
        )
        
        # 选中列预览
        self.selected_preview_frame = ttk.Frame(self.stats_frame)
        
        self.selected_label = ttk.Label(
            self.selected_preview_frame,
            text="已选择的列:",
            font=("Arial", 9, "bold")
        )
        
        # 创建选中列的滚动文本框
        self.selected_text_frame = ttk.Frame(self.selected_preview_frame)
        
        self.selected_text = tk.Text(
            self.selected_text_frame,
            height=2,
            width=60,
            wrap=tk.WORD,
            state="disabled"
        )
        
        self.selected_text_scrollbar = ttk.Scrollbar(
            self.selected_text_frame,
            orient="vertical",
            command=self.selected_text.yview
        )
        self.selected_text.configure(yscrollcommand=self.selected_text_scrollbar.set)
    
    def _create_button_area(self):
        """
        创建按钮区域
        """
        self.button_frame = ttk.Frame(self.main_frame)
        
        # 确认按钮
        self.confirm_btn = ttk.Button(
            self.button_frame,
            text="确认选择",
            command=self._confirm_selection,
            width=12
        )
        
        # 取消按钮
        self.cancel_btn = ttk.Button(
            self.button_frame,
            text="取消",
            command=self._cancel_selection,
            width=8
        )
        
        # 预览按钮
        self.preview_btn = ttk.Button(
            self.button_frame,
            text="预览效果",
            command=self._preview_result,
            width=10
        )
    
    def _setup_layout(self):
        """
        设置界面布局
        """
        # 主框架
        self.main_frame.pack(fill="both", expand=True)
        
        # 标题和说明
        self.title_label.pack(pady=(0, 5))
        self.desc_label.pack(pady=(0, 15))
        
        # 模板区域
        self.template_frame.pack(fill="x", pady=(0, 10))
        self.template_combo.pack(side="left", padx=(0, 10))
        self.template_buttons_frame.pack(side="left")
        self.apply_template_btn.pack(side="left", padx=(0, 5))
        self.save_template_btn.pack(side="left", padx=(0, 5))
        self.delete_template_btn.pack(side="left")
        
        # 搜索区域
        self.search_frame.pack(fill="x", pady=(0, 10))
        self.search_entry.pack(side="left", padx=(0, 10))
        self.search_btn.pack(side="left", padx=(0, 5))
        self.clear_search_btn.pack(side="left")
        
        # 列选择区域
        self.selection_frame.pack(fill="x", pady=(0, 10))
        
        # 选择操作按钮
        self.selection_buttons_frame.pack(fill="x", pady=(0, 10))
        self.select_all_btn.pack(side="left", padx=(0, 5))
        self.select_none_btn.pack(side="left", padx=(0, 5))
        self.invert_selection_btn.pack(side="left")
        
        # 列列表
        self.columns_list_frame.pack(fill="x")
        self.columns_canvas.pack(side="left", fill="both", expand=True)
        self.columns_scrollbar.pack(side="right", fill="y")
        
        # 选择统计区域
        self.stats_frame.pack(fill="x", pady=(0, 10))
        self.stats_label.pack(pady=(0, 10))
        
        self.selected_preview_frame.pack(fill="x")
        self.selected_label.pack(anchor="w", pady=(0, 5))
        
        self.selected_text_frame.pack(fill="x")
        self.selected_text.pack(side="left", fill="both", expand=True)
        self.selected_text_scrollbar.pack(side="right", fill="y")
        
        # 按钮区域
        self.button_frame.pack(fill="x", pady=(20, 15))
        self.cancel_btn.pack(side="right")
        self.confirm_btn.pack(side="right", padx=(0, 10))
        self.preview_btn.pack(side="right", padx=(0, 10))
    
    def _load_templates(self):
        """
        加载模板列表
        """
        try:
            templates = self.config_manager.load_templates()
            template_names = list(templates.keys())
            
            self.template_combo['values'] = template_names
            
            # 设置默认模板
            default_template = self.config_manager.get_setting("default_template", "")
            if default_template in template_names:
                self.template_var.set(default_template)
            
        except Exception as e:
            logger.error(f"加载模板失败: {e}")
    
    def _populate_columns(self):
        """
        填充列选择列表 - 多列布局，每列5个复选框
        """
        try:
            # 清除现有组件
            for widget in self.columns_inner_frame.winfo_children():
                widget.destroy()
            
            self.column_vars = {}
            
            # 计算需要的列数（每列5个复选框）
            items_per_column = 5
            total_items = len(self.headers)
            num_columns = (total_items + items_per_column - 1) // items_per_column  # 向上取整
            
            # 创建列框架列表
            column_frames = []
            for col_idx in range(num_columns):
                col_frame = ttk.Frame(self.columns_inner_frame)
                col_frame.pack(side="left", fill="y", padx=10, pady=5)
                column_frames.append(col_frame)
            
            # 创建列复选框
            for i, header in enumerate(self.headers):
                # 计算当前复选框应该放在哪一列
                col_idx = i // items_per_column
                
                # 创建复选框变量
                var = tk.BooleanVar()
                var.set(header in self.selected_columns)
                self.column_vars[header] = var
                
                # 创建复选框框架
                checkbox_frame = ttk.Frame(column_frames[col_idx])
                checkbox_frame.pack(fill="x", padx=2, pady=2)
                
                # Excel列标识
                excel_col = self._get_excel_column_name(i)
                col_label = ttk.Label(
                    checkbox_frame,
                    text=f"{excel_col}:",
                    width=4,
                    font=("Arial", 9, "bold")
                )
                col_label.pack(side="left")
                
                # 复选框
                checkbox = ttk.Checkbutton(
                    checkbox_frame,
                    text=header,
                    variable=var,
                    command=self._on_selection_changed
                )
                checkbox.pack(side="left", fill="x", expand=True, padx=(5, 0))
                
                # 绑定变量变化事件
                var.trace("w", self._on_selection_changed)
            
            # 更新滚动区域
            self.columns_inner_frame.update_idletasks()
            self.columns_canvas.configure(scrollregion=self.columns_canvas.bbox("all"))
            
            # 绑定鼠标滚轮事件
            self._bind_mousewheel()
            
        except Exception as e:
            logger.error(f"填充列选择列表失败: {e}")
    
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
    
    def _bind_mousewheel(self):
        """
        绑定鼠标滚轮事件
        """
        def _on_mousewheel(event):
            self.columns_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        self.columns_canvas.bind("<MouseWheel>", _on_mousewheel)
    
    def _on_selection_changed(self, *args):
        """
        选择变化事件处理
        """
        self._update_selection_count()
        self._update_selected_preview()
    
    def _update_selection_count(self):
        """
        更新选择统计
        """
        try:
            selected_count = sum(1 for var in self.column_vars.values() if var.get())
            total_count = len(self.headers)
            remaining_count = total_count - selected_count
            
            self.stats_var.set(
                f"总列数: {total_count} | 已选择删除: {selected_count} | 将保留: {remaining_count}"
            )
            
        except Exception as e:
            logger.error(f"更新选择统计失败: {e}")
    
    def _update_selected_preview(self):
        """
        更新选中列预览
        """
        try:
            selected_columns = [header for header, var in self.column_vars.items() if var.get()]
            
            # 更新文本框
            self.selected_text.config(state="normal")
            self.selected_text.delete(1.0, tk.END)
            
            if selected_columns:
                preview_text = ", ".join(selected_columns)
                self.selected_text.insert(1.0, preview_text)
            else:
                self.selected_text.insert(1.0, "(未选择任何列)")
            
            self.selected_text.config(state="disabled")
            
        except Exception as e:
            logger.error(f"更新选中列预览失败: {e}")
    
    def _on_search_changed(self, *args):
        """
        搜索内容变化事件
        """
        # 实时搜索
        self._filter_columns()
    
    def _filter_columns(self):
        """
        根据搜索条件筛选列 - 适配多列布局
        """
        try:
            search_text = self.search_var.get().lower().strip()
            
            # 遍历所有列框架
            for col_frame in self.columns_inner_frame.winfo_children():
                if isinstance(col_frame, ttk.Frame):
                    # 遍历每列中的复选框框架
                    for checkbox_frame in col_frame.winfo_children():
                        if isinstance(checkbox_frame, ttk.Frame):
                            # 查找复选框组件
                            for child in checkbox_frame.winfo_children():
                                if isinstance(child, ttk.Checkbutton):
                                    checkbox_text = child.cget("text").lower()
                                    if not search_text or search_text in checkbox_text:
                                        checkbox_frame.pack(fill="x", padx=2, pady=2)
                                    else:
                                        checkbox_frame.pack_forget()
                                    break
            
            # 更新滚动区域
            self.columns_inner_frame.update_idletasks()
            self.columns_canvas.configure(scrollregion=self.columns_canvas.bbox("all"))
            
        except Exception as e:
            logger.error(f"筛选列失败: {e}")
    
    def _clear_search(self):
        """
        清除搜索
        """
        self.search_var.set("")
        self._filter_columns()
    
    def _select_all(self):
        """
        全选所有列
        """
        for var in self.column_vars.values():
            var.set(True)
    
    def _select_none(self):
        """
        取消选择所有列
        """
        for var in self.column_vars.values():
            var.set(False)
    
    def _invert_selection(self):
        """
        反选
        """
        for var in self.column_vars.values():
            var.set(not var.get())
    
    def _apply_template(self):
        """
        应用选中的模板
        """
        try:
            template_name = self.template_var.get()
            if not template_name:
                messagebox.showwarning("警告", "请先选择一个模板")
                return
            
            templates = self.config_manager.load_templates()
            if template_name not in templates:
                messagebox.showerror("错误", "模板不存在")
                return
            
            template = templates[template_name]
            template_columns = template.get("columns_to_delete", [])
            
            # 先清除所有选择
            self._select_none()
            
            # 应用模板选择（模糊匹配）
            matched_count = 0
            for template_col in template_columns:
                for header in self.headers:
                    if (template_col.lower() in header.lower() or 
                        header.lower() in template_col.lower()):
                        if header in self.column_vars:
                            self.column_vars[header].set(True)
                            matched_count += 1
                        break
            
            messagebox.showinfo(
                "模板应用完成", 
                f"已应用模板 '{template_name}'\n匹配到 {matched_count} 个列"
            )
            
            logger.info(f"应用模板: {template_name}, 匹配 {matched_count} 列")
            
        except Exception as e:
            logger.error(f"应用模板失败: {e}")
            messagebox.showerror("错误", f"应用模板失败: {str(e)}")
    
    def _save_template(self):
        """
        保存当前选择为模板
        """
        try:
            # 获取当前选择的列
            selected_columns = [header for header, var in self.column_vars.items() if var.get()]
            
            if not selected_columns:
                messagebox.showwarning("警告", "请先选择要删除的列")
                return
            
            # 输入模板名称
            template_name = simpledialog.askstring(
                "保存模板",
                "请输入模板名称:",
                parent=self.window
            )
            
            if not template_name:
                return
            
            # 输入模板描述
            description = simpledialog.askstring(
                "模板描述",
                "请输入模板描述（可选）:",
                parent=self.window
            )
            
            if description is None:
                description = ""
            
            # 保存模板
            self.config_manager.add_template(template_name, selected_columns, description)
            
            # 重新加载模板列表
            self._load_templates()
            self.template_var.set(template_name)
            
            messagebox.showinfo(
                "保存成功", 
                f"模板 '{template_name}' 已保存\n包含 {len(selected_columns)} 个列"
            )
            
            logger.info(f"保存模板: {template_name}, {len(selected_columns)} 列")
            
        except Exception as e:
            logger.error(f"保存模板失败: {e}")
            messagebox.showerror("错误", f"保存模板失败: {str(e)}")
    
    def _delete_template(self):
        """
        删除选中的模板
        """
        try:
            template_name = self.template_var.get()
            if not template_name:
                messagebox.showwarning("警告", "请先选择一个模板")
                return
            
            # 确认删除
            result = messagebox.askyesno(
                "确认删除",
                f"确定要删除模板 '{template_name}' 吗？",
                parent=self.window
            )
            
            if result:
                if self.config_manager.delete_template(template_name):
                    # 重新加载模板列表
                    self._load_templates()
                    self.template_var.set("")
                    
                    messagebox.showinfo("删除成功", f"模板 '{template_name}' 已删除")
                    logger.info(f"删除模板: {template_name}")
                else:
                    messagebox.showerror("删除失败", "无法删除该模板（可能是默认模板）")
            
        except Exception as e:
            logger.error(f"删除模板失败: {e}")
            messagebox.showerror("错误", f"删除模板失败: {str(e)}")
    
    def _preview_result(self):
        """
        预览删除效果
        """
        try:
            selected_columns = [header for header, var in self.column_vars.items() if var.get()]
            remaining_columns = [header for header in self.headers if header not in selected_columns]
            
            # 创建预览窗口
            preview_window = tk.Toplevel(self.window)
            preview_window.title("删除效果预览")
            preview_window.geometry("600x400")
            
            main_frame = ttk.Frame(preview_window, padding="15")
            main_frame.pack(fill="both", expand=True)
            
            # 标题
            title_label = ttk.Label(main_frame, text="删除效果预览", font=("Arial", 12, "bold"))
            title_label.pack(pady=(0, 15))
            
            # 统计信息
            stats_frame = ttk.LabelFrame(main_frame, text="统计信息", padding="10")
            stats_frame.pack(fill="x", pady=(0, 15))
            
            ttk.Label(stats_frame, text=f"原始列数: {len(self.headers)}").pack(anchor="w")
            ttk.Label(stats_frame, text=f"删除列数: {len(selected_columns)}").pack(anchor="w")
            ttk.Label(stats_frame, text=f"保留列数: {len(remaining_columns)}").pack(anchor="w")
            
            # 创建标签页
            notebook = ttk.Notebook(main_frame)
            notebook.pack(fill="both", expand=True)
            
            # 删除的列标签页
            deleted_frame = ttk.Frame(notebook)
            notebook.add(deleted_frame, text=f"将删除的列 ({len(selected_columns)})")
            
            deleted_text = tk.Text(deleted_frame, wrap=tk.WORD)
            deleted_scrollbar = ttk.Scrollbar(deleted_frame, orient="vertical", command=deleted_text.yview)
            deleted_text.configure(yscrollcommand=deleted_scrollbar.set)
            
            deleted_text.pack(side="left", fill="both", expand=True)
            deleted_scrollbar.pack(side="right", fill="y")
            
            if selected_columns:
                deleted_text.insert(1.0, "\n".join(selected_columns))
            else:
                deleted_text.insert(1.0, "(没有选择要删除的列)")
            
            deleted_text.config(state="disabled")
            
            # 保留的列标签页
            remaining_frame = ttk.Frame(notebook)
            notebook.add(remaining_frame, text=f"将保留的列 ({len(remaining_columns)})")
            
            remaining_text = tk.Text(remaining_frame, wrap=tk.WORD)
            remaining_scrollbar = ttk.Scrollbar(remaining_frame, orient="vertical", command=remaining_text.yview)
            remaining_text.configure(yscrollcommand=remaining_scrollbar.set)
            
            remaining_text.pack(side="left", fill="both", expand=True)
            remaining_scrollbar.pack(side="right", fill="y")
            
            if remaining_columns:
                remaining_text.insert(1.0, "\n".join(remaining_columns))
            else:
                remaining_text.insert(1.0, "(所有列都将被删除 - 这是不允许的!)")
                remaining_text.config(fg="red")
            
            remaining_text.config(state="disabled")
            
            # 关闭按钮
            close_btn = ttk.Button(main_frame, text="关闭", command=preview_window.destroy)
            close_btn.pack(pady=(15, 0))
            
        except Exception as e:
            logger.error(f"预览效果失败: {e}")
            messagebox.showerror("错误", f"预览效果失败: {str(e)}")
    
    def _confirm_selection(self):
        """
        确认选择
        """
        try:
            selected_columns = [header for header, var in self.column_vars.items() if var.get()]
            
            # 验证选择
            if not selected_columns:
                messagebox.showwarning("警告", "请至少选择一列要删除")
                return
            
            remaining_columns = [header for header in self.headers if header not in selected_columns]
            if not remaining_columns:
                messagebox.showerror("错误", "不能删除所有列，请至少保留一列数据")
                return
            
            # 确认对话框
            result = messagebox.askyesno(
                "确认选择",
                f"确定要删除选中的 {len(selected_columns)} 列吗？\n\n"
                f"删除后将保留 {len(remaining_columns)} 列数据。",
                parent=self.window
            )
            
            if result:
                # 调用回调函数
                if self.callback:
                    self.callback(selected_columns)
                
                # 关闭窗口
                self.window.destroy()
                
                logger.info(f"确认选择 {len(selected_columns)} 列删除")
            
        except Exception as e:
            logger.error(f"确认选择失败: {e}")
            messagebox.showerror("错误", f"确认选择失败: {str(e)}")
    
    def _cancel_selection(self):
        """
        取消选择
        """
        self.window.destroy()
        logger.info("取消列选择")
    
    def show(self):
        """
        显示选择器窗口
        """
        if self.window:
            self.window.deiconify()
            self.window.lift()
            self.window.focus_force()