# -*- coding: utf-8 -*-
"""
主窗口界面
提供文件选择、工作表选择、列选择等主要功能界面
"""

import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import ttk, filedialog, messagebox
import os
from typing import Optional, Dict, Any

from src.core.excel_reader import ExcelReader
from src.core.data_processor import DataProcessor
from src.core.file_handler import FileHandler
from src.utils.config import ConfigManager
from src.utils.logger import get_logger
from src.ui.worksheet_selector import WorksheetSelector
from src.ui.column_selector import ColumnSelector
from src.ui.sum_column_selector import SumColumnSelector
from src.ui.progress_dialog import ProgressDialog

logger = get_logger("MainWindow")

class MainWindow:
    """
    主窗口类
    管理整个应用程序的主界面和核心功能
    """
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.excel_reader = ExcelReader()
        self.data_processor = DataProcessor()
        self.file_handler = FileHandler()
        self.config_manager = ConfigManager()
        
        # 子窗口组件
        self.worksheet_selector = None
        self.column_selector = None
        self.progress_dialog = None
        
        # 当前状态
        self.current_file_path = None
        self.current_worksheet = None
        self.selected_columns_to_delete = []
        self.selected_columns_to_recalculate = []  # 选择的求和列
        
        # 批量处理相关
        self.batch_files = []  # 批量文件列表
        self.is_batch_mode = False  # 是否为批量模式
        
        # 初始化界面
        self._setup_window()
        self._create_widgets()
        self._setup_layout()
        self._bind_events()
        
        logger.info("主窗口初始化完成")
    
    def _setup_window(self):
        """
        设置主窗口属性
        """
        self.root.title("Excel发票数据处理软件 v1.0.0")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 设置窗口图标（如果有的话）
        try:
            logger.info(f"拖放事件数据: {event.data}")
            # self.root.iconbitmap("icon.ico")
            pass
        except:
            pass
        
        # 居中显示窗口
        self._center_window()
    
    def _center_window(self):
        """
        将窗口居中显示
        """
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def _create_widgets(self):
        """
        创建界面组件
        """
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        
        # 标题标签
        self.title_label = ttk.Label(
            self.main_frame, 
            text="Excel发票数据处理软件", 
            font=("Arial", 16, "bold")
        )
        
        # 文件选择区域
        self._create_file_selection_area()
        
        # 文件信息显示区域
        self._create_file_info_area()
        
        # 数据统计信息区域
        self._create_data_stats_area()
        
        # 工作表选择区域
        self._create_worksheet_area()
        
        # 操作按钮区域
        self._create_action_buttons()
        
        # 状态栏
        self._create_status_bar()
    
    def _create_file_selection_area(self):
        """
        创建文件选择区域
        """
        # 文件选择框架
        self.file_frame = ttk.LabelFrame(self.main_frame, text="文件选择", padding="10")
        
        # 文件路径显示
        self.file_path_var = tk.StringVar(value="请选择Excel文件...")
        self.file_path_label = ttk.Label(
            self.file_frame, 
            textvariable=self.file_path_var,
            foreground="gray",
            font=("Arial", 9)
        )
        
        # 按钮框架
        self.file_buttons_frame = ttk.Frame(self.file_frame)
        
        # 选择文件按钮
        self.select_file_btn = ttk.Button(
            self.file_buttons_frame,
            text="选择文件",
            command=self._select_file,
            width=12
        )
        
        # 清除文件按钮
        self.clear_file_btn = ttk.Button(
            self.file_buttons_frame,
            text="清除",
            command=self._clear_file,
            width=8,
            state="disabled"
        )
        
        # 拖拽提示标签
        self.drag_label = ttk.Label(
            self.file_frame,
            text="或将Excel文件拖拽到此处",
            foreground="gray",
            font=("Arial", 9, "italic")
        )
    
    def _create_file_info_area(self):
        """
        创建文件信息显示区域
        """
        self.info_frame = ttk.LabelFrame(self.main_frame, text="文件信息", padding="10")
        
        # 创建信息显示的树形视图
        self.info_tree = ttk.Treeview(
            self.info_frame,
            columns=("value",),
            show="tree headings",
            height=6
        )
        
        self.info_tree.heading("#0", text="属性")
        self.info_tree.heading("value", text="值")
        self.info_tree.column("#0", width=150)
        self.info_tree.column("value", width=300)
        
        # 滚动条
        self.info_scrollbar = ttk.Scrollbar(self.info_frame, orient="vertical", command=self.info_tree.yview)
        self.info_tree.configure(yscrollcommand=self.info_scrollbar.set)
    
    def _create_data_stats_area(self):
        """
        创建数据统计信息显示区域
        """
        self.stats_frame = ttk.LabelFrame(self.main_frame, text="数据统计", padding="10")
        
        # 统计信息显示标签
        self.stats_info_var = tk.StringVar(value="请先选择Excel文件和工作表")
        self.stats_info_label = ttk.Label(
            self.stats_frame,
            textvariable=self.stats_info_var,
            foreground="blue",
            font=("Arial", 10, "bold")
        )
        self.stats_info_label.pack(anchor="w")
        
        # 详细统计信息文本框
        self.stats_text_frame = ttk.Frame(self.stats_frame)
        self.stats_text = tk.Text(
            self.stats_text_frame,
            height=4,
            wrap=tk.WORD,
            state="disabled",
            font=("Arial", 9)
        )
        
        # 统计信息滚动条
        self.stats_scrollbar = ttk.Scrollbar(
            self.stats_text_frame,
            orient="vertical",
            command=self.stats_text.yview
        )
        self.stats_text.configure(yscrollcommand=self.stats_scrollbar.set)
        
        # 默认隐藏详细统计信息
        self.stats_detail_visible = False
    
    def _create_worksheet_area(self):
        """
        创建工作表选择区域
        """
        self.worksheet_frame = ttk.LabelFrame(self.main_frame, text="工作表选择", padding="10")
        
        # 工作表选择下拉框
        self.worksheet_var = tk.StringVar()
        self.worksheet_combo = ttk.Combobox(
            self.worksheet_frame,
            textvariable=self.worksheet_var,
            state="readonly",
            width=30
        )
        self.worksheet_combo.bind("<<ComboboxSelected>>", self._on_worksheet_selected)
        
        # 工作表信息标签
        self.worksheet_info_var = tk.StringVar(value="请先选择Excel文件")
        self.worksheet_info_label = ttk.Label(
            self.worksheet_frame,
            textvariable=self.worksheet_info_var,
            foreground="gray"
        )
    
    def _create_action_buttons(self):
        """
        创建操作按钮区域
        """
        self.action_frame = ttk.Frame(self.main_frame)
        
        # 列选择按钮
        self.column_select_btn = ttk.Button(
            self.action_frame,
            text="选择要删除的列",
            command=self._open_column_selector,
            state="disabled",
            width=15
        )
        
        # 求和列选择按钮
        self.sum_select_btn = ttk.Button(
            self.action_frame,
            text="选择求和列",
            command=self._open_sum_column_selector,
            state="disabled",
            width=12
        )
        
        # 数据预览按钮
        self.preview_btn = ttk.Button(
            self.action_frame,
            text="数据预览",
            command=self._preview_data,
            state="disabled",
            width=12
        )
        
        # 开始处理按钮
        self.process_btn = ttk.Button(
            self.action_frame,
            text="开始处理",
            command=self._start_processing,
            state="disabled",
            width=12
        )
        
        # 边框选项复选框
        self.border_var = tk.BooleanVar(value=True)  # 默认添加边框
        self.border_checkbox = ttk.Checkbutton(
            self.action_frame,
            text="添加表格边框",
            variable=self.border_var
        )
        
        # 跨工作表数据关联复选框
        self.cross_sheet_var = tk.BooleanVar(value=False)  # 默认不启用
        self.cross_sheet_checkbox = ttk.Checkbutton(
            self.action_frame,
            text="启用跨工作表数据关联",
            variable=self.cross_sheet_var
        )
        
        # 设置按钮
        self.settings_btn = ttk.Button(
            self.action_frame,
            text="设置",
            command=self._open_settings,
            width=8
        )
    
    def _create_status_bar(self):
        """
        创建状态栏
        """
        self.status_frame = ttk.Frame(self.root)
        
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(
            self.status_frame,
            textvariable=self.status_var,
            relief="sunken",
            anchor="w"
        )
        
        # 进度条（初始隐藏）
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.status_frame,
            variable=self.progress_var,
            maximum=100
        )
    
    def _setup_layout(self):
        """
        设置界面布局
        """
        # 主框架
        self.main_frame.pack(fill="both", expand=True)
        
        # 标题
        self.title_label.pack(pady=(0, 20))
        
        # 文件选择区域
        self.file_frame.pack(fill="x", pady=(0, 10))
        self.file_path_label.pack(fill="x", pady=(0, 10))
        self.file_buttons_frame.pack(fill="x")
        self.select_file_btn.pack(side="left")
        self.clear_file_btn.pack(side="left", padx=(10, 0))
        self.drag_label.pack(pady=(10, 0))
        
        # 文件信息区域
        self.info_frame.pack(fill="x", pady=(0, 10))
        self.info_tree.pack(side="left", fill="both", expand=True)
        self.info_scrollbar.pack(side="right", fill="y")
        
        # 数据统计区域
        self.stats_frame.pack(fill="x", pady=(0, 10))
        
        # 工作表选择区域
        self.worksheet_frame.pack(fill="x", pady=(0, 10))
        self.worksheet_combo.pack(pady=(0, 5))
        self.worksheet_info_label.pack()
        
        # 操作按钮区域
        self.action_frame.pack(fill="x", pady=(0, 10))
        self.column_select_btn.pack(side="left", padx=(0, 10))
        self.sum_select_btn.pack(side="left", padx=(0, 10))
        self.preview_btn.pack(side="left", padx=(0, 10))
        self.process_btn.pack(side="left", padx=(0, 10))
        self.border_checkbox.pack(side="left", padx=(20, 10))
        self.cross_sheet_checkbox.pack(side="left", padx=(10, 10))
        self.settings_btn.pack(side="right")
        
        # 状态栏
        self.status_frame.pack(fill="x", side="bottom")
        self.status_label.pack(side="left", fill="x", expand=True)
    
    def _on_drop(self, event):
        """
        处理文件拖拽事件
        """
        try:
            # 获取拖拽的文件列表
            files = self.root.tk.splitlist(event.data)
            
            # 过滤出Excel文件
            excel_files = []
            for file_path in files:
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    excel_files.append(file_path)
            
            if not excel_files:
                messagebox.showwarning("文件类型错误", "请拖拽Excel文件（.xlsx 或 .xls）")
                return
            
            logger.info(f"拖拽事件数据: {event.data}")
            logger.info(f"检测到Excel文件: {excel_files}")
            
            if len(excel_files) == 1:
                # 单文件模式
                self._load_file(excel_files[0])
            else:
                # 批量处理模式
                self._load_batch_files(excel_files)
                
        except Exception as e:
            logger.error(f"处理拖拽事件时发生错误: {str(e)}")
            messagebox.showerror("拖拽错误", f"处理拖拽文件时发生错误:\n{str(e)}")

    def _bind_events(self):
        """
        绑定事件
        """
        # 窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # 绑定拖拽事件到文件选择区域
        self.file_frame.drop_target_register(DND_FILES)
        self.file_frame.dnd_bind('<<Drop>>', self._on_drop)

    def _select_file(self):
        """
        选择Excel文件（支持多选）
        """
        try:
            # 获取上次使用的目录
            initial_dir = self.config_manager.get_setting("last_input_directory", "")
            
            file_paths = filedialog.askopenfilenames(
                title="选择Excel文件（可多选）",
                initialdir=initial_dir,
                filetypes=[
                    ("Excel文件", "*.xlsx *.xls"),
                    ("Excel 2007+", "*.xlsx"),
                    ("Excel 97-2003", "*.xls"),
                    ("所有文件", "*.*")
                ]
            )
            
            if file_paths:
                if len(file_paths) == 1:
                    # 单文件模式
                    self._load_file(file_paths[0])
                else:
                    # 批量处理模式
                    self._load_batch_files(file_paths)
                
                # 保存目录到配置
                self.config_manager.update_setting("last_input_directory", os.path.dirname(file_paths[0]))
                
        except Exception as e:
            logger.error(f"选择文件失败: {e}")
            messagebox.showerror("错误", f"选择文件失败: {str(e)}")
    
    def _load_file(self, file_path: str):
        """
        加载Excel文件
        
        Args:
            file_path: 文件路径
        """
        try:
            self._update_status("正在加载文件...")
            
            # 加载文件
            if self.excel_reader.load_file(file_path):
                self.current_file_path = file_path
                self.is_batch_mode = False
                
                # 更新界面
                self._update_file_info()
                self._update_worksheet_list()
                self._enable_controls(True)
                
                # 自动选择目标工作表
                target_worksheet = self.excel_reader.get_target_worksheet()
                if target_worksheet:
                    self.worksheet_var.set(target_worksheet)
                    self._on_worksheet_selected()
                
                self._update_status("文件加载成功")
                logger.info(f"文件加载成功: {file_path}")
            else:
                self._update_status("文件加载失败")
                messagebox.showerror("错误", "无法加载Excel文件，请检查文件格式是否正确")
                
        except Exception as e:
            logger.error(f"加载文件失败: {e}")
            self._update_status("文件加载失败")
            messagebox.showerror("错误", f"加载文件失败: {str(e)}")
    
    def _load_batch_files(self, file_paths: list):
        """
        加载批量文件
        
        Args:
            file_paths: 文件路径列表
        """
        try:
            self._update_status("正在加载批量文件...")
            
            self.batch_files = file_paths
            self.is_batch_mode = True
            
            # 更新界面显示
            self.file_path_var.set(f"已选择 {len(file_paths)} 个文件进行批量处理")
            
            # 更新文件信息显示
            self._update_batch_file_info()
            
            # 读取第一个文件的工作表信息
            self._load_batch_worksheet_info(file_paths[0])
            
            # 禁用工作表选择（批量模式下自动处理）
            self.worksheet_combo.config(state="disabled")
            
            # 启用相关控件
            self.clear_file_btn.config(state="normal")
            self.column_select_btn.config(state="normal")
            self.process_btn.config(state="normal")
            
            self._update_status(f"批量文件加载成功，共 {len(file_paths)} 个文件")
            logger.info(f"批量文件加载成功: {len(file_paths)} 个文件")
            
        except Exception as e:
            logger.error(f"加载批量文件失败: {e}")
            self._update_status("批量文件加载失败")
            messagebox.showerror("错误", f"加载批量文件失败: {str(e)}")
    
    def _clear_file(self):
        """
        清除当前文件
        """
        self.excel_reader.close()
        self.current_file_path = None
        self.current_worksheet = None
        self.selected_columns_to_delete = []
        self.batch_files = []
        self.is_batch_mode = False
        
        # 重置界面
        self.file_path_var.set("请选择Excel文件...")
        self._clear_file_info()
        self._clear_data_stats()
        self._clear_worksheet_list()
        self._enable_controls(False)
        
        self._update_status("就绪")
        logger.info("文件已清除")
    
    def _update_file_info(self):
        """
        更新文件信息显示
        """
        try:
            # 清除现有信息
            for item in self.info_tree.get_children():
                self.info_tree.delete(item)
            
            # 获取文件信息
            file_info = self.excel_reader.get_file_info()
            
            if file_info:
                # 更新文件路径显示
                self.file_path_var.set(file_info['file_path'])
                
                # 添加文件信息到树形视图
                self.info_tree.insert("", "end", text="文件名", values=(file_info['file_name'],))
                self.info_tree.insert("", "end", text="文件大小", values=(f"{file_info['file_size_mb']} MB",))
                self.info_tree.insert("", "end", text="工作表数量", values=(file_info['worksheet_count'],))
                
                # 添加工作表列表
                worksheets_item = self.info_tree.insert("", "end", text="工作表列表", values=("",))
                for ws_name in file_info['worksheets']:
                    self.info_tree.insert(worksheets_item, "end", text=f"  {ws_name}", values=("",))
                
                # 展开工作表列表
                self.info_tree.item(worksheets_item, open=True)
                
        except Exception as e:
            logger.error(f"更新文件信息失败: {e}")
    
    def _update_batch_file_info(self):
        """
        更新批量文件信息显示
        """
        try:
            # 清除现有信息
            for item in self.info_tree.get_children():
                self.info_tree.delete(item)
            
            # 添加批量文件信息
            self.info_tree.insert("", "end", text="处理模式", values=("批量处理",))
            self.info_tree.insert("", "end", text="文件数量", values=(len(self.batch_files),))
            
            # 添加文件列表
            files_item = self.info_tree.insert("", "end", text="文件列表", values=("",))
            for i, file_path in enumerate(self.batch_files, 1):
                file_name = os.path.basename(file_path)
                self.info_tree.insert(files_item, "end", text=f"  {i}. {file_name}", values=("",))
            
            # 展开文件列表
            self.info_tree.item(files_item, open=True)
            
        except Exception as e:
            logger.error(f"更新批量文件信息失败: {e}")
    
    def _load_batch_worksheet_info(self, first_file_path: str):
        """
        加载批量模式下的工作表信息
        
        Args:
            first_file_path: 第一个文件的路径
        """
        try:
            # 临时加载第一个文件获取工作表信息
            temp_reader = ExcelReader()
            if temp_reader.load_file(first_file_path):
                worksheets = temp_reader.get_worksheets_list()
                
                if worksheets:
                    # 尝试找到"发票基础信息"工作表
                    target_worksheet = None
                    for ws in worksheets:
                        if "发票基础信息" in ws['name']:
                            target_worksheet = ws['name']
                            break
                    
                    # 如果没找到，使用第一个工作表
                    if not target_worksheet:
                        target_worksheet = worksheets[0]['name']
                    
                    # 更新工作表选择显示
                    worksheet_names = [ws['name'] for ws in worksheets]
                    self.worksheet_combo['values'] = worksheet_names
                    self.worksheet_var.set(target_worksheet)
                    
                    # 设置批量模式提示信息
                    self.worksheet_info_var.set(f"批量模式：将处理所有文件的'{target_worksheet}'工作表")
                    
                    logger.info(f"批量模式：检测到工作表 {target_worksheet}")
                else:
                    self.worksheet_info_var.set("批量模式：未检测到工作表")
                    logger.warning("批量模式：第一个文件中未找到工作表")
            else:
                self.worksheet_info_var.set("批量模式：无法读取工作表信息")
                logger.error(f"批量模式：无法加载第一个文件 {first_file_path}")
                
        except Exception as e:
            logger.error(f"加载批量工作表信息失败: {e}")
            self.worksheet_info_var.set("批量模式：工作表信息加载失败")
    
    def _clear_file_info(self):
        """
        清除文件信息显示
        """
        for item in self.info_tree.get_children():
            self.info_tree.delete(item)
    
    def _update_worksheet_list(self):
        """
        更新工作表列表
        """
        try:
            worksheets = self.excel_reader.get_worksheets_list()
            worksheet_names = [ws['name'] for ws in worksheets]
            
            self.worksheet_combo['values'] = worksheet_names
            
            if worksheet_names:
                self.worksheet_info_var.set(f"共 {len(worksheet_names)} 个工作表")
            else:
                self.worksheet_info_var.set("没有找到工作表")
                
        except Exception as e:
            logger.error(f"更新工作表列表失败: {e}")
    
    def _clear_worksheet_list(self):
        """
        清除工作表列表
        """
        self.worksheet_combo['values'] = ()
        self.worksheet_var.set("")
        self.worksheet_info_var.set("请先选择Excel文件")
    
    def _update_data_stats(self):
        """
        更新数据统计信息
        """
        try:
            if not self.current_worksheet or not self.current_file_path:
                self.stats_info_var.set("请先选择Excel文件和工作表")
                return
            
            # 读取当前工作表的数据
            data = self.excel_reader.read_full_data(self.current_worksheet)
            if data.empty:
                self.stats_info_var.set("当前工作表没有数据")
                return
            
            # 计算数值列统计
            stats_result = self.data_processor.calculate_all_numeric_sums(data)
            
            if stats_result['success'] and stats_result['sums']:
                # 查找价税合计列
                price_tax_total_col = None
                for col_name in stats_result['sums'].keys():
                    if '价税合计' in col_name or '合计' in col_name:
                        price_tax_total_col = col_name
                        break
                
                if price_tax_total_col:
                    # 显示价税合计信息
                    col_stats = stats_result['sums'][price_tax_total_col]
                    self.stats_info_var.set(
                        f"价税合计: {col_stats['formatted_sum']} 元 "
                        f"(共 {col_stats['valid_count']} 条记录)"
                    )
                else:
                    # 显示总体统计信息
                    total_numeric_cols = stats_result['total_numeric_columns']
                    self.stats_info_var.set(f"共找到 {total_numeric_cols} 个数值列，点击'数据预览'查看详细统计")
                
                logger.info(f"数据统计更新成功: {len(stats_result['sums'])} 个数值列")
            else:
                self.stats_info_var.set("未找到数值列或计算统计失败")
                
        except Exception as e:
            logger.error(f"更新数据统计失败: {e}")
            self.stats_info_var.set("统计信息计算失败")
    
    def _clear_data_stats(self):
        """
        清除数据统计信息
        """
        self.stats_info_var.set("请先选择Excel文件和工作表")
        
        # 清除详细统计信息
        if hasattr(self, 'stats_text'):
            self.stats_text.config(state="normal")
            self.stats_text.delete(1.0, tk.END)
            self.stats_text.config(state="disabled")
    
    def _on_worksheet_selected(self, event=None):
        """
        工作表选择事件处理
        """
        try:
            selected_worksheet = self.worksheet_var.get()
            if selected_worksheet:
                # 选择工作表
                if self.excel_reader.select_worksheet(selected_worksheet):
                    self.current_worksheet = selected_worksheet
                    
                    # 更新工作表信息
                    worksheets_info = self.excel_reader.worksheets_info
                    if selected_worksheet in worksheets_info:
                        info = worksheets_info[selected_worksheet]
                        self.worksheet_info_var.set(
                            f"'{selected_worksheet}' - {info['max_row']}行 x {info['max_column']}列"
                        )
                    
                    # 启用相关按钮
                    self.column_select_btn.config(state="normal")
                    self.sum_select_btn.config(state="normal")
                    self.preview_btn.config(state="normal")
                    
                    # 检查是否可以启用处理按钮
                    self._check_process_button_state()
                    
                    # 更新数据统计信息
                    self._update_data_stats()
                    
                    self._update_status(f"已选择工作表: {selected_worksheet}")
                    logger.info(f"选择工作表: {selected_worksheet}")
                    
        except Exception as e:
            logger.error(f"选择工作表失败: {e}")
            messagebox.showerror("错误", f"选择工作表失败: {str(e)}")
    
    def _enable_controls(self, enabled: bool):
        """
        启用/禁用控件
        
        Args:
            enabled: 是否启用
        """
        state = "normal" if enabled else "disabled"
        
        self.clear_file_btn.config(state=state)
        
        if not enabled:
            self.worksheet_combo.config(state="disabled")
            self.column_select_btn.config(state="disabled")
            self.sum_select_btn.config(state="disabled")
            self.preview_btn.config(state="disabled")
            self.process_btn.config(state="disabled")
        else:
            self.worksheet_combo.config(state="readonly")
    
    def _check_process_button_state(self):
        """
        检查处理按钮的启用状态
        """
        # 基本条件：文件已加载且工作表已选择
        if self.excel_reader and self.current_worksheet:
            self.process_btn.config(state="normal")
        else:
            self.process_btn.config(state="disabled")
    
    def _open_column_selector(self):
        """
        打开列选择器
        """
        try:
            headers = None
            
            if self.is_batch_mode:
                # 批量模式：从第一个文件读取表头
                if not self.batch_files:
                    messagebox.showwarning("警告", "没有选择批量文件")
                    return
                
                # 临时加载第一个文件获取表头
                temp_reader = ExcelReader()
                if temp_reader.load_file(self.batch_files[0]):
                    # 尝试找到"发票基础信息"工作表
                    worksheets = temp_reader.get_worksheets_list()
                    target_worksheet = None
                    
                    for ws in worksheets:
                        if "发票基础信息" in ws['name']:
                            target_worksheet = ws['name']
                            break
                    
                    if not target_worksheet and worksheets:
                        target_worksheet = worksheets[0]['name']
                    
                    if target_worksheet:
                        headers = temp_reader.read_headers(target_worksheet)
                
                if not headers:
                    messagebox.showerror("错误", "无法从批量文件中读取表头信息")
                    return
            else:
                # 单文件模式
                if not self.current_worksheet:
                    messagebox.showwarning("警告", "请先选择工作表")
                    return
                
                headers = self.excel_reader.read_headers(self.current_worksheet)
                if not headers:
                    messagebox.showerror("错误", "无法读取表头信息")
                    return
            
            # 创建列选择器窗口
            self.column_selector = ColumnSelector(
                self.root,
                headers,
                self.selected_columns_to_delete,
                self.config_manager,
                self._on_columns_selected
            )
            
        except Exception as e:
            logger.error(f"打开列选择器失败: {e}")
            messagebox.showerror("错误", f"打开列选择器失败: {str(e)}")
    
    def _on_columns_selected(self, selected_columns: list):
        """
        列选择完成回调
        
        Args:
            selected_columns: 选中的列名列表
        """
        self.selected_columns_to_delete = selected_columns
        
        if selected_columns:
            self._update_status(f"已选择 {len(selected_columns)} 列待删除")
        else:
            self._update_status("未选择要删除的列")
        
        # 检查是否可以启用处理按钮（不再强制要求选择删除列）
        self._check_process_button_state()
        
        logger.info(f"选择了 {len(selected_columns)} 列待删除")
    
    def _open_sum_column_selector(self):
        """
        打开求和列选择器
        """
        try:
            if not self.current_worksheet:
                messagebox.showwarning("警告", "请先选择工作表")
                return
            
            # 读取表头
            headers = self.excel_reader.read_headers(self.current_worksheet)
            if not headers:
                messagebox.showerror("错误", "无法读取表头")
                return
            
            # 获取可用于求和的数值列
            full_data = self.excel_reader.read_full_data(self.current_worksheet)
            numeric_columns = self.data_processor.get_numeric_columns_for_summary(full_data)
            
            if not numeric_columns:
                messagebox.showinfo("提示", "当前工作表没有可用于求和的数值列")
                return
            
            # 过滤出数值列的表头
            numeric_headers = [header for header in headers if header in numeric_columns]
            
            if not numeric_headers:
                messagebox.showinfo("提示", "当前工作表没有可用于求和的数值列")
                return
            
            # 创建求和列选择器
            self.sum_column_selector = SumColumnSelector(
                self.root,
                numeric_columns,
                self.selected_columns_to_recalculate.copy(),
                self._on_sum_columns_selected
            )
            
            # 显示选择器
            self.sum_column_selector.show()
            
        except Exception as e:
            logger.error(f"打开求和列选择器失败: {e}")
            messagebox.showerror("错误", f"打开求和列选择器失败: {str(e)}")
    
    def _on_sum_columns_selected(self, selected_columns: list):
        """
        求和列选择回调
        
        Args:
            selected_columns: 选择的列名列表
        """
        try:
            self.selected_columns_to_recalculate = selected_columns.copy()
            
            # 设置到数据处理器
            if self.selected_columns_to_recalculate:
                self.data_processor.set_columns_to_recalculate(self.selected_columns_to_recalculate)
                
                # 更新状态显示
                self._update_status(f"已选择 {len(selected_columns)} 列用于重新计算合计")
                
                logger.info(f"选择了 {len(selected_columns)} 列用于重新计算合计")
                logger.info(f"求和列: {', '.join(selected_columns)}")
            else:
                self._update_status("未选择求和列")
                logger.info("未选择求和列")
            
        except Exception as e:
            logger.error(f"求和列选择回调失败: {e}")
            messagebox.showerror("错误", f"求和列选择失败: {str(e)}")
    
    def _preview_data(self):
        """
        预览数据
        """
        try:
            if not self.current_worksheet:
                messagebox.showwarning("警告", "请先选择工作表")
                return
            
            # 读取预览数据
            preview_data = self.excel_reader.read_data_preview(self.current_worksheet, preview_rows=10)
            
            if preview_data.empty:
                messagebox.showwarning("警告", "工作表中没有数据")
                return
            
            # 创建预览窗口
            self._show_data_preview(preview_data)
            
        except Exception as e:
            logger.error(f"预览数据失败: {e}")
            messagebox.showerror("错误", f"预览数据失败: {str(e)}")
    
    def _show_data_preview(self, data):
        """
        显示数据预览窗口
        
        Args:
            data: 要预览的数据
        """
        # 创建预览窗口
        preview_window = tk.Toplevel(self.root)
        preview_window.title("数据预览")
        preview_window.geometry("900x600")
        
        # 主框架
        main_frame = ttk.Frame(preview_window, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 数据表格框架
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # 创建Treeview
        columns = list(data.columns)
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        
        # 设置列标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        
        # 添加数据
        for _, row in data.iterrows():
            tree.insert("", "end", values=list(row))
        
        # 添加滚动条
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # 布局表格
        tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # 统计信息框架
        stats_frame = ttk.LabelFrame(main_frame, text="数值列统计信息", padding="10")
        stats_frame.pack(fill="x", pady=(0, 10))
        
        # 计算数值列统计
        try:
            stats_result = self.data_processor.calculate_all_numeric_sums(data)
            
            if stats_result['success'] and stats_result['sums']:
                # 创建统计信息显示
                stats_text = tk.Text(stats_frame, height=6, wrap=tk.WORD, state="disabled")
                stats_scrollbar = ttk.Scrollbar(stats_frame, orient="vertical", command=stats_text.yview)
                stats_text.configure(yscrollcommand=stats_scrollbar.set)
                
                # 添加统计信息
                stats_text.config(state="normal")
                stats_text.insert(tk.END, f"共找到 {stats_result['total_numeric_columns']} 个数值列:\n\n")
                
                for col_name, col_stats in stats_result['sums'].items():
                    stats_text.insert(tk.END, f"【{col_name}】\n")
                    stats_text.insert(tk.END, f"  合计: {col_stats['formatted_sum']}\n")
                    stats_text.insert(tk.END, f"  有效数据: {col_stats['valid_count']}/{col_stats['total_count']} 行\n")
                    if col_stats['null_count'] > 0:
                        stats_text.insert(tk.END, f"  空值: {col_stats['null_count']} 行\n")
                    stats_text.insert(tk.END, "\n")
                
                stats_text.config(state="disabled")
                
                # 布局统计信息
                stats_text.pack(side="left", fill="both", expand=True)
                stats_scrollbar.pack(side="right", fill="y")
            else:
                # 没有数值列或计算失败
                no_stats_label = ttk.Label(stats_frame, text="未找到数值列或计算统计信息失败")
                no_stats_label.pack()
                
        except Exception as e:
            logger.error(f"计算预览统计信息失败: {e}")
            error_label = ttk.Label(stats_frame, text=f"计算统计信息时出错: {str(e)}")
            error_label.pack()
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        
        ttk.Button(button_frame, text="关闭", command=preview_window.destroy).pack(side="right")
    
    def _start_processing(self):
        """
        开始处理数据
        """
        try:
            # 检查是否有删除列的操作，如果有则需要确认
            if self.selected_columns_to_delete:
                # 确认对话框
                result = messagebox.askyesno(
                    "确认处理",
                    f"确定要删除选中的 {len(self.selected_columns_to_delete)} 列吗？\n\n"
                    f"删除的列: {', '.join(self.selected_columns_to_delete[:5])}"
                    f"{'...' if len(self.selected_columns_to_delete) > 5 else ''}"
                )
                
                if not result:
                    return
            
            # 开始处理数据
            self._process_data()
                
        except Exception as e:
            logger.error(f"开始处理失败: {e}")
            messagebox.showerror("错误", f"开始处理失败: {str(e)}")
    
    def _process_data(self):
        """
        处理数据
        """
        try:
            # 创建进度对话框
            self.progress_dialog = ProgressDialog(self.root, "正在处理数据...")
            
            # 在后台线程中处理数据
            import threading
            processing_thread = threading.Thread(target=self._do_processing)
            processing_thread.daemon = True
            processing_thread.start()
            
        except Exception as e:
            logger.error(f"处理数据失败: {e}")
            messagebox.showerror("错误", f"处理数据失败: {str(e)}")
    
    def _do_processing(self):
        """
        在后台线程中执行数据处理
        """
        try:
            if self.is_batch_mode:
                self._do_batch_processing()
            else:
                self._do_single_processing()
                
        except Exception as e:
            logger.error(f"数据处理失败: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"数据处理失败: {str(e)}"))
            self.root.after(0, lambda: self.progress_dialog.close())
    
    def _do_single_processing(self):
        """
        单文件处理
        """
        try:
            # 更新进度
            self.root.after(0, lambda: self.progress_dialog.update_progress(10, "读取数据..."))
            
            # 读取完整数据
            full_data = self.excel_reader.read_full_data(self.current_worksheet)
            if full_data.empty:
                self.root.after(0, lambda: messagebox.showerror("错误", "无法读取数据"))
                return
            
            self.root.after(0, lambda: self.progress_dialog.update_progress(30, "加载数据处理器..."))
            
            # 加载数据到处理器
            if not self.data_processor.load_data(full_data):
                self.root.after(0, lambda: messagebox.showerror("错误", "数据加载失败"))
                return
            
            self.root.after(0, lambda: self.progress_dialog.update_progress(50, "设置删除列..."))
            
            # 设置要删除的列
            if not self.data_processor.set_columns_to_delete(self.selected_columns_to_delete):
                self.root.after(0, lambda: messagebox.showerror("错误", "设置删除列失败"))
                return
            
            # 设置需要重新计算的求和列
            if self.selected_columns_to_recalculate:
                if not self.data_processor.set_columns_to_recalculate(self.selected_columns_to_recalculate):
                    logger.warning("设置求和列失败，将跳过合计行重新计算")
            
            # 检查是否启用跨工作表数据关联
            if self.cross_sheet_var.get():
                self.root.after(0, lambda: self.progress_dialog.update_progress(65, "执行跨工作表数据关联..."))
                
                try:
                    # 获取所有工作表数据
                    all_sheets = self.excel_reader.get_all_worksheets_data()
                    if all_sheets and self.data_processor.load_cross_sheet_data(all_sheets):
                        # 尝试自动识别发票基础信息表和明细表
                        invoice_sheet = self.current_worksheet  # 当前选择的工作表作为发票基础信息表
                        detail_sheet = None
                        
                        # 查找包含更多列的工作表作为明细表
                        for sheet_name in all_sheets.keys():
                            if sheet_name != invoice_sheet and len(all_sheets[sheet_name].columns) > len(all_sheets[invoice_sheet].columns):
                                detail_sheet = sheet_name
                                break
                        
                        if detail_sheet:
                            success = self.data_processor.process_cross_sheet_association(
                                invoice_sheet, detail_sheet
                            )
                            if not success:
                                logger.warning("跨工作表数据关联失败，将继续常规处理")
                        else:
                            logger.warning("未找到合适的明细表，跳过跨工作表关联")
                    else:
                        logger.warning("无法加载跨工作表数据，跳过关联处理")
                except Exception as cross_error:
                    logger.warning(f"跨工作表数据关联出错: {cross_error}，将继续常规处理")
            
            self.root.after(0, lambda: self.progress_dialog.update_progress(70, "处理数据..."))
            
            # 处理数据
            if not self.data_processor.process_data():
                self.root.after(0, lambda: messagebox.showerror("错误", "数据处理失败"))
                return
            
            self.root.after(0, lambda: self.progress_dialog.update_progress(90, "保存文件..."))
            
            # 保存处理后的数据
            self._save_processed_data()
            
            self.root.after(0, lambda: self.progress_dialog.update_progress(100, "处理完成"))
            
            # 显示结果
            self.root.after(0, lambda: self._show_processing_result())
            
        except Exception as e:
            error_msg = str(e)
            logger.error(f"单文件处理失败: {error_msg}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"单文件处理失败: {error_msg}"))
        finally:
            self.root.after(0, lambda: self.progress_dialog.close())
    
    def _do_batch_processing(self):
        """
        批量文件处理
        """
        try:
            total_files = len(self.batch_files)
            processed_files = 0
            failed_files = []
            batch_stats = []  # 存储每个文件的统计信息
            
            self.root.after(0, lambda: self.progress_dialog.update_progress(0, f"开始批量处理 {total_files} 个文件..."))
            
            for i, file_path in enumerate(self.batch_files):
                try:
                    file_name = os.path.basename(file_path)
                    progress = int((i / total_files) * 100)
                    self.root.after(0, lambda p=progress, f=file_name: self.progress_dialog.update_progress(p, f"处理文件: {f}"))
                    
                    # 创建临时读取器
                    temp_reader = ExcelReader()
                    if not temp_reader.load_file(file_path):
                        failed_files.append(f"{file_name}: 文件加载失败")
                        continue
                    
                    # 找到目标工作表
                    worksheets = temp_reader.get_worksheets_list()
                    target_worksheet = None
                    
                    for ws in worksheets:
                        if "发票基础信息" in ws['name']:
                            target_worksheet = ws['name']
                            break
                    
                    if not target_worksheet and worksheets:
                        target_worksheet = worksheets[0]['name']
                    
                    if not target_worksheet:
                        failed_files.append(f"{file_name}: 未找到工作表")
                        continue
                    
                    # 读取数据
                    full_data = temp_reader.read_full_data(target_worksheet)
                    if full_data.empty:
                        failed_files.append(f"{file_name}: 数据为空")
                        continue
                    
                    # 处理数据
                    temp_processor = DataProcessor()
                    if not temp_processor.load_data(full_data):
                        failed_files.append(f"{file_name}: 数据加载失败")
                        continue
                    
                    if not temp_processor.set_columns_to_delete(self.selected_columns_to_delete):
                        failed_files.append(f"{file_name}: 设置删除列失败")
                        continue
                    
                    if not temp_processor.process_data():
                        failed_files.append(f"{file_name}: 数据处理失败")
                        continue
                    
                    # 计算统计信息
                    processed_data = temp_processor.get_processed_data()
                    stats_result = temp_processor.calculate_all_numeric_sums(processed_data)
                    
                    file_stats = {
                        'file_name': file_name,
                        'file_path': file_path,
                        'stats': stats_result if stats_result['success'] else None
                    }
                    batch_stats.append(file_stats)
                    
                    # 保存文件（保留原始格式）
                    temp_handler = FileHandler()
                    temp_handler.set_original_file(file_path)
                    
                    # 生成输出文件路径
                    output_path = temp_handler.generate_output_filename(file_path)
                    
                    # 获取边框选项
                    add_border = self.border_var.get()
                    
                    if not temp_handler.save_excel_with_format_and_border(file_path, output_path, processed_data, target_worksheet, add_border):
                        failed_files.append(f"{file_name}: 文件保存失败")
                        continue
                    
                    processed_files += 1
                    
                except Exception as e:
                    failed_files.append(f"{file_name}: {str(e)}")
                    logger.error(f"处理文件 {file_path} 失败: {e}")
            
            # 显示批量处理结果
            self.root.after(0, lambda: self.progress_dialog.update_progress(100, "批量处理完成"))
            self.root.after(0, lambda: self._show_batch_processing_result(processed_files, total_files, failed_files, batch_stats))
            
        except Exception as e:
            logger.error(f"批量处理失败: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"批量处理失败: {str(e)}"))
        finally:
            self.root.after(0, lambda: self.progress_dialog.close())
    
    def _show_batch_processing_result(self, processed_files: int, total_files: int, failed_files: list):
        """
        显示批量处理结果
        """
        try:
            # 创建结果窗口
            result_window = tk.Toplevel(self.root)
            result_window.title("批量处理结果")
            result_window.geometry("600x400")
            result_window.resizable(True, True)
            
            # 居中显示
            result_window.transient(self.root)
            result_window.grab_set()
            
            # 主框架
            main_frame = ttk.Frame(result_window, padding="10")
            main_frame.pack(fill="both", expand=True)
            
            # 统计信息
            stats_frame = ttk.LabelFrame(main_frame, text="处理统计", padding="10")
            stats_frame.pack(fill="x", pady=(0, 10))
            
            ttk.Label(stats_frame, text=f"总文件数: {total_files}").pack(anchor="w")
            ttk.Label(stats_frame, text=f"成功处理: {processed_files}", foreground="green").pack(anchor="w")
            ttk.Label(stats_frame, text=f"处理失败: {len(failed_files)}", foreground="red").pack(anchor="w")
            ttk.Label(stats_frame, text=f"成功率: {processed_files/total_files*100:.1f}%").pack(anchor="w")
            
            # 失败文件列表
            if failed_files:
                failed_frame = ttk.LabelFrame(main_frame, text="失败文件详情", padding="10")
                failed_frame.pack(fill="both", expand=True, pady=(0, 10))
                
                # 创建文本框显示失败信息
                text_frame = ttk.Frame(failed_frame)
                text_frame.pack(fill="both", expand=True)
                
                failed_text = tk.Text(text_frame, wrap=tk.WORD, state="disabled")
                scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=failed_text.yview)
                failed_text.configure(yscrollcommand=scrollbar.set)
                
                failed_text.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")
                
                # 添加失败信息
                failed_text.config(state="normal")
                for failed_file in failed_files:
                    failed_text.insert(tk.END, f"• {failed_file}\n")
                failed_text.config(state="disabled")
            
            # 按钮框架
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x")
            
            ttk.Button(button_frame, text="确定", command=result_window.destroy).pack(side="right")
            
            # 更新状态
            self._update_status(f"批量处理完成: {processed_files}/{total_files} 个文件处理成功")
            
        except Exception as e:
            logger.error(f"显示批量处理结果失败: {e}")
            messagebox.showerror("错误", f"显示批量处理结果失败: {str(e)}")
    
    def _save_processed_data(self):
        """
        保存处理后的数据
        """
        try:
            # 获取处理后的数据
            processed_data = self.data_processor.get_processed_data()
            if processed_data is None or processed_data.empty:
                raise Exception("没有处理后的数据")
            
            # 设置文件处理器
            self.file_handler.set_original_file(self.current_file_path)
            
            # 生成输出文件名
            output_path = self.file_handler.generate_output_filename(self.current_file_path)
            
            # 获取边框选项
            add_border = self.border_var.get()
            
            # 保存文件（保留原始格式，可选边框）
            if self.file_handler.save_excel_with_format_and_border(self.current_file_path, output_path, processed_data, self.current_worksheet, add_border):
                self.output_file_path = output_path
                border_info = "（含边框）" if add_border else "（无边框）"
                logger.info(f"文件保存成功（保留格式）{border_info}: {output_path}")
            else:
                raise Exception("文件保存失败")
                
        except Exception as e:
            logger.error(f"保存处理后的数据失败: {e}")
            raise
    
    def _show_processing_result(self):
        """
        显示处理结果
        """
        try:
            # 获取处理摘要
            summary = self.data_processor.get_processing_summary()
            
            # 创建结果窗口
            result_window = tk.Toplevel(self.root)
            result_window.title("处理完成")
            result_window.geometry("600x600")
            result_window.resizable(True, True)
            
            # 主框架
            main_frame = ttk.Frame(result_window, padding="20")
            main_frame.pack(fill="both", expand=True)
            
            # 标题
            title_label = ttk.Label(main_frame, text="数据处理完成！", font=("Arial", 14, "bold"))
            title_label.pack(pady=(0, 20))
            
            # 处理结果信息
            info_frame = ttk.LabelFrame(main_frame, text="处理结果", padding="10")
            info_frame.pack(fill="x", pady=(0, 10))
            
            ttk.Label(info_frame, text=f"原始列数: {summary['original_columns_count']}").pack(anchor="w")
            ttk.Label(info_frame, text=f"删除列数: {summary['deleted_columns_count']}").pack(anchor="w")
            ttk.Label(info_frame, text=f"保留列数: {summary['remaining_columns_count']}").pack(anchor="w")
            ttk.Label(info_frame, text=f"数据行数: {summary['data_rows_count']}").pack(anchor="w")
            
            # 数值列统计信息
            stats_frame = ttk.LabelFrame(main_frame, text="数值列统计", padding="10")
            stats_frame.pack(fill="both", expand=True, pady=(0, 10))
            
            try:
                # 获取处理后的数据进行统计
                processed_data = self.data_processor.get_processed_data()
                stats_result = self.data_processor.calculate_all_numeric_sums(processed_data)
                
                if stats_result['success'] and stats_result['sums']:
                    # 创建统计信息显示
                    stats_text = tk.Text(stats_frame, height=8, wrap=tk.WORD, state="disabled")
                    stats_scrollbar = ttk.Scrollbar(stats_frame, orient="vertical", command=stats_text.yview)
                    stats_text.configure(yscrollcommand=stats_scrollbar.set)
                    
                    # 添加统计信息
                    stats_text.config(state="normal")
                    stats_text.insert(tk.END, f"共找到 {stats_result['total_numeric_columns']} 个数值列:\n\n")
                    
                    for col_name, col_stats in stats_result['sums'].items():
                        stats_text.insert(tk.END, f"【{col_name}】\n")
                        stats_text.insert(tk.END, f"  合计: {col_stats['formatted_sum']}\n")
                        stats_text.insert(tk.END, f"  有效数据: {col_stats['valid_count']}/{col_stats['total_count']} 行\n")
                        if col_stats['null_count'] > 0:
                            stats_text.insert(tk.END, f"  空值: {col_stats['null_count']} 行\n")
                        stats_text.insert(tk.END, "\n")
                    
                    stats_text.config(state="disabled")
                    
                    # 布局统计信息
                    stats_text.pack(side="left", fill="both", expand=True)
                    stats_scrollbar.pack(side="right", fill="y")
                else:
                    # 没有数值列或计算失败
                    no_stats_label = ttk.Label(stats_frame, text="未找到数值列或计算统计信息失败")
                    no_stats_label.pack()
                    
            except Exception as e:
                logger.error(f"计算结果统计信息失败: {e}")
                error_label = ttk.Label(stats_frame, text=f"计算统计信息时出错: {str(e)}")
                error_label.pack()
            
            # 输出文件信息
            file_frame = ttk.LabelFrame(main_frame, text="输出文件", padding="10")
            file_frame.pack(fill="x", pady=(0, 10))
            
            file_path_label = ttk.Label(file_frame, text=self.output_file_path, wraplength=450)
            file_path_label.pack(anchor="w")
            
            # 操作按钮
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x")
            
            ttk.Button(
                button_frame,
                text="打开文件",
                command=lambda: self.file_handler.open_file(self.output_file_path)
            ).pack(side="left", padx=(0, 10))
            
            ttk.Button(
                button_frame,
                text="打开文件夹",
                command=lambda: self.file_handler.open_file_location(self.output_file_path)
            ).pack(side="left", padx=(0, 10))
            
            ttk.Button(
                button_frame,
                text="关闭",
                command=result_window.destroy
            ).pack(side="right")
            
            # 居中显示
            result_window.transient(self.root)
            result_window.grab_set()
            
            self._update_status("数据处理完成")
            
        except Exception as e:
            logger.error(f"显示处理结果失败: {e}")
    
    def _open_settings(self):
        """
        打开设置窗口
        """
        # TODO: 实现设置窗口
        messagebox.showinfo("提示", "设置功能正在开发中...")
    
    def _update_status(self, message: str):
        """
        更新状态栏
        
        Args:
            message: 状态消息
        """
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def _on_closing(self):
        """
        窗口关闭事件处理
        """
        try:
            # 清理资源
            self.excel_reader.close()
            
            # 关闭窗口
            self.root.destroy()
            
            logger.info("应用程序正常退出")
            
        except Exception as e:
            logger.error(f"关闭应用程序失败: {e}")
            self.root.destroy()