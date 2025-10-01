# -*- coding: utf-8 -*-
"""
进度对话框
显示处理进度和状态信息
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable
from src.utils.logger import get_logger

logger = get_logger("ProgressDialog")

class ProgressDialog:
    """
    进度对话框类
    显示处理进度、状态信息和取消功能
    """
    
    def __init__(self, parent: tk.Tk, title: str = "处理中...", 
                 cancelable: bool = True, on_cancel: Optional[Callable] = None):
        """
        初始化进度对话框
        
        Args:
            parent: 父窗口
            title: 对话框标题
            cancelable: 是否可取消
            on_cancel: 取消回调函数
        """
        self.parent = parent
        self.title = title
        self.cancelable = cancelable
        self.on_cancel = on_cancel
        self.is_cancelled = False
        self.is_closed = False
        
        # 创建对话框窗口
        self.dialog = None
        self.progress_var = None
        self.status_var = None
        self.progress_bar = None
        self.status_label = None
        self.cancel_button = None
        self.detail_text = None
        
        self._create_dialog()
        
    def _create_dialog(self):
        """
        创建对话框界面
        """
        try:
            # 创建顶级窗口
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title(self.title)
            self.dialog.geometry("400x200")
            self.dialog.resizable(False, False)
            
            # 设置为模态对话框
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            
            # 居中显示
            self._center_dialog()
            
            # 创建主框架
            main_frame = ttk.Frame(self.dialog, padding="20")
            main_frame.pack(fill="both", expand=True)
            
            # 状态标签
            self.status_var = tk.StringVar(value="正在初始化...")
            self.status_label = ttk.Label(
                main_frame, 
                textvariable=self.status_var,
                font=("Arial", 10)
            )
            self.status_label.pack(pady=(0, 10))
            
            # 进度条
            self.progress_var = tk.DoubleVar()
            self.progress_bar = ttk.Progressbar(
                main_frame,
                variable=self.progress_var,
                maximum=100,
                length=300,
                mode='determinate'
            )
            self.progress_bar.pack(pady=(0, 10))
            
            # 进度百分比标签
            self.percent_var = tk.StringVar(value="0%")
            self.percent_label = ttk.Label(
                main_frame,
                textvariable=self.percent_var,
                font=("Arial", 9)
            )
            self.percent_label.pack(pady=(0, 10))
            
            # 详细信息文本框（可选）
            self.detail_frame = ttk.LabelFrame(main_frame, text="详细信息", padding="5")
            self.detail_text = tk.Text(
                self.detail_frame,
                height=4,
                width=40,
                wrap=tk.WORD,
                state="disabled",
                font=("Arial", 8)
            )
            self.detail_scrollbar = ttk.Scrollbar(
                self.detail_frame,
                orient="vertical",
                command=self.detail_text.yview
            )
            self.detail_text.configure(yscrollcommand=self.detail_scrollbar.set)
            
            # 默认隐藏详细信息
            self.detail_visible = False
            
            # 按钮框架
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x", pady=(10, 0))
            
            # 显示/隐藏详细信息按钮
            self.detail_button = ttk.Button(
                button_frame,
                text="显示详细信息",
                command=self._toggle_detail
            )
            self.detail_button.pack(side="left")
            
            # 取消按钮
            if self.cancelable:
                self.cancel_button = ttk.Button(
                    button_frame,
                    text="取消",
                    command=self._on_cancel_clicked
                )
                self.cancel_button.pack(side="right")
            
            # 绑定关闭事件
            self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
            
            logger.info("进度对话框创建完成")
            
        except Exception as e:
            logger.error(f"创建进度对话框失败: {e}")
    
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
    
    def _toggle_detail(self):
        """
        切换详细信息显示状态
        """
        try:
            if self.detail_visible:
                # 隐藏详细信息
                self.detail_frame.pack_forget()
                self.detail_button.config(text="显示详细信息")
                self.dialog.geometry("400x200")
                self.detail_visible = False
            else:
                # 显示详细信息
                self.detail_frame.pack(fill="both", expand=True, pady=(10, 0))
                self.detail_text.pack(side="left", fill="both", expand=True)
                self.detail_scrollbar.pack(side="right", fill="y")
                self.detail_button.config(text="隐藏详细信息")
                self.dialog.geometry("400x350")
                self.detail_visible = True
            
            # 重新居中
            self._center_dialog()
            
        except Exception as e:
            logger.error(f"切换详细信息显示失败: {e}")
    
    def _on_cancel_clicked(self):
        """
        取消按钮点击事件
        """
        try:
            self.is_cancelled = True
            
            if self.on_cancel:
                self.on_cancel()
            
            self.update_status("正在取消...")
            
            if self.cancel_button:
                self.cancel_button.config(state="disabled")
            
            logger.info("用户取消了操作")
            
        except Exception as e:
            logger.error(f"处理取消事件失败: {e}")
    
    def _on_close(self):
        """
        窗口关闭事件
        """
        if self.cancelable and not self.is_cancelled:
            self._on_cancel_clicked()
        else:
            self.close()
    
    def update_progress(self, progress: float, status: str = None):
        """
        更新进度
        
        Args:
            progress: 进度值（0-100）
            status: 状态信息
        """
        try:
            if self.is_closed:
                return
            
            # 更新进度条
            if self.progress_var:
                self.progress_var.set(max(0, min(100, progress)))
            
            # 更新百分比显示
            if self.percent_var:
                self.percent_var.set(f"{progress:.1f}%")
            
            # 更新状态信息
            if status and self.status_var:
                self.status_var.set(status)
            
            # 刷新界面
            if self.dialog:
                self.dialog.update_idletasks()
            
        except Exception as e:
            logger.error(f"更新进度失败: {e}")
    
    def update_status(self, status: str):
        """
        更新状态信息
        
        Args:
            status: 状态信息
        """
        try:
            if self.is_closed:
                return
            
            if self.status_var:
                self.status_var.set(status)
            
            # 刷新界面
            if self.dialog:
                self.dialog.update_idletasks()
            
        except Exception as e:
            logger.error(f"更新状态失败: {e}")
    
    def add_detail(self, message: str):
        """
        添加详细信息
        
        Args:
            message: 详细信息
        """
        try:
            if self.is_closed or not self.detail_text:
                return
            
            # 启用文本框
            self.detail_text.config(state="normal")
            
            # 添加消息
            self.detail_text.insert(tk.END, f"{message}\n")
            
            # 滚动到底部
            self.detail_text.see(tk.END)
            
            # 禁用文本框
            self.detail_text.config(state="disabled")
            
            # 刷新界面
            if self.dialog:
                self.dialog.update_idletasks()
            
        except Exception as e:
            logger.error(f"添加详细信息失败: {e}")
    
    def set_indeterminate(self, indeterminate: bool = True):
        """
        设置进度条为不确定模式
        
        Args:
            indeterminate: 是否为不确定模式
        """
        try:
            if self.is_closed or not self.progress_bar:
                return
            
            if indeterminate:
                self.progress_bar.config(mode='indeterminate')
                self.progress_bar.start()
                if self.percent_var:
                    self.percent_var.set("处理中...")
            else:
                self.progress_bar.stop()
                self.progress_bar.config(mode='determinate')
            
        except Exception as e:
            logger.error(f"设置进度条模式失败: {e}")
    
    def close(self):
        """
        关闭对话框
        """
        try:
            if self.is_closed:
                return
            
            self.is_closed = True
            
            if self.dialog:
                self.dialog.grab_release()
                self.dialog.destroy()
                self.dialog = None
            
            logger.info("进度对话框已关闭")
            
        except Exception as e:
            logger.error(f"关闭进度对话框失败: {e}")
    
    def is_canceled(self) -> bool:
        """
        检查是否已取消
        
        Returns:
            bool: 是否已取消
        """
        return self.is_cancelled
    
    def show(self):
        """
        显示对话框
        """
        try:
            if self.dialog and not self.is_closed:
                self.dialog.deiconify()
                self.dialog.lift()
                self.dialog.focus_set()
            
        except Exception as e:
            logger.error(f"显示进度对话框失败: {e}")
    
    def hide(self):
        """
        隐藏对话框
        """
        try:
            if self.dialog and not self.is_closed:
                self.dialog.withdraw()
            
        except Exception as e:
            logger.error(f"隐藏进度对话框失败: {e}")

# 简化的进度对话框函数
def show_progress_dialog(parent: tk.Tk, title: str = "处理中...", 
                        cancelable: bool = True, on_cancel: Optional[Callable] = None) -> ProgressDialog:
    """
    显示进度对话框
    
    Args:
        parent: 父窗口
        title: 对话框标题
        cancelable: 是否可取消
        on_cancel: 取消回调函数
        
    Returns:
        ProgressDialog: 进度对话框实例
    """
    return ProgressDialog(parent, title, cancelable, on_cancel)

# 上下文管理器版本
class ProgressContext:
    """
    进度对话框上下文管理器
    """
    
    def __init__(self, parent: tk.Tk, title: str = "处理中...", 
                 cancelable: bool = True, on_cancel: Optional[Callable] = None):
        self.dialog = ProgressDialog(parent, title, cancelable, on_cancel)
    
    def __enter__(self) -> ProgressDialog:
        return self.dialog
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.dialog.close()
        return False