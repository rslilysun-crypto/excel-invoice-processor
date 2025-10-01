# -*- coding: utf-8 -*-
"""
日志记录器
提供统一的日志记录功能
"""

import logging
import os
from datetime import datetime
from typing import Optional

# 全局日志配置
LOG_LEVEL = logging.INFO
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# 日志文件配置
LOG_DIR = 'logs'
LOG_FILE_PREFIX = 'excel_processor'

def setup_logger(log_level: int = LOG_LEVEL, log_to_file: bool = True) -> logging.Logger:
    """
    设置主日志记录器
    
    Args:
        log_level: 日志级别
        log_to_file: 是否记录到文件
        
    Returns:
        logging.Logger: 配置好的日志记录器
    """
    # 创建根日志记录器
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    
    # 清除现有的处理器
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    # 创建格式化器
    formatter = logging.Formatter(LOG_FORMAT, DATE_FORMAT)
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)
    
    # 文件处理器
    if log_to_file:
        try:
            # 确保日志目录存在
            if not os.path.exists(LOG_DIR):
                os.makedirs(LOG_DIR)
            
            # 生成日志文件名
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_filename = f"{LOG_FILE_PREFIX}_{timestamp}.log"
            log_filepath = os.path.join(LOG_DIR, log_filename)
            
            # 创建文件处理器
            file_handler = logging.FileHandler(log_filepath, encoding='utf-8')
            file_handler.setLevel(log_level)
            file_handler.setFormatter(formatter)
            root_logger.addHandler(file_handler)
            
            root_logger.info(f"日志文件: {log_filepath}")
            
        except Exception as e:
            root_logger.error(f"创建日志文件失败: {e}")
    
    root_logger.info("日志系统初始化完成")
    return root_logger

def get_logger(name: str) -> logging.Logger:
    """
    获取指定名称的日志记录器
    
    Args:
        name: 日志记录器名称
        
    Returns:
        logging.Logger: 日志记录器
    """
    return logging.getLogger(name)

def set_log_level(level: int):
    """
    设置全局日志级别
    
    Args:
        level: 日志级别
    """
    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    
    for handler in root_logger.handlers:
        handler.setLevel(level)

def log_exception(logger: logging.Logger, message: str, exception: Exception):
    """
    记录异常信息
    
    Args:
        logger: 日志记录器
        message: 错误消息
        exception: 异常对象
    """
    logger.error(f"{message}: {str(exception)}", exc_info=True)

def log_function_call(logger: logging.Logger, func_name: str, *args, **kwargs):
    """
    记录函数调用信息
    
    Args:
        logger: 日志记录器
        func_name: 函数名称
        *args: 位置参数
        **kwargs: 关键字参数
    """
    args_str = ', '.join([str(arg) for arg in args])
    kwargs_str = ', '.join([f"{k}={v}" for k, v in kwargs.items()])
    
    params = []
    if args_str:
        params.append(args_str)
    if kwargs_str:
        params.append(kwargs_str)
    
    params_str = ', '.join(params)
    logger.debug(f"调用函数: {func_name}({params_str})")

def log_performance(logger: logging.Logger, operation: str, duration: float):
    """
    记录性能信息
    
    Args:
        logger: 日志记录器
        operation: 操作名称
        duration: 耗时（秒）
    """
    logger.info(f"性能统计: {operation} 耗时 {duration:.2f} 秒")

def cleanup_old_logs(days_to_keep: int = 30):
    """
    清理旧的日志文件
    
    Args:
        days_to_keep: 保留天数
    """
    try:
        if not os.path.exists(LOG_DIR):
            return
        
        current_time = datetime.now()
        
        for filename in os.listdir(LOG_DIR):
            if filename.startswith(LOG_FILE_PREFIX) and filename.endswith('.log'):
                filepath = os.path.join(LOG_DIR, filename)
                
                # 获取文件修改时间
                file_time = datetime.fromtimestamp(os.path.getmtime(filepath))
                
                # 计算文件年龄
                age_days = (current_time - file_time).days
                
                # 删除过期文件
                if age_days > days_to_keep:
                    os.remove(filepath)
                    print(f"删除过期日志文件: {filename}")
    
    except Exception as e:
        print(f"清理日志文件失败: {e}")

class LoggerMixin:
    """
    日志记录器混入类
    为类提供日志记录功能
    """
    
    @property
    def logger(self) -> logging.Logger:
        """
        获取当前类的日志记录器
        
        Returns:
            logging.Logger: 日志记录器
        """
        if not hasattr(self, '_logger'):
            self._logger = get_logger(self.__class__.__name__)
        return self._logger
    
    def log_info(self, message: str):
        """记录信息日志"""
        self.logger.info(message)
    
    def log_warning(self, message: str):
        """记录警告日志"""
        self.logger.warning(message)
    
    def log_error(self, message: str, exception: Optional[Exception] = None):
        """记录错误日志"""
        if exception:
            log_exception(self.logger, message, exception)
        else:
            self.logger.error(message)
    
    def log_debug(self, message: str):
        """记录调试日志"""
        self.logger.debug(message)

def create_performance_logger(name: str) -> logging.Logger:
    """
    创建性能日志记录器
    
    Args:
        name: 记录器名称
        
    Returns:
        logging.Logger: 性能日志记录器
    """
    perf_logger = logging.getLogger(f"performance.{name}")
    
    if not perf_logger.handlers:
        # 创建性能日志文件处理器
        try:
            if not os.path.exists(LOG_DIR):
                os.makedirs(LOG_DIR)
            
            timestamp = datetime.now().strftime('%Y%m%d')
            perf_log_file = os.path.join(LOG_DIR, f"performance_{timestamp}.log")
            
            handler = logging.FileHandler(perf_log_file, encoding='utf-8')
            formatter = logging.Formatter('%(asctime)s - %(message)s', DATE_FORMAT)
            handler.setFormatter(formatter)
            
            perf_logger.addHandler(handler)
            perf_logger.setLevel(logging.INFO)
            
        except Exception as e:
            print(f"创建性能日志记录器失败: {e}")
    
    return perf_logger

# 装饰器：自动记录函数执行时间
def log_execution_time(logger_name: str = None):
    """
    装饰器：记录函数执行时间
    
    Args:
        logger_name: 日志记录器名称
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            import time
            
            # 获取日志记录器
            if logger_name:
                logger = get_logger(logger_name)
            else:
                logger = get_logger(func.__module__)
            
            # 记录开始时间
            start_time = time.time()
            
            try:
                # 执行函数
                result = func(*args, **kwargs)
                
                # 记录执行时间
                duration = time.time() - start_time
                log_performance(logger, func.__name__, duration)
                
                return result
                
            except Exception as e:
                # 记录异常和执行时间
                duration = time.time() - start_time
                log_exception(logger, f"函数 {func.__name__} 执行失败（耗时 {duration:.2f} 秒）", e)
                raise
        
        return wrapper
    return decorator

# 上下文管理器：记录代码块执行时间
class LogExecutionTime:
    """
    上下文管理器：记录代码块执行时间
    """
    
    def __init__(self, logger: logging.Logger, operation_name: str):
        self.logger = logger
        self.operation_name = operation_name
        self.start_time = None
    
    def __enter__(self):
        import time
        self.start_time = time.time()
        self.logger.info(f"开始执行: {self.operation_name}")
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        import time
        duration = time.time() - self.start_time
        
        if exc_type is None:
            log_performance(self.logger, self.operation_name, duration)
        else:
            self.logger.error(f"执行失败: {self.operation_name}（耗时 {duration:.2f} 秒）")
        
        return False