#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel发票数据处理软件 - Streamlit Web版本
基于原有Tkinter版本改造，保持核心功能不变
"""

import streamlit as st
import pandas as pd
import io
import sys
import os
from pathlib import Path
from typing import Optional, List, Dict, Any

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# 导入现有的核心模块（复用原有逻辑）
from src.core.excel_reader import ExcelReader
from src.core.data_processor import DataProcessor
from src.core.file_handler import FileHandler
from src.utils.config import ConfigManager
from src.utils.logger import setup_logger, get_logger

# 设置页面配置
st.set_page_config(
    page_title="Excel发票数据处理软件",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 添加自定义CSS样式
st.markdown("""
<style>
/* 强制隐藏Streamlit顶部工具栏 */
.stApp > header {
    display: none !important;
}

/* 强制隐藏Streamlit顶部菜单 */
.stApp > div[data-testid="stToolbar"] {
    display: none !important;
}

/* 强制设置主容器的顶部间距 */
.stApp > div[data-testid="stAppViewContainer"] > .main {
    padding-top: 0 !important;
    margin-top: 0 !important;
}

/* 强制设置主内容区域的顶部间距 */
.stApp > div[data-testid="stAppViewContainer"] > .main > .block-container {
    padding-top: 1rem !important;
    margin-top: 0 !important;
}

/* 确保页面有足够的顶部空间，整体下移 */
.main > div {
    padding-top: 0.5rem !important;
}

/* 确保标题有足够的顶部间距 */
h1 {
    margin-top: 1rem !important;
    margin-bottom: 0.5rem !important;
}

/* 减少分隔线间距 */
hr {
    margin-top: 0.5rem;
    margin-bottom: 1rem;
}

/* 优化侧边栏顶部间距 */
.css-1d391kg {
    padding-top: 1rem;
}



/* 减少expander间距 */
div[data-testid="stExpander"] {
    margin-bottom: 0.3rem;
}

/* 进一步优化侧边栏expander间距 */
.css-1d391kg div[data-testid="stExpander"] {
    margin-bottom: 0.25rem !important;
    margin-top: 0 !important;
}

/* 减少侧边栏标题间距 */
.css-1d391kg h4 {
    margin-top: 0;
    margin-bottom: 0.5rem;
}

/* 减少markdown标题间距 */
.markdown-text-container h3 {
    margin-top: 0;
    margin-bottom: 0.5rem;
}

/* 优化整体页面布局 */
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 0.8rem !important;
}

/* 优化subheader间距 */
h3 {
    margin-top: 0;
    margin-bottom: 0.5rem;
}

/* 减少列布局间距 */
div[data-testid="column"] {
    padding-top: 0;
}

/* 优化数据预览区域 */
div[data-testid="stDataFrame"] {
    max-height: 400px;
    overflow-y: auto;
}

/* 强制设置所有标题的顶部空间 */
.main h1, .main h2, .main h3 {
    margin-top: 0.5rem !important;
    padding-top: 0.5rem !important;
}

/* 缩短文件上传区域高度 */
div[data-testid="stFileUploader"] {
    padding: 0.5rem !important;
    min-height: auto !important;
}

/* 缩短文件上传拖拽区域 */
div[data-testid="stFileUploader"] > div {
    padding: 0.5rem !important;
    min-height: 60px !important;
}

/* 优化文件上传按钮区域 */
div[data-testid="stFileUploader"] button {
    margin: 0.25rem 0 !important;
}

/* 减少文件上传区域的内边距 */
div[data-testid="stFileUploader"] > div > div {
    padding: 0.5rem !important;
}

/* 优化已上传文件显示区域 */
div[data-testid="stFileUploader"] div[data-testid="fileUploadedFile"] {
    padding: 0.15rem 0.5rem !important;
    margin: 0.1rem 0 !important;
    min-height: auto !important;
}

/* 缩小已上传文件的图标和文本间距 */
div[data-testid="fileUploadedFile"] > div {
    padding: 0.25rem !important;
    gap: 0.5rem !important;
}

/* 优化文件删除按钮 */
div[data-testid="fileUploadedFile"] button {
    padding: 0.25rem !important;
    min-height: auto !important;
    height: auto !important;
}

/* 减少已上传文件列表的整体间距 */
div[data-testid="stFileUploader"] > div:last-child {
    margin-top: 0.1rem !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
}

/* 进一步优化文件列表容器 */
div[data-testid="stFileUploader"] > div {
    gap: 0.1rem !important;
}

/* 优化文件上传区域和文件列表之间的间距 */
div[data-testid="stFileUploader"] > div:first-child {
    margin-bottom: 0.1rem !important;
}

/* 优化文件信息文本显示 */
div[data-testid="fileUploadedFile"] span {
    line-height: 1.2 !important;
    font-size: 0.9rem !important;
}

/* 优化expander内部内容间距 */
div[data-testid="stExpander"] > div > div {
    padding-top: 0.5rem !important;
    padding-bottom: 0.5rem !important;
}

/* 优化侧边栏内容间距 */
.css-1d391kg .stSelectbox > div {
    margin-bottom: 0.5rem !important;
}

.css-1d391kg .stCheckbox {
    margin-bottom: 0.3rem !important;
}

.css-1d391kg .stButton > button {
    margin-top: 0.5rem !important;
    margin-bottom: 0.3rem !important;
}

/* 优化下拉框和其他表单元素间距 */
div[data-testid="stSelectbox"] {
    margin-bottom: 0.5rem !important;
}

/* 减少文本元素间距 */
.css-1d391kg p, .css-1d391kg div {
    margin-bottom: 0.3rem !important;
}

/* 优化主内容区域间距 */
.main .block-container > div {
    margin-bottom: 0.5rem !important;
}

/* 减少标题下方间距 */
.main h3 + hr {
    margin-top: 0.3rem !important;
    margin-bottom: 0.8rem !important;
}

/* 优化数据表格显示区域 */
.main div[data-testid="stDataFrame"] {
    margin-top: 0.5rem !important;
    margin-bottom: 0.5rem !important;
}

/* 减少列布局的间距 */
.main div[data-testid="column"] > div {
    padding-top: 0.3rem !important;
}

/* 优化主内容区域按钮间距 */
.main div[data-testid="stButton"] {
    margin-top: 0.8rem !important;
    margin-bottom: 0.5rem !important;
}

.main div[data-testid="stButton"] > button {
    padding: 0.4rem 1rem !important;
}

/* 优化复选框间距 */
.main div[data-testid="stCheckbox"] {
    margin-bottom: 0.3rem !important;
    margin-top: 0.3rem !important;
}

/* 优化表单元素组合间距 */
.main div[data-testid="stSelectbox"] + div[data-testid="stSelectbox"] {
    margin-top: 0.3rem !important;
}

/* 减少成功/错误消息间距 */
.main div[data-testid="stAlert"] {
    margin-top: 0.5rem !important;
    margin-bottom: 0.5rem !important;
}
</style>
""", unsafe_allow_html=True)

# 初始化日志
logger = setup_logger()
app_logger = get_logger("StreamlitApp")

# 初始化session state
def init_session_state():
    """初始化session state变量"""
    if 'excel_reader' not in st.session_state:
        st.session_state.excel_reader = ExcelReader()
    if 'data_processor' not in st.session_state:
        st.session_state.data_processor = DataProcessor()
    if 'file_handler' not in st.session_state:
        st.session_state.file_handler = FileHandler()
    if 'config_manager' not in st.session_state:
        st.session_state.config_manager = ConfigManager()
    
    # 状态变量
    if 'current_file_path' not in st.session_state:
        st.session_state.current_file_path = None
    if 'current_worksheet' not in st.session_state:
        st.session_state.current_worksheet = None
    if 'selected_columns_to_delete' not in st.session_state:
        st.session_state.selected_columns_to_delete = []
    if 'selected_columns_to_recalculate' not in st.session_state:
        st.session_state.selected_columns_to_recalculate = []
    if 'current_data' not in st.session_state:
        st.session_state.current_data = None
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'show_column_selector' not in st.session_state:
        st.session_state.show_column_selector = False

def main():
    """主函数"""
    app_logger.info("Excel发票数据处理软件 - Streamlit版本启动")
    
    # 初始化session state
    init_session_state()
    
    # 页面标题 - 紧凑显示
    st.markdown("### 📊 Excel发票数据处理软件")
    st.markdown("---")
    
    # 侧边栏
    with st.sidebar:
        st.markdown("#### 🔧 操作面板")
        
        # 文件上传 - 始终显示
        with st.expander("📁 文件上传", expanded=True):
            uploaded_file = st.file_uploader(
                "选择Excel文件",
                type=['xlsx', 'xls'],
                help="支持.xlsx和.xls格式的Excel文件",
                label_visibility="collapsed"
            )
            
            # 处理文件上传
            if uploaded_file is not None:
                handle_file_upload(uploaded_file)
        
        # 工作表选择
        if st.session_state.current_file_path:
            with st.expander("📋 工作表选择", expanded=True):
                handle_worksheet_selection()
        
        # 列选择
        if st.session_state.current_worksheet:
            with st.expander("🎯 列选择", expanded=True):
                handle_column_selection()
        
        # 处理选项和执行
        if st.session_state.current_data is not None:
            with st.expander("⚙️ 处理选项", expanded=True):
                handle_processing_options()
    
    # 处理列选择界面 - 优先显示
    if st.session_state.get('show_column_selector', False):
        show_column_selector_interface()
    else:
        # 主内容区域
        if st.session_state.current_file_path:
            display_main_content()
        else:
            display_welcome_message()

def handle_file_upload(uploaded_file):
    """处理文件上传"""
    try:
        # 保存上传的文件到临时位置
        temp_file_path = f"temp_{uploaded_file.name}"
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # 使用现有的ExcelReader加载文件
        if st.session_state.excel_reader.load_file(temp_file_path):
            st.session_state.current_file_path = temp_file_path
            st.session_state.current_worksheet = None
            st.session_state.current_data = None
            st.session_state.processed_data = None
            
            st.success(f"✅ {uploaded_file.name}")
            app_logger.info(f"文件上传成功: {uploaded_file.name}")
        else:
            st.error("❌ 文件加载失败")
            
    except Exception as e:
        st.error(f"❌ 上传失败")
        app_logger.error(f"文件上传失败: {e}")

def handle_worksheet_selection():
    """处理工作表选择"""
    # 获取工作表列表
    worksheets = st.session_state.excel_reader.get_worksheets_list()
    
    if worksheets:
        worksheet_names = [ws['name'] for ws in worksheets]
        
        # 确定默认选择的索引
        default_index = 0
        if st.session_state.current_worksheet is None:
            # 如果没有当前选择，优先选择"发票基础信息"
            if "发票基础信息" in worksheet_names:
                default_index = worksheet_names.index("发票基础信息")
            else:
                default_index = 0
        else:
            # 如果有当前选择，保持当前选择
            default_index = worksheet_names.index(st.session_state.current_worksheet) if st.session_state.current_worksheet in worksheet_names else 0
        
        selected_worksheet = st.selectbox(
            "选择工作表",
            worksheet_names,
            index=default_index
        )
        
        if selected_worksheet != st.session_state.current_worksheet:
            st.session_state.current_worksheet = selected_worksheet
            load_worksheet_data(selected_worksheet)
        
        # 显示工作表信息
        worksheet_info = next((ws for ws in worksheets if ws['name'] == selected_worksheet), None)
        if worksheet_info:
            st.sidebar.info(f"📊 行数: {worksheet_info['rows']}\n📊 列数: {worksheet_info['columns']}")

def load_worksheet_data(worksheet_name):
    """加载工作表数据"""
    try:
        data = st.session_state.excel_reader.read_full_data(worksheet_name)
        if data is not None and not data.empty:
            st.session_state.current_data = data
            st.session_state.data_processor.load_data(data)
            st.session_state.processed_data = None
            app_logger.info(f"工作表数据加载成功: {worksheet_name}")
        else:
            st.sidebar.error("❌ 工作表数据加载失败")
    except Exception as e:
        st.sidebar.error(f"❌ 加载工作表失败: {str(e)}")
        app_logger.error(f"加载工作表失败: {e}")

def handle_column_selection():
    """处理列选择"""
    if st.session_state.current_data is None:
        return
    
    columns = st.session_state.current_data.columns.tolist()
    current_selected = st.session_state.selected_columns_to_delete
    
    # 删除列选择 - 单行紧凑显示
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col1:
        st.markdown("**删除列:**")
    with col2:
        if current_selected:
            st.caption(f"✅ 已选 {len(current_selected)} 列")
        else:
            st.caption("❌ 未选择")
    with col3:
        if st.button("选择", key="select_del_cols", use_container_width=True):
            st.session_state.show_column_selector = True
            st.rerun()
    
    # 求和列选择 - 与标题在同一行
    col1, col2 = st.columns([1, 3])
    with col1:
        st.markdown("**求和列:**")
    with col2:
        remaining_columns = [col for col in columns if col not in current_selected]
        selected_sum_columns = st.multiselect(
            "",
            remaining_columns,
            default=st.session_state.selected_columns_to_recalculate,
            key="sum_columns",
            placeholder="选择需要重新计算求和的列",
            label_visibility="collapsed"
        )
        st.session_state.selected_columns_to_recalculate = selected_sum_columns

def handle_processing_options():
    """处理处理选项"""
    # 选项布局
    col1, col2 = st.columns(2)
    with col1:
        add_border = st.checkbox("添加边框", value=True)
    with col2:
        enable_cross_sheet = st.checkbox("跨表关联", value=False)
    
    # 处理按钮
    if st.button("🚀 开始处理", type="primary", use_container_width=True):
        process_data(add_border, enable_cross_sheet)

def process_data(add_border=True, enable_cross_sheet=False):
    """处理数据"""
    try:
        with st.spinner("正在处理数据..."):
            # 跨工作表数据关联处理
            if enable_cross_sheet:
                st.info("🔗 正在执行跨工作表数据关联...")
                try:
                    # 获取所有工作表数据
                    all_sheets = st.session_state.excel_reader.get_all_worksheets_data()
                    
                    if all_sheets and st.session_state.data_processor.load_cross_sheet_data(all_sheets):
                        # 查找发票基础信息表和信息汇总表
                        invoice_sheet = None
                        summary_sheet = None
                        
                        for sheet_name in all_sheets.keys():
                            if "发票基础信息" in sheet_name or "基础信息" in sheet_name:
                                invoice_sheet = sheet_name
                            elif "信息汇总表" in sheet_name or "汇总表" in sheet_name or "信息汇总" in sheet_name:
                                summary_sheet = sheet_name
                        
                        if invoice_sheet and summary_sheet:
                            st.info(f"📋 找到发票基础信息表: {invoice_sheet}")
                            st.info(f"📋 找到信息汇总表: {summary_sheet}")
                            
                            success = st.session_state.data_processor.process_cross_sheet_association(
                                invoice_sheet, summary_sheet
                            )
                            if success:
                                st.success("✅ 跨工作表数据关联完成！已将货物名称添加到发票基础信息表")
                                # 更新当前数据为关联后的数据
                                updated_data = st.session_state.data_processor.get_original_data()
                                if updated_data is not None:
                                    st.session_state.current_data = updated_data
                                    st.session_state.data_processor.load_data(updated_data)
                                    st.info("📊 已更新当前数据，包含关联的货物名称")
                            else:
                                st.warning("⚠️ 跨工作表数据关联失败，将继续常规处理")
                        else:
                            missing_sheets = []
                            if not invoice_sheet:
                                missing_sheets.append("发票基础信息表")
                            if not summary_sheet:
                                missing_sheets.append("信息汇总表")
                            st.warning(f"⚠️ 未找到必要的工作表: {', '.join(missing_sheets)}，跳过跨工作表关联")
                            st.info("💡 提示：请确保Excel文件包含'发票基础信息'和'信息汇总表'工作表")
                    else:
                        st.warning("⚠️ 无法加载跨工作表数据，跳过关联处理")
                        
                except Exception as cross_error:
                    st.warning(f"⚠️ 跨工作表数据关联出错: {cross_error}，将继续常规处理")
                    app_logger.warning(f"跨工作表数据关联出错: {cross_error}")
            
            # 设置要删除的列
            st.session_state.data_processor.set_columns_to_delete(
                st.session_state.selected_columns_to_delete
            )
            
            # 设置求和列
            st.session_state.data_processor.set_columns_to_recalculate(
                st.session_state.selected_columns_to_recalculate
            )
            
            # 执行处理
            success = st.session_state.data_processor.process_data()
            
            if success:
                st.session_state.processed_data = st.session_state.data_processor.get_processed_data()
                st.success("✅ 数据处理完成！")
                app_logger.info("数据处理成功")
            else:
                st.error("❌ 数据处理失败")
                app_logger.error("数据处理失败")
                
    except Exception as e:
        st.error(f"❌ 处理过程中发生错误: {str(e)}")
        app_logger.error(f"数据处理异常: {e}")

def display_main_content():
    """显示主内容区域"""
    # 文件信息、数据统计和下载功能在同一行显示 - 使用紧凑布局
    col1, col2, col3 = st.columns([1.5, 1, 1.5])
    
    with col1:
        st.markdown("#### 📁 文件信息")
        if st.session_state.current_file_path:
            file_name = os.path.basename(st.session_state.current_file_path)
            st.info(f"**当前文件:** {file_name}")
            
            if st.session_state.current_worksheet:
                st.info(f"**当前工作表:** {st.session_state.current_worksheet}")
    
    with col2:
        st.markdown("#### 📊 数据统计")
        if st.session_state.current_data is not None:
            st.metric("总行数", len(st.session_state.current_data))
            st.metric("总列数", len(st.session_state.current_data.columns))
    
    with col3:
        st.markdown("#### 💾 下载处理后的文件")
        if st.session_state.processed_data is not None:
            # 生成Excel文件
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.processed_data.to_excel(writer, index=False, sheet_name='处理结果')
            
            st.download_button(
                label="📥 下载Excel文件",
                data=output.getvalue(),
                file_name=f"处理结果_{st.session_state.current_worksheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("处理数据后可下载文件")
    
    # 数据预览 - 紧凑布局
    if st.session_state.current_data is not None:
        st.markdown("#### 👀 数据预览")
        
        # 选择预览模式
        preview_mode = st.radio(
            "选择预览模式",
            ["原始数据", "处理后数据"],
            horizontal=True
        )
        
        if preview_mode == "原始数据":
            # 使用HTML表格避免pyarrow依赖
            preview_data = st.session_state.current_data.head(10)  # 进一步减少显示行数
            st.write(f"显示前 {len(preview_data)} 行数据（共 {len(st.session_state.current_data)} 行）")
            
            # 添加CSS样式
            st.markdown("""
            <style>
            .compact-table {
                font-size: 12px;
                border-collapse: collapse;
                width: 100%;
                margin: 10px 0;
            }
            .compact-table th, .compact-table td {
                border: 1px solid #ddd;
                padding: 4px 8px;
                text-align: left;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                max-width: 150px;
            }
            .compact-table th {
                background-color: #f2f2f2;
                font-weight: bold;
            }
            .compact-table tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            .compact-table tr:hover {
                background-color: #f5f5f5;
            }
            </style>
            """, unsafe_allow_html=True)
            
            # 转换为HTML表格显示
            html_table = preview_data.to_html(classes='compact-table', escape=False, index=False)
            st.markdown(html_table, unsafe_allow_html=True)
        elif preview_mode == "处理后数据" and st.session_state.processed_data is not None:
            # 使用HTML表格避免pyarrow依赖
            preview_data = st.session_state.processed_data.head(20)  # 减少显示行数
            st.write(f"显示前 {len(preview_data)} 行数据（共 {len(st.session_state.processed_data)} 行）")
            
            # 添加CSS样式（与原始数据显示保持一致）
            st.markdown("""
            <style>
            .compact-table {
                font-size: 12px;
                border-collapse: collapse;
                width: 100%;
                margin: 10px 0;
            }
            .compact-table th, .compact-table td {
                border: 1px solid #ddd;
                padding: 4px 8px;
                text-align: left;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                max-width: 150px;
            }
            .compact-table th {
                background-color: #f2f2f2;
                font-weight: bold;
            }
            .compact-table tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            .compact-table tr:hover {
                background-color: #f5f5f5;
            }
            </style>
            """, unsafe_allow_html=True)
            
            # 转换为HTML表格显示
            html_table = preview_data.to_html(classes='compact-table', escape=False, index=False)
            st.markdown(html_table, unsafe_allow_html=True)
        elif preview_mode == "处理后数据":
            st.info("请先执行数据处理")

def display_welcome_message():
    """显示欢迎信息"""
    st.markdown("""
    ## 👋 欢迎使用Excel发票数据处理软件
    
    ### 🚀 功能特点
    - 📁 **文件上传**: 支持拖拽上传Excel文件
    - 📋 **工作表选择**: 从Excel文件中选择特定工作表
    - 🎯 **列管理**: 灵活选择要删除的列和求和列
    - 👀 **数据预览**: 实时预览处理前后的数据
    - 💾 **文件下载**: 一键下载处理后的Excel文件
    
    ### 📝 使用步骤
    1. 在左侧面板上传Excel文件
    2. 选择要处理的工作表
    3. 选择要删除的列和求和列
    4. 点击"开始处理"按钮
    5. 预览处理结果并下载文件
    
    ### 💡 提示
    - 支持.xlsx和.xls格式的Excel文件
    - 处理过程中会保留原始数据，确保数据安全
    - 可以随时切换预览模式查看处理前后的对比
    """)

def update_column_selection(column_name):
    """更新列选择状态"""
    checkbox_key = f"checkbox_{column_name}_{hash(column_name) % 10000}"
    
    if st.session_state.get(checkbox_key, False):
        # 选中时添加到列表
        if column_name not in st.session_state.temp_selected_columns:
            st.session_state.temp_selected_columns.append(column_name)
    else:
        # 未选中时从列表移除
        if column_name in st.session_state.temp_selected_columns:
            st.session_state.temp_selected_columns.remove(column_name)

def show_column_selector_interface():
    """显示列选择界面"""
    # 创建一个紧凑的容器来显示列选择界面
    with st.container():
        st.markdown("## 🔧 选择要删除的列")
        
        if st.session_state.current_data is None:
            st.error("没有可用的数据")
            return
        
        columns = st.session_state.current_data.columns.tolist()
        
        # 初始化临时选择状态
        if 'temp_selected_columns' not in st.session_state:
            st.session_state.temp_selected_columns = st.session_state.selected_columns_to_delete.copy()
        
        # 确保临时选择状态与当前选择状态同步
        if not hasattr(st.session_state, '_temp_initialized'):
            st.session_state.temp_selected_columns = st.session_state.selected_columns_to_delete.copy()
            st.session_state._temp_initialized = True
        
        # 顶部工具栏 - 简化布局
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col1:
            # 预设模板
            template_options = ["自定义选择", "发票数据标准模板"]
            selected_template = st.selectbox("📋 选择预设模板", template_options, key="template_selector", index=1)
        
        with col2:
            if st.button("应用模板", key="apply_template", use_container_width=True):
                if selected_template == "发票数据标准模板":
                    try:
                        # 从配置文件中读取模板
                        config_manager = ConfigManager()
                        templates = config_manager.load_templates()
                        
                        if selected_template in templates:
                            template_columns = templates[selected_template].get("columns_to_delete", [])
                            # 只选择存在的列
                            template_selected = [col for col in template_columns if col in columns]
                            st.session_state.temp_selected_columns = template_selected
                            st.success(f"✅ 已应用模板：{selected_template}，匹配到 {len(template_selected)} 列")
                        else:
                            st.error("❌ 模板不存在")
                    except Exception as e:
                        st.error(f"❌ 应用模板失败: {str(e)}")
                    st.rerun()
        
        with col3:
            # 显示统计信息
            selected_count = len(st.session_state.temp_selected_columns)
            total_count = len(columns)
            st.metric("选择统计", f"{selected_count}/{total_count}")
        
        # 使用所有列（移除筛选功能）
        filtered_columns = columns
        
        # 快速操作按钮 - 紧凑布局
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("全选", key="select_all", use_container_width=True):
                st.session_state.temp_selected_columns = filtered_columns.copy()
                # 更新所有复选框的状态
                for column_name in filtered_columns:
                    checkbox_key = f"checkbox_{column_name}_{hash(column_name) % 10000}"
                    st.session_state[checkbox_key] = True
                st.rerun()
        
        with col2:
            if st.button("全不选", key="select_none", use_container_width=True):
                st.session_state.temp_selected_columns = []
                # 更新所有复选框的状态
                for column_name in filtered_columns:
                    checkbox_key = f"checkbox_{column_name}_{hash(column_name) % 10000}"
                    st.session_state[checkbox_key] = False
                st.rerun()
        
        with col3:
            if st.button("反选", key="invert_selection", use_container_width=True):
                current_selected = st.session_state.get('temp_selected_columns', [])
                new_selected = []
                for col in filtered_columns:
                    checkbox_key = f"checkbox_{col}_{hash(col) % 10000}"
                    if col not in current_selected:
                        new_selected.append(col)
                        st.session_state[checkbox_key] = True
                    else:
                        st.session_state[checkbox_key] = False
                st.session_state.temp_selected_columns = new_selected
                st.rerun()
        
        # 列选择区域 - 使用固定高度的滚动容器
        st.markdown("**可删除的列:**")
        
        # 创建一个固定高度的容器用于列选择
        with st.container():
            # 计算列数（每列显示6个复选框以节省空间）
            items_per_column = 6
            num_columns = min(4, (len(filtered_columns) + items_per_column - 1) // items_per_column)
            
            if num_columns > 0:
                cols = st.columns(num_columns)
                
                # 使用CSS样式限制高度并添加滚动
                st.markdown("""
                <style>
                .column-selector-container {
                    max-height: 300px;
                    overflow-y: auto;
                    border: 1px solid #e0e0e0;
                    border-radius: 5px;
                    padding: 10px;
                    margin: 10px 0;
                }
                </style>
                """, unsafe_allow_html=True)
                
                # 使用容器来管理复选框状态
                checkbox_container = st.container()
                
                with checkbox_container:
                    for i, column_name in enumerate(filtered_columns):
                        col_idx = i % num_columns
                        
                        # Excel列标识
                        excel_col = chr(65 + (columns.index(column_name) % 26))  # A, B, C...
                        
                        with cols[col_idx]:
                            # 检查是否已选中
                            is_selected = column_name in st.session_state.temp_selected_columns
                            
                            # 创建复选框，使用唯一的key
                            checkbox_key = f"checkbox_{column_name}_{hash(column_name) % 10000}"
                            
                            # 使用on_change回调来处理状态变化
                            checkbox_value = st.checkbox(
                                f"{excel_col}: {column_name}",
                                value=is_selected,
                                key=checkbox_key,
                                on_change=update_column_selection,
                                args=(column_name,)
                            )
        
        # 显示已选择的列 - 使用紧凑的展开面板
        if st.session_state.temp_selected_columns:
            with st.expander(f"已选择的列 ({len(st.session_state.temp_selected_columns)})", expanded=False):
                # 分列显示已选择的列，节省空间
                selected_cols = st.columns(3)
                for i, col in enumerate(st.session_state.temp_selected_columns):
                    with selected_cols[i % 3]:
                        st.write(f"• {col}")
        
        # 底部按钮区域 - 紧凑布局
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            if st.button("🔍 预览效果", key="preview_result", use_container_width=True):
                st.info("预览功能将在后续版本中实现")
        
        with col2:
            if st.button("✅ 确认选择", key="confirm_selection", use_container_width=True, type="primary"):
                # 应用选择
                st.session_state.selected_columns_to_delete = st.session_state.temp_selected_columns.copy()
                st.session_state.show_column_selector = False
                # 清理临时状态
                if 'temp_selected_columns' in st.session_state:
                    del st.session_state.temp_selected_columns
                if '_temp_initialized' in st.session_state:
                    del st.session_state._temp_initialized
                st.success(f"已选择 {selected_count} 列待删除")
                st.rerun()
        
        with col3:
            if st.button("❌ 取消", key="cancel_selection", use_container_width=True):
                st.session_state.show_column_selector = False
                # 清理临时状态
                if 'temp_selected_columns' in st.session_state:
                    del st.session_state.temp_selected_columns
                if '_temp_initialized' in st.session_state:
                    del st.session_state._temp_initialized
                st.rerun()

if __name__ == "__main__":
    main()
