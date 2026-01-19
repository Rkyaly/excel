import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import streamlit as st


def load_css():
    """加载统一的CSS样式"""
    css = """
    <style>
        .main {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }
        .stApp {
            background: transparent;
        }
        .header {
            background: rgba(255, 255, 255, 0.95);
            padding: 2rem;
            border-radius: 15px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            backdrop-filter: blur(10px);
            margin-bottom: 2rem;
        }
        .chart-container {                          #分割线
            background: rgba(255, 255, 255, 0.95);
            padding: 1rem 2rem;  /* 减小上下内边距 */
            border-radius: 15px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            backdrop-filter: blur(10px);
            margin-bottom: 1rem;  /* 减小容器间距 */
        }
        .sidebar {                                   #侧边栏
            background: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 1rem;
            margin: 1rem;
        }
        h1 {
            color: #4a5568;
            font-weight: 800;
        }
        h2 {
            color: #2d3748;
            font-weight: 700;
        }
        /* 保留其他样式，删除自定义上传按钮相关样式 */
        /* 自定义成功提示样式 */
        .stAlert.success {
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border: 2px solid #28a745;
            border-radius: 8px;
            padding: 1rem;
            color: #155724;
            font-weight: 600;
        }
        /* 自定义信息提示样式 */
        .stAlert.info {
            background: linear-gradient(135deg, #d1ecf1 0%, #bee5eb 100%);
            border: 2px solid #17a2b8;
            border-radius: 8px;
            padding: 1rem;
            color: #0c5460;
        }
        /* 自定义错误提示样式 */
        .stAlert.error {
            background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
            border: 2px solid #dc3545;
            border-radius: 8px;
            padding: 1rem;
            color: #721c24;
        }
        /* 统一全局按钮样式 */
        .stButton > button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 0.5rem 1.5rem;
            font-weight: 600;
            transition: all 0.3s ease;
        }
        .stButton > button:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


def render_header():
    """渲染主页面标题"""
    # 移除独立的header容器，改为在main函数中统一管理
    pass


def render_sidebar():                #渲染侧边栏
    """渲染侧边栏"""
    st.sidebar.markdown('<div class="sidebar">', unsafe_allow_html=True)
    st.sidebar.header("菜单栏")


    
    return None


def read_file(uploaded_file):
    """读取上传的文件并返回数据"""
    try:
        with st.spinner('🔄 正在读取文件...'):
            if uploaded_file.name.endswith('.csv'):
                # CSV文件只有一个表
                df = pd.read_csv(uploaded_file)
                sheet_names = ["Sheet1"]
                sheet_dfs = {"Sheet1": df}
            else:
                # Excel文件可能有多个子表
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                # 读取所有子表数据
                sheet_dfs = {}
                for sheet_name in sheet_names:
                    sheet_dfs[sheet_name] = excel_file.parse(sheet_name)
                # 默认使用第一个子表数据
                df = sheet_dfs[sheet_names[0]]
        return df, sheet_names, sheet_dfs
    except Exception as e:
        st.error(f"读取文件时出错: {str(e)}")
        return None, None, None


def show_file_info(uploaded_file, df, sheet_names):
    """显示文件信息（在上传组件下方）"""
    # 在上传组件下方显示包含完整信息的成功提示
    st.success(f"成功读取文件: {uploaded_file.name}\n📊 数据行数: {df.shape[0]}\n📋 数据列数: {df.shape[1]}\n📁 子表数量: {len(sheet_names)}\n📝 子表名称: {', '.join(sheet_names)}")


def filter_data(df):
    """筛选数据"""
    st.subheader("🔍 数据筛选")
    
    # 初始化筛选后的数据
    filtered_df = df.copy()
    
    # 筛选功能：选择列索引和对应值进行筛选
    try:
        # 选择筛选列（按索引）
        column_indices = list(range(len(df.columns)))
        selected_col_index = st.selectbox(
            "选择筛选列索引",
            options=column_indices,
            index=0,
            help="选择要筛选的列索引，从0开始计数"
        )
        
        # 获取选中列的名称
        selected_column = df.columns[selected_col_index]
        
        # 获取选中列的唯一值并排序
        unique_values = df[selected_column].dropna().unique().tolist()
        unique_values.sort()
        
        # 选择筛选值
        selected_value = st.selectbox(
            f"选择{selected_column}的值",
            options=unique_values,
            index=0,
            help="可输入搜索值"
        )
        
        # 执行筛选
        filtered_df = df[df[selected_column] == selected_value]
        st.info(f"筛选后数据行数: {len(filtered_df)}")
    except Exception as e:
        st.error(f"筛选数据时出错: {str(e)}")
        # 如果筛选出错，使用原始数据
        filtered_df = df.copy()
    return filtered_df


def configure_radar_chart(df):
    """配置雷达图"""
    st.sidebar.subheader("图表配置")
    
    # 雷达图配置
    st.sidebar.markdown("<h4 style='margin-bottom: 10px;'>📊 雷达图配置</h4>", unsafe_allow_html=True)
    # 初始化顶点列列表
    numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
    
    # 添加雷达图坐标模式选择
    invert_radar_coords = st.sidebar.checkbox(
        "反转雷达图坐标",
        value=True,
        help="勾选后：外部边线为0点，数据越大顶点越靠近中心；取消勾选：中心为0点，数据越大顶点越高"
    )
    
    if len(numeric_columns) >= 2:
        # 直接使用所有数值列作为顶点列
        vertex_cols = numeric_columns
    else:
        st.sidebar.warning("数据中至少需要2个数值列来创建雷达图")
        vertex_cols = []
        invert_radar_coords = False
    
    return vertex_cols, invert_radar_coords


def render_radar_chart(vertex_cols, current_name, sheet_names, sheet_dfs, invert_coordinates=False):
    """渲染雷达图
    
    Args:
        vertex_cols (list): 雷达图顶点列名列表
        current_name (str): 当前选中的名称
        sheet_names (list): 子表名称列表
        sheet_dfs (dict): 子表数据字典
        invert_coordinates (bool): 是否反转坐标（True: 外部为0，数据越大越靠近中心；False: 中心为0，数据越大越远离中心）
    """
    try:
        # 检查是否已选择顶点列
        if len(vertex_cols) >= 2:
            # 使用第一个子表数据
            if len(sheet_names) > 0:
                radar_df = sheet_dfs[sheet_names[0]]
                # 确保第一列存在
                if len(radar_df.columns) > 0:
                    radar_data = radar_df[radar_df.iloc[:, 0] == current_name]
                    
                    if not radar_data.empty:
                        # 使用用户选择的顶点列
                        row_data = radar_data.iloc[0]
                        original_values = [row_data[col] for col in vertex_cols]
                        
                        # 检查数据是否为零
                        total_value = sum(original_values)
                        
                        # 使用第一个子表名称
                        radar_sheet_name = sheet_names[0]
                        
                        # 显示图表标题
                        st.markdown(f"<h4>📊 雷达图</h4>", unsafe_allow_html=True)
                        st.markdown(f"<h5 style='margin-top: 5px; margin-bottom: 15px;'>{current_name} - {radar_sheet_name}</h5>", unsafe_allow_html=True)
                        
                        if total_value == 0:
                            # 数据为零时显示统一文本
                            st.markdown(f"<h2 style='color: #e53e3e; text-align: center; margin-top: 80px;'>数据为0</h2>", unsafe_allow_html=True)
                        else:
                            # 数据处理：根据invert_coordinates参数决定是否反转
                            if invert_coordinates:
                                # 反转坐标：外部为0，数据越大越靠近中心
                                max_value = max(original_values) * 1.2  # 增加20%作为缓冲
                                vertex_values = [max_value - val for val in original_values]
                                
                                radial_axis_config = dict(
                                    visible=True,
                                    range=[0, max_value],
                                    # 不显示具体刻度，只显示轴线
                                    tickvals=[],
                                    ticktext=[]
                                )
                                chart_title = f"{current_name} - 雷达图（反转坐标）"
                            else:
                                # 原始坐标：中心为0，数据越大越远离中心
                                vertex_values = original_values
                                max_value = max(vertex_values) * 1.2
                                
                                radial_axis_config = dict(
                                    visible=True,
                                    range=[0, max_value],
                                    # 不显示具体刻度，只显示轴线
                                    tickvals=[],
                                    ticktext=[]
                                )
                                chart_title = f"{current_name} - 雷达图"
                            
                            # 创建带数值的标签：在数据名称后显示对应数值
                            theta_labels = [f"{col}: {val:.0f}" for col, val in zip(vertex_cols, original_values)]
                            
                            fig_radar = go.Figure()
                            fig_radar.add_trace(go.Scatterpolar(
                                r=vertex_values + [vertex_values[0]],
                                theta=theta_labels + [theta_labels[0]],  # 使用带数值的标签
                                fill='toself',
                                name='雷达图数据',
                                line_color='rgba(102, 126, 234, 1)',
                                fillcolor='rgba(102, 126, 234, 0.3)',
                                line=dict(width=2)
                            ))
                            # 使用第一个子表名称
                            radar_sheet_name = sheet_names[0]
                            
                            fig_radar.update_layout(
                                polar=dict(radialaxis=radial_axis_config),
                                height=350,  # 增加雷达图高度，避免名称被遮挡
                                margin=dict(l=20, r=20, t=50, b=40),  # 增加top边距，让雷达图整体下移
                                template="plotly_white",
                                font=dict(size=10)  # 图表内部字体比h4小两号
                            )
                            st.plotly_chart(fig_radar, use_container_width=True)
                else:
                    st.info("雷达图子表没有数据列")
            else:
                st.info("未找到雷达图数据")
        else:
            st.info("请先在左侧侧边栏配置雷达图顶点")
    except Exception as e:
        st.error(f"生成雷达图时出错: {str(e)}")


def render_pie_chart(current_name, sheet_names, sheet_dfs):
    """渲染饼图"""
    try:
        if len(sheet_names) > 1:
            # 使用第二个子表数据
            pie_sheet_name = sheet_names[1]
            pie_df = sheet_dfs[pie_sheet_name]
            # 确保第一列存在
            if len(pie_df.columns) > 0:
                # 筛选当前人名的数据
                pie_data = pie_df[pie_df.iloc[:, 0] == current_name]
                
                if not pie_data.empty:
                    numeric_cols = pie_data.select_dtypes(include=[np.number]).columns.tolist()
                    if len(numeric_cols) > 0:
                        row_data = pie_data.iloc[0]
                        # 准备数据：名称为数值列名称，数值为当前行对应列的值
                        chart_data = pd.DataFrame({
                            "数据列": numeric_cols,
                            "数值": [row_data[col] for col in numeric_cols]
                        })
                        
                        # 显示图表标题
                        # st.markdown(f"<h4>🥧 饼图</h4>", unsafe_allow_html=True)
                        st.markdown(f"<h5 style='margin-top: -15px; margin-bottom: 15px;'>{current_name} - {pie_sheet_name}</h5>", unsafe_allow_html=True)
                        
                        # 检查数据是否为零
                        total_value = chart_data["数值"].sum()
                        if total_value == 0:
                            # 数据为零时显示统一文本
                            st.markdown(f"<h2 style='color: #e53e3e; text-align: center; margin-top: 80px;'>数据为0</h2>", unsafe_allow_html=True)
                        else:
                            # 数据不为零时绘制饼图
                            fig_pie = px.pie(
                                chart_data, names="数据列", values="数值",
                                color_discrete_sequence=px.colors.sequential.RdBu,
                                template="plotly_white",
                                hole=0.3
                            )
                            # 添加数据标签
                            fig_pie.update_traces(
                                textinfo='label+value+percent',  # 显示标签、数值和百分比
                                textfont_size=10  # 调整为比h4小两号
                            )
                            fig_pie.update_layout(
                                plot_bgcolor="rgba(0,0,0,0)",
                                paper_bgcolor="rgba(0,0,0,0)",
                                margin=dict(l=10, r=10, t=10, b=10),
                                height=300,
                                font=dict(size=10)  # 图表内部字体比h4小两号
                            )
                            st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("未找到当前人名的饼图数据")
            else:
                st.info("饼图子表没有数据列")
        else:
            st.info("请上传包含至少2个子表的Excel文件")
    except Exception as e:
        st.error(f"生成饼图时出错: {str(e)}")


def render_bar_chart(chart_num, current_name, sheet_names, sheet_dfs, color_sequence):
    """渲染柱状图"""
    try:
        sheet_index = chart_num + 2  # 柱状图1使用第3个子表（索引2），柱状图2使用第4个子表（索引3）
        if len(sheet_names) > sheet_index:
            # 使用对应子表数据
            bar_sheet_name = sheet_names[sheet_index]
            bar_df = sheet_dfs[bar_sheet_name]
            # 确保第一列存在
            if len(bar_df.columns) > 0:
                # 筛选当前人名的数据
                bar_data = bar_df[bar_df.iloc[:, 0] == current_name]
                
                if not bar_data.empty:
                        numeric_cols = bar_data.select_dtypes(include=[np.number]).columns.tolist()
                        if len(numeric_cols) > 0:
                            row_data = bar_data.iloc[0]
                            # 准备数据：x轴为数值列名称，y轴为当前行对应列的值
                            chart_data = pd.DataFrame({
                                "数据列": numeric_cols,
                                "数值": [row_data[col] for col in numeric_cols]
                            })
                            
                            # 显示图表标题
                            # st.markdown(f"<h4>📊 柱状图</h4>", unsafe_allow_html=True)
                            st.markdown(f"<h5 style='margin-top: -15px; margin-bottom: 15px;'>{current_name} - {bar_sheet_name}</h5>", unsafe_allow_html=True)
                            
                            # 检查数据是否为零
                            total_value = chart_data["数值"].sum()
                            if total_value == 0:
                                # 数据为零时显示统一文本
                                st.markdown(f"<h2 style='color: #e53e3e; text-align: center; margin-top: 80px;'>数据为0</h2>", unsafe_allow_html=True)
                            else:
                                fig_bar = px.bar(
                                    chart_data, x="数据列", y="数值",
                                    color_discrete_sequence=color_sequence,
                                    template="plotly_white",
                                    barmode='group',
                                    text="数值"  # 在柱子上显示数值
                                )
                                # 计算y轴范围，确保y轴为自然数（从0开始的正整数）
                                max_value = chart_data["数值"].max()
                                # 为y轴范围添加一些缓冲，确保最大值能够完整显示
                                y_max = max(max_value * 1.2, 1)  # 确保y轴至少显示到1
                                y_min = 0  # 自然数从0开始
                                
                                # 计算y轴刻度，只显示自然数
                                y_ticks = list(range(0, int(y_max) + 2))  # 生成0到y_max+1的自然数刻度
                                
                                fig_bar.update_traces(
                                    textposition='auto',  # 自动调整数值位置，避免重叠
                                    textfont_size=10  # 设置数值字体大小为比h4小两号
                                )
                                fig_bar.update_layout(
                                    plot_bgcolor="rgba(0,0,0,0)",
                                    paper_bgcolor="rgba(0,0,0,0)",
                                    margin=dict(l=10, r=10, t=10, b=50),
                                    xaxis=dict(tickmode='linear', title="数据列", tickangle=45, title_font=dict(size=10), tickfont=dict(size=9)),  # 旋转x轴标签，避免重叠
                                    yaxis=dict(
                                        title="数据高度", 
                                        range=[y_min, y_max],  # 设置y轴范围从0开始
                                        tickmode='array',  # 使用自定义刻度
                                        tickvals=y_ticks,  # 只显示自然数刻度
                                        automargin=True,
                                        title_font=dict(size=10),
                                        tickfont=dict(size=9)
                                    ),
                                    font=dict(size=10),  # 图表内部字体比h4小两号
                                    bargap=0.5,
                                    bargroupgap=0.3,
                                    height=350  # 增加高度以容纳数值标签
                                )
                                st.plotly_chart(fig_bar, use_container_width=True)
                else:
                    st.info("未找到当前人名的柱状图数据")
            else:
                st.info("柱状图子表没有数据列")
        else:
            st.info(f"请上传包含至少{sheet_index + 1}个子表的Excel文件")
    except Exception as e:
        st.error(f"生成柱状图{chart_num + 1}时出错: {str(e)}")


def render_line_chart(current_name, sheet_names, sheet_dfs):
    """渲染折线图"""
    try:
        if len(sheet_names) > 4:
            # 使用第五个子表数据
            line_sheet_name = sheet_names[4]
            line_df = sheet_dfs[line_sheet_name]
            # 确保第一列存在
            if len(line_df.columns) > 0:
                # 筛选当前人名的数据
                line_data = line_df[line_df.iloc[:, 0] == current_name]
                
                if not line_data.empty:
                    numeric_cols = line_data.select_dtypes(include=[np.number]).columns.tolist()
                    if len(numeric_cols) > 0:
                        row_data = line_data.iloc[0]
                        # 准备数据：x轴为数据列，y轴为对应值
                        chart_data = pd.DataFrame({
                            "数据列": numeric_cols,
                            "数值": [row_data[col] for col in numeric_cols]
                        })
                        
                        # 显示图表标题
                        # st.markdown(f"<h4>📈 折线图</h4>", unsafe_allow_html=True)
                        st.markdown(f"<h5 style='margin-top: -15px; margin-bottom: 15px;'>{current_name} - {line_sheet_name}</h5>", unsafe_allow_html=True)
                        
                        # 检查数据是否为零
                        total_value = chart_data["数值"].sum()
                        if total_value == 0:
                            # 数据为零时显示统一文本
                            st.markdown(f"<h2 style='color: #e53e3e; text-align: center; margin-top: 80px;'>数据为0</h2>", unsafe_allow_html=True)
                        else:
                            fig_line = px.line(
                                chart_data, x="数据列", y="数值",
                                color_discrete_sequence=px.colors.sequential.Plasma,
                                template="plotly_white",
                                markers=True,
                                text="数值"  # 在数据点上显示数值
                            )
                            fig_line.update_traces(
                                textposition='top center',  # 数值显示在数据点上方
                                textfont_size=10,  # 设置数值字体大小为比h4小两号
                                marker=dict(size=8)  # 增大数据点大小
                            )
                            # 计算y轴范围，确保y轴为自然数（从0开始的正整数）
                            max_value = chart_data["数值"].max()
                            # 为y轴范围添加一些缓冲，确保最大值能够完整显示
                            y_max = max(max_value * 1.2, 1)  # 确保y轴至少显示到1
                            y_min = 0  # 自然数从0开始
                            
                            # 计算y轴刻度，只显示自然数
                            y_ticks = list(range(0, int(y_max) + 2))  # 生成0到y_max+1的自然数刻度
                            
                            fig_line.update_layout(
                                plot_bgcolor="rgba(0,0,0,0)",
                                paper_bgcolor="rgba(0,0,0,0)",
                                margin=dict(l=20, r=20, t=10, b=50),
                                xaxis=dict(tickmode='linear', title="数据列", title_font=dict(size=10), tickfont=dict(size=9)),
                                yaxis=dict(
                                    title="数据高度", 
                                    range=[y_min, y_max],  # 设置y轴范围从0开始
                                    tickmode='array',  # 使用自定义刻度
                                    tickvals=y_ticks,  # 只显示自然数刻度
                                    automargin=True,
                                    title_font=dict(size=10),
                                    tickfont=dict(size=9)
                                ),
                                font=dict(size=10),  # 图表内部字体比h4小两号
                                height=450  # 增加高度以容纳数值标签
                            )
                            st.plotly_chart(fig_line, use_container_width=True)
                else:
                    st.info("未找到当前人名的折线图数据")
            else:
                st.info("折线图子表没有数据列")
        else:
            st.info("请上传包含至少5个子表的Excel文件")
    except Exception as e:
        st.error(f"生成折线图时出错: {str(e)}")


def show_data_preview(sheet_names, sheet_dfs, current_name):
    """显示数据预览"""
    # 第四行：所有子表数据预览（只显示筛选后的数据）
    for sheet_name in sheet_names:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        # 以子表名称命名数据预览
        st.subheader(f"📋 {sheet_name} 数据预览")
        
        # 获取当前子表数据
        current_sheet_df = sheet_dfs[sheet_name]
        
        try:
            # 应用与主筛选相同的条件：使用第一列进行筛选
            # 检查当前子表是否有第一列
            if len(current_sheet_df.columns) > 0:
                # 使用当前筛选的人名进行筛选
                filtered_sheet_df = current_sheet_df[current_sheet_df.iloc[:, 0] == current_name]
                # 显示筛选后的数据
                st.dataframe(filtered_sheet_df, width='stretch', height=200)  # 降低高度
            else:
                st.info(f"{sheet_name} 子表没有数据列")
        except Exception as e:
            st.error(f"筛选{sheet_name}数据时出错: {str(e)}")
            # 出错时显示完整数据
            st.dataframe(current_sheet_df, width='stretch', height=200)  # 降低高度
        
        st.markdown('</div>', unsafe_allow_html=True)


def show_quick_start():
    """显示快速开始指南"""
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.subheader("📝 欢迎使用数据可视化应用")
    
    st.markdown("<h4 style='margin-top: 2rem; margin-bottom: 15px;'>💡 快速开始</h4>", unsafe_allow_html=True)
    st.markdown("<ul>", unsafe_allow_html=True)
    st.markdown("<li style='margin: 0.5rem 0;'><strong>上传数据</strong>：在主页面点击上传按钮，选择Excel或CSV文件（支持.xlsx, .xls, .csv格式，最大不超过200M）</li>", unsafe_allow_html=True)
    st.markdown("<li style='margin: 0.5rem 0;'><strong>筛选数据</strong>：在'🔍 数据筛选'中选择筛选条件</li>", unsafe_allow_html=True)
    st.markdown("<li style='margin: 0.5rem 0;'><strong>配置图表</strong>：在左侧侧边栏的'图表配置'中调整雷达图参数</li>", unsafe_allow_html=True)
    st.markdown("<li style='margin: 0.5rem 0;'><strong>查看结果</strong>：浏览生成的各类图表和数据预览</li>", unsafe_allow_html=True)
    st.markdown("</ul>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


def main():
    # 设置页面配置
    st.set_page_config(
        page_title="Excel数据可视化",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # 加载CSS
    load_css()
    
    # 渲染侧边栏
    render_sidebar()
    
    # 1. 主页面标题 - 简洁显示
    # 以下代码已注释，如需恢复带样式的标题容器可取消注释
    # st.markdown('<div class="header">', unsafe_allow_html=True)
    st.title("📊 Excel数据可视化")
    # st.markdown('</div>', unsafe_allow_html=True)
    
    # 2. 上传组件区域 - 与可视化部分明确分隔
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.subheader("上传文件")
    
    # 3. 上传数据文件
    st.markdown("<h6>📁 选择要分析的数据文件</h6>", unsafe_allow_html=True)
    st.markdown("<p style='color: #718096; font-size: 14px; margin-top: 10px;'>支持格式：Excel (.xlsx, .xls) 和 CSV (.csv)，最大不超过200M</p>", unsafe_allow_html=True)
      
    # 4. 上传组件
    uploaded_file = st.file_uploader(
        label="上传数据文件", # 简化标签，实际提示已在页面显示
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed"  # 隐藏默认标签
    )
    
    # 添加一个空白分隔区域
    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
    
    if uploaded_file is not None:
        # 读取文件
        df, sheet_names, sheet_dfs = read_file(uploaded_file)
        
        if df is not None:
            # 显示文件信息
            show_file_info(uploaded_file, df, sheet_names)
            
            # 配置雷达图
            vertex_cols, invert_radar_coords = configure_radar_chart(df)
            
            # 检查初始数据是否为空
            if len(df) > 0:
                # 第一行：左侧显示数据筛选，中间留空白，右侧显示雷达图（保持原始大小）
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                # 使用三列布局，中间列作为空白分隔
                row1_col1, row1_col_space, row1_col2 = st.columns([0.7, 0.3, 1])  # 中间0.3列为空白
                
                with row1_col1:
                    # 调用筛选数据函数，获取筛选后的结果
                    filtered_df = filter_data(df)
                    
                    # 检查筛选后的数据是否为空
                    if len(filtered_df) > 0:
                        # 显示较大的人名（筛选后的数据），向右调整
                        # 显示较大的人名（筛选后的数据），向右并向下调整
                        st.markdown(f"<h1 style='color: #4a5568; font-weight: 800; margin-bottom: 0; margin-left: 100px; margin-top: 30px;'>{filtered_df.iloc[0].iloc[0]}</h1>", unsafe_allow_html=True)
                        st.markdown(f"<p style='color: #718096; margin-top: 0; margin-bottom: 2rem; margin-left: 100px; '>当前选中人员数据</p>", unsafe_allow_html=True)
                
                # 中间列为空白，不显示任何内容
                with row1_col_space:
                    st.write(" ")  # 只添加一个空格作为分隔
                
                with row1_col2:
                    # 检查筛选后的数据是否为空
                    if len(filtered_df) > 0:
                        # 生成雷达图（使用用户选择的顶点列）
                        render_radar_chart(vertex_cols, filtered_df.iloc[0].iloc[0], sheet_names, sheet_dfs, invert_coordinates=invert_radar_coords)
                st.markdown('</div>', unsafe_allow_html=True)
                
                # 检查筛选后的数据是否为空
                if len(filtered_df) > 0:
                    # 第二行：饼图、柱状图1、柱状图2放置在同一行
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.subheader("📊 数据分析")
                    st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)
                    
                    # 创建三列布局，将三个图表放在同一行
                    row2_col1, row2_col2, row2_col3 = st.columns(3)
                    
                    with row2_col1:
                        # 生成饼图（使用第二个子表数据）
                        st.markdown("<h4 style='margin-bottom: 15px;'>🥧 饼图</h4>", unsafe_allow_html=True)
                        render_pie_chart(filtered_df.iloc[0].iloc[0], sheet_names, sheet_dfs)
                    
                    with row2_col2:
                        # 第一个柱状图：使用第三个子表数据
                        st.markdown("<h4 style='margin-bottom: 15px;'>📊 柱状图1</h4>", unsafe_allow_html=True)
                        render_bar_chart(0, filtered_df.iloc[0].iloc[0], sheet_names, sheet_dfs, px.colors.sequential.Viridis)
                    
                    with row2_col3:
                        # 第二个柱状图：使用第四个子表数据
                        st.markdown("<h4 style='margin-bottom: 15px;'>📊 柱状图2</h4>", unsafe_allow_html=True)
                        render_bar_chart(1, filtered_df.iloc[0].iloc[0], sheet_names, sheet_dfs, px.colors.sequential.Plasma)
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # 第三行：折线图（单独一行，使用第五个子表数据）
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.subheader("📈 折线图")
                    render_line_chart(filtered_df.iloc[0].iloc[0], sheet_names, sheet_dfs)
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # 显示数据预览
                    show_data_preview(sheet_names, sheet_dfs, filtered_df.iloc[0].iloc[0])
                else:
                    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                    st.info("请先选择数据行")
                    st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.info("请先选择数据行")
                st.markdown('</div>', unsafe_allow_html=True)
    else:
        # 显示快速开始指南
        show_quick_start()
    
    st.sidebar.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
