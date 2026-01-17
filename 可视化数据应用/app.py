import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import streamlit as st

def main():
    # 设置页面配置
    st.set_page_config(
        page_title="Excel数据可视化",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # 自定义CSS样式
    st.markdown("""
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
        .chart-container {
            background: rgba(255, 255, 255, 0.95);
            padding: 2rem;
            border-radius: 15px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            backdrop-filter: blur(10px);
            margin-bottom: 2rem;
        }
        .sidebar {
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
        .stFileUploader > label {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 8px;
            padding: 1rem;
            border: 2px dashed #667eea;
        }
    </style>
    """, unsafe_allow_html=True)

    # 主页面标题
    st.markdown('<div class="header">', unsafe_allow_html=True)
    st.title("📊 Excel数据可视化")
    st.subheader("上传Excel文件，生成可视化图表")
    st.markdown('</div>', unsafe_allow_html=True)

    # 侧边栏
    st.sidebar.markdown('<div class="sidebar">', unsafe_allow_html=True)
    st.sidebar.header("设置")

    # 文件上传
    uploaded_file = st.sidebar.file_uploader("上传Excel文件", type=["xlsx", "xls", "csv"])

    if uploaded_file is not None:
        # 读取文件
        try:
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
            
            st.sidebar.success(f"成功读取文件: {uploaded_file.name}")
            st.sidebar.write(f"数据行数: {df.shape[0]}")
            st.sidebar.write(f"数据列数: {df.shape[1]}")
            st.sidebar.write(f"子表数量: {len(sheet_names)}")
            st.sidebar.write(f"子表名称: {', '.join(sheet_names)}")
            
            # 数据筛选区域
            st.sidebar.subheader("🔍 数据筛选")
            
            # 初始化筛选后的数据
            filtered_df = df.copy()
            
            # 筛选功能：选择列索引和对应值进行筛选
            try:
                # 选择筛选列（按索引）
                column_indices = list(range(len(df.columns)))
                selected_col_index = st.sidebar.selectbox(
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
                selected_value = st.sidebar.selectbox(
                    f"选择{selected_column}的值",
                    options=unique_values,
                    index=0,
                    help="可输入搜索值"
                )
                
                # 执行筛选
                filtered_df = df[df[selected_column] == selected_value]
                st.sidebar.info(f"筛选后数据行数: {len(filtered_df)}")
            except Exception as e:
                st.sidebar.error(f"筛选数据时出错: {str(e)}")
                # 如果筛选出错，使用原始数据
                filtered_df = df.copy()
            
            # 选择数据列
            numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
            non_numeric_columns = df.select_dtypes(exclude=[np.number]).columns.tolist()
            
            # 图表配置
            st.sidebar.subheader("图表配置")
            
            # 雷达图配置
            st.sidebar.markdown("<h4 style='margin-bottom: 10px;'>📊 雷达图配置</h4>", unsafe_allow_html=True)
            # 初始化顶点列列表
            vertex_cols = []
            if len(numeric_columns) >= 2:
                # 为雷达图选择顶点列（至少2个）
                available_cols = numeric_columns.copy()
                
                # 允许用户选择任意数量的顶点列（2-10个）
                num_vertices = st.sidebar.slider(
                    "顶点数量",
                    min_value=2,
                    max_value=10,
                    value=6,
                    key="num_vertices"
                )
                
                for i in range(num_vertices):
                    vertex_col = st.sidebar.selectbox(
                        f"顶点{i+1}",
                        options=available_cols,
                        key=f"vertex_col_{i}"
                    )
                    vertex_cols.append(vertex_col)
                    # 从可用列中移除已选择的列
                    available_cols.remove(vertex_col)
            else:
                st.sidebar.warning("数据中至少需要2个数值列来创建雷达图")
            
            # 显示区域布局
            if len(filtered_df) > 0:
                # 获取当前选择的人名（第一列的值）
                current_name = filtered_df.iloc[0].iloc[0]
                
                # 第一行：左侧显示人名，右侧显示缩小的雷达图（使用第一个子表数据）
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                row1_col1, row1_col2 = st.columns([1, 2])
                
                with row1_col1:
                    # 显示较大的人名
                    st.markdown(f"<h1 style='color: #4a5568; font-weight: 800; margin-bottom: 0;'>{current_name}</h1>", unsafe_allow_html=True)
                    st.markdown(f"<p style='color: #718096; margin-top: 0;'>当前选中人员数据</p>", unsafe_allow_html=True)
                
                with row1_col2:
                    # 生成缩小的雷达图（使用用户选择的顶点列）
                    st.markdown("<h3 style='margin-bottom: 10px;'>📊 雷达图</h3>", unsafe_allow_html=True)
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
                                        vertex_values = [row_data[col] for col in vertex_cols]
                                        
                                        fig_radar = go.Figure()
                                        fig_radar.add_trace(go.Scatterpolar(
                                            r=vertex_values + [vertex_values[0]],
                                            theta=vertex_cols + [vertex_cols[0]],
                                            fill='toself',
                                            name='雷达图数据',
                                            line_color='rgba(102, 126, 234, 1)',
                                            fillcolor='rgba(102, 126, 234, 0.3)',
                                            line=dict(width=2)
                                        ))
                                        fig_radar.update_layout(
                                            polar=dict(radialaxis=dict(visible=True, range=[0, max(vertex_values) * 1.2])),
                                            height=250,  # 增加雷达图高度，避免名称被遮挡
                                            margin=dict(l=10, r=10, t=40, b=30),  # 调整边距，增加底部边距
                                            template="plotly_white"
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
                st.markdown('</div>', unsafe_allow_html=True)
                
                # 第二行：柱状图和饼图（同一行）
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("📊 数据分析")
                st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)
                row2_col1, row2_col2 = st.columns(2)
                
                # 柱状图使用第二个子表数据
                with row2_col1:
                    st.markdown("<h4 style='margin-bottom: 15px;'>📊 柱状图</h4>", unsafe_allow_html=True)
                    try:
                        if len(sheet_names) > 1:
                            # 使用第二个子表数据
                            bar_sheet_name = sheet_names[1]
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
                                        
                                        fig_bar = px.bar(
                                            chart_data, x="数据列", y="数值",
                                            color_discrete_sequence=px.colors.sequential.Viridis,
                                            template="plotly_white",
                                            barmode='group'
                                        )
                                        fig_bar.update_layout(
                                            title=f"{current_name} - {bar_sheet_name}",
                                            plot_bgcolor="rgba(0,0,0,0)",
                                            paper_bgcolor="rgba(0,0,0,0)",
                                            margin=dict(l=10, r=10, t=40, b=50),
                                            xaxis=dict(tickmode='linear', title="数据列"),
                                            yaxis=dict(title="数据高度"),
                                            bargap=0.5,
                                            bargroupgap=0.3,
                                            height=300
                                        )
                                        st.plotly_chart(fig_bar, use_container_width=True)
                                else:
                                    st.info("未找到当前人名的柱状图数据")
                            else:
                                st.info("柱状图子表没有数据列")
                        else:
                            st.info("请上传包含至少2个子表的Excel文件")
                    except Exception as e:
                        st.error(f"生成柱状图时出错: {str(e)}")
                
                # 饼图使用第三个子表数据
                with row2_col2:
                    st.markdown("<h4 style='margin-bottom: 15px;'>🥧 饼图</h4>", unsafe_allow_html=True)
                    try:
                        if len(sheet_names) > 2:
                            # 使用第三个子表数据
                            pie_sheet_name = sheet_names[2]
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
                                        
                                        # 检查数据是否为零
                                        total_value = chart_data["数值"].sum()
                                        if total_value == 0:
                                            # 数据为零时显示指定文本
                                            st.markdown(f"<h2 style='color: #e53e3e; text-align: center; margin-top: 80px;'>故障失格律0%</h2>", unsafe_allow_html=True)
                                        else:
                                            # 数据不为零时绘制饼图
                                            fig_pie = px.pie(
                                                chart_data, names="数据列", values="数值",
                                                color_discrete_sequence=px.colors.sequential.RdBu,
                                                template="plotly_white",
                                                hole=0.3
                                            )
                                            fig_pie.update_layout(
                                                title=f"{current_name} - {pie_sheet_name}",
                                                plot_bgcolor="rgba(0,0,0,0)",
                                                paper_bgcolor="rgba(0,0,0,0)",
                                                margin=dict(l=10, r=10, t=40, b=10),
                                                height=300
                                            )
                                            st.plotly_chart(fig_pie, use_container_width=True)
                                else:
                                    st.info("未找到当前人名的饼图数据")
                            else:
                                st.info("饼图子表没有数据列")
                        else:
                            st.info("请上传包含至少3个子表的Excel文件")
                    except Exception as e:
                        st.error(f"生成饼图时出错: {str(e)}")
                st.markdown('</div>', unsafe_allow_html=True)
                
                # 第三行：折线图（单独一行，使用第四个子表数据）
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.subheader("📈 折线图")
                try:
                    if len(sheet_names) > 3:
                        # 使用第四个子表数据
                        line_sheet_name = sheet_names[3]
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
                                    
                                    fig_line = px.line(
                                        chart_data, x="数据列", y="数值",
                                        color_discrete_sequence=px.colors.sequential.Plasma,
                                        template="plotly_white",
                                        markers=True
                                    )
                                    fig_line.update_layout(
                                        title=f"{current_name} - {line_sheet_name}",
                                        plot_bgcolor="rgba(0,0,0,0)",
                                        paper_bgcolor="rgba(0,0,0,0)",
                                        margin=dict(l=20, r=20, t=40, b=50),
                                        xaxis=dict(tickmode='linear', title="数据列"),
                                        yaxis=dict(title="数据高度"),
                                        height=400
                                    )
                                    st.plotly_chart(fig_line, use_container_width=True)
                            else:
                                st.info("未找到当前人名的折线图数据")
                        else:
                            st.info("折线图子表没有数据列")
                    else:
                        st.info("请上传包含至少4个子表的Excel文件")
                except Exception as e:
                    st.error(f"生成折线图时出错: {str(e)}")
                st.markdown('</div>', unsafe_allow_html=True)
                
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
            else:
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)
                st.info("请先选择数据行")
                st.markdown('</div>', unsafe_allow_html=True)
                
        except Exception as e:
            st.error(f"读取文件时出错: {str(e)}")
    else:
        # 示例数据
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.subheader("📝 示例数据")
        
        # 创建示例数据 - 包含4个子表，用于展示当前界面的所有功能
        
        # 第一个子表：雷达图数据
        radar_data = {
            "姓名": ["张三", "李四", "王五"],
            "能力1": [85, 90, 75],
            "能力2": [78, 82, 88],
            "能力3": [92, 85, 70],
            "能力4": [88, 95, 80],
            "能力5": [75, 80, 85],
            "能力6": [80, 88, 92]
        }
        radar_df = pd.DataFrame(radar_data)
        
        # 第二个子表：柱状图数据
        bar_data = {
            "姓名": ["张三", "李四", "王五"],
            "项目1": [120, 150, 90],
            "项目2": [180, 200, 160],
            "项目3": [90, 130, 110],
            "项目4": [210, 190, 180]
        }
        bar_df = pd.DataFrame(bar_data)
        
        # 第三个子表：饼图数据（包含零值情况）
        pie_data = {
            "姓名": ["张三", "李四", "王五"],
            "完成": [80, 100, 0],  # 王五的数据为0，用于测试零值显示
            "未完成": [20, 0, 0]
        }
        pie_df = pd.DataFrame(pie_data)
        
        # 第四个子表：折线图数据
        line_data = {
            "姓名": ["张三", "李四", "王五"],
            "一月": [25, 30, 20],
            "二月": [35, 40, 30],
            "三月": [30, 45, 25],
            "四月": [40, 50, 35],
            "五月": [45, 55, 40],
            "六月": [50, 60, 45]
        }
        line_df = pd.DataFrame(line_data)
        
        # 显示示例数据说明
        st.markdown("<h4 style='margin-bottom: 10px;'>📊 示例数据说明</h4>", unsafe_allow_html=True)
        st.markdown("<p>该示例包含4个子表，可用于展示当前界面的所有功能：</p>", unsafe_allow_html=True)
        st.markdown("<ul>", unsafe_allow_html=True)
        st.markdown("<li><strong>子表1</strong>：雷达图数据 - 包含能力评分</li>", unsafe_allow_html=True)
        st.markdown("<li><strong>子表2</strong>：柱状图数据 - 包含项目数据</li>", unsafe_allow_html=True)
        st.markdown("<li><strong>子表3</strong>：饼图数据 - 包含完成情况（王五的数据为0，用于测试零值显示）</li>", unsafe_allow_html=True)
        st.markdown("<li><strong>子表4</strong>：折线图数据 - 包含月度数据</li>", unsafe_allow_html=True)
        st.markdown("</ul>", unsafe_allow_html=True)
        
        # 显示子表1：雷达图数据
        st.markdown("<h5 style='margin-top: 20px; margin-bottom: 10px;'>1. 雷达图数据</h5>", unsafe_allow_html=True)
        st.dataframe(radar_df)
        
        # 显示子表2：柱状图数据
        st.markdown("<h5 style='margin-top: 20px; margin-bottom: 10px;'>2. 柱状图数据</h5>", unsafe_allow_html=True)
        st.dataframe(bar_df)
        
        # 显示子表3：饼图数据
        st.markdown("<h5 style='margin-top: 20px; margin-bottom: 10px;'>3. 饼图数据</h5>", unsafe_allow_html=True)
        st.dataframe(pie_df)
        
        # 显示子表4：折线图数据
        st.markdown("<h5 style='margin-top: 20px; margin-bottom: 10px;'>4. 折线图数据</h5>", unsafe_allow_html=True)
        st.dataframe(line_df)
        
        # 显示使用说明
        st.markdown("<h4 style='margin-top: 20px; margin-bottom: 10px;'>💡 使用说明</h4>", unsafe_allow_html=True)
        st.markdown("<p>1. 上传包含多个子表的Excel文件，或使用示例数据</p>", unsafe_allow_html=True)
        st.markdown("<p>2. 在左侧侧边栏的'🔍 数据筛选'中选择筛选列索引和对应值</p>", unsafe_allow_html=True)  # 修复双引号嵌套问题
        st.markdown("<p>3. 在'图表配置'中配置雷达图的顶点数量和顶点列</p>", unsafe_allow_html=True)  # 修复双引号嵌套问题
        st.markdown("<p>4. 查看各个图表：雷达图、柱状图、饼图、折线图</p>", unsafe_allow_html=True)
        st.markdown("<p>5. 查看各个子表的筛选后数据预览</p>", unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 使用示例数据生成图表
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.subheader("📊 示例图表")
        
        # 默认使用李四的数据
        current_name = "李四"
        
        # 1. 显示雷达图
        st.markdown("<h3 style='margin-bottom: 10px;'>📊 雷达图</h3>", unsafe_allow_html=True)
        radar_data = radar_df[radar_df["姓名"] == current_name]
        if not radar_data.empty:
            # 使用前6个数值列作为雷达图顶点
            numeric_cols = radar_df.select_dtypes(include=[np.number]).columns.tolist()[:6]
            if len(numeric_cols) >= 2:
                row_data = radar_data.iloc[0]
                vertex_values = [row_data[col] for col in numeric_cols]
                
                fig_radar = go.Figure()
                fig_radar.add_trace(go.Scatterpolar(
                    r=vertex_values + [vertex_values[0]],
                    theta=numeric_cols + [numeric_cols[0]],
                    fill='toself',
                    name='雷达图数据',
                    line_color='rgba(102, 126, 234, 1)',
                    fillcolor='rgba(102, 126, 234, 0.3)',
                    line=dict(width=2)
                ))
                fig_radar.update_layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0, max(vertex_values) * 1.2])),
                    height=250,
                    margin=dict(l=10, r=10, t=40, b=30),
                    template="plotly_white"
                )
                st.plotly_chart(fig_radar, use_container_width=True)
        
        # 2. 显示柱状图和饼图（同一行）
        st.markdown("<h3 style='margin-top: 20px; margin-bottom: 10px;'>📊 数据分析</h3>", unsafe_allow_html=True)
        row1_col1, row1_col2 = st.columns(2)
        
        with row1_col1:
            # 柱状图使用第二个子表数据
            st.markdown("<h4 style='margin-bottom: 15px;'>📊 柱状图</h4>", unsafe_allow_html=True)
            bar_data = bar_df[bar_df["姓名"] == current_name]
            if not bar_data.empty:
                numeric_cols = bar_df.select_dtypes(include=[np.number]).columns.tolist()
                if len(numeric_cols) > 0:
                    row_data = bar_data.iloc[0]
                    chart_data = pd.DataFrame({
                        "数据列": numeric_cols,
                        "数值": [row_data[col] for col in numeric_cols]
                    })
                    
                    fig_bar = px.bar(
                        chart_data, x="数据列", y="数值",
                        color_discrete_sequence=px.colors.sequential.Viridis,
                        template="plotly_white",
                        barmode='group'
                    )
                    fig_bar.update_layout(
                        title=f"{current_name} - 柱状图数据",
                        plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)",
                        margin=dict(l=10, r=10, t=40, b=50),
                        xaxis=dict(tickmode='linear', title="数据列"),
                        yaxis=dict(title="数据高度"),
                        bargap=0.5,
                        bargroupgap=0.3,
                        height=300
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
        
        with row1_col2:
            # 饼图使用第三个子表数据
            st.markdown("<h4 style='margin-bottom: 15px;'>🥧 饼图</h4>", unsafe_allow_html=True)
            pie_data = pie_df[pie_df["姓名"] == current_name]
            if not pie_data.empty:
                numeric_cols = pie_df.select_dtypes(include=[np.number]).columns.tolist()
                if len(numeric_cols) > 0:
                    row_data = pie_data.iloc[0]
                    chart_data = pd.DataFrame({
                        "数据列": numeric_cols,
                        "数值": [row_data[col] for col in numeric_cols]
                    })
                    
                    total_value = chart_data["数值"].sum()
                    if total_value == 0:
                        st.markdown(f"<h2 style='color: #e53e3e; text-align: center; margin-top: 80px;'>故障失格律0%</h2>", unsafe_allow_html=True)
                    else:
                        fig_pie = px.pie(
                            chart_data, names="数据列", values="数值",
                            color_discrete_sequence=px.colors.sequential.RdBu,
                            template="plotly_white",
                            hole=0.3
                        )
                        fig_pie.update_layout(
                            title=f"{current_name} - 饼图数据",
                            plot_bgcolor="rgba(0,0,0,0)",
                            paper_bgcolor="rgba(0,0,0,0)",
                            margin=dict(l=10, r=10, t=40, b=10),
                            height=300
                        )
                        st.plotly_chart(fig_pie, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 3. 显示折线图（单独一行）
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.subheader("📈 折线图")
        line_data = line_df[line_df["姓名"] == current_name]
        if not line_data.empty:
            numeric_cols = line_df.select_dtypes(include=[np.number]).columns.tolist()
            if len(numeric_cols) > 0:
                row_data = line_data.iloc[0]
                chart_data = pd.DataFrame({
                    "数据列": numeric_cols,
                    "数值": [row_data[col] for col in numeric_cols]
                })
                
                fig_line = px.line(
                    chart_data, x="数据列", y="数值",
                    color_discrete_sequence=px.colors.sequential.Plasma,
                    template="plotly_white",
                    markers=True
                )
                fig_line.update_layout(
                    title=f"{current_name} - 折线图数据",
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    margin=dict(l=20, r=20, t=40, b=50),
                    xaxis=dict(tickmode='linear', title="数据列"),
                    yaxis=dict(title="数据高度"),
                    height=400
                )
                st.plotly_chart(fig_line, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 4. 显示示例数据预览
        st.subheader("📋 示例数据预览")
        
        # 模拟sheet_dfs和sheet_names
        sample_sheet_dfs = {
            "雷达图数据": radar_df,
            "柱状图数据": bar_df,
            "饼图数据": pie_df,
            "折线图数据": line_df
        }
        sample_sheet_names = list(sample_sheet_dfs.keys())
        
        # 显示每个子表的筛选后数据预览
        for sheet_name in sample_sheet_names:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader(f"📋 {sheet_name} 数据预览")
            
            current_sheet_df = sample_sheet_dfs[sheet_name]
            try:
                if len(current_sheet_df.columns) > 0:
                    filtered_sheet_df = current_sheet_df[current_sheet_df.iloc[:, 0] == current_name]
                    st.dataframe(filtered_sheet_df, width='stretch', height=200)
                else:
                    st.info(f"{sheet_name} 子表没有数据列")
            except Exception as e:
                st.error(f"筛选{sheet_name}数据时出错: {str(e)}")
                st.dataframe(current_sheet_df, width='stretch', height=200)
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.info("💡 提示：在左侧上传Excel文件可使用您自己的数据")
        st.markdown('</div>', unsafe_allow_html=True)

    st.sidebar.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()