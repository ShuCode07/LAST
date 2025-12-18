import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 设置页面标题
st.set_page_config(page_title="企业数字化转型指数查询", layout="wide")

def load_data():
    """数据加载函数，支持本地和部署环境"""
    try:
        # 尝试读取主要数据文件的不同可能路径
        possible_paths = [
            r"APP ALL\两版合并后的年报数据_完整版.xlsx",
            r"两版合并后的年报数据_完整版.xlsx",
            r"app\两版合并后的年报数据_完整版.xlsx"
        ]
        
        df = None
        for data_path in possible_paths:
            if os.path.exists(data_path):
                df = pd.read_excel(data_path, engine='openpyxl')
                st.success(f"成功加载数据文件: {data_path}")
                break
        
        if df is None:
            # 如果所有路径都失败，提供示例数据
            st.warning("未找到数据文件，使用示例数据")
            # 创建示例数据，包含战略、技术、组织、业务四个维度的指标
            sample_data = {
                '股票代码': ['000001', '000001', '000002', '000002'],
                '企业名称': ['平安银行', '平安银行', '万科A', '万科A'],
                '年份': [2020, 2021, 2020, 2021],
                # 战略维度指标
                '数字化转型战略词频数': [15, 25, 10, 20],
                '数字化转型愿景词频数': [10, 15, 8, 12],
                '数字化转型目标词频数': [12, 20, 9, 16],
                '数字化转型投资词频数': [20, 30, 15, 25],
                # 技术维度指标
                '人工智能词频数': [10, 15, 8, 12],
                '大数据词频数': [15, 20, 10, 18],
                '云计算词频数': [20, 25, 12, 20],
                '区块链词频数': [5, 10, 3, 8],
                '数字技术运用词频数': [50, 70, 33, 58],
                # 组织维度指标
                '数字化组织词频数': [8, 15, 6, 12],
                '数字化人才词频数': [12, 20, 9, 16],
                '数字化文化词频数': [10, 18, 7, 14],
                '数字化治理词频数': [15, 25, 10, 20],
                # 业务维度指标
                '数字化产品词频数': [12, 20, 8, 16],
                '数字化服务词频数': [15, 25, 10, 20],
                '数字化营销词频数': [10, 18, 7, 14],
                '数字化运营词频数': [18, 30, 12, 24],
                # 综合指数
                '数字化转型指数': [60, 80, 45, 70]
            }
            df = pd.DataFrame(sample_data)
        
        # 处理股票代码
        if '股票代码' in df.columns:
            df['股票代码'] = df['股票代码'].astype(str).str.zfill(6)
        
        # 尝试读取行业数据文件
        try:
            industry_paths = [
                r"APP ALL\最终数据dta格式-上市公司年度行业代码至2021.xlsx",
                r"最终数据dta格式-上市公司年度行业代码至2021.xlsx",
                r"app\最终数据dta格式-上市公司年度行业代码至2021.xlsx"
            ]
            
            df_industry = None
            for industry_path in industry_paths:
                if os.path.exists(industry_path):
                    df_industry = pd.read_excel(industry_path, engine='openpyxl')
                    st.success(f"成功加载行业数据文件: {industry_path}")
                    break
            
            if df_industry is not None:
                # 处理行业数据的股票代码和年份
                if '股票代码全称' in df_industry.columns:
                    df_industry['股票代码'] = df_industry['股票代码全称'].astype(str).str.zfill(6)
                
                if '年度' in df_industry.columns:
                    df_industry.rename(columns={'年度': '年份'}, inplace=True)
                
                # 合并数据
                if '股票代码' in df_industry.columns and '年份' in df_industry.columns:
                    merge_columns = ['股票代码', '年份']
                    if '行业名称' in df_industry.columns:
                        merge_columns.append('行业名称')
                    elif '行业代码' in df_industry.columns:
                        merge_columns.append('行业代码')
                    
                    df = pd.merge(df, df_industry[merge_columns], on=['股票代码', '年份'], how='left')
                
        except Exception as e:
            st.warning(f"读取行业数据失败，仅使用主要数据: {str(e)}")
        
        return df
        
    except Exception as e:
        st.error(f"数据加载错误: {str(e)}")
        st.info("建议检查数据文件是否存在于正确位置，或使用应用内的数据上传功能")
        return None

# 加载数据
df = load_data()

if df is not None:
    st.title("企业数字化转型指数查询系统")
    
    # 定义指标分类体系
    index_categories = {
        '战略维度': ['数字化转型战略词频数', '数字化转型愿景词频数', '数字化转型目标词频数', '数字化转型投资词频数'],
        '技术维度': ['人工智能词频数', '大数据词频数', '云计算词频数', '区块链词频数', '数字技术运用词频数'],
        '组织维度': ['数字化组织词频数', '数字化人才词频数', '数字化文化词频数', '数字化治理词频数'],
        '业务维度': ['数字化产品词频数', '数字化服务词频数', '数字化营销词频数', '数字化运营词频数'],
        '综合指数': ['数字化转型指数']
    }
    
    # 侧边栏查询条件
    st.sidebar.header("查询条件")
    
    # 行业选择（如果有行业数据）
    if '行业名称' in df.columns:
        industries = ['全部行业'] + sorted(df['行业名称'].dropna().unique())
        selected_industry = st.sidebar.selectbox("选择行业", options=industries)
    else:
        selected_industry = '全部行业'
    
    # 股票代码选择
    if selected_industry != '全部行业' and '行业名称' in df.columns:
        filtered_df = df[df['行业名称'] == selected_industry]
        stock_codes = sorted(filtered_df['股票代码'].unique())
    else:
        stock_codes = sorted(df['股票代码'].unique())
    
    # 获取企业名称映射
    stock_name_map = {}
    if '企业名称' in df.columns:
        for code in stock_codes:
            names = df[df['股票代码'] == code]['企业名称'].dropna().unique()
            if len(names) > 0:
                stock_name_map[code] = names[0]
    
    selected_stock = st.sidebar.selectbox(
        "选择股票代码",
        options=stock_codes,
        format_func=lambda x: f"{x} - {stock_name_map.get(x, '')}"
    )
    
    # 年份选择
    stock_years = sorted(df[df['股票代码'] == selected_stock]['年份'].unique())
    
    # 添加"全部显示"选项
    year_options = ['全部显示'] + stock_years
    selected_year = st.sidebar.selectbox("选择年份", options=year_options)
    
    # 获取所选数据
    if selected_year == '全部显示':
        selected_data = df[df['股票代码'] == selected_stock]
    else:
        selected_data = df[(df['股票代码'] == selected_stock) & (df['年份'] == selected_year)]
    
    # 添加数据上传功能
    st.sidebar.header("数据上传")
    uploaded_file = st.sidebar.file_uploader("上传自定义数据文件（Excel格式）", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            custom_df = pd.read_excel(uploaded_file, engine='openpyxl')
            st.sidebar.success("自定义数据文件上传成功！")
            # 可以在这里添加自定义数据处理逻辑
        except Exception as e:
            st.sidebar.error(f"数据上传错误: {str(e)}")
    
    if not selected_data.empty:
        # 显示企业基本信息
        st.header("企业基本信息")
        
        if '企业名称' in selected_data.columns:
            stock_name = selected_data['企业名称'].iloc[0]
            st.subheader(f"{stock_name} ({selected_stock})")
        else:
            st.subheader(f"{selected_stock}")
        
        # 显示行业信息
        if '行业名称' in selected_data.columns:
            industry_name = selected_data['行业名称'].iloc[0]
            if pd.notna(industry_name):
                st.write(f"行业: {industry_name}")
        elif '行业代码' in selected_data.columns:
            industry_code = selected_data['行业代码'].iloc[0]
            if pd.notna(industry_code):
                st.write(f"行业代码: {industry_code}")
        
        # 按维度展示数字化转型指标
        st.header("数字化转型指标分析")
        
        # 综合指数展示
        st.subheader("综合指数")
        if '数字化转型指数' in selected_data.columns:
            col1, col2 = st.columns(2)
            company_info = selected_data.iloc[0]
            col1.metric("数字化转型指数", value=round(company_info['数字化转型指数'], 4))
            
            # 综合指数趋势图
            if len(selected_data) > 1:
                fig_data = selected_data.sort_values('年份')
                fig = px.line(
                    fig_data,
                    x='年份',
                    y='数字化转型指数',
                    title=f"{stock_name_map.get(selected_stock, selected_stock)} 数字化转型指数趋势",
                    markers=True,
                    labels={'value': '指数值', '年份': '年份'}
                )
                fig.update_layout(
                    template='plotly_white',
                    hovermode='x unified'
                )
                col2.plotly_chart(fig, use_container_width=True)
            
        # 分维度展示指标
        for category, indicators in index_categories.items():
            if category != '综合指数':  # 已经单独展示过综合指数
                st.subheader(f"{category}")
                
                # 筛选该维度下实际存在的指标
                available_indicators = [ind for ind in indicators if ind in selected_data.columns]
                
                if available_indicators:
                    # 指标数值卡片展示
                    col1, col2, col3, col4 = st.columns(4)
                    cols = [col1, col2, col3, col4]
                    
                    company_info = selected_data.iloc[0]
                    for i, ind in enumerate(available_indicators):
                        if pd.notna(company_info[ind]):
                            cols[i % 4].metric(label=ind, value=round(company_info[ind], 4))
                    
                    # 维度内指标对比柱状图
                    if len(selected_data) == 1:
                        # 单一年份：该维度下各指标对比
                        fig = px.bar(
                            x=available_indicators,
                            y=[company_info[ind] for ind in available_indicators],
                            title=f"{category}各指标对比 ({selected_year})",
                            labels={'x': '指标', 'y': '数值'}
                        )
                        fig.update_layout(
                            template='plotly_white',
                            xaxis_tickangle=-45
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        # 多年份：该维度下各指标趋势对比
                        fig_data = selected_data.sort_values('年份')
                        fig = px.line(
                            fig_data,
                            x='年份',
                            y=available_indicators,
                            title=f"{category}各指标趋势对比",
                            markers=True,
                            labels={'value': '数值', '年份': '年份'}
                        )
                        fig.update_layout(
                            template='plotly_white',
                            hovermode='x unified',
                            legend_title='指标'
                        )
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info(f"未找到{category}相关指标数据")
        
        # 四大维度综合对比
        st.header("四大维度综合对比")
        
        # 计算各维度的平均得分
        dimension_scores = {}
        for category, indicators in index_categories.items():
            if category != '综合指数':
                available_indicators = [ind for ind in indicators if ind in selected_data.columns]
                if available_indicators:
                    if len(selected_data) == 1:
                        # 单一年份：取该维度下所有指标的平均值
                        dimension_scores[category] = selected_data[available_indicators].mean(axis=1).iloc[0]
                    else:
                        # 多年份：取最新年份的数据
                        latest_data = selected_data[selected_data['年份'] == selected_data['年份'].max()]
                        dimension_scores[category] = latest_data[available_indicators].mean(axis=1).iloc[0]
        
        if dimension_scores:
            # 雷达图展示四大维度
            fig = px.line_polar(
                r=list(dimension_scores.values()),
                theta=list(dimension_scores.keys()),
                line_close=True,
                title="四大维度综合对比雷达图",
                markers=True
            )
            fig.update_layout(
                template='plotly_white'
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 柱状图展示四大维度
            fig = px.bar(
                x=list(dimension_scores.keys()),
                y=list(dimension_scores.values()),
                title="四大维度综合对比柱状图",
                labels={'x': '维度', 'y': '平均得分'}
            )
            fig.update_layout(
                template='plotly_white'
            )
            st.plotly_chart(fig, use_container_width=True)
            
        # 显示原始数据
        st.header("原始数据")
        
        # 确保所有数据都能显示
        st.dataframe(selected_data, use_container_width=True, height=800)
        
        # 添加数据下载功能
        st.header("数据下载")
        
        # 生成文件名
        if '企业名称' in selected_data.columns:
            company_name = selected_data['企业名称'].iloc[0] if pd.notna(selected_data['企业名称'].iloc[0]) else selected_stock
        else:
            company_name = selected_stock
        
        if selected_year == '全部显示':
            file_name = f"{company_name}_全部年份_数字化转型数据.csv"
        else:
            file_name = f"{company_name}_{selected_year}_数字化转型数据.csv"
        
        st.download_button(
            label="下载原始数据（CSV）",
            data=selected_data.to_csv(index=False, encoding='utf-8-sig'),
            file_name=file_name,
            mime="text/csv"
        )
    else:
        st.warning("未找到所选条件的数据")
    
    # 数据概览
    st.header("数据概览")
    
    # 基本统计信息
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("总股票数", len(df['股票代码'].unique()))
    if '行业名称' in df.columns:
        col2.metric("行业数量", len(df['行业名称'].dropna().unique()))
    else:
        col2.metric("行业数量", "无数据")
    col3.metric("年份范围", f"{df['年份'].min()} - {df['年份'].max()}")
    col4.metric("数据总量", len(df))
    
    # 数字化转型指数分布
    if '数字化转型指数' in df.columns:
        st.subheader("数字化转型指数分布")
        
        # 直方图展示指数分布
        fig = px.histogram(
            df,
            x='数字化转型指数',
            nbins=20,
            title="数字化转型指数整体分布",
            labels={'数字化转型指数': '指数值', 'count': '企业数量'}
        )
        fig.update_layout(template='plotly_white')
        st.plotly_chart(fig, use_container_width=True)
        
        # 年份维度的指数变化
        st.subheader("数字化转型指数年度趋势")
        year_trend = df.groupby('年份')['数字化转型指数'].mean().reset_index()
        fig = px.line(
            year_trend,
            x='年份',
            y='数字化转型指数',
            markers=True,
            title="各年份平均数字化转型指数",
            labels={'数字化转型指数': '平均指数值', '年份': '年份'}
        )
        fig.update_layout(template='plotly_white')
        st.plotly_chart(fig, use_container_width=True)
    
    # 行业维度分析（如果有行业数据）
    if '行业名称' in df.columns:
        st.subheader("行业数字化转型分析")
        
        # 各行业平均数字化转型指数
        industry_avg = df.groupby('行业名称')['数字化转型指数'].mean().reset_index().sort_values('数字化转型指数', ascending=False)
        fig = px.bar(
            industry_avg,
            x='行业名称',
            y='数字化转型指数',
            title="各行业平均数字化转型指数排名",
            labels={'数字化转型指数': '平均指数值', '行业名称': '行业'}
        )
        fig.update_layout(template='plotly_white', xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
else:
    st.error("无法加载数据，请检查文件路径和格式")