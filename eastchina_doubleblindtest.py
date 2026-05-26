# 飞行员双盲测试评估分析
# Streamlit + Plotly + python-docx

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="飞行员双盲测试评估分析", layout="wide")

# ------------------ 数据读取 ------------------
@st.cache_data
def load_template(file):
    """读取统一模板Excel，处理重复列名"""
    raw = pd.read_excel(file, header=None)
    
    h1 = raw.iloc[0].fillna("")
    h2 = raw.iloc[1].fillna("")
    
    cols_raw = []
    for a, b in zip(h1, h2):
        if pd.notna(a) and str(a).strip() != "":
            cols_raw.append(str(a).replace("\n", ""))
        else:
            cols_raw.append(str(b).replace("\n", ""))
    
    cols = []
    col_count = {}
    for c in cols_raw:
        if c == "" or c.isspace():
            c = "unnamed"
        if c in col_count:
            col_count[c] += 1
            cols.append(f"{c}_{col_count[c]}")
        else:
            col_count[c] = 1
            cols.append(c)
    
    df = raw.iloc[3:].copy()
    df.columns = cols
    df = df.dropna(how='all')
    
    if "姓名" in df.columns:
        df = df[df["姓名"].notna()]
    else:
        name_col = None
        for c in df.columns:
            if "姓名" in c:
                name_col = c
                break
        if name_col:
            df = df[df[name_col].notna()]
    
    return df


# ------------------ 科目定义 ------------------
SUBJECTS = {
    '大坡度盘旋': ['高度偏差', '速度偏差', '改出航向', '坡度保持', '滚转速率'],
    '大侧风目视起落': ['起飞抬头率', '一边航迹', 'BANK ANGLE', '三边高度', '三边航迹', '三边宽度', '四边航迹', '四转弯', '下滑线', '五边速度', '稳定意识', 'SINK RATE', '入口高度', '着陆'],
    '非精密进近_中断着陆': ['滑跑方向', '稳定进近', '高距比', '下滑线控制', '中断动作', '中断程序'],
    '中断着陆后发动机失效': ['单发初始姿态', '单发推力', '航迹误差', '坡度控制', '航迹控制', '通信', 'PANPAN'],
    '单发ILS无指引落地': ['1000ft以下', 'LOC', 'G/S', 'Vapp', '剖面控制']
}

COMPREHENSIVE_ITEMS = ['技术性复飞', '决策失误', 'PULL UP', '盲目蛮干', '冲偏出跑道']


def categorize_deduction_item(col_name):
    """将扣分项归类到具体科目"""
    col_lower = col_name.lower()
    
    for subject, keywords in SUBJECTS.items():
        for keyword in keywords:
            if keyword.lower() in col_lower:
                return subject
    
    for item in COMPREHENSIVE_ITEMS:
        if item.lower() in col_lower:
            return '综合考评'
    
    return '其他'


def extract_deductions_with_subjects(df, company_name):
    """提取所有扣分项并归类到科目"""
    df = df.copy()
    
    deduction_cols = []
    for col in df.columns:
        numeric_vals = pd.to_numeric(df[col], errors='coerce')
        if (numeric_vals < 0).any():
            deduction_cols.append(col)
    
    total_deduction = pd.Series(0, index=df.index)
    
    for subject in SUBJECTS.keys():
        df[f'{subject}_扣分'] = 0
    df['综合考评_扣分'] = 0
    df['其他_扣分'] = 0
    
    for col in deduction_cols:
        numeric_vals = pd.to_numeric(df[col], errors='coerce').fillna(0)
        col_deduction = numeric_vals * (numeric_vals < 0)
        total_deduction += col_deduction
        
        subject = categorize_deduction_item(col)
        if subject in SUBJECTS.keys():
            df[f'{subject}_扣分'] += col_deduction
        elif subject == '综合考评':
            df['综合考评_扣分'] += col_deduction
        else:
            df['其他_扣分'] += col_deduction
    
    df['扣分总和'] = total_deduction
    df['最终得分'] = 100 + df['扣分总和']
    df['扣分项数量'] = (df[deduction_cols] < 0).sum(axis=1) if deduction_cols else 0
    
    for subject in SUBJECTS.keys():
        df[f'{subject}_得分'] = 100 + df[f'{subject}_扣分']
    df['综合考评_得分'] = 100 + df['综合考评_扣分']
    
    return df, deduction_cols


def calculate_subject_stats(df):
    """计算各科目统计数据"""
    subject_stats = {}
    
    for subject in SUBJECTS.keys():
        if f'{subject}_得分' in df.columns:
            subject_stats[subject] = {
                '平均得分': df[f'{subject}_得分'].mean(),
                '平均扣分': -df[f'{subject}_扣分'].mean(),
                '最低分': df[f'{subject}_得分'].min(),
                '最高分': df[f'{subject}_得分'].max(),
                '扣分人数': (df[f'{subject}_扣分'] < 0).sum(),
                '总扣分': -df[f'{subject}_扣分'].sum()
            }
    
    return subject_stats


def analyze_operator_performance(df):
    """按操纵者分析"""
    if '技术等级' not in df.columns:
        return None
    
    operator_stats = {}
    
    captains = df[df['技术等级'].str.contains('机长|教员', na=False)]
    fos = df[df['技术等级'].str.contains('副驾驶', na=False)]
    
    operator_stats['机长'] = {
        '人数': len(captains),
        '平均分': captains['最终得分'].mean() if len(captains) > 0 else 0,
        '最高分': captains['最终得分'].max() if len(captains) > 0 else 0,
        '最低分': captains['最终得分'].min() if len(captains) > 0 else 0
    }
    
    operator_stats['副驾驶'] = {
        '人数': len(fos),
        '平均分': fos['最终得分'].mean() if len(fos) > 0 else 0,
        '最高分': fos['最终得分'].max() if len(fos) > 0 else 0,
        '最低分': fos['最终得分'].min() if len(fos) > 0 else 0
    }
    
    return operator_stats


def analyze_comprehensive_issues(df):
    """分析综合考评问题"""
    comprehensive_data = []
    
    for col in df.columns:
        numeric_vals = pd.to_numeric(df[col], errors='coerce')
        deduction_count = (numeric_vals < 0).sum()
        if deduction_count > 0:
            for item in COMPREHENSIVE_ITEMS:
                if item.lower() in col.lower():
                    comprehensive_data.append({
                        '问题类型': item,
                        '出现次数': deduction_count
                    })
                    break
    
    if comprehensive_data:
        df_issues = pd.DataFrame(comprehensive_data)
        return df_issues.groupby('问题类型').sum().reset_index().sort_values('出现次数', ascending=False)
    return pd.DataFrame()


def get_top_deductions_by_subject(df, subject, n=3):
    """获取某科目TOP扣分项"""
    deductions = []
    for col in df.columns:
        numeric_vals = pd.to_numeric(df[col], errors='coerce')
        if (numeric_vals < 0).any():
            if categorize_deduction_item(col) == subject:
                deductions.append({
                    '扣分项': col,
                    '出现次数': (numeric_vals < 0).sum(),
                    '总扣分': -numeric_vals[numeric_vals < 0].sum()
                })
    
    return sorted(deductions, key=lambda x: x['出现次数'], reverse=True)[:n]


# ------------------ 可视化函数 ------------------
def create_company_radar(company_stats):
    """公司雷达图"""
    if company_stats.empty:
        return None
    
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=company_stats['平均分'].tolist(),
        theta=company_stats['单位名称'].tolist(),
        fill='toself',
        name='平均分',
        line_color='#2ecc71'
    ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        title="各单位平均分雷达图",
        height=450
    )
    return fig


def create_operator_comparison_chart(operator_stats):
    """机长副驾驶对比图"""
    if not operator_stats:
        return None
    
    df = pd.DataFrame([
        {'操纵者': role, '平均分': stats['平均分'], '人数': stats['人数']}
        for role, stats in operator_stats.items() if stats['人数'] > 0
    ])
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df['操纵者'], y=df['平均分'],
        text=df['平均分'].round(1), textposition='auto',
        marker_color=['#2ecc71', '#3498db']
    ))
    fig.update_layout(
        title="机长与副驾驶得分对比",
        xaxis_title="操纵者", yaxis_title="平均得分",
        height=400
    )
    return fig


# ------------------ 主页面 ------------------
st.title("✈️ 飞行员双盲测试评估分析平台")
st.markdown("基于《华东飞行员双盲测试数据采集表》统一模板，客观扣分计算")

with st.sidebar:
    st.markdown("### 📂 数据上传")
    uploaded_files = st.file_uploader(
        "上传一个或多个Excel文件（每个文件代表一个公司）", 
        type=["xlsx"], 
        accept_multiple_files=True
    )
    
    st.markdown("---")
    st.markdown("### 📌 评分规则")
    st.markdown("""
    - 基础分: **100分**
    - 每出现一次扣分项，扣除相应分数
    - 最终得分 = 100 + 扣分总和
    """)

if uploaded_files:
    all_data_list = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, file in enumerate(uploaded_files):
        status_text.text(f"正在处理: {file.name}")
        df = load_template(file)
        company_name = file.name.replace('.xlsx', '').replace('_', ' ')
        df, _ = extract_deductions_with_subjects(df, company_name)
        df['所属单位'] = company_name
        all_data_list.append(df)
        progress_bar.progress((i + 1) / len(uploaded_files))
    
    status_text.text("数据处理完成！")
    all_data = pd.concat(all_data_list, ignore_index=True)
    
    # 计算统计
    subject_stats = calculate_subject_stats(all_data)
    operator_stats = analyze_operator_performance(all_data)
    comprehensive_issues = analyze_comprehensive_issues(all_data)
    company_stats = all_data.groupby('所属单位')['最终得分'].agg(
        人数='count', 平均分='mean', 最高分='max', 最低分='min'
    ).round(2).reset_index()
    company_stats.columns = ['单位名称', '人数', '平均分', '最高分', '最低分']
    
    # 筛选器
    companies = ["全部"] + sorted(all_data['所属单位'].unique().tolist())
    
    with st.sidebar:
        selected_company = st.selectbox("单位筛选", companies, key="company_filter")
    
    filtered_data = all_data.copy()
    if selected_company != "全部":
        filtered_data = filtered_data[filtered_data['所属单位'] == selected_company]
    
    st.success(f"✅ 成功加载 {len(all_data)} 条飞行员记录，来自 {len(company_stats)} 个单位")
    
    # 主界面
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 整体情况", "🏢 单位对比", "📈 科目分析", "📄 报告输出"
    ])
    
    with tab1:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("参加单位数", len(company_stats))
        col2.metric("参加人数", len(all_data))
        col3.metric("整体平均分", f"{all_data['最终得分'].mean():.1f}")
        col4.metric("最高/最低分", f"{all_data['最终得分'].max():.0f}/{all_data['最终得分'].min():.0f}")
        
        if operator_stats:
            fig_operator = create_operator_comparison_chart(operator_stats)
            if fig_operator:
                st.plotly_chart(fig_operator, use_container_width=True)
        
        fig_hist = px.histogram(all_data, x='最终得分', nbins=20, title="整体得分分布",
                                labels={'最终得分': '得分', 'count': '人数'},
                                color_discrete_sequence=['#2ecc71'])
        st.plotly_chart(fig_hist, use_container_width=True)
        
        st.markdown("### 统计摘要")
        stats_df = pd.DataFrame({
            "指标": ["平均分", "中位数", "标准差", "最高分", "最低分"],
            "数值": [
                f"{all_data['最终得分'].mean():.2f}",
                f"{all_data['最终得分'].median():.2f}",
                f"{all_data['最终得分'].std():.2f}",
                f"{all_data['最终得分'].max():.0f}",
                f"{all_data['最终得分'].min():.0f}"
            ]
        })
        st.dataframe(stats_df, use_container_width=True)
    
    with tab2:
        st.markdown("### 各单位对比分析")
        st.dataframe(company_stats, use_container_width=True, hide_index=True)
        
        fig_bar = px.bar(company_stats, x='单位名称', y='平均分', title="各单位平均分对比",
                        text='平均分', color_discrete_sequence=['#2ecc71'])
        fig_bar.update_traces(textposition='outside')
        st.plotly_chart(fig_bar, use_container_width=True)
        
        fig_radar = create_company_radar(company_stats)
        if fig_radar:
            st.plotly_chart(fig_radar, use_container_width=True)
    
    with tab3:
        st.markdown("### 科目分析")
        
        if subject_stats:
            subject_df = pd.DataFrame([
                {'科目': k, '平均分': v['平均得分'], '平均扣分': v['平均扣分'], '扣分人数': v['扣分人数']}
                for k, v in subject_stats.items()
            ])
            st.dataframe(subject_df, use_container_width=True, hide_index=True)
            
            # 各科目详细分析
            for subject in SUBJECTS.keys():
                with st.expander(f"📌 {subject}科目分析"):
                    if subject in subject_stats:
                        stats = subject_stats[subject]
                        c1, c2, c3 = st.columns(3)
                        c1.metric("平均得分", f"{stats['平均得分']:.1f}")
                        c2.metric("平均扣分", f"{stats['平均扣分']:.1f}")
                        c3.metric("扣分人数", stats['扣分人数'])
                        
                        top3 = get_top_deductions_by_subject(all_data, subject, 3)
                        if top3:
                            st.markdown("**高频扣分项TOP3**")
                            st.dataframe(pd.DataFrame(top3), use_container_width=True, hide_index=True)
        
        if not comprehensive_issues.empty:
            st.markdown("### 综合考评分析")
            st.dataframe(comprehensive_issues, use_container_width=True, hide_index=True)
            st.info("综合考评部分，技术性复飞占比最多，多数由于前期未建立适合的落地条件所致。")
    
    with tab4:
        st.markdown("### 📄 生成分析报告")
        
        report_text = f"""
        ═══════════════════════════════════════════════════════════
                       双盲测试数据分析报告
        ═══════════════════════════════════════════════════════════
        
        生成时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
        
        一、整体情况
        
        1. 参加测试人数
        辖区内C909机型参加本次双盲测试的共有{len(company_stats)}家单位，共计{len(all_data)}人参加了测试。
        
        2. 测试分数
        本次测试满分100分，整体平均分: {all_data['最终得分'].mean():.2f}分
        最高分: {all_data['最终得分'].max():.0f}分，最低分: {all_data['最终得分'].min():.0f}分
        
        二、各单位概况
        """
        for _, row in company_stats.iterrows():
            report_text += f"\n  {row['单位名称']}: {row['人数']}人, 平均分{row['平均分']:.1f}分"
        
        st.text_area("报告预览", report_text, height=400)
        
        csv_data = all_data.to_csv(index=False, encoding='utf-8-sig')
        st.download_button("📥 下载CSV数据", csv_data, "飞行员评估数据.csv", mime="text/csv")

else:
    st.info("👈 请先从左侧上传一个或多个Excel文件（每个文件代表一个公司）")
    
    with st.expander("📖 文件格式说明"):
        st.markdown("""
        **支持的文件格式**：
        - 按《华东飞行员双盲测试数据采集表》模板填写的Excel文件
        - 每个文件代表一个公司的数据
        """)

st.markdown("---")
st.caption("飞行员双盲测试评估分析 | 基于华东局飞行技能评估标准")