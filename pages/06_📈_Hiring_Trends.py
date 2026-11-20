import streamlit as st
import plotly.express as px
from src.analyser import get_monthly_trends, get_role_monthly_trend
from src.ui import render_hero_header, render_insight_box

st.set_page_config(page_title="Hiring Trends Over Time", page_icon="📈", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None or len(df) == 0:
    st.warning("No data matches current filters. Please adjust the sidebar filters.")
    st.stop()

render_hero_header("Longitudinal Hiring Trends & Seasonality (2022–2024)", "Examine monthly volume cycles to time recruitment campaigns and job applications (OQ-04 / BR-05, BR-12)", "UI-06")

tab1, tab2 = st.tabs(["📊 Aggregate Monthly Timeline", "🎯 Single Role Trajectory"])

with tab1:
    timeline_df = get_monthly_trends(df)
    fig = px.line(
        timeline_df,
        x='posted_month',
        y='posting_count',
        markers=True,
        line_shape='spline',
        color_discrete_sequence=['#2E74B5']
    )
    fig.update_layout(xaxis_title="Calendar Month", yaxis_title="Active Job Postings")
    st.plotly_chart(fig, use_container_width=True)
    
    render_insight_box("Recruitment displays distinct seasonal surges in Q1 (Jan–Mar) and Q3 (Jul–Sep) driven by corporate budget releases and campus placement drives.", "Seasonality Insight")

with tab2:
    role_options = sorted(list(df['role_title'].unique()))
    chosen_role = st.selectbox("Select Target Role for Trendline", role_options)
    
    role_trend_df = get_role_monthly_trend(df, chosen_role)
    fig_role = px.bar(
        role_trend_df,
        x='posted_month',
        y='posting_count',
        color='posting_count',
        color_continuous_scale='Blues'
    )
    fig_role.update_layout(xaxis_title="Month", yaxis_title="Openings")
    st.plotly_chart(fig_role, use_container_width=True)

csv = timeline_df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Monthly Trend Data (CSV)", csv, "hiring_trends_monthly.csv", "text/csv")
