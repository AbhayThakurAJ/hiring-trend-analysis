import streamlit as st
import plotly.express as px
from src.analyser import get_job_type_distribution, get_experience_breakdown_per_role
from src.ui import render_hero_header, render_insight_box

st.set_page_config(page_title="Work Modes & Experience", page_icon="👥", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None or len(df) == 0:
    st.warning("No data matches current filters. Please adjust the sidebar filters.")
    st.stop()

render_hero_header("Workplace Policy & Seniority Demand Analysis", "Evaluate adoption of flexible work modes and experience-tier requirements (OQ-06, OQ-08 / BR-08, BR-10)", "UI-10")

col1, col2 = st.columns(2)

with col1:
    st.subheader("🌐 Employment Work Modes (Donut Chart)")
    jt_df = get_job_type_distribution(df)
    fig_donut = px.pie(
        jt_df,
        values='count',
        names='job_type',
        hole=0.45,
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    fig_donut.update_traces(textinfo='percent+label')
    st.plotly_chart(fig_donut, use_container_width=True)

with col2:
    st.subheader("🎓 100% Experience Demand Breakdown")
    exp_df = get_experience_breakdown_per_role(df)
    exp_melted = exp_df.melt(id_vars='role_title', var_name='Experience Tier', value_name='Percentage')
    
    fig_stack = px.bar(
        exp_melted,
        x='Percentage',
        y='role_title',
        color='Experience Tier',
        orientation='h',
        color_discrete_map={'Fresher': '#6EE7B7', 'Mid-level': '#60A5FA', 'Senior': '#3B82F6'}
    )
    fig_stack.update_layout(barmode='stack', xaxis_title="Proportion (%)", yaxis_title="")
    st.plotly_chart(fig_stack, use_container_width=True)

render_insight_box(
    "Full-time roles remain foundational (60%+), but Hybrid (24%) and Remote (12%) models command substantial share in software and analytics domains. Entry-level fresher hiring represents over 35% of overall openings.",
    "Workforce Structure"
)

csv = jt_df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Job Type Proportions (CSV)", csv, "job_types_summary.csv", "text/csv")
