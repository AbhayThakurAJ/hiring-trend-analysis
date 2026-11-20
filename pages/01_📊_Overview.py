import streamlit as st
import plotly.express as px
from src.analyser import get_kpi_metrics, get_role_demand, get_monthly_trends
from src.ui import render_hero_header, render_kpi_cards, render_insight_box

st.set_page_config(page_title="Executive Overview", page_icon="📊", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None:
    st.warning("Please launch the application from `app.py` first to initialize dataset filters.")
    st.stop()

render_hero_header("Executive Overview & Macro KPIs", "Aggregated talent metrics across Indian recruitment corridors", "UI-01")

kpis = get_kpi_metrics(df)
render_kpi_cards(kpis['total_jobs'], kpis['avg_salary'], kpis['top_city'], kpis['top_role'])

col1, col2 = st.columns(2)
with col1:
    st.subheader("📈 Hiring Trajectory Over Time")
    trend_df = get_monthly_trends(df)
    if len(trend_df) > 0:
        fig = px.line(trend_df, x='posted_month', y='posting_count', markers=True, color_discrete_sequence=['#1F3564'])
        fig.update_layout(xaxis_title="Calendar Month", yaxis_title="Open Openings", margin=dict(l=50, r=20, t=25, b=40))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("No postings match selected filters.")

with col2:
    st.subheader("💼 Top Demanded Roles")
    roles_df = get_role_demand(df, top_n=6)
    if len(roles_df) > 0:
        fig = px.bar(roles_df, x='posting_count', y='role_title', orientation='h', color='role_title', color_discrete_sequence=['#2563EB'])
        fig.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Postings Count", yaxis_title="", margin=dict(l=150, r=20, t=25, b=40), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("No postings match selected filters.")

st.markdown("### 📋 Recent Job Postings")
st.dataframe(df[['job_id', 'role_title', 'company_name', 'industry_name', 'city_name', 'job_type', 'salary_max', 'posted_date']].head(10), use_container_width=True)

csv = df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Filtered Dataset (CSV)", csv, "hiring_trends_filtered.csv", "text/csv")
