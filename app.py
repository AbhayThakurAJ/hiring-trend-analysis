"""
app.py — Primary Entrypoint & Global State Management
Handles sidebar filters, global session state, and redirects to 01_Overview.
"""
import streamlit as st
import pandas as pd
from src.data_loader import load_data
from src.cleaner import clean_data
from src.validator import validate_dataframe

st.set_page_config(
    page_title="Hiring Trend Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 1. Load and Clean Raw Data
try:
    raw_df = load_data()
    validate_dataframe(raw_df)
    df = clean_data(raw_df)
except Exception as e:
    st.error(f"⚠️ Data Engine Error: {str(e)}")
    st.stop()

# 2. Initialize Session State Filter Defaults
if "date_range" not in st.session_state:
    st.session_state.date_range = (df['posted_date'].min().date(), df['posted_date'].max().date())
if "selected_industries" not in st.session_state:
    st.session_state.selected_industries = []
if "selected_cities" not in st.session_state:
    st.session_state.selected_cities = []
if "selected_exp" not in st.session_state:
    st.session_state.selected_exp = []

# 3. Sidebar Filter Controls
st.sidebar.title("🎛️ Global Slicers")
st.sidebar.caption("Filters automatically synchronize across all 10 pages.")

# Date Range Slicer
min_date = df['posted_date'].min().date()
max_date = df['posted_date'].max().date()
st.session_state.date_range = st.sidebar.date_input(
    "📅 Posting Date Window",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

# Industry Slicer
industries_list = sorted(list(df['industry_name'].unique()))
st.session_state.selected_industries = st.sidebar.multiselect(
    "🏢 Industry Sector",
    options=industries_list,
    default=[]
)

# City Slicer
cities_list = sorted(list(df['city_name'].unique()))
st.session_state.selected_cities = st.sidebar.multiselect(
    "📍 Metro Location",
    options=cities_list,
    default=[]
)

# Experience Slicer
exp_list = ["Fresher", "Mid-level", "Senior"]
st.session_state.selected_exp = st.sidebar.multiselect(
    "🎓 Experience Tier",
    options=exp_list,
    default=[]
)

# Reset Button
if st.sidebar.button("🔄 Reset Global Filters", use_container_width=True):
    st.session_state.selected_industries = []
    st.session_state.selected_cities = []
    st.session_state.selected_exp = []
    st.session_state.date_range = (min_date, max_date)
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("**Project P01 — MCA Data Analytics**")
st.sidebar.caption("Student: Abhayjeet Singh | ABES")

# 4. Filter Application Function for Pages
def get_filtered_df(df: pd.DataFrame) -> pd.DataFrame:
    filtered = df.copy()
    
    # Date filtering
    if isinstance(st.session_state.date_range, tuple) and len(st.session_state.date_range) == 2:
        start_d, end_d = st.session_state.date_range
        filtered = filtered[(filtered['posted_date'].dt.date >= start_d) & (filtered['posted_date'].dt.date <= end_d)]
        
    # Industry filtering
    if st.session_state.selected_industries:
        filtered = filtered[filtered['industry_name'].isin(st.session_state.selected_industries)]
        
    # City filtering
    if st.session_state.selected_cities:
        filtered = filtered[filtered['city_name'].isin(st.session_state.selected_cities)]
        
    # Experience filtering
    if st.session_state.selected_exp:
        filtered = filtered[filtered['experience_level'].isin(st.session_state.selected_exp)]
        
    return filtered

# Store in session state for pages to access
st.session_state.filtered_df = get_filtered_df(df)

# Default landing page content
st.info("👈 Use the left sidebar to navigate to any of the 10 analysis screens, or inspect the Overview dashboard below.")

# Render Executive Overview directly on root launch
from src.ui import render_hero_header, render_kpi_cards, render_insight_box
from src.analyser import get_kpi_metrics, get_role_demand, get_monthly_trends
import plotly.express as px

render_hero_header(
    "Hiring Trend Analysis Dashboard",
    "Real-time Indian Recruitment Intelligence, Compensation Benchmarks & Skill Analytics",
    "P01 | ABES MCA 2nd Yr"
)

kpis = get_kpi_metrics(st.session_state.filtered_df)
render_kpi_cards(kpis['total_jobs'], kpis['avg_salary'], kpis['top_city'], kpis['top_role'])

col1, col2 = st.columns(2)
with col1:
    st.subheader("📈 Monthly Hiring Volume")
    trend_df = get_monthly_trends(st.session_state.filtered_df)
    if len(trend_df) > 0:
        fig_trend = px.line(trend_df, x='posted_month', y='posting_count', markers=True, color_discrete_sequence=['#1F3564'])
        fig_trend.update_layout(xaxis_title="Month", yaxis_title="Open Postings", margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.warning("No data matches current filters.")

with col2:
    st.subheader("💼 Top In-Demand Roles")
    roles_df = get_role_demand(st.session_state.filtered_df, top_n=6)
    if len(roles_df) > 0:
        fig_role = px.bar(roles_df, x='posting_count', y='role_title', orientation='h', color='posting_count', color_continuous_scale='Blues')
        fig_role.update_layout(xaxis_title="Postings Count", yaxis_title="", yaxis={'categoryorder':'total ascending'}, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig_role, use_container_width=True)
    else:
        st.warning("No data matches current filters.")

render_insight_box(
    f"Current dataset contains {kpis['total_jobs']} postings with an average maximum compensation of ₹{kpis['avg_salary']/100000:.1f} LPA. {kpis['top_city']} remains the leading tech corridor.",
    "Macro Market Summary"
)
