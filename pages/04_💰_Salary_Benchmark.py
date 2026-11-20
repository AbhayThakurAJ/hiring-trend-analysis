import streamlit as st
import plotly.express as px
from src.analyser import get_salary_benchmarks_by_industry, get_salary_by_exp_and_city
from src.ui import render_hero_header, render_insight_box

st.set_page_config(page_title="Salary Benchmark", page_icon="💰", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None or len(df) == 0:
    st.warning("No data matches current filters. Please adjust the sidebar filters.")
    st.stop()

render_hero_header("Compensation & Salary Benchmarking", "Evaluate market compensation by industry, location, and seniority tier (OQ-02, OQ-07 / BR-03, BR-06)", "UI-04")

tab1, tab2 = st.tabs(["🏢 Cross-Industry Comparison", "📍 Experience & Location Slicer"])

with tab1:
    st.subheader("Industry Compensation Spectrum (Min, Mean, Max)")
    ind_sal_df = get_salary_benchmarks_by_industry(df)
    
    fig = px.bar(
        ind_sal_df,
        x='industry_name',
        y=['min_salary', 'avg_salary', 'max_salary'],
        barmode='group',
        labels={'value': 'Salary (₹ INR)', 'industry_name': 'Industry Sector'},
        color_discrete_map={'min_salary': '#94A3B8', 'avg_salary': '#2E74B5', 'max_salary': '#1E293B'}
    )
    st.plotly_chart(fig, use_container_width=True)
    
    st.subheader("Salary Spread & Quartiles (Box Plot)")
    fig_box = px.box(df, x='industry_name', y='salary_max', color='industry_name', points='outliers')
    fig_box.update_layout(showlegend=False, xaxis_title="", yaxis_title="Max Offered Salary (₹ INR)")
    st.plotly_chart(fig_box, use_container_width=True)

with tab2:
    st.subheader("Filtered Salary Cohort")
    exp_filter = st.selectbox("Filter Seniority", ["All", "Fresher", "Mid-level", "Senior"])
    city_filter = st.selectbox("Filter Metro", ["All"] + sorted(list(df['city_name'].unique())))
    
    cohort_df = get_salary_by_exp_and_city(df, exp_level=exp_filter, city=city_filter)
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Min CTC", f"₹{cohort_df['salary_min'].min()/100000:.1f} LPA" if len(cohort_df)>0 else "N/A")
    col2.metric("Average CTC", f"₹{cohort_df['salary_max'].mean()/100000:.1f} LPA" if len(cohort_df)>0 else "N/A")
    col3.metric("Max CTC", f"₹{cohort_df['salary_max'].max()/100000:.1f} LPA" if len(cohort_df)>0 else "N/A")
    
    st.dataframe(cohort_df, use_container_width=True)

csv = ind_sal_df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Industry Salary Data (CSV)", csv, "salary_benchmarks.csv", "text/csv")
