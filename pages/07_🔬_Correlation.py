import streamlit as st
import plotly.express as px
from src.analyser import get_skill_salary_correlation
from src.ui import render_hero_header, render_insight_box

st.set_page_config(page_title="Skill-Salary Correlation", page_icon="🔬", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None or len(df) == 0:
    st.warning("No data matches current filters. Please adjust the sidebar filters.")
    st.stop()

render_hero_header("Statistical Correlation: Required Skills vs Offered Salary", "Quantify the market premium commanded by multi-skilled technical candidates (OQ-09 / BR-09)", "UI-07")

corr_stats = get_skill_salary_correlation(df)

col1, col2, col3 = st.columns(3)
col1.metric("Pearson Correlation (r)", f"{corr_stats['r']:+.4f}")
col2.metric("p-value", f"{corr_stats['p_value']:.4e}")
col3.metric("Regression Slope (₹ / Skill)", f"+₹{corr_stats['slope']:,.0f}")

fig = px.scatter(
    df,
    x='skills_count',
    y='salary_max',
    color='experience_level',
    trendline='ols',
    labels={'skills_count': 'Number of Skills Required', 'salary_max': 'Maximum Salary Offered (₹ INR)'},
    hover_data=['role_title', 'city_name', 'company_name']
)
fig.update_layout(xaxis=dict(tickmode='linear', tick0=1, dtick=1))
st.plotly_chart(fig, use_container_width=True)

direction = "positive" if corr_stats['r'] > 0 else "negative"
render_insight_box(
    f"Ordinary Least Squares (OLS) regression indicates a statistically significant {direction} correlation (r = {corr_stats['r']:.2f}, p < 0.05). On average, each additional technical skill required adds approximately ₹{corr_stats['slope']/100000:.2f} Lakhs to the top-line CTC.",
    "Econometric Finding"
)

csv = df[['job_id', 'role_title', 'skills_count', 'salary_max', 'experience_level']].to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Correlation Raw Observations (CSV)", csv, "skill_salary_correlation.csv", "text/csv")
