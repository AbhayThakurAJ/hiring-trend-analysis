import streamlit as st
import plotly.express as px
from src.analyser import get_company_leaderboard
from src.ui import render_hero_header, render_insight_box

st.set_page_config(page_title="Top Recruiting Companies", page_icon="🏢", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None or len(df) == 0:
    st.warning("No data matches current filters. Please adjust the sidebar filters.")
    st.stop()

render_hero_header("Top Recruiting Employers & Industry Leaderboard", "Identify high-volume enterprise recruiters to prioritize placement drives (OQ-10 / BR-04)", "UI-08")

domains_list = ["All"] + sorted(list(df['industry_name'].unique()))
selected_domain = st.selectbox("Filter Company Leaderboard by Industry Domain", domains_list)

comp_df = get_company_leaderboard(df, domain=selected_domain, top_n=10)

fig = px.bar(
    comp_df,
    x='openings_count',
    y='company_name',
    orientation='h',
    color='openings_count',
    color_continuous_scale='Greens',
    text='openings_count'
)
fig.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Total Active Postings", yaxis_title="Enterprise Employer")
st.plotly_chart(fig, use_container_width=True)

if len(comp_df) > 0:
    top_employer = comp_df.iloc[0]['company_name']
    top_count = comp_df.iloc[0]['openings_count']
    render_insight_box(f"'{top_employer}' ranks as the most active employer with {top_count} active listings in this segment.", "Employer Intelligence")

st.markdown("### 📋 Ranked Employer Directory")
st.dataframe(comp_df, use_container_width=True)

csv = comp_df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Leaderboard (CSV)", csv, "company_leaderboard.csv", "text/csv")
