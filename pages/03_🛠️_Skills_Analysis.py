import streamlit as st
import plotly.express as px
from src.analyser import get_skill_frequencies
from src.ui import render_hero_header, render_insight_box

st.set_page_config(page_title="Skills Demand Analysis", page_icon="🛠️", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None or len(df) == 0:
    st.warning("No data matches current filters. Please adjust the sidebar filters.")
    st.stop()

render_hero_header("Technical Skills Demand Frequency", "Identify in-demand competencies to optimize job descriptions & curriculum (OQ-03 / BR-02)", "UI-03")

top_n = st.slider("Select Number of Top Skills to Display", min_value=5, max_value=20, value=10)
skills_df = get_skill_frequencies(df, top_n=top_n)

fig = px.bar(
    skills_df,
    x='frequency',
    y='skill',
    orientation='h',
    color='frequency',
    color_continuous_scale='Teal',
    text='frequency'
)
fig.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Times Requested Across Postings", yaxis_title="")
st.plotly_chart(fig, use_container_width=True)

if len(skills_df) > 0:
    top_skill = skills_df.iloc[0]['skill']
    top_freq = skills_df.iloc[0]['frequency']
    render_insight_box(f"'{top_skill}' is the most requested technical requirement appearing in {top_freq} postings ({top_freq/len(df)*100:.1f}% penetration).", "Top Skill Insight")

st.markdown("### 📋 Skill Frequency Breakdown")
st.dataframe(skills_df, use_container_width=True)

csv = skills_df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Skills Breakdown (CSV)", csv, "skills_demand.csv", "text/csv")
