import streamlit as st
import plotly.express as px
from src.analyser import get_role_demand
from src.ui import render_hero_header, render_insight_box

st.set_page_config(page_title="Role Demand Analysis", page_icon="💼", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None or len(df) == 0:
    st.warning("No data matches current filters. Please adjust the sidebar filters.")
    st.stop()

render_hero_header("Job Role Demand & Volume Ranking", "Identify which roles command highest posting activity (OQ-01 / BR-01)", "UI-02")

top_n = st.slider("Select Number of Top Roles to Display", min_value=5, max_value=12, value=10)
roles_df = get_role_demand(df, top_n=top_n)

fig = px.bar(
    roles_df,
    x='posting_count',
    y='role_title',
    orientation='h',
    color='posting_count',
    color_continuous_scale='Blues',
    text='posting_count'
)
fig.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Total Job Postings", yaxis_title="")
st.plotly_chart(fig, use_container_width=True)

top_role_name = roles_df.iloc[0]['role_title'] if len(roles_df) > 0 else "N/A"
top_role_vol = roles_df.iloc[0]['posting_count'] if len(roles_df) > 0 else 0
render_insight_box(f"'{top_role_name}' dominates current hiring demand with {top_role_vol} postings ({top_role_vol/len(df)*100:.1f}% of total active openings).", "Role Concentration")

st.markdown("### 📋 Role Demand Summary Table")
st.dataframe(roles_df, use_container_width=True)

csv = roles_df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download Role Demand Data (CSV)", csv, "role_demand.csv", "text/csv")
