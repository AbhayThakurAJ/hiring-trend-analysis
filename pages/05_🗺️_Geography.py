import streamlit as st
import plotly.express as px
from src.analyser import get_geographic_hotspots
from src.ui import render_hero_header, render_insight_box

st.set_page_config(page_title="Geographic Hotspots", page_icon="🗺️", layout="wide")

df = st.session_state.get('filtered_df', None)
if df is None or len(df) == 0:
    st.warning("No data matches current filters. Please adjust the sidebar filters.")
    st.stop()

render_hero_header("Geographic Talent Concentration & Hotspots", "Identify prime hiring hubs across Indian metropolitan centers (OQ-05 / BR-07)", "UI-05")

city_df = get_geographic_hotspots(df, top_n=15)

fig = px.bar(
    city_df,
    x='posting_count',
    y='city_name',
    orientation='h',
    color='posting_count',
    color_continuous_scale='Purples',
    text='posting_count'
)
fig.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Total Openings", yaxis_title="City")
st.plotly_chart(fig, use_container_width=True)

top_city = city_df.iloc[0]['city_name'] if len(city_df) > 0 else "N/A"
top_city_vol = city_df.iloc[0]['posting_count'] if len(city_df) > 0 else 0
render_insight_box(f"{top_city} leads national hiring volume with {top_city_vol} postings, accounting for {top_city_vol/len(df)*100:.1f}% of aggregate opportunities.", "Regional Dominance")

st.markdown("### 📋 Metro Concentration Summary")
st.dataframe(city_df, use_container_width=True)

csv = city_df.to_csv(index=False).encode('utf-8')
st.download_button("📥 Download City Breakdown (CSV)", csv, "city_hotspots.csv", "text/csv")
