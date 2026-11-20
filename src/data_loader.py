"""
src/data_loader.py — Cached SQLite Ingestion Module
"""
import os, sqlite3
import pandas as pd
import streamlit as st

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "app.db")

@st.cache_data(ttl=3600, show_spinner="Loading hiring data from SQLite...")
def load_data(db_path: str = DB_PATH) -> pd.DataFrame:
    """
    Load data from the pre-joined analytical view vw_hiring_analysis.
    Cached in memory to ensure sub-second response on subsequent filter clicks.
    """
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"Database file not found at {db_path}. Please run src/generate_data.py first.")
    
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM vw_hiring_analysis;", conn)
    conn.close()
    return df
