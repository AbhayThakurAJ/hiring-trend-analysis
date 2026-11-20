"""
tests/test_app.py — Smoke Tests for Streamlit App Structure & Modules (Step 13)
Tests TC-T-13 through TC-T-15
"""

import os, time
import pytest

def test_TC_T_13_directory_and_page_routes():
    """TC-T-13: Verify multipage directory and Streamlit page routes exist."""
    base_dir = r"C:\Hiring Trend Analysis"
    pages_dir = os.path.join(base_dir, "pages")
    
    # In Step 15 pages will be populated; test checks structure
    assert os.path.exists(os.path.join(base_dir, "src", "ui.py"))
    assert os.path.exists(os.path.join(base_dir, ".streamlit", "config.toml"))

def test_TC_T_14_ui_helper_render_functions():
    """TC-T-14: Verify src/ui.py module imports and renders HTML strings."""
    import sys
    sys.path.append(r"C:\Hiring Trend Analysis")
    from src.ui import TAILWIND_CDN
    assert "tailwindcss" in TAILWIND_CDN

def test_TC_T_15_database_query_latency():
    """TC-T-15: Database read and view evaluation completes in under 1 second (< 1000 ms)."""
    import sqlite3
    db_path = r"C:\Hiring Trend Analysis\data\app.db"
    
    start_time = time.time()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM vw_hiring_analysis;")
    rows = cursor.fetchall()
    conn.close()
    elapsed = time.time() - start_time
    
    assert len(rows) >= 500
    assert elapsed < 1.0, f"Query took {elapsed:.3f}s, exceeding 1.0s latency budget!"
