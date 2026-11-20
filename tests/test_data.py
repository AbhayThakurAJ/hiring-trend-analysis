"""
tests/test_data.py — Unit Tests for Data Integrity & Database Schema (Step 13)
Tests TC-T-01 through TC-T-06
"""

import os, sqlite3
import pandas as pd
import pytest

DB_PATH = r"C:\Hiring Trend Analysis\data\app.db"

@pytest.fixture
def db_conn():
    """Fixture providing an active SQLite database connection."""
    assert os.path.exists(DB_PATH), f"Database file {DB_PATH} not found!"
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    yield conn
    conn.close()

def test_TC_T_01_db_file_exists():
    """TC-T-01: Verify SQLite database file exists and is populated."""
    assert os.path.exists(DB_PATH)
    assert os.path.getsize(DB_PATH) > 10000, "Database file is empty or suspiciously small!"

def test_TC_T_02_table_and_view_presence(db_conn):
    """TC-T-02: Verify all 5 tables and 1 analytical view exist."""
    cursor = db_conn.cursor()
    cursor.execute("SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view');")
    entities = {row[0]: row[1] for row in cursor.fetchall()}
    
    expected_tables = ["dim_industry", "dim_city", "dim_job_role", "fact_job_postings", "job_skills"]
    for tbl in expected_tables:
        assert tbl in entities, f"Expected table {tbl} missing from database!"
    assert "vw_hiring_analysis" in entities, "Analytical view vw_hiring_analysis missing!"

def test_TC_T_03_fact_row_count(db_conn):
    """TC-T-03: Fact table row count must meet or exceed 500 rows."""
    cursor = db_conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM fact_job_postings;")
    count = cursor.fetchone()[0]
    assert count >= 500, f"Expected at least 500 job postings, found {count}!"

def test_TC_T_04_master_tables_counts(db_conn):
    """TC-T-04: Verify master dimension tables have exact planned row counts."""
    cursor = db_conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM dim_industry;")
    assert cursor.fetchone()[0] == 8
    
    cursor.execute("SELECT COUNT(*) FROM dim_city;")
    assert cursor.fetchone()[0] == 15
    
    cursor.execute("SELECT COUNT(*) FROM dim_job_role;")
    assert cursor.fetchone()[0] == 12

def test_TC_T_05_foreign_key_integrity(db_conn):
    """TC-T-05: PRAGMA foreign_key_check must return 0 violations."""
    cursor = db_conn.cursor()
    cursor.execute("PRAGMA foreign_key_check;")
    violations = cursor.fetchall()
    assert len(violations) == 0, f"Referential integrity violations found: {violations}"

def test_TC_T_06_view_column_count(db_conn):
    """TC-T-06: Analytical view must contain at least 25 pre-joined columns."""
    df = pd.read_sql_query("SELECT * FROM vw_hiring_analysis LIMIT 1;", db_conn)
    assert len(df.columns) >= 25, f"Expected >= 25 columns in analytical view, got {len(df.columns)}"
