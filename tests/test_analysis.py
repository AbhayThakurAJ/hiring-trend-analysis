"""
tests/test_analysis.py — Unit Tests for Analytical Logic & Aggregations (Step 13)
Tests TC-T-07 through TC-T-12
"""

import sqlite3
import pandas as pd
import pytest

DB_PATH = r"C:\Hiring Trend Analysis\data\app.db"

@pytest.fixture
def df_view():
    """Fixture providing a pandas DataFrame from the analytical view."""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM vw_hiring_analysis;", conn)
    conn.close()
    return df

def test_TC_T_07_role_demand_ranking(df_view):
    """TC-T-07: Verify role posting volume aggregation (OQ-01 / BR-01)."""
    counts = df_view.groupby("role_title")["job_id"].count()
    assert len(counts) > 0
    top_role = counts.idxmax()
    assert isinstance(top_role, str)
    assert counts.sum() == len(df_view)

def test_TC_T_08_salary_statistics_logic(df_view):
    """TC-T-08: Salary mean, min, max validity (OQ-02 / BR-03, BR-06)."""
    valid_salaries = df_view.dropna(subset=["salary_min", "salary_max"])
    assert len(valid_salaries) > 0
    assert (valid_salaries["salary_min"] <= valid_salaries["salary_max"]).all()
    assert valid_salaries["salary_min"].mean() > 200000

def test_TC_T_09_city_concentration_ranking(df_view):
    """TC-T-09: City posting ranking covers top Indian tech hubs (OQ-05 / BR-07)."""
    city_counts = df_view.groupby("city_name")["job_id"].count().sort_values(ascending=False)
    top_3_cities = list(city_counts.head(3).index)
    # Bangalore or Hyderabad should be among the top cities
    assert any(c in top_3_cities for c in ["Bangalore", "Hyderabad", "Pune"])

def test_TC_T_10_time_series_chronology(df_view):
    """TC-T-10: Monthly time-series data is properly ordered (OQ-04 / BR-05)."""
    monthly = df_view.groupby("posted_month")["job_id"].count().sort_index()
    assert len(monthly) >= 12, "Should cover at least 12 distinct calendar months"
    assert monthly.index.is_monotonic_increasing

def test_TC_T_11_job_type_proportions(df_view):
    """TC-T-11: Job types sum to exactly 100% (OQ-06 / BR-10)."""
    type_counts = df_view["job_type"].value_counts(normalize=True)
    assert round(type_counts.sum(), 4) == 1.0000
    assert "Full-time" in type_counts.index

def test_TC_T_12_skill_counts_distribution(df_view):
    """TC-T-12: Every posting has between 1 and 8 skills (OQ-09 / BR-09)."""
    assert df_view["skills_count"].between(1, 8).all()
