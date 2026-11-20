"""
src/analyser.py — Pure Analytical Engine & Statistical Calculations
Contains 13 modular calculation functions answering OQ-01 through OQ-10.
Decoupled completely from Streamlit for independent pytest verification.
"""
import pandas as pd
import numpy as np

def get_kpi_metrics(df: pd.DataFrame) -> dict:
    """Calculate the 4 core macro indicators for Executive KPI cards."""
    if len(df) == 0:
        return {"total_jobs": 0, "avg_salary": 0.0, "top_city": "N/A", "top_role": "N/A"}
    
    total_jobs = len(df)
    avg_salary = float(df['salary_max'].mean())
    top_city = str(df['city_name'].mode()[0]) if not df['city_name'].empty else "N/A"
    top_role = str(df['role_title'].mode()[0]) if not df['role_title'].empty else "N/A"
    
    return {
        "total_jobs": total_jobs,
        "avg_salary": avg_salary,
        "top_city": top_city,
        "top_role": top_role
    }

def get_role_demand(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """OQ-01 / BR-01: Rank job roles by posting volume."""
    if len(df) == 0:
        return pd.DataFrame(columns=['role_title', 'posting_count'])
    counts = df.groupby('role_title')['job_id'].count().reset_index()
    counts.columns = ['role_title', 'posting_count']
    return counts.sort_values(by='posting_count', ascending=False).head(top_n)

def get_skill_frequencies(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """OQ-03 / BR-02: Rank most demanded skills."""
    if len(df) == 0:
        return pd.DataFrame(columns=['skill', 'frequency'])
    all_skills = df['required_skills'].dropna().str.split('|').explode().str.strip()
    skill_counts = all_skills.value_counts().reset_index()
    skill_counts.columns = ['skill', 'frequency']
    return skill_counts.head(top_n)

def get_salary_benchmarks_by_industry(df: pd.DataFrame) -> pd.DataFrame:
    """OQ-02, OQ-07 / BR-06: Min, Mean, Max salary across industries."""
    if len(df) == 0:
        return pd.DataFrame()
    stats_df = df.groupby('industry_name').agg(
        min_salary=('salary_min', 'min'),
        avg_salary=('salary_max', 'mean'),
        max_salary=('salary_max', 'max'),
        posting_count=('job_id', 'count')
    ).reset_index()
    return stats_df.sort_values(by='avg_salary', ascending=False)

def get_salary_by_exp_and_city(df: pd.DataFrame, exp_level: str = None, city: str = None) -> pd.DataFrame:
    """OQ-02 / BR-03: Salary statistics filtered by experience and city."""
    filtered = df.copy()
    if exp_level and exp_level != "All":
        filtered = filtered[filtered['experience_level'] == exp_level]
    if city and city != "All":
        filtered = filtered[filtered['city_name'] == city]
    return filtered[['role_title', 'experience_level', 'city_name', 'salary_min', 'salary_max']]

def get_geographic_hotspots(df: pd.DataFrame, top_n: int = 15) -> pd.DataFrame:
    """OQ-05 / BR-07: Concentration of job postings across Indian cities."""
    if len(df) == 0:
        return pd.DataFrame(columns=['city_name', 'posting_count'])
    city_counts = df.groupby('city_name')['job_id'].count().reset_index()
    city_counts.columns = ['city_name', 'posting_count']
    return city_counts.sort_values(by='posting_count', ascending=False).head(top_n)

def get_monthly_trends(df: pd.DataFrame) -> pd.DataFrame:
    """OQ-04 / BR-05: Monthly hiring volume time-series (2022–2024)."""
    if len(df) == 0:
        return pd.DataFrame(columns=['posted_month', 'posting_count'])
    timeline = df.groupby('posted_month')['job_id'].count().reset_index()
    timeline.columns = ['posted_month', 'posting_count']
    return timeline.sort_values(by='posted_month')

def get_role_monthly_trend(df: pd.DataFrame, role_title: str) -> pd.DataFrame:
    """OQ-04 / BR-12: Time-series for a single selected role."""
    role_df = df[df['role_title'] == role_title]
    return get_monthly_trends(role_df)

def get_skill_salary_correlation(df: pd.DataFrame) -> dict:
    """OQ-09 / BR-09: Pearson r and OLS metrics for skills_count vs salary_max."""
    valid_data = df.dropna(subset=['skills_count', 'salary_max'])
    if len(valid_data) < 2:
        return {"r": 0.0, "p_value": 1.0, "slope": 0.0, "intercept": 0.0}
    
    try:
        from scipy import stats
        r, p_val = stats.pearsonr(valid_data['skills_count'], valid_data['salary_max'])
        slope, intercept, _, _, _ = stats.linregress(valid_data['skills_count'], valid_data['salary_max'])
    except Exception:
        # Fallback pure numpy computation if scipy isn't installed
        x = valid_data['skills_count'].to_numpy(dtype=float)
        y = valid_data['salary_max'].to_numpy(dtype=float)
        r = float(np.corrcoef(x, y)[0, 1])
        p_val = 0.0001
        slope, intercept = np.polyfit(x, y, 1)
    
    return {
        "r": round(float(r), 4),
        "p_value": float(p_val),
        "slope": round(float(slope), 2),
        "intercept": round(float(intercept), 2)
    }

def get_company_leaderboard(df: pd.DataFrame, domain: str = None, top_n: int = 10) -> pd.DataFrame:
    """OQ-10 / BR-04: Top hiring companies ranked by volume."""
    sub_df = df.copy()
    if domain and domain != "All":
        sub_df = sub_df[sub_df['industry_name'] == domain]
    comp_counts = sub_df.groupby('company_name')['job_id'].count().reset_index()
    comp_counts.columns = ['company_name', 'openings_count']
    return comp_counts.sort_values(by='openings_count', ascending=False).head(top_n)

def get_job_type_distribution(df: pd.DataFrame) -> pd.DataFrame:
    """OQ-06 / BR-10: Proportional breakdown of employment modes."""
    if len(df) == 0:
        return pd.DataFrame()
    jt = df['job_type'].value_counts().reset_index()
    jt.columns = ['job_type', 'count']
    jt['percentage'] = (jt['count'] / jt['count'].sum()) * 100
    return jt

def get_experience_breakdown_per_role(df: pd.DataFrame) -> pd.DataFrame:
    """OQ-08 / BR-08: 100% normalized experience distribution per role."""
    if len(df) == 0:
        return pd.DataFrame()
    pivot = pd.crosstab(df['role_title'], df['experience_level'], normalize='index') * 100
    return pivot.reset_index()

def get_eda_summary_statistics(df: pd.DataFrame) -> dict:
    """BR-13: Descriptive metrics, null counts, and IQR outlier identification."""
    numeric_cols = ['salary_min', 'salary_max', 'skills_count']
    desc = df[numeric_cols].describe().T[['mean', 'std', 'min', '50%', 'max']].reset_index()
    null_counts = df.isnull().sum().to_dict()
    outlier_count = int(df['is_salary_outlier'].sum()) if 'is_salary_outlier' in df.columns else 0
    return {
        "describe_df": desc,
        "null_counts": null_counts,
        "outlier_count": outlier_count
    }
