"""
src/validator.py — Schema & Business Rule Enforcement
Enforces BR-R-01 through BR-R-10.
"""
import pandas as pd

REQUIRED_COLUMNS = [
    'job_id', 'role_title', 'company_name', 'industry_name', 'city_name',
    'job_type', 'experience_level', 'salary_min', 'salary_max', 'required_skills',
    'skills_count', 'posted_date', 'posted_month', 'posted_year'
]
VALID_EXP_LEVELS = {'Fresher', 'Mid-level', 'Senior'}
VALID_JOB_TYPES = {'Full-time', 'Part-time', 'Remote', 'Hybrid', 'Contract'}

def validate_dataframe(df: pd.DataFrame) -> None:
    """
    Validate that the DataFrame satisfies foundational data constraints.
    Raises ValueError on breach.
    """
    if len(df) < 500:
        raise ValueError(f"Dataset has {len(df)} rows; minimum 500 records required (BR-R-10).")
    
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Required schema columns missing: {missing_cols}")
    
    # Check experience level values
    invalid_exp = set(df['experience_level'].dropna().unique()) - VALID_EXP_LEVELS
    if invalid_exp:
        raise ValueError(f"Invalid experience levels detected: {invalid_exp} (BR-R-04)")
        
    # Check job type values
    invalid_jt = set(df['job_type'].dropna().unique()) - VALID_JOB_TYPES
    if invalid_jt:
        raise ValueError(f"Invalid job types detected: {invalid_jt} (BR-R-05)")
        
    return True
