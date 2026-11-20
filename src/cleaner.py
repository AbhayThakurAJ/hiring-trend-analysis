"""
src/cleaner.py — Data Cleaning, Imputation & Feature Engineering
Handles deliberate missing values and outlier normalization for Step 15.
"""
import pandas as pd
import numpy as np

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Perform cleaning operations:
    1. Impute missing salary_min using role-based median or 70% of salary_max
    2. Convert posted_date to datetime
    3. Ensure salary_mid is calculated
    4. Flag IQR outliers in salary_max
    """
    clean_df = df.copy()
    
    # 1. Missing Value Imputation (12 deliberate nulls in salary_min)
    if clean_df['salary_min'].isnull().any():
        clean_df['salary_min'] = clean_df['salary_min'].fillna(clean_df['salary_max'] * 0.70)
        
    # 2. DateTime cast
    clean_df['posted_date'] = pd.to_datetime(clean_df['posted_date'])
    
    # 3. Ensure derived salary_mid
    if 'salary_mid' not in clean_df.columns:
        clean_df['salary_mid'] = (clean_df['salary_min'] + clean_df['salary_max']) / 2.0
        
    # 4. Outlier Flagging via IQR on salary_max
    q1 = clean_df['salary_max'].quantile(0.25)
    q3 = clean_df['salary_max'].quantile(0.75)
    iqr = q3 - q1
    upper_bound = q3 + (1.5 * iqr)
    clean_df['is_salary_outlier'] = clean_df['salary_max'] > upper_bound
    
    return clean_df

def explode_skills(df: pd.DataFrame) -> pd.DataFrame:
    """
    Decompose the pipe-delimited required_skills column into a tidy long-form DataFrame.
    """
    skills_df = df[['job_id', 'role_title', 'industry_name', 'city_name', 'salary_max', 'required_skills']].copy()
    skills_df['skill'] = skills_df['required_skills'].str.split('|')
    skills_exploded = skills_df.explode('skill')
    skills_exploded['skill'] = skills_exploded['skill'].str.strip()
    return skills_exploded
