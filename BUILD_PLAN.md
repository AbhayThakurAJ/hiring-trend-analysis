# 🛠️ BUILD_PLAN.md — End-to-End Build & Implementation Blueprint

**PROJECT:** P071 – Hiring Trend Analysis | **DOMAIN:** Data Analytics  
**BATCH ID:** ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01  
**STUDENT:** Abhayjeet Singh | **ROLL:** 2500320140004 | **MENTOR:** Dinesh Sereno | **INSTITUTION:** ABES MCA 2nd Yr DA  
**STACK:** Python 3.12, Streamlit Multipage, Plotly Express, Tailwind CSS CDN, SQLite 3 (`data/app.db`), Pytest  

---

## 1. Project Summary & Deliverable Traceability

The **Hiring Trend Analysis (P071)** application is a full-stack data analytics product designed to eliminate fragmented spreadsheets and intuition-driven hiring decisions. Developed on a structured SQLite database containing 652 job postings across 15 Indian metropolitan areas and 8 industries, it delivers interactive exploratory data analysis, salary benchmarks, and demand forecasts. The system empowers HR leaders, technical recruiters, placement officers, and job seekers to extract actionable talent market intelligence in real time. It resolves 10 foundational Operational Questions (OQ-01 through OQ-10) with complete bidirectional traceability across all SDLC artifacts. Every user requirement is linked from business needs to database entities, UI screens, and automated test cases.

### Project Document Registry:
- [Step 01: Project Understanding (DOCX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step01_ProjectUnderstanding.docx)
- [Step 02: How-Might-We Statements (DOCX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step02_HowMightWe.docx)
- [Step 03: Needs Analysis & Business Requirements (DOCX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step03_NeedsAnalysis_BR.docx)
- [Step 04: Requirements Traceability Matrix (RTM - XLSX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step04_RTM.xlsx)
- [Step 05: Business Requirements Document (BRD - DOCX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step05_BRD.docx)
- [Step 07: Technical Requirements Document (TRD - DOCX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step07_TechBriefDocument.docx)
- [Step 08: Database Design Workbook (XLSX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step08_DatabaseDesign.xlsx)
- [Step 09: ER Diagram Specification (DOCX & PNG)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step09_ERDiagram.docx)
- [Step 10: Synthetic Dataset (Excel & SQLite)](file:///C:/Hiring%20Trend%20Analysis/data/hiring_trends_data.xlsx)
- [Step 11: User Stories & Sprint Plan (DOCX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step11_UserStories.docx)
- [Step 12: UI Framework & Wireframes (DOCX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step12_UIFramework.docx)
- [Step 13: Test Cases Specification & Pytest Suite (DOCX)](file:///C:/Hiring%20Trend%20Analysis/docs/P01_HTA_Step13_TestCases.docx)

---

## 2. Final Repository & Codebase Structure

```text
C:\Hiring Trend Analysis\
│
├── app.py                          # Primary Streamlit Entrypoint (Filters & Overview Redirection)
│
├── .streamlit/
│   └── config.toml                 # Custom Color Palette & Server Configuration
│
├── pages/
│   ├── 01_📊_Overview.py           # UI-01: Executive KPIs, Hero Header, Recent Highlights
│   ├── 02_💼_Role_Demand.py        # UI-02: Ranked Job Role Demands (OQ-01 / BR-01)
│   ├── 03_🛠️_Skills_Analysis.py    # UI-03: In-Demand Technical Skills (OQ-03 / BR-02)
│   ├── 04_💰_Salary_Benchmark.py   # UI-04: Compensation Stats & Box Plots (OQ-02 / BR-03, 06)
│   ├── 05_🗺️_Geography.py          # UI-05: Urban Concentration & Hotspots (OQ-05 / BR-07)
│   ├── 06_📈_Hiring_Trends.py      # UI-06: Longitudinal Time-Series 2022-2024 (OQ-04 / BR-05, 12)
│   ├── 07_🔬_Correlation.py        # UI-07: Skills vs Salary OLS Regression (OQ-09 / BR-09)
│   ├── 08_🏢_Companies.py          # UI-08: Top Recruiting Employers Leaderboard (OQ-10 / BR-04)
│   ├── 09_🧹_EDA_Quality.py        # UI-09: Data Cleaning, Imputation & Outlier Audit (BR-13)
│   └── 10_👥_Job_Types_Exp.py      # UI-10: Work Modes & Seniority Ratios (OQ-06, 08 / BR-08, 10, 11)
│
├── src/
│   ├── __init__.py
│   ├── data_loader.py              # SQLite Data Access Layer with @st.cache_data
│   ├── cleaner.py                  # Cleaning, Missing Value Imputation & Feature Engineering
│   ├── validator.py                # Schema & Business Rule Enforcement
│   ├── analyser.py                 # Pure Aggregation & Statistical Functions (13 functions)
│   ├── ui.py                       # Tailwind CSS Hero Banners & KPI Cards
│   └── generate_data.py            # Synthetic Data Generation Engine (Step 10)
│
├── data/
│   ├── app.db                      # Primary SQLite Relational Database (652 Postings)
│   ├── hiring_trends_data.xlsx     # Multi-Sheet Excel Master Dataset
│   └── *.csv                       # Individual CSV Table Exports
│
├── tests/
│   ├── conftest.py                 # Shared Pytest Fixtures
│   ├── test_data.py                # TC-T-01..06: Database Schema & Referential Integrity
│   ├── test_analysis.py            # TC-T-07..12: Analytical Formulas & Normalization Logic
│   └── test_app.py                 # TC-T-13..15: Component Imports & Query Latency Benchmark
│
├── docs/                           # Steps 01–18 Project Deliverables (Word, Excel, Images)
├── requirements.txt                # Pinned Dependency File
├── runtime.txt                     # Target Python Runtime (python-3.12)
└── README.md                       # Comprehensive User Guide & Architecture Documentation
```

---

## 3. Step-by-Step Build Tasks (BUILD-01 to BUILD-10)

### 📌 BUILD-01: Environment Configuration & Pinned Dependencies
- **Goal:** Establish a clean Python 3.12 virtual environment and freeze pinned dependencies.
- **Files:** `requirements.txt`, `runtime.txt`, `.streamlit/config.toml`
- **Linked Epics / Stories:** EP-01 / US-01, FR-01
- **Exact Execution Prompt:**
  > "Freeze `requirements.txt` with pinned versions of streamlit, pandas, numpy, plotly, scipy, statsmodels, faker, openpyxl, and pytest. Create `runtime.txt` containing `python-3.12` and confirm `.streamlit/config.toml` has primaryColor='#2E74B5'."
- **Verification:** Run `python -m pip check` to ensure zero conflicting packages.
- **Tests Required:** `tests/test_app.py::test_TC_T_13_directory_and_page_routes`

---

### 📌 BUILD-02: Data Ingestion, Cleaning & Validation Layer
- **Goal:** Implement the data ingestion module reading from SQLite, clean data nuances, and validate schema integrity.
- **Files:** `src/data_loader.py`, `src/cleaner.py`, `src/validator.py`
- **Linked Epics / Stories:** EP-01 / US-01, US-02; FR-01, FR-02
- **Exact Execution Prompt:**
  > "Implement `load_data()` in `src/data_loader.py` with `@st.cache_data`. In `src/cleaner.py`, implement missing value median imputation for `salary_min`, outlier clipping flags, and `explode_skills()`. In `src/validator.py`, enforce BR-R-01 through BR-R-10."
- **Verification:** Execute `python -c "from src.data_loader import load_data; df = load_data(); print(df.shape)"`.
- **Tests Required:** `tests/test_data.py::test_TC_T_01_db_file_exists`, `test_TC_T_03_fact_row_count`, `test_TC_T_05_foreign_key_integrity`

---

### 📌 BUILD-03: Analytical Calculation Engine (`src/analyser.py`)
- **Goal:** Author 13 pure Python/pandas calculation functions answering all Operational Questions (OQ-01 through OQ-10).
- **Files:** `src/analyser.py`
- **Linked Epics / Stories:** EP-03, EP-04, EP-05 / US-05 through US-16
- **Exact Execution Prompt:**
  > "Build `src/analyser.py` containing pure functions: `get_role_demand()`, `get_skill_frequencies()`, `get_salary_benchmarks()`, `get_geographic_hotspots()`, `get_monthly_trends()`, `get_skill_salary_correlation()`, and `get_company_leaderboard()`. Do not import streamlit in analyser.py."
- **Verification:** Verify that all return values are clean DataFrames or statistical tuples with Pearson $r$ and $p$-value.
- **Tests Required:** `tests/test_analysis.py::test_TC_T_07` through `test_TC_T_12`

---

### 📌 BUILD-04: Reusable Tailwind UI Component Integration
- **Goal:** Finalize reusable Tailwind CSS iframe widgets in `src/ui.py` for headers, KPI metrics, and insight banners.
- **Files:** `src/ui.py`
- **Linked Epics / Stories:** EP-02 / US-03, FR-15
- **Exact Execution Prompt:**
  > "Confirm `src/ui.py` renders `render_hero_header()`, `render_kpi_cards()`, and `render_insight_box()` via `st.components.v1.html` with responsive grid layouts and clean blue-indigo styling."
- **Verification:** Run `tests/test_app.py::test_TC_T_14_ui_helper_render_functions`.
- **Tests Required:** `tests/test_app.py::test_TC_T_14_ui_helper_render_functions`

---

### 📌 BUILD-05: App Entrypoint & Global State Filters (`app.py`)
- **Goal:** Construct `app.py` containing synchronized sidebar filters (`date_range`, `industry`, `city`, `exp_level`), state initialization, and overview redirection.
- **Files:** `app.py`
- **Linked Epics / Stories:** EP-02 / US-04, FR-14
- **Exact Execution Prompt:**
  > "Construct `app.py` with `st.set_page_config(layout='wide')`. Initialize `st.session_state` with default filter values. Build sidebar inputs for date range, industry multiselect, city multiselect, and experience level. Provide a 'Reset Filters' button and render the primary overview page."
- **Verification:** Verify sidebar controls correctly update session state across page changes.
- **Tests Required:** `TC-F-14`, `TC-F-15`

---

### 📌 BUILD-06: Multipage Implementations (Pages 01 to 05)
- **Goal:** Implement the first 5 interactive Streamlit pages for Overview, Role Demand, Skills, Salary, and Geography.
- **Files:** `pages/01_📊_Overview.py`, `pages/02_💼_Role_Demand.py`, `pages/03_🛠️_Skills_Analysis.py`, `pages/04_💰_Salary_Benchmark.py`, `pages/05_🗺️_Geography.py`
- **Linked Epics / Stories:** EP-02, EP-03, EP-04 / US-03, US-05, US-06, US-09, US-10, US-12; UI-01 to UI-05
- **Exact Execution Prompt:**
  > "Implement pages 01 to 05 using Plotly Express with Blues color palettes. In 01_Overview, render the Tailwind hero and 4 KPI cards. In 02_Role_Demand, show top-10 horizontal bar chart. In 03_Skills, explode skills and display top skills. In 04_Salary, render grouped min/mean/max bars and box plots. In 05_Geography, render city hiring rankings."
- **Verification:** Navigate to each page and ensure charts render without error.
- **Tests Required:** `TC-F-01`, `TC-F-02`, `TC-F-03`, `TC-F-06`, `TC-F-07`

---

### 📌 BUILD-07: Multipage Implementations (Pages 06 to 10)
- **Goal:** Complete the remaining 5 Streamlit pages for Trends, Correlation, Companies, Data Quality, and Job Types.
- **Files:** `pages/06_📈_Hiring_Trends.py`, `pages/07_🔬_Correlation.py`, `pages/08_🏢_Companies.py`, `pages/09_🧹_EDA_Quality.py`, `pages/10_👥_Job_Types_Exp.py`
- **Linked Epics / Stories:** EP-01, EP-03, EP-04, EP-05 / US-02, US-07, US-08, US-11, US-13, US-14, US-15, US-16; UI-06 to UI-10
- **Exact Execution Prompt:**
  > "Implement pages 06 to 10. Page 06 renders monthly timeline and role trend slicer. Page 07 renders OLS scatter plot, Pearson r, and p-value. Page 08 renders top-10 employer leaderboard with sortable table. Page 09 presents pandas describe() statistics and null audit. Page 10 renders work mode donut chart and 100% stacked bar for experience."
- **Verification:** Ensure Page 07 displays valid regression statistics and Page 09 shows the 12 nulls and 3 outliers.
- **Tests Required:** `TC-F-04`, `TC-F-05`, `TC-F-08`, `TC-F-09`, `TC-F-10`, `TC-F-11`, `TC-F-12`, `TC-F-13`

---

### 📌 BUILD-08: Data Export & Edge Case Hardening
- **Goal:** Add CSV download buttons (`st.download_button`) on all analysis pages and handle empty filter results gracefully.
- **Files:** All `pages/*.py`
- **Linked Epics / Stories:** FR-16, US-16 / TC-F-15, TC-F-16
- **Exact Execution Prompt:**
  > "Add `st.download_button` under every primary data table to export the active filtered DataFrame. Enforce empty state guard clauses: `if len(filtered_df) == 0: st.warning('No postings match selected filters'); st.stop()`."
- **Verification:** Apply an impossible filter and verify that a warning is displayed instead of a script traceback.
- **Tests Required:** `TC-F-15`, `TC-F-16`

---

### 📌 BUILD-09: Automated Quality Assurance Suite
- **Goal:** Execute the full 15-test pytest suite and achieve a 100% pass rate.
- **Files:** `tests/test_data.py`, `tests/test_analysis.py`, `tests/test_app.py`
- **Linked Epics / Stories:** All Epics / SC-04, KPI O-03
- **Exact Execution Prompt:**
  > "Run `python -m pytest tests/ -v`. Fix any regression or schema mismatches until 15/15 tests pass in under 10 seconds."
- **Verification:** Check pytest terminal output displays `15 passed in <10s`.
- **Tests Required:** All `TC-T-01` through `TC-T-15`

---

### 📌 BUILD-10: Documentation & Deployment Preparation
- **Goal:** Author the comprehensive `README.md` guide and configure for Streamlit Community Cloud hosting.
- **Files:** `README.md`, `.gitignore`, `runtime.txt`
- **Linked Epics / Stories:** Step 16 / DR-01 through DR-08
- **Exact Execution Prompt:**
  > "Write a professional `README.md` containing project title, badges, architecture diagram (Mermaid), installation instructions, local run commands, test execution guide, and Streamlit Community Cloud URL link."
- **Verification:** Verify all markdown links and mermaid syntax render cleanly.
- **Tests Required:** `TC-T-15` (Latency < 1.0s)

---

## 4. Definition of Done (DoD)

The project build is officially marked **COMPLETE** only when:
1. ✅ **All 13 Business Requirements (BR-01..13):** Fully implemented and functional in the UI.
2. ✅ **All 10 Operational Questions (OQ-01..10):** Visualized with appropriate Plotly charts.
3. ✅ **100% Pytest Pass Rate:** All 15 unit tests pass (`15 passed in < 10s`).
4. ✅ **Zero Open Critical Defects:** Zero unhandled tracebacks or blank charts.
5. ✅ **Traceability Intact:** All RTM columns populated and verified.
6. ✅ **Clean Code Quality:** Modular layout with zero Streamlit imports inside `src/analyser.py`.
7. ✅ **Deployment Ready:** App runs locally via `streamlit run app.py` with $<1.0\text{ s}$ query latency.

---

## 5. Standard Operating Commands

```bash
# 1. Install Dependencies
pip install -r requirements.txt

# 2. Run Synthetic Dataset Generator (if regenerating DB)
python src/generate_data.py

# 3. Execute Pytest Test Suite
python -m pytest tests/ -v

# 4. Launch Streamlit Application Locally
streamlit run app.py
```

---

## 6. Comprehensive Viva Examination Guide (10 Likely Questions)

### Q1: What is the business problem your project solves?
> **Answer:** Organizations and job seekers currently rely on fragmented job portal listings and static surveys. This project consolidates recruitment data into an interactive analytics tool that automates salary benchmarking, skill demand tracking, and geographic hiring hotspot identification in real time.

### Q2: Why did you choose SQLite over MySQL or PostgreSQL?
> **Answer:** SQLite is serverless, zero-configuration, and fully self-contained in a single file (`data/app.db`). This makes it optimal for academic demonstration, version control, and instant deployment on Streamlit Community Cloud without needing managed database hosting.

### Q3: How is data normalized in your database design?
> **Answer:** It uses a Star Schema with 3 master dimension tables (`dim_industry`, `dim_city`, `dim_job_role`), 1 core fact table (`fact_job_postings`), and 1 junction bridge table (`job_skills`) resolving the M:N skill relationship to First Normal Form (1NF).

### Q4: What is the purpose of `vw_hiring_analysis`?
> **Answer:** It is a pre-joined analytical view combining dimensions, facts, and derived attributes (such as `salary_mid`, `quarter`, and `skill_complexity`). It simplifies Streamlit data access by allowing the dashboard to read clean analytical data in one cached query.

### Q5: How do you prevent Streamlit from reloading data on every user click?
> **Answer:** By decorating `load_data()` with `@st.cache_data(ttl=3600)`. This caches the loaded DataFrame in memory, enabling sub-second filtering across sidebar widgets without hitting the SQLite disk file repeatedly.

### Q6: Why did you introduce deliberate missing values and outliers in the dataset?
> **Answer:** Real-world recruitment data contains data quality issues. Introducing 12 nulls in `salary_min` and 3 executive compensation outliers allowed us to build and demonstrate a dedicated Data Quality & EDA module (Page 09) featuring imputation and IQR outlier analysis.

### Q7: What statistical method did you use for OQ-09 (Skills vs Salary)?
> **Answer:** We performed Ordinary Least Squares (OLS) linear regression and computed the Pearson correlation coefficient ($r$) along with its $p$-value to statistically verify whether roles demanding more technical skills command higher market compensation.

### Q8: How did you implement Tailwind CSS without Node.js or npm?
> **Answer:** We injected the Tailwind CSS CDN stylesheet directly inside isolated iframes using `st.components.v1.html` in `src/ui.py`. This provides modern cards and gradients without complex build pipelines.

### Q9: What is the difference between functional and technical tests in Step 13?
> **Answer:** Functional tests (`TC-F`) test user-facing workflows, widget interactions, and expected UI behaviors from a stakeholder perspective. Technical tests (`TC-T`) are automated code-level unit tests written in pytest to verify database constraints, math invariants, and query latency.

### Q10: How does the Requirements Traceability Matrix (RTM) safeguard project quality?
> **Answer:** The RTM guarantees bidirectional traceability. Every operational question traces forward through HMWs, business requirements, database columns, user stories, UI screens, and test cases, ensuring zero orphan requirements and full test coverage.
