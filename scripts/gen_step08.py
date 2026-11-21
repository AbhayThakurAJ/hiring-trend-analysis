"""
Step 08 – Database Design Excel + RTM Update
Project : P071 – Hiring Trend Analysis
"""
import os, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT_DIR  = r"C:\Hiring Trend Analysis\docs"
DB_FILE  = os.path.join(OUT_DIR, "P01_HTA_Step08_DatabaseDesign.xlsx")
RTM_FILE = os.path.join(OUT_DIR, "P01_HTA_Step04_RTM.xlsx")
os.makedirs(OUT_DIR, exist_ok=True)

# ── styles ────────────────────────────────────────────────────────────────────
def F(h): return PatternFill('solid',fgColor=h)
def FN(b=False,c='000000',s=10,i=False): return Font(bold=b,color=c,size=s,italic=i,name='Calibri')
def AL(h='left',v='center',w=True): return Alignment(horizontal=h,vertical=v,wrap_text=w)
def BD(c='2E74B5'):
    s=Side(style='thin',color=c)
    return Border(left=s,right=s,top=s,bottom=s)

HDR_FILL=F('1F3564'); HDR_FONT=FN(True,'FFFFFF',10)
ALT=F('EEF4FB'); WHT=F('FFFFFF'); OK=F('E5FFE5'); WARN=F('FFF3CD'); ERR=F('FFE5E5')
MUST=F('D6E4F0'); SHOULD=F('E2EFDA'); COULD=F('FFF3CD')
PK_FILL=F('D6E4F0'); FK_FILL=F('E2EFDA'); REQ_FILL=F('FFF8F0')
CODE_FILL=F('F0F4FF')
bdr=BD()

def style(cell,fill=None,font=None,align=None):
    if fill: cell.fill=fill
    if font: cell.font=font
    if align: cell.alignment=align
    cell.border=bdr

def write_header(ws,row,cols,widths=None):
    for i,c in enumerate(cols,1):
        cell=ws.cell(row=row,column=i,value=c)
        style(cell,HDR_FILL,HDR_FONT,AL('center'))
    if widths:
        for i,w in enumerate(widths,1):
            ws.column_dimensions[get_column_letter(i)].width=w

def write_row(ws,row_num,vals,fills=None,bolds=None):
    for i,v in enumerate(vals,1):
        cell=ws.cell(row=row_num,column=i,value=v)
        fl=fills[i-1] if fills else (ALT if row_num%2==0 else WHT)
        bo=bolds[i-1] if bolds else False
        style(cell,fl,FN(bo),AL())

def banner(ws,text,cols=8):
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=cols)
    c=ws.cell(1,1,value=text)
    style(c,F('1F3564'),FN(True,'FFFFFF',14),AL('center'))
    ws.row_dimensions[1].height=28

# ═══════════════════════════════════════════════════════════════════════════════
wb=Workbook()

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — Overview
# ══════════════════════════════════════════════════════════════════════════════
ws=wb.active; ws.title="Overview"
ws.sheet_view.showGridLines=False
banner(ws,'P071 – Hiring Trend Analysis  |  Database Design  |  Step 08',8)

meta=[
    ('Project','P071 – HTA'),('DB Engine','SQLite 3'),
    ('DB File','data/app.db'),('Tables','5'),
    ('Relationships','4 FK'),('Analysis View','1 (30+ cols)'),
    ('Date',datetime.date.today().strftime('%d %b %Y')),('Author','Abhayjeet Singh'),
]
for i,(k,v) in enumerate(meta):
    r=3+i
    ws.cell(r,1,value=k); style(ws.cell(r,1),ALT if i%2==0 else WHT,FN(True))
    ws.cell(r,2,value=v); style(ws.cell(r,2),ALT if i%2==0 else WHT,FN())
ws.column_dimensions['A'].width=18; ws.column_dimensions['B'].width=35

# design overview paragraph
r_start=3+len(meta)+1
design_text=[
    "5-table normalised schema: 3 master (dimension) tables + 1 bridge + 1 fact table.",
    "Master tables: dim_industry (8 rows), dim_city (15 rows), dim_job_role (12 rows).",
    "Bridge table: job_skills — many-to-many between postings and skills.",
    "Fact table: fact_job_postings — one row per job posting (500–1000 rows).",
    "Analysis view: vw_hiring_analysis — 30+ columns joining all 5 tables.",
    "Every OQ (OQ-01 to OQ-10) is answerable from vw_hiring_analysis.",
    "All BR-R business rules enforced via CHECK and NOT NULL constraints.",
]
for i,line in enumerate(design_text):
    ws.cell(r_start+i,1,value=f"• {line}")
    ws.merge_cells(start_row=r_start+i,start_column=1,end_row=r_start+i,end_column=8)
    style(ws.cell(r_start+i,1),WHT,FN(False,'333333',9.5),AL())

ws.sheet_properties.tabColor='1F3564'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — Tables Summary
# ══════════════════════════════════════════════════════════════════════════════
ws2=wb.create_sheet("Tables_Summary")
ws2.sheet_view.showGridLines=False
banner(ws2,'Table Summary — 5 Tables + 1 Analysis View',8)

hdrs=['Table Name','Type','Grain (1 row =)','Primary Key','Foreign Keys','Column Count','Planned Rows','Key OQs / BRs Served']
write_header(ws2,3,hdrs,[20,12,22,14,24,12,12,30])

tables=[
    ('dim_industry','Master (Dimension)','One industry sector','industry_id','—','3','8',
     'OQ-02,07 · BR-02,06,10,11'),
    ('dim_city','Master (Dimension)','One city-state pair','city_id','—','4','15',
     'OQ-05 · BR-03,07'),
    ('dim_job_role','Master (Dimension)','One job role template','role_id','—','4','12',
     'OQ-01,08 · BR-01,08'),
    ('fact_job_postings','Fact (Transactional)','One job posting','job_id',
     'industry_id → dim_industry\ncity_id → dim_city\nrole_id → dim_job_role','14','500–1000',
     'ALL OQs · ALL BRs'),
    ('job_skills','Bridge (M:N)','One skill per posting','skill_id',
     'job_id → fact_job_postings','4','1500–4000',
     'OQ-03,09 · BR-02,09'),
    ('vw_hiring_analysis','View (JOIN)','One job posting (denormalised)','—','—','32','= fact rows',
     'ALL OQs · ALL BRs (read-only)'),
]
for i,row in enumerate(tables,4):
    fl=[PK_FILL if 'Master' in row[1] else (MUST if 'Fact' in row[1] else (FK_FILL if 'Bridge' in row[1] else WARN))] * 8
    write_row(ws2,i,row,fills=fl,bolds=[True]+[False]*7)
    ws2.row_dimensions[i].height=36

ws2.sheet_properties.tabColor='2E74B5'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — dim_industry
# ══════════════════════════════════════════════════════════════════════════════
ws3=wb.create_sheet("dim_industry")
ws3.sheet_view.showGridLines=False
banner(ws3,'dim_industry — Industry Sector Master',8)

dd_hdrs=['Column','Data Type','Key','Nullable','Allowed Values / Range','Description','OQ/BR Supported']
write_header(ws3,3,dd_hdrs,[16,12,10,10,30,30,20])

dim_ind_cols=[
    ('industry_id','INTEGER','PK','NO','Auto-increment','Unique industry identifier','All industry-related OQs'),
    ('industry_name','TEXT','UNIQUE','NO','IT|Healthcare|Finance|Education|Manufacturing|Retail|Telecom|Media','Industry sector name','OQ-02,07 · BR-02,06,10,11'),
    ('sector_category','TEXT','—','NO','Services|Production|Knowledge','Broad sector grouping','Filtering / groupby'),
]
for i,row in enumerate(dim_ind_cols,4):
    fl=PK_FILL if 'PK' in row[2] else (ALT if i%2==0 else WHT)
    write_row(ws3,i,row,fills=[fl]*7,bolds=[True]+[False]*6)

ws3.sheet_properties.tabColor='4A86C8'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 — dim_city
# ══════════════════════════════════════════════════════════════════════════════
ws4=wb.create_sheet("dim_city")
ws4.sheet_view.showGridLines=False
banner(ws4,'dim_city — City-State Geography Master',8)
write_header(ws4,3,dd_hdrs,[16,12,10,10,30,30,20])

dim_city_cols=[
    ('city_id','INTEGER','PK','NO','Auto-increment','Unique city identifier','OQ-05 · BR-03,07'),
    ('city_name','TEXT','UNIQUE','NO','Mumbai|Delhi|Bangalore|Hyderabad|Chennai|Pune|Kolkata|Ahmedabad|Jaipur|Lucknow|Noida|Gurgaon|Chandigarh|Indore|Kochi',
     'City name','OQ-05 · BR-07'),
    ('state_name','TEXT','—','NO','Maharashtra|Delhi|Karnataka|Telangana|Tamil Nadu|Rajasthan|UP|Gujarat|Punjab|MP|Kerala',
     'State or UT name','OQ-05 · BR-07'),
    ('region','TEXT','—','NO','North|South|West|East|Central','Geographic region for grouping','Filtering'),
]
for i,row in enumerate(dim_city_cols,4):
    fl=PK_FILL if 'PK' in row[2] else (ALT if i%2==0 else WHT)
    write_row(ws4,i,row,fills=[fl]*7,bolds=[True]+[False]*6)

ws4.sheet_properties.tabColor='4A86C8'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5 — dim_job_role
# ══════════════════════════════════════════════════════════════════════════════
ws5=wb.create_sheet("dim_job_role")
ws5.sheet_view.showGridLines=False
banner(ws5,'dim_job_role — Job Role Master',8)
write_header(ws5,3,dd_hdrs,[16,12,10,10,30,30,20])

dim_role_cols=[
    ('role_id','INTEGER','PK','NO','Auto-increment','Unique role identifier','OQ-01,08 · BR-01,08'),
    ('role_title','TEXT','UNIQUE','NO','Data Analyst|Software Engineer|Data Scientist|Business Analyst|DevOps Engineer|Product Manager|UI/UX Designer|Full Stack Developer|Cloud Architect|ML Engineer|QA Engineer|Cyber Security Analyst',
     'Standard job role title','OQ-01 · BR-01'),
    ('role_category','TEXT','—','NO','Analytics|Engineering|Management|Design|Security','Broad role family','Grouping / filtering'),
    ('seniority_band','TEXT','—','NO','Fresher|Mid-level|Senior','Default seniority expectation for the role','OQ-08 · BR-08'),
]
for i,row in enumerate(dim_role_cols,4):
    fl=PK_FILL if 'PK' in row[2] else (ALT if i%2==0 else WHT)
    write_row(ws5,i,row,fills=[fl]*7,bolds=[True]+[False]*6)

ws5.sheet_properties.tabColor='4A86C8'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 6 — fact_job_postings
# ══════════════════════════════════════════════════════════════════════════════
ws6=wb.create_sheet("fact_job_postings")
ws6.sheet_view.showGridLines=False
banner(ws6,'fact_job_postings — Main Fact Table (1 row = 1 job posting)',8)
write_header(ws6,3,dd_hdrs,[18,12,10,10,30,30,20])

fact_cols=[
    ('job_id','INTEGER','PK','NO','Auto-increment','Unique job posting ID','All OQs'),
    ('role_id','INTEGER','FK → dim_job_role','NO','Existing role_id','FK to dim_job_role','OQ-01,08 · BR-01,08'),
    ('company_name','TEXT','—','NO','50 synthetic companies','Hiring company (synthetic)','OQ-10 · BR-04'),
    ('industry_id','INTEGER','FK → dim_industry','NO','Existing industry_id','FK to dim_industry','OQ-02,07 · BR-02,06,10,11'),
    ('city_id','INTEGER','FK → dim_city','NO','Existing city_id','FK to dim_city','OQ-05 · BR-03,07'),
    ('job_type','TEXT','—','NO','Full-time|Part-time|Remote|Hybrid|Contract','Employment type (BR-R-05)','OQ-06 · BR-10'),
    ('experience_level','TEXT','—','NO','Fresher|Mid-level|Senior','Required experience (BR-R-04)','OQ-08 · BR-03,08'),
    ('salary_min','REAL','—','NO','200000–1500000','Min salary ₹ INR','OQ-02 · BR-03,06'),
    ('salary_max','REAL','—','NO','300000–3500000; must be > salary_min','Max salary ₹ INR (BR-R-09)','OQ-02,09 · BR-03,06,09'),
    ('required_skills','TEXT','—','NO','Pipe-delimited; 1–8 skills per row (BR-R-08)','Skill list for the posting','OQ-03,09 · BR-02,09'),
    ('skills_count','INTEGER','—','NO','1–8','Derived: len(required_skills.split("|"))','OQ-09 · BR-09'),
    ('posted_date','TEXT','—','NO','2022-01-01 to 2024-12-31 (BR-R-03)','ISO 8601 posting date','OQ-04 · BR-05,12'),
    ('posted_month','TEXT','—','NO','2022-01 to 2024-12','Derived: YYYY-MM','OQ-04 · BR-05'),
    ('posted_year','INTEGER','—','NO','2022|2023|2024','Derived: YYYY','Filtering'),
]
for i,row in enumerate(fact_cols,4):
    if 'PK' in row[2]: fl=PK_FILL
    elif 'FK' in row[2]: fl=FK_FILL
    else: fl=ALT if i%2==0 else WHT
    write_row(ws6,i,row,fills=[fl]*7,bolds=[True]+[False]*6)

ws6.sheet_properties.tabColor='1F5C2E'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 7 — job_skills
# ══════════════════════════════════════════════════════════════════════════════
ws7=wb.create_sheet("job_skills")
ws7.sheet_view.showGridLines=False
banner(ws7,'job_skills — Bridge Table (M:N posting↔skill)',8)
write_header(ws7,3,dd_hdrs,[16,12,10,10,30,30,20])

bridge_cols=[
    ('skill_id','INTEGER','PK','NO','Auto-increment','Unique row identifier','—'),
    ('job_id','INTEGER','FK → fact_job_postings','NO','Existing job_id','FK to fact table','OQ-03,09 · BR-02,09'),
    ('skill_name','TEXT','—','NO','Python|SQL|Tableau|Power BI|Excel|Java|AWS|Docker|R|TensorFlow|React|Node.js|Kubernetes|Spark|Hadoop|Figma|Jira|Git|Linux|Azure|GCP|MongoDB|PostgreSQL|Kafka|Airflow',
     'Individual skill extracted from pipe-delimited string','OQ-03 · BR-02'),
    ('is_primary','INTEGER','—','YES','0|1','Flag: 1 if skill is first in the original list','Filtering priority skills'),
]
for i,row in enumerate(bridge_cols,4):
    if 'PK' in row[2]: fl=PK_FILL
    elif 'FK' in row[2]: fl=FK_FILL
    else: fl=ALT if i%2==0 else WHT
    write_row(ws7,i,row,fills=[fl]*7,bolds=[True]+[False]*6)

ws7.sheet_properties.tabColor='7B5700'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 8 — Relationships
# ══════════════════════════════════════════════════════════════════════════════
ws8=wb.create_sheet("Relationships")
ws8.sheet_view.showGridLines=False
banner(ws8,'Table Relationships (FK Constraints)',7)

rel_hdrs=['#','Parent Table','Parent Key','Child Table','Child Key','Cardinality','Delete Rule']
write_header(ws8,3,rel_hdrs,[5,20,14,22,14,14,14])

rels=[
    ('1','dim_industry','industry_id','fact_job_postings','industry_id','1 : N','RESTRICT'),
    ('2','dim_city','city_id','fact_job_postings','city_id','1 : N','RESTRICT'),
    ('3','dim_job_role','role_id','fact_job_postings','role_id','1 : N','RESTRICT'),
    ('4','fact_job_postings','job_id','job_skills','job_id','1 : N','CASCADE'),
]
for i,row in enumerate(rels,4):
    fl=FK_FILL
    write_row(ws8,i,row,fills=[fl]*7,bolds=[True]+[False]*6)

ws8.sheet_properties.tabColor='2E74B5'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 9 — KPIs & Measures
# ══════════════════════════════════════════════════════════════════════════════
ws9=wb.create_sheet("KPIs_Measures")
ws9.sheet_view.showGridLines=False
banner(ws9,'KPIs, Calculated Measures & Aggregations',8)

kpi_hdrs=['Measure ID','Measure Name','SQL / pandas Formula','OQ','BR','Chart Type','Page']
write_header(ws9,3,kpi_hdrs,[10,24,42,8,8,18,22])

kpis=[
    ('KPI-01','Total Postings','COUNT(job_id)','All','All','st.metric / KPI card','01_Overview'),
    ('KPI-02','Average Max Salary','AVG(salary_max)','OQ-02','BR-06','st.metric','01_Overview'),
    ('KPI-03','Top City by Postings','city_name WHERE COUNT(job_id) = MAX','OQ-05','BR-07','st.metric','01_Overview'),
    ('KPI-04','Top Role by Postings','role_title WHERE COUNT(job_id) = MAX','OQ-01','BR-01','st.metric','01_Overview'),
    ('M-01','Role Posting Count','GROUP BY role_title → COUNT(job_id) ORDER BY DESC','OQ-01','BR-01','px.bar (horizontal)','02_Role_Demand'),
    ('M-02','Skill Frequency','GROUP BY skill_name → COUNT(skill_id) → TOP N','OQ-03','BR-02','px.bar (horizontal)','03_Skills'),
    ('M-03','Salary Stats (Avg/Min/Max)','GROUP BY industry → AVG/MIN/MAX(salary_min, salary_max)','OQ-02','BR-03,06','px.bar (grouped) + px.box','04_Salary'),
    ('M-04','City Posting Count','GROUP BY city_name → COUNT(job_id) DESC','OQ-05','BR-07','px.bar','05_Geography'),
    ('M-05','Monthly Hiring Volume','GROUP BY posted_month → COUNT(job_id)','OQ-04','BR-05','px.line','06_Trends'),
    ('M-06','Role Monthly Trend','WHERE role_title = X → GROUP BY posted_month → COUNT','OQ-04','BR-12','px.line','06_Trends'),
    ('M-07','Pearson r (skills ↔ salary)','scipy.stats.pearsonr(skills_count, salary_max)','OQ-09','BR-09','px.scatter + OLS','07_Correlation'),
    ('M-08','Company Posting Count','GROUP BY company_name → COUNT(job_id) → TOP 10','OQ-10','BR-04','px.bar + table','08_Companies'),
    ('M-09','EDA Summary Stats','pandas df.describe() + null count + IQR outliers','—','BR-13','st.dataframe','09_EDA'),
    ('M-10','Job Type %','GROUP BY job_type → COUNT / SUM * 100','OQ-06','BR-10','px.pie (donut)','10_JobType'),
    ('M-11','Exp Level % per Role','GROUP BY role_title, exp_level → normalise to 100%','OQ-08','BR-08','px.bar (stack 100%)','10_JobType'),
    ('M-12','Industry by Skill Keyword','WHERE skill_name LIKE X → GROUP BY industry → COUNT','OQ-07','BR-11','px.bar','10_JobType'),
]
for i,row in enumerate(kpis,4):
    fl=MUST if 'KPI' in row[0] else (ALT if i%2==0 else WHT)
    write_row(ws9,i,row,fills=[fl]*7,bolds=[True]+[False]*6)

ws9.sheet_properties.tabColor='1F5C2E'

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 10 — SQLite DDL
# ══════════════════════════════════════════════════════════════════════════════
ws10=wb.create_sheet("SQLite_DDL")
ws10.sheet_view.showGridLines=False
banner(ws10,'SQLite DDL — CREATE TABLE + VIEW Statements',3)
ws10.column_dimensions['A'].width=5
ws10.column_dimensions['B'].width=90
ws10.column_dimensions['C'].width=18

ddl_blocks=[
    ("dim_industry","""CREATE TABLE IF NOT EXISTS dim_industry (
    industry_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    industry_name   TEXT    NOT NULL UNIQUE,
    sector_category TEXT    NOT NULL
        CHECK(sector_category IN ('Services','Production','Knowledge'))
);"""),
    ("dim_city","""CREATE TABLE IF NOT EXISTS dim_city (
    city_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    city_name  TEXT    NOT NULL UNIQUE,
    state_name TEXT    NOT NULL,
    region     TEXT    NOT NULL
        CHECK(region IN ('North','South','West','East','Central'))
);"""),
    ("dim_job_role","""CREATE TABLE IF NOT EXISTS dim_job_role (
    role_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    role_title    TEXT    NOT NULL UNIQUE,
    role_category TEXT    NOT NULL
        CHECK(role_category IN ('Analytics','Engineering','Management','Design','Security')),
    seniority_band TEXT  NOT NULL
        CHECK(seniority_band IN ('Fresher','Mid-level','Senior'))
);"""),
    ("fact_job_postings","""CREATE TABLE IF NOT EXISTS fact_job_postings (
    job_id           INTEGER PRIMARY KEY AUTOINCREMENT,
    role_id          INTEGER NOT NULL REFERENCES dim_job_role(role_id),
    company_name     TEXT    NOT NULL,
    industry_id      INTEGER NOT NULL REFERENCES dim_industry(industry_id),
    city_id          INTEGER NOT NULL REFERENCES dim_city(city_id),
    job_type         TEXT    NOT NULL
        CHECK(job_type IN ('Full-time','Part-time','Remote','Hybrid','Contract')),
    experience_level TEXT    NOT NULL
        CHECK(experience_level IN ('Fresher','Mid-level','Senior')),
    salary_min       REAL    NOT NULL CHECK(salary_min > 0),
    salary_max       REAL    NOT NULL CHECK(salary_max > salary_min),
    required_skills  TEXT    NOT NULL,
    skills_count     INTEGER NOT NULL CHECK(skills_count BETWEEN 1 AND 8),
    posted_date      TEXT    NOT NULL,
    posted_month     TEXT    NOT NULL,
    posted_year      INTEGER NOT NULL CHECK(posted_year BETWEEN 2022 AND 2024),
    FOREIGN KEY (role_id)     REFERENCES dim_job_role(role_id),
    FOREIGN KEY (industry_id) REFERENCES dim_industry(industry_id),
    FOREIGN KEY (city_id)     REFERENCES dim_city(city_id)
);"""),
    ("job_skills","""CREATE TABLE IF NOT EXISTS job_skills (
    skill_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id     INTEGER NOT NULL REFERENCES fact_job_postings(job_id)
        ON DELETE CASCADE,
    skill_name TEXT    NOT NULL,
    is_primary INTEGER DEFAULT 0
        CHECK(is_primary IN (0, 1)),
    FOREIGN KEY (job_id) REFERENCES fact_job_postings(job_id)
);"""),
    ("vw_hiring_analysis","""CREATE VIEW IF NOT EXISTS vw_hiring_analysis AS
SELECT
    f.job_id,
    r.role_title,
    r.role_category,
    r.seniority_band,
    f.company_name,
    i.industry_name,
    i.sector_category,
    c.city_name,
    c.state_name,
    c.region,
    f.job_type,
    f.experience_level,
    f.salary_min,
    f.salary_max,
    ROUND((f.salary_min + f.salary_max) / 2.0, 0) AS salary_mid,
    f.required_skills,
    f.skills_count,
    f.posted_date,
    f.posted_month,
    f.posted_year,
    -- Derived measures for OQ analysis
    ROUND(f.salary_max - f.salary_min, 0) AS salary_range,
    CASE
        WHEN f.salary_max >= 2000000 THEN 'Premium'
        WHEN f.salary_max >= 1000000 THEN 'Mid-Range'
        ELSE 'Entry-Level'
    END AS salary_band,
    CASE
        WHEN f.skills_count >= 5 THEN 'High-Skill'
        WHEN f.skills_count >= 3 THEN 'Mid-Skill'
        ELSE 'Low-Skill'
    END AS skill_complexity,
    SUBSTR(f.posted_date, 1, 7) AS ym_key,
    CAST(SUBSTR(f.posted_date, 6, 2) AS INTEGER) AS post_month_num,
    CASE
        WHEN CAST(SUBSTR(f.posted_date, 6, 2) AS INTEGER) BETWEEN 1 AND 3 THEN 'Q1'
        WHEN CAST(SUBSTR(f.posted_date, 6, 2) AS INTEGER) BETWEEN 4 AND 6 THEN 'Q2'
        WHEN CAST(SUBSTR(f.posted_date, 6, 2) AS INTEGER) BETWEEN 7 AND 9 THEN 'Q3'
        ELSE 'Q4'
    END AS quarter,
    f.posted_year || '-' ||
    CASE
        WHEN CAST(SUBSTR(f.posted_date, 6, 2) AS INTEGER) BETWEEN 1 AND 3 THEN 'Q1'
        WHEN CAST(SUBSTR(f.posted_date, 6, 2) AS INTEGER) BETWEEN 4 AND 6 THEN 'Q2'
        WHEN CAST(SUBSTR(f.posted_date, 6, 2) AS INTEGER) BETWEEN 7 AND 9 THEN 'Q3'
        ELSE 'Q4'
    END AS year_quarter
FROM fact_job_postings f
JOIN dim_job_role r   ON f.role_id     = r.role_id
JOIN dim_industry i   ON f.industry_id = i.industry_id
JOIN dim_city c       ON f.city_id     = c.city_id;"""),
    ("INDEXES","""-- Performance indexes
CREATE INDEX IF NOT EXISTS idx_fact_role     ON fact_job_postings(role_id);
CREATE INDEX IF NOT EXISTS idx_fact_industry ON fact_job_postings(industry_id);
CREATE INDEX IF NOT EXISTS idx_fact_city     ON fact_job_postings(city_id);
CREATE INDEX IF NOT EXISTS idx_fact_date     ON fact_job_postings(posted_date);
CREATE INDEX IF NOT EXISTS idx_fact_month    ON fact_job_postings(posted_month);
CREATE INDEX IF NOT EXISTS idx_skills_job    ON job_skills(job_id);
CREATE INDEX IF NOT EXISTS idx_skills_name   ON job_skills(skill_name);"""),
]

r=3
for name, ddl in ddl_blocks:
    ws10.cell(r,1,value=''); style(ws10.cell(r,1),F('1F5C2E'),FN(True,'FFFFFF',10),AL('center'))
    ws10.cell(r,2,value=name); style(ws10.cell(r,2),F('1F5C2E'),FN(True,'FFFFFF',10),AL())
    ws10.cell(r,3,value='DDL'); style(ws10.cell(r,3),F('1F5C2E'),FN(True,'FFFFFF',10),AL('center'))
    r+=1
    for line in ddl.split('\n'):
        ws10.cell(r,1,value='')
        ws10.cell(r,2,value=line); style(ws10.cell(r,2),CODE_FILL,Font(name='Courier New',size=9,color='1A1A2E'),AL())
        ws10.cell(r,3,value='')
        r+=1
    r+=1  # blank line separator

ws10.sheet_properties.tabColor='7B5700'

# ── Save DB Design ───────────────────────────────────────────────────────────
wb.save(DB_FILE)
print(f"SAVED: {DB_FILE}")

# ══════════════════════════════════════════════════════════════════════════════
# UPDATE RTM — Fill column J (DB Table.Column)
# ══════════════════════════════════════════════════════════════════════════════
rtm_wb=load_workbook(RTM_FILE)
ws_rtm=rtm_wb['RTM']

# BR → DB Table.Column mapping
db_map={
    'BR-01': 'fact_job_postings.role_id → dim_job_role.role_title\nfact_job_postings.posted_date',
    'BR-02': 'job_skills.skill_name\nfact_job_postings.industry_id → dim_industry.industry_name',
    'BR-03': 'fact_job_postings.salary_min, salary_max\nfact_job_postings.experience_level\nfact_job_postings.city_id → dim_city.city_name',
    'BR-04': 'fact_job_postings.company_name\nfact_job_postings.industry_id → dim_industry.industry_name',
    'BR-05': 'fact_job_postings.posted_month\nfact_job_postings.job_id (COUNT)',
    'BR-06': 'fact_job_postings.salary_min, salary_max\nfact_job_postings.industry_id → dim_industry.industry_name',
    'BR-07': 'fact_job_postings.city_id → dim_city.city_name\nfact_job_postings.job_id (COUNT)',
    'BR-08': 'fact_job_postings.experience_level\nfact_job_postings.role_id → dim_job_role.role_title',
    'BR-09': 'fact_job_postings.skills_count\nfact_job_postings.salary_max',
    'BR-10': 'fact_job_postings.job_type\nfact_job_postings.industry_id → dim_industry.industry_name',
    'BR-11': 'job_skills.skill_name\nfact_job_postings.industry_id → dim_industry.industry_name\nfact_job_postings.posted_month',
    'BR-12': 'fact_job_postings.role_id → dim_job_role.role_title\nfact_job_postings.posted_month',
    'BR-13': 'fact_job_postings.salary_min, salary_max\nfact_job_postings.skills_count (all numeric)',
}

# Column C = BR ID (col 3), Column J = DB Table.Column (col 10)
for row_num in range(4, 17):  # rows 4-16 = data rows
    br_cell = ws_rtm.cell(row=row_num, column=3).value
    if br_cell and br_cell in db_map:
        cell = ws_rtm.cell(row=row_num, column=10)
        cell.value = db_map[br_cell]
        style(cell, F('E2EFDA'), FN(False,'000000',9), AL(w=True))

# Update status column H → "DB Mapped"
for row_num in range(4, 17):
    cell = ws_rtm.cell(row=row_num, column=8)
    cell.value = 'DB Mapped'
    style(cell, F('E5FFE5'), FN(True,'1F5C2E',9), AL('center'))

# Update changelog
ws_ch = rtm_wb['Changelog']
# Find row for Step 08
for r in range(3, 10):
    if ws_ch.cell(r,1).value and 'Step 08' in str(ws_ch.cell(r,1).value):
        ws_ch.cell(r,2).value = datetime.date.today().strftime('%d %b %Y')
        ws_ch.cell(r,5).value = 'DB Table.Column filled for all 13 BRs. Status → DB Mapped.'
        break

rtm_wb.save(RTM_FILE)
print(f"UPDATED: {RTM_FILE}")
