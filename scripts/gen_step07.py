"""
Step 07 – Technical Requirements Document (TRD) Generator
Project : P01 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES, MCA 2nd Yr DA
NOTE    : Step 06 (PRD) skipped by learner — TRD proceeds from BRD (Step 05)
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os, datetime

OUT_DIR  = r"C:\Hiring Trend Analysis\docs"
OUT_FILE = os.path.join(OUT_DIR, "P01_HTA_Step07_TechBriefDocument.docx")
os.makedirs(OUT_DIR, exist_ok=True)

DARK  = RGBColor(0x1F,0x35,0x64); MID   = RGBColor(0x2E,0x74,0xB5)
STEEL = RGBColor(0x4A,0x86,0xC8); WHITE = RGBColor(0xFF,0xFF,0xFF)
BLACK = RGBColor(0x00,0x00,0x00); GREEN = RGBColor(0x37,0x86,0x36)
AMBER = RGBColor(0xC5,0x5A,0x11); RED   = RGBColor(0xC0,0x00,0x00)
CODE  = RGBColor(0x1A,0x1A,0x2E)

HEX = dict(
    dark='1F3564',mid='2E74B5',steel='4A86C8',
    alt='EEF4FB',white='FFFFFF',ok='E5FFE5',
    warn='FFF3CD',err='FFE5E5',grey='F2F2F2',
    code='1A1A2E',code_lt='F0F4FF',fr='D6E4F0',
    must='D6E4F0',should='E2EFDA',could='FFF3CD',
)

def bg(cell,h):
    tc=cell._tc;pr=tc.get_or_add_tcPr()
    s=OxmlElement('w:shd');s.set(qn('w:val'),'clear')
    s.set(qn('w:color'),'auto');s.set(qn('w:fill'),h);pr.append(s)

def borders(table,color='2E74B5'):
    for row in table.rows:
        for cell in row.cells:
            tc=cell._tc;pr=tc.get_or_add_tcPr()
            tb=OxmlElement('w:tcBorders')
            for e in ('top','left','bottom','right','insideH','insideV'):
                x=OxmlElement(f'w:{e}');x.set(qn('w:val'),'single')
                x.set(qn('w:sz'),'4');x.set(qn('w:space'),'0')
                x.set(qn('w:color'),color);tb.append(x)
            pr.append(tb)

def hdr(table,*texts,hbg='1F3564'):
    row=table.rows[0]
    for cell,text in zip(row.cells,texts):
        bg(cell,hbg);p=cell.paragraphs[0]
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r=p.add_run(text);r.bold=True
        r.font.color.rgb=WHITE;r.font.size=Pt(9.5)

def frow(table,ri,vals,bgs=None,bolds=None,aligns=None,size=9.5):
    row=table.rows[ri]
    for i,(cell,val) in enumerate(zip(row.cells,vals)):
        b =bgs[i]    if bgs    else HEX['white']
        bo=bolds[i]  if bolds  else False
        al=aligns[i] if aligns else WD_ALIGN_PARAGRAPH.LEFT
        bg(cell,b);p=cell.paragraphs[0];p.alignment=al
        r=p.add_run(str(val));r.bold=bo;r.font.size=Pt(size)

def H(doc,text,lv=1):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(14 if lv==1 else 10)
    p.paragraph_format.space_after=Pt(4)
    r=p.add_run(text);r.bold=True
    r.font.size=Pt(14 if lv==1 else(12 if lv==2 else 11))
    r.font.color.rgb=DARK if lv==1 else(MID if lv==2 else STEEL)

def P(doc,text,bold=False,italic=False,size=10,color=BLACK,ind=0):
    p=doc.add_paragraph()
    p.paragraph_format.left_indent=Cm(ind)
    p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(2)
    r=p.add_run(text);r.bold=bold;r.italic=italic
    r.font.size=Pt(size);r.font.color.rgb=color

def code_block(doc,text):
    p=doc.add_paragraph()
    p.paragraph_format.left_indent=Cm(0.5)
    p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(4)
    r=p.add_run(text);r.font.name='Courier New';r.font.size=Pt(8.5)
    r.font.color.rgb=CODE
    # light background shading on the paragraph
    pPr=p._p.get_or_add_pPr()
    shd=OxmlElement('w:shd');shd.set(qn('w:val'),'clear')
    shd.set(qn('w:color'),'auto');shd.set(qn('w:fill'),'F0F4FF')
    pPr.append(shd)

def bullet(doc,text,lv=0):
    p=doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent=Cm(0.5+lv*0.4)
    p.paragraph_format.space_after=Pt(1)
    p.add_run(text).font.size=Pt(9.5)

def rule(doc):
    p=doc.add_paragraph();pr=p._p.get_or_add_pPr()
    pb=OxmlElement('w:pBdr');bt=OxmlElement('w:bottom')
    bt.set(qn('w:val'),'single');bt.set(qn('w:sz'),'6')
    bt.set(qn('w:space'),'1');bt.set(qn('w:color'),'2E74B5')
    pb.append(bt);pr.append(pb)

# ═══════════════════════════════════════════════════════════════════════════════
doc=Document()
for s in doc.sections:
    s.top_margin=Cm(2.0);s.bottom_margin=Cm(2.0)
    s.left_margin=Cm(2.5);s.right_margin=Cm(2.5)

# ── BANNER ──────────────────────────────────────────────────────────────────
ban=doc.add_table(rows=1,cols=1);ban.alignment=WD_TABLE_ALIGNMENT.CENTER
bc=ban.cell(0,0);bg(bc,'1F3564')
bp=bc.paragraphs[0];bp.alignment=WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_before=Pt(14);bp.paragraph_format.space_after=Pt(14)
r1=bp.add_run('P01 – HIRING TREND ANALYSIS\n')
r1.bold=True;r1.font.size=Pt(20);r1.font.color.rgb=WHITE
r2=bp.add_run('Technical Requirements Document (TRD)\n')
r2.font.size=Pt(14);r2.font.color.rgb=RGBColor(0xBD,0xD5,0xEE)
r3=bp.add_run('Step 07 of 18  |  Version 1.0  |  Status: Draft\n')
r3.font.size=Pt(10);r3.font.color.rgb=RGBColor(0xBD,0xD5,0xEE)
r4=bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | ABES MCA 2nd Yr DA | Mentor: Dinesh Sereno')
r4.font.size=Pt(9);r4.font.color.rgb=RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

# meta strip
meta=doc.add_table(rows=2,cols=4);meta.alignment=WD_TABLE_ALIGNMENT.CENTER
mdata=[('Project','P01 – Hiring Trend Analysis'),('Step','07 of 18'),
       ('Date',datetime.date.today().strftime('%d %b %Y')),('Status','Draft'),
       ('Stack','Python 3.12 · Streamlit · Plotly · SQLite'),('DB','data/app.db'),
       ('Hosting','Streamlit Community Cloud'),('Note','Step 06 PRD skipped by learner')]
for i,row in enumerate(meta.rows):
    fill=HEX['alt'] if i%2==0 else HEX['white']
    for j,cell in enumerate(row.cells):
        bg(cell,fill);p=cell.paragraphs[0]
        txt=mdata[i*4+j]
        r=p.add_run(f'{txt[0]}: ');r.bold=True;r.font.size=Pt(9)
        p.add_run(txt[1]).font.size=Pt(9)
borders(meta);doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §1  ARCHITECTURE
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'1. System Architecture',1)
P(doc,
  'The HTA application follows a layered data-flow architecture: '
  'raw SQLite data → loading & caching → validation & cleaning → '
  'analysis modules → Streamlit multipage UI. '
  'The diagram below is rendered as a Mermaid flowchart in the README.md.',
  size=10)
doc.add_paragraph()

H(doc,'1.1  Architecture Mermaid Diagram (copy into README.md)',2)
mermaid_code="""```mermaid
flowchart TD
    DB[(SQLite\\ndata/app.db)]
    DB --> DL[src/data_loader.py\\nload_data]
    DL --> CA[st.cache_data\\nIn-memory cache]
    CA --> VL[src/validator.py\\nvalidate_dataframe]
    VL --> CL[src/cleaner.py\\nclean + type-cast]
    CL --> AN[src/analyser.py\\naggregation functions]
    AN --> PG{Streamlit Pages}
    PG --> P1[01_Overview.py\\nKPI cards + summary]
    PG --> P2[02_Role_Demand.py\\nOQ-01 BR-01]
    PG --> P3[03_Skills.py\\nOQ-03 BR-02]
    PG --> P4[04_Salary.py\\nOQ-02 OQ-07 BR-03 BR-06]
    PG --> P5[05_Geography.py\\nOQ-05 BR-07]
    PG --> P6[06_Trends.py\\nOQ-04 BR-05 BR-12]
    PG --> P7[07_Correlation.py\\nOQ-09 BR-09]
    PG --> P8[08_Companies.py\\nOQ-10 BR-04]
    PG --> P9[09_EDA.py\\nBR-13]
    PG --> P10[10_Job_Types.py\\nOQ-06 OQ-08 BR-08 BR-10 BR-11]
    SB[Sidebar Filters\\ndate · industry · city · exp_level] --> PG
    TW[Tailwind CDN\\nst.components.v1.html\\nHeader + KPI cards] --> P1
```"""
code_block(doc,mermaid_code)
doc.add_paragraph()

H(doc,'1.2  Layer Descriptions',2)
layers=[
    ('Data Layer','data/app.db (SQLite)',
     'Single source of truth. 1 table: hiring_trends. ≥500 rows. '
     'All fields validated against business rules BR-R-01 to BR-R-10.'),
    ('Loading Layer','src/data_loader.py',
     'Reads SQLite via pandas.read_sql_query. Decorated with @st.cache_data(ttl=3600) '
     'to avoid re-querying on every Streamlit re-run.'),
    ('Validation Layer','src/validator.py',
     'Checks: row count ≥500, no nulls in required fields, date range 2022–2024, '
     'salary_min < salary_max, experience_level in {Fresher, Mid-level, Senior}. '
     'Raises ValueError with descriptive message on failure.'),
    ('Cleaning Layer','src/cleaner.py',
     'Converts posted_date to datetime, derives month/year columns, '
     'explodes required_skills pipe-delimited string into a tidy skills DataFrame, '
     'derives skills_count per row.'),
    ('Analysis Layer','src/analyser.py',
     'Pure Python/pandas functions — no Streamlit imports. '
     'Returns processed DataFrames ready for Plotly. '
     'Fully unit-testable with pytest without launching a browser.'),
    ('Presentation Layer','pages/*.py',
     'Streamlit multipage files. Each page imports from src/analyser.py and '
     'calls Plotly Express / Graph Objects for charts. '
     'Sidebar filters live in app.py and are accessed via st.session_state.'),
    ('UI Enhancement','Tailwind CDN via st.components.v1.html',
     'Hero header banner and KPI cards on the Overview page. '
     'Injected as raw HTML with Tailwind CSS CDN — no build step required.'),
]
lt=doc.add_table(rows=len(layers)+1,cols=3);lt.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(lt,'Layer','Module / Location','Responsibility')
for i,(l,m,r_) in enumerate(layers,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(lt,i,[l,m,r_],bgs=[b]*3,bolds=[True,True,False])
borders(lt)
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §2  FOLDER STRUCTURE
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'2. Folder & File Structure',1)
P(doc,'The repository root is P01_HiringTrendAnalysis/. All paths are relative to this root.',size=9.5)
doc.add_paragraph()

folder_tree="""P01_HiringTrendAnalysis/
│
├── app.py                        # Streamlit entry point; global sidebar + session state
│
├── pages/
│   ├── 01_Overview.py            # KPI cards, Tailwind header, dataset summary
│   ├── 02_Role_Demand.py         # OQ-01 · BR-01 · FR-03
│   ├── 03_Skills_Analysis.py     # OQ-03 · BR-02 · FR-05
│   ├── 04_Salary_Benchmark.py    # OQ-02 · BR-03 · BR-06 · FR-04
│   ├── 05_Geography.py           # OQ-05 · BR-07 · FR-07
│   ├── 06_Hiring_Trends.py       # OQ-04 · BR-05 · BR-12 · FR-06
│   ├── 07_Correlation.py         # OQ-09 · BR-09 · FR-11
│   ├── 08_Companies.py           # OQ-10 · BR-04 · FR-12
│   ├── 09_EDA_Summary.py         # BR-13 · FR-13
│   └── 10_Job_Type_ExpLevel.py   # OQ-06 · OQ-07 · OQ-08 · BR-08 · BR-10 · BR-11 · FR-08/09/10
│
├── src/
│   ├── __init__.py
│   ├── data_loader.py            # load_data() with @st.cache_data
│   ├── validator.py              # validate_dataframe() — raises on rule violations
│   ├── cleaner.py                # clean_data(), explode_skills()
│   └── analyser.py               # 13 pure analysis functions (one per BR)
│
├── data/
│   ├── app.db                    # SQLite database (committed to repo)
│   ├── generate_data.py          # Faker-based synthetic data generator (Step 10)
│   └── hiring_trends.csv         # CSV companion for quick inspection
│
├── tests/
│   ├── conftest.py               # Shared fixtures: sample DataFrame, DB path
│   ├── test_data_loader.py       # Test load_data() output shape, types
│   ├── test_validator.py         # Test all 10 business rules (BR-R-01 to BR-R-10)
│   ├── test_analyser.py          # Test all 13 analysis functions vs expected output
│   └── test_pages.py             # Streamlit AppTest smoke tests per page
│
├── docs/                         # Steps 01–18 deliverables (Word, Excel, PPT)
│
├── .streamlit/
│   └── config.toml               # Theme: primaryColor, backgroundColor, font
│
├── requirements.txt              # Pinned dependencies
├── .gitignore
└── README.md                     # Setup guide + Mermaid architecture diagram"""

code_block(doc,folder_tree)
doc.add_paragraph()

H(doc,'2.1  Key File Responsibilities',2)
files=[
    ('app.py',          'Global Streamlit config, page registry, sidebar filter widgets, session_state initialisation'),
    ('src/data_loader.py','Single function load_data() that opens app.db and returns a cached DataFrame'),
    ('src/validator.py', 'validate_dataframe(df) — checks all BR-R-01..10 rules; called on every startup'),
    ('src/cleaner.py',   'clean_data(df) → typed df; explode_skills(df) → tidy skills df; derive_features(df)'),
    ('src/analyser.py',  '13 functions: get_role_demand, get_skill_freq, get_salary_stats … (one per BR)'),
    ('generate_data.py', 'Step 10 script: uses Faker + custom distributions to populate app.db'),
    ('conftest.py',      'pytest fixtures: sample_df (50-row in-memory df), db_path (test SQLite)'),
    ('config.toml',      'primaryColor=#2E74B5, backgroundColor=#FFFFFF, font=sans serif'),
]
ft=doc.add_table(rows=len(files)+1,cols=2);ft.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(ft,'File','Responsibility')
for i,(f,r_) in enumerate(files,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(ft,i,[f,r_],bgs=[b,b],bolds=[True,False])
    # monospace for file name
    ft.rows[i].cells[0].paragraphs[0].runs[0].font.name='Courier New'
    ft.rows[i].cells[0].paragraphs[0].runs[0].font.size=Pt(9)
borders(ft)
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §3  FUNCTIONAL REQUIREMENTS (FR)
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'3. Functional Requirements (FR)',1)
P(doc,
  'Each FR defines a specific system behaviour. FR IDs are used in the RTM (Step 04) to fill '
  'the Feature ID column. FR ↔ BR traceability is mandatory. '
  'All FRs are Must unless marked.',size=9.5)
doc.add_paragraph()

frs=[
    # (FR_ID, description, linked_BR, linked_OQ, module, priority)
    ('FR-01','Load the hiring_trends SQLite table into a pandas DataFrame on app start using @st.cache_data',
     'BR-01..13','All OQs','src/data_loader.py','Must'),
    ('FR-02','Validate the loaded DataFrame against all 10 business rules (BR-R-01..10); display error and halt if any rule fails',
     'BR-R-01..10','—','src/validator.py','Must'),
    ('FR-03','Render an interactive ranked bar chart of job posting volume by role for a user-selected date range',
     'BR-01','OQ-01','pages/02_Role_Demand.py','Must'),
    ('FR-04','Render grouped salary statistics (avg / min / max) per industry and per exp-level + city combination',
     'BR-03, BR-06','OQ-02','pages/04_Salary_Benchmark.py','Must'),
    ('FR-05','Explode pipe-delimited skills, count frequency by industry filter, render top-N horizontal bar chart',
     'BR-02','OQ-03','pages/03_Skills_Analysis.py','Must'),
    ('FR-06','Resample postings by calendar month, render an interactive line chart of total hiring volume over time',
     'BR-05','OQ-04','pages/06_Hiring_Trends.py','Must'),
    ('FR-07','Group postings by city, render a ranked bar chart of geographic posting concentration',
     'BR-07','OQ-05','pages/05_Geography.py','Must'),
    ('FR-08','Compute job-type proportions per industry, render a donut/pie chart; allow industry filter',
     'BR-10','OQ-06','pages/10_Job_Type_ExpLevel.py','Must'),
    ('FR-09','Rank industries by posting count for a user-selected skill keyword; render horizontal bar chart',
     'BR-11','OQ-07','pages/10_Job_Type_ExpLevel.py','Must'),
    ('FR-10','Pivot experience-level counts per job role, render a 100% stacked bar chart',
     'BR-08','OQ-08','pages/10_Job_Type_ExpLevel.py','Must'),
    ('FR-11','Derive skills_count per row; compute Pearson r; render OLS scatter plot with trendline; display r and p-value',
     'BR-09','OQ-09','pages/07_Correlation.py','Should'),
    ('FR-12','Rank companies by posting count within a user-selected domain; render bar chart + sortable table',
     'BR-04','OQ-10','pages/08_Companies.py','Must'),
    ('FR-13','Render EDA panel: pandas describe() table, null count bar, IQR outlier count per numeric field',
     'BR-13','—','pages/09_EDA_Summary.py','Could'),
    ('FR-14','Provide global sidebar filters: date-range slider, industry multiselect, city multiselect, experience-level multiselect; apply to all pages via st.session_state',
     'BR-01..13','All OQs','app.py','Must'),
    ('FR-15','Inject Tailwind CDN HTML via st.components.v1.html to render hero header and 4 KPI metric cards on Overview page',
     'BR-01,BR-05,BR-07,BR-04','OQ-01,04,05,10','pages/01_Overview.py','Must'),
    ('FR-16','Provide CSV download button (st.download_button) for currently filtered DataFrame on each analysis page',
     'BR-01..13','All OQs','All pages','Should'),
]

fr_tbl=doc.add_table(rows=len(frs)+1,cols=6)
fr_tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(fr_tbl,'FR ID','Functional Requirement','Linked BR(s)','Linked OQ(s)','Module','Priority')
prio_bg={'Must':HEX['must'],'Should':HEX['should'],'Could':HEX['could']}
prio_fg={'Must':'1F3564','Should':'1F5C2E','Could':'7B5700'}
for i,(fid,desc,br,oq,mod,prio) in enumerate(frs,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    pb=prio_bg.get(prio,HEX['white'])
    pf=prio_fg.get(prio,'000000')
    row=fr_tbl.rows[i]
    data=[fid,desc,br,oq,mod,prio]
    for ci,(cell,val) in enumerate(zip(row.cells,data)):
        cbg=pb if ci==5 else b
        cfg=pf if ci==5 else '000000'
        bg(cell,cbg);p=cell.paragraphs[0]
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER if ci in (0,2,3,5) else WD_ALIGN_PARAGRAPH.LEFT
        r=p.add_run(str(val));r.font.size=Pt(9)
        if ci==0: r.bold=True;r.font.name='Courier New'
        if ci==4: r.font.name='Courier New'
        if ci==5: r.bold=True
        if ci==5:
            r.font.color.rgb=RGBColor(*[int(pf[k:k+2],16) for k in (0,2,4)])
borders(fr_tbl)
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §4  DATA LAYER
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'4. Data Layer Design',1)

H(doc,'4.1  data_loader.py — load_data()',2)
code_block(doc,
"""# src/data_loader.py
import sqlite3, pandas as pd
import streamlit as st

DB_PATH = "data/app.db"

@st.cache_data(ttl=3600, show_spinner="Loading hiring data…")
def load_data(db_path: str = DB_PATH) -> pd.DataFrame:
    \"\"\"Load hiring_trends table from SQLite into a DataFrame.\"\"\"
    con = sqlite3.connect(db_path)
    df  = pd.read_sql_query("SELECT * FROM hiring_trends", con)
    con.close()
    return df""")
doc.add_paragraph()

H(doc,'4.2  validator.py — validate_dataframe()',2)
code_block(doc,
"""# src/validator.py
REQUIRED_FIELDS = ['job_title','company_name','industry','city','state',
                   'job_type','experience_level','required_skills',
                   'salary_min','salary_max','posted_date']
EXP_LEVELS  = {'Fresher','Mid-level','Senior'}
JOB_TYPES   = {'Full-time','Part-time','Remote','Hybrid','Contract'}
DATE_MIN    = '2022-01-01'
DATE_MAX    = '2024-12-31'

def validate_dataframe(df: pd.DataFrame) -> None:
    if len(df) < 500:
        raise ValueError(f"Dataset has {len(df)} rows — minimum 500 required (BR-R-10)")
    for f in REQUIRED_FIELDS:
        if df[f].isna().any():
            raise ValueError(f"Null values found in required field: {f} (BR-R-01)")
    if not set(df['experience_level'].unique()).issubset(EXP_LEVELS):
        raise ValueError("Invalid experience_level values (BR-R-04)")
    if not set(df['job_type'].unique()).issubset(JOB_TYPES):
        raise ValueError("Invalid job_type values (BR-R-05)")
    if (df['salary_min'] >= df['salary_max']).any():
        raise ValueError("salary_min >= salary_max violation (BR-R-09)")""")
doc.add_paragraph()

H(doc,'4.3  cleaner.py — clean_data() & explode_skills()',2)
code_block(doc,
"""# src/cleaner.py
import pandas as pd

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['posted_date']   = pd.to_datetime(df['posted_date'])
    df['month']         = df['posted_date'].dt.to_period('M').astype(str)
    df['year']          = df['posted_date'].dt.year
    df['skills_count']  = df['required_skills'].str.split('|').str.len()
    df['salary_mid']    = (df['salary_min'] + df['salary_max']) / 2
    return df

def explode_skills(df: pd.DataFrame) -> pd.DataFrame:
    \"\"\"Return tidy long-form skills DataFrame.\"\"\"
    return (df[['job_id','industry','required_skills']]
              .assign(skill=lambda x: x['required_skills'].str.split('|'))
              .explode('skill')
              .assign(skill=lambda x: x['skill'].str.strip()))""")
doc.add_paragraph()

H(doc,'4.4  Caching Strategy',2)
cache_data=[
    ('load_data()','@st.cache_data(ttl=3600)','Avoids re-querying SQLite on every Streamlit re-run; invalidated hourly'),
    ('clean_data()','@st.cache_data','Memoised per input DataFrame hash; cleaning is deterministic'),
    ('explode_skills()','@st.cache_data','Skills explosion is expensive on 1 000 rows; cache saves ~200 ms per re-run'),
    ('All analyser fns','@st.cache_data','Each analysis function decorated; filters invalidate cache automatically'),
]
ct=doc.add_table(rows=len(cache_data)+1,cols=3);ct.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(ct,'Function','Cache Decorator','Rationale')
for i,row in enumerate(cache_data,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(ct,i,row,bgs=[b]*3,bolds=[True,False,False])
    ct.rows[i].cells[0].paragraphs[0].runs[0].font.name='Courier New'
borders(ct)
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §5  STATISTICAL METHODS PER OQ
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'5. Statistical Method per Operational Question',1)
P(doc,'This section maps each OQ to its pandas transformation, statistical method, and Plotly chart type.',size=9.5)
doc.add_paragraph()

stats=[
    ('OQ-01','Role demand ranking',
     'groupby("job_title").size().sort_values(ascending=False)',
     'Frequency count','plotly.express.bar (horizontal, top-15)','BR-01','FR-03'),
    ('OQ-02','Salary stats by industry & exp+city',
     'groupby(["industry"]).agg(avg=("salary_max","mean"),\nmin=("salary_min","min"),max=("salary_max","max"))',
     'Descriptive stats (mean, min, max)','px.bar (grouped) + px.box','BR-03,BR-06','FR-04'),
    ('OQ-03','Skill frequency by industry',
     'explode_skills(df).groupby("skill").size().nlargest(n)',
     'Frequency count after pipe-delimited explode','px.bar (horizontal)','BR-02','FR-05'),
    ('OQ-04','Monthly hiring volume trend',
     'groupby("month").size().sort_index()',
     'Time series aggregation (monthly resampling)','px.line + px.bar (dual)','BR-05,BR-12','FR-06'),
    ('OQ-05','City-wise posting hotspots',
     'groupby("city").size().sort_values(ascending=False)',
     'Frequency count','px.bar (horizontal, top-15)','BR-07','FR-07'),
    ('OQ-06','Job-type distribution',
     'groupby("job_type").size() / len(df)  # proportions',
     'Proportion / percentage','px.pie (donut, hole=0.4)','BR-10','FR-08'),
    ('OQ-07','Industry hiring volume by skill',
     'explode_skills → filter skill keyword → groupby("industry").size()',
     'Filtered frequency count','px.bar (horizontal)','BR-11','FR-09'),
    ('OQ-08','Exp-level demand per role',
     'groupby(["job_title","experience_level"]).size().unstack(fill_value=0)\n.apply(lambda r: r/r.sum()*100, axis=1)',
     'Normalised proportion (100% stacked)','px.bar (barmode="stack", barnorm="percent")','BR-08','FR-10'),
    ('OQ-09','Skill count ↔ salary correlation',
     'scipy.stats.pearsonr(df["skills_count"], df["salary_max"])\nstatsmodels OLS: salary_max ~ skills_count',
     'Pearson r + OLS linear regression','px.scatter + px.scatter trendline="ols"','BR-09','FR-11'),
    ('OQ-10','Top recruiting companies',
     'groupby("company_name").size().nlargest(10)',
     'Frequency count + rank','px.bar + st.dataframe (sortable)','BR-04','FR-12'),
]

st_tbl=doc.add_table(rows=len(stats)+1,cols=7)
st_tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(st_tbl,'OQ','Topic','pandas Transform','Statistical Method','Plotly Chart','BR(s)','FR')
for i,row in enumerate(stats,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(st_tbl,i,row,bgs=[b]*7,bolds=[True,False,False,False,False,False,False],size=8.5)
    # monospace for code column
    st_tbl.rows[i].cells[2].paragraphs[0].runs[0].font.name='Courier New'
    st_tbl.rows[i].cells[2].paragraphs[0].runs[0].font.size=Pt(8)
borders(st_tbl)
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §6  UI NOTES — TAILWIND, THEME, LAYOUT
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'6. UI Design Notes',1)

H(doc,'6.1  Tailwind CDN Integration',2)
P(doc,
  'Tailwind CSS is injected via st.components.v1.html on the Overview page only. '
  'No build toolchain (npm/vite) is required — the CDN script tag is embedded in the HTML string.',
  size=10)
code_block(doc,
"""# pages/01_Overview.py  — Tailwind hero header + KPI cards
import streamlit.components.v1 as components

HERO_HTML = \"\"\"
<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
<div class="bg-blue-900 text-white p-6 rounded-xl shadow-lg mb-4">
  <h1 class="text-3xl font-bold">📊 Hiring Trend Analysis</h1>
  <p class="text-blue-200 mt-1">P01 · Abhayjeet Singh · ABES MCA 2nd Yr DA</p>
</div>
<div class="grid grid-cols-4 gap-4 mb-6">
  <div class="bg-blue-100 p-4 rounded-lg shadow text-center">
    <div class="text-2xl font-bold text-blue-800">{total_jobs}</div>
    <div class="text-sm text-gray-600">Total Postings</div>
  </div>
  <!-- repeat for avg_salary, top_city, top_role -->
</div>
\"\"\"
components.html(HERO_HTML.format(total_jobs=len(df), ...), height=250)""")
doc.add_paragraph()

H(doc,'6.2  Streamlit Theme (config.toml)',2)
code_block(doc,
"""# .streamlit/config.toml
[theme]
primaryColor       = "#2E74B5"
backgroundColor    = "#FFFFFF"
secondaryBackgroundColor = "#EEF4FB"
textColor          = "#1A1A2E"
font               = "sans serif"

[server]
headless   = true
enableCORS = false""")
doc.add_paragraph()

H(doc,'6.3  Page Layout Standards',2)
layout=[
    ('Page width',       'st.set_page_config(layout="wide") on all pages'),
    ('Sidebar',          'Global filters: date range (slider), industry (multiselect), city, exp_level'),
    ('Chart container',  'st.plotly_chart(fig, use_container_width=True) on all charts'),
    ('Metric cards',     'st.metric() for quick KPIs on each page; Tailwind cards on Overview only'),
    ('Data table',       'st.dataframe(df, use_container_width=True) below each chart; CSV download button'),
    ('Expander',         'st.expander("ℹ️ How to read this chart") with one-line interpretation per chart'),
    ('Error display',    'st.error(msg) with actionable text if validator raises ValueError'),
    ('Spinner',          'st.spinner("Loading…") wraps load_data() call on first render'),
    ('Colour palette',   'Plotly charts use color_discrete_sequence=px.colors.sequential.Blues_r'),
    ('Responsiveness',   'wide layout + use_container_width handles most screen sizes; min 1024 px width'),
]
lt2=doc.add_table(rows=len(layout)+1,cols=2);lt2.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(lt2,'UI Element','Implementation Standard')
for i,(el,std) in enumerate(layout,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(lt2,i,[el,std],bgs=[b,b],bolds=[True,False])
borders(lt2)
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §7  PERFORMANCE, ERROR HANDLING, SECURITY
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'7. Performance, Error Handling and Security',1)

H(doc,'7.1  Performance Requirements',2)
perf=[
    ('PF-01','Initial page load','< 3 s on broadband (BR KPI O-04)','@st.cache_data on load_data() + clean_data()'),
    ('PF-02','Chart re-render after filter','< 1 s','Cached analysis functions; Plotly renders client-side'),
    ('PF-03','Dataset size supported','≤ 10 000 rows without degradation','pandas in-memory; SQLite single-file'),
    ('PF-04','Memory footprint','< 200 MB RSS on Streamlit cloud','1 000-row dataset ≈ 500 KB in memory'),
    ('PF-05','Cold start (Streamlit cloud)','< 10 s (cloud container spin-up)','requirements.txt kept minimal; no heavy ML libs'),
]
pt=doc.add_table(rows=len(perf)+1,cols=4);pt.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(pt,'ID','Requirement','Target','Implementation')
for i,row in enumerate(perf,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(pt,i,row,bgs=[b]*4,bolds=[True,False,False,False])
borders(pt)
doc.add_paragraph()

H(doc,'7.2  Error Handling Strategy',2)
err=[
    ('EH-01','Database not found',
     'FileNotFoundError → st.error("Database file not found. Run data/generate_data.py first.") + st.stop()'),
    ('EH-02','Validation failure',
     'ValueError from validator → st.error(str(e)) + st.stop(); prevents blank / misleading charts'),
    ('EH-03','Empty DataFrame after filter',
     'if len(filtered_df) == 0: st.warning("No data matches the current filters. Adjust the sidebar.") + st.stop()'),
    ('EH-04','SQLite connection error',
     'try/except sqlite3.Error → st.error("Database connection failed. Check data/app.db.") + st.stop()'),
    ('EH-05','Unexpected column missing',
     'KeyError in analyser → caught per function; st.warning with column name + st.stop() on page'),
]
et=doc.add_table(rows=len(err)+1,cols=3);et.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(et,'ID','Scenario','Handling')
for i,row in enumerate(err,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(et,i,row,bgs=[b]*3,bolds=[True,False,False])
borders(et)
doc.add_paragraph()

H(doc,'7.3  Security Considerations',2)
sec=[
    ('SC-01','No real personal data','Dataset is 100% synthetic (Faker). No PII committed to repo. (BR-R-07)'),
    ('SC-02','No credentials in repo','.gitignore excludes .env; no API keys or passwords used in this project'),
    ('SC-03','Read-only DB access','app.db is opened with sqlite3.connect in read mode (check_same_thread=False); no INSERT/UPDATE in app'),
    ('SC-04','Input sanitisation','Sidebar widgets (st.selectbox, st.multiselect, st.date_input) return validated Python types; no raw SQL user input'),
    ('SC-05','Dependency hygiene','All packages pinned in requirements.txt; no unpinned transitive dependencies'),
    ('SC-06','Public repo disclosure','README explicitly states dataset is synthetic; no business-sensitive information is present'),
]
st2=doc.add_table(rows=len(sec)+1,cols=3);st2.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(st2,'ID','Area','Control')
for i,row in enumerate(sec,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(st2,i,row,bgs=[b]*3,bolds=[True,False,False])
borders(st2)
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §8  REQUIREMENTS.TXT
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'8. requirements.txt (Pinned)',1)
P(doc,'All versions pinned to ensure reproducible builds on Streamlit Community Cloud.',size=9.5)
doc.add_paragraph()
code_block(doc,
"""# requirements.txt  —  P01 Hiring Trend Analysis
# Generated: """ + datetime.date.today().strftime('%d %b %Y') + """

# ── Core app ──────────────────────────────────────────────────────────────
streamlit==1.35.0
pandas==2.2.2
numpy==1.26.4
plotly==5.22.0

# ── Statistics ────────────────────────────────────────────────────────────
scipy==1.13.1
statsmodels==0.14.2

# ── Data generation (Step 10 — not needed at runtime) ────────────────────
faker==24.11.0

# ── Testing ───────────────────────────────────────────────────────────────
pytest==8.2.2
pytest-cov==5.0.0

# ── Utilities ─────────────────────────────────────────────────────────────
# (sqlite3 is part of Python stdlib — no entry needed)""")

P(doc,
  'Note: faker is included in requirements.txt so generate_data.py can be run on any '
  'clone of the repo. It is not imported by the Streamlit app itself.',
  italic=True, size=9, color=AMBER)
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §9  DEPLOYMENT REQUIREMENTS
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'9. Deployment Requirements',1)

dep_steps=[
    ('DR-01','GitHub Repository',
     'Public GitHub repo; name: P01_HiringTrendAnalysis',
     'Required by Streamlit Community Cloud for free-tier deployment'),
    ('DR-02','Streamlit Community Cloud',
     'Account at share.streamlit.io; link GitHub repo; set Main file: app.py',
     'Free tier; no credit card; ≤1 app per account on free plan'),
    ('DR-03','data/app.db in repo',
     'SQLite file must be committed to the repo (not in .gitignore)',
     'Streamlit Cloud clones the repo — DB must be part of it'),
    ('DR-04','requirements.txt at root',
     'Streamlit Cloud auto-detects and installs from root-level requirements.txt',
     'Exact path: /requirements.txt (not /src/ or /app/)'),
    ('DR-05','.streamlit/config.toml',
     'Committed to repo; contains theme and server settings',
     'Applied automatically by Streamlit Cloud on deployment'),
    ('DR-06','Python version',
     'Set runtime.txt with content: python-3.12',
     'Ensures Streamlit Cloud uses Python 3.12 (not default 3.9)'),
    ('DR-07','App URL',
     'Format: https://<username>-p01-hiringtrendanalysis-app.streamlit.app',
     'Share URL in Step 16 Deployment Guide and Step 17 Project Document'),
    ('DR-08','Local run command',
     'streamlit run app.py  (from repo root with venv activated)',
     'Backup demo fallback if cloud is unavailable during presentation'),
]
dt2=doc.add_table(rows=len(dep_steps)+1,cols=4)
dt2.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(dt2,'DR ID','Requirement','Implementation','Notes')
for i,row in enumerate(dep_steps,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(dt2,i,row,bgs=[b]*4,bolds=[True,False,False,False])
borders(dt2)
doc.add_paragraph()

H(doc,'9.1  runtime.txt',2)
code_block(doc,'python-3.12')
doc.add_paragraph()

H(doc,'9.2  Deployment Checklist',2)
deploy_checks=[
    ('☐','GitHub repo is public'),
    ('☐','app.py is at repo root'),
    ('☐','requirements.txt is at repo root (all packages pinned)'),
    ('☐','data/app.db is committed and not in .gitignore'),
    ('☐','.streamlit/config.toml is committed'),
    ('☐','runtime.txt contains "python-3.12"'),
    ('☐','Streamlit Cloud app linked to correct branch (main)'),
    ('☐','App loads without error on first deploy'),
    ('☐','All 10 pages render correctly on the cloud URL'),
    ('☐','CSV download buttons work on the cloud URL'),
]
for sym,item in deploy_checks:
    bullet(doc,f'{sym}  {item}')
doc.add_paragraph();rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §10  TECHNICAL RISKS
# ══════════════════════════════════════════════════════════════════════════════
H(doc,'10. Technical Risks',1)
tech_risks=[
    ('TR-01','Streamlit version API change',  'M','M','M',
     'Pin streamlit==1.35.0 in requirements.txt; test on exact version'),
    ('TR-02','st.cache_data miss on filter change','M','M','M',
     'Use @st.cache_data with explicit hash_funcs for DataFrame inputs where needed'),
    ('TR-03','Tailwind CDN blocked by browser CSP','L','M','L',
     'Provide non-Tailwind fallback using st.metric() if CDN fails; test on Chrome+Firefox'),
    ('TR-04','SQLite WAL mode conflict on cloud','L','H','M',
     'Open DB in read-only mode; app never writes to DB during runtime'),
    ('TR-05','statsmodels OLS fails on small filtered datasets','M','M','M',
     'Guard with: if len(filtered_df) < 10: st.warning("Too few records for regression") + st.stop()'),
    ('TR-06','Plotly chart not rendering in AppTest','M','M','M',
     'Use Streamlit AppTest for page-load smoke tests only; chart correctness tested via analyser unit tests'),
    ('TR-07','pandas 2.x breaking change from 1.x','L','H','M',
     'All code written for pandas 2.2.x from start; no legacy .append() or .iteritems() used'),
    ('TR-08','GitHub repo size limit (100 MB)',  'L','L','L',
     'app.db with 1 000 rows ≈ 150 KB; well within GitHub 100 MB file limit'),
]
rt=doc.add_table(rows=len(tech_risks)+1,cols=7)
rt.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(rt,'TR ID','Risk','Prob','Impact','Rating','Mitigation','Residual')
risk_bg_={'H':'FFE5E5','M':'FFF3CD','L':'E5FFE5'}
risk_fg_={'H':'C00000','M':'7B5700','L':'1F5C2E'}
for i,(rid,desc,prob,imp,rat,mit) in enumerate(tech_risks,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    row=rt.rows[i]
    for ci,(cell,val) in enumerate(zip(row.cells,[rid,desc,prob,imp,rat,mit,'Low']),0):
        is_r=(ci in (2,3,4,6))
        cb=risk_bg_.get(val,b) if is_r else b
        fg=risk_fg_.get(val,'000000') if is_r else '000000'
        bg(cell,cb);p=cell.paragraphs[0]
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER if is_r or ci==0 else WD_ALIGN_PARAGRAPH.LEFT
        r=p.add_run(str(val));r.font.size=Pt(9)
        if is_r or ci==0: r.bold=True
        if is_r:
            r.font.color.rgb=RGBColor(*[int(fg[k:k+2],16) for k in (0,2,4)])
borders(rt)
doc.add_paragraph();rule(doc)

# ── DOC CONTROL ───────────────────────────────────────────────────────────────
H(doc,'Document Control',2)
dc=doc.add_table(rows=2,cols=4);dc.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(dc,'Version','Date','Author','Summary')
frow(dc,1,['1.0',datetime.date.today().strftime('%d %b %Y'),
            'Abhayjeet Singh','Initial TRD — Step 07 (Step 06 PRD skipped)'],
     bgs=[HEX['alt']]*4)
borders(dc)
doc.add_paragraph()
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run('P01 – Hiring Trend Analysis  |  TRD v1.0  |  Step 07 of 18  |  ABES MCA 2nd Yr DA  |  Mentor: Dinesh Sereno')
r.font.size=Pt(8);r.font.color.rgb=RGBColor(0x88,0x88,0x88);r.italic=True

doc.save(OUT_FILE)
print(f"SAVED: {OUT_FILE}")
