"""
Step 11 – User Stories Document Generator & RTM Updater
Project : P01 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES MCA 2nd Yr DA
"""

import os, datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOCS_DIR = r"C:\Hiring Trend Analysis\docs"
OUT_DOCX = os.path.join(DOCS_DIR, "P01_HTA_Step11_UserStories.docx")
RTM_FILE = os.path.join(DOCS_DIR, "P01_HTA_Step04_RTM.xlsx")

# ── 1. USER STORIES DEFINITION (16 STORIES ACROSS 5 EPICS) ───────────────────

USER_STORIES = [
    # Epic 1: Data & Pipeline Infrastructure
    {
        "us_id": "US-01",
        "epic": "EP-01: Data & Pipeline Infrastructure",
        "persona": "Data Analyst (SH-02)",
        "statement": "As a Data Analyst, I want to automatically load and validate the SQLite hiring dataset with in-memory caching so that dashboard queries respond in under 2 seconds without database reload overhead.",
        "feature_id": "FR-01, FR-02",
        "br_id": "BR-13",
        "oq_id": "All",
        "priority": "Must",
        "story_points": 3,
        "given": "the Streamlit application starts or receives a user session,",
        "when": "the data loading function is invoked,",
        "then": "the dataset is read from data/app.db, validated for required schema integrity, and stored in cache for sub-second reuse."
    },
    {
        "us_id": "US-02",
        "epic": "EP-01: Data & Pipeline Infrastructure",
        "persona": "Data Analyst (SH-02)",
        "statement": "As a Data Analyst, I want an EDA summary panel showing missing values, descriptive statistics, and IQR outliers so that I can audit data quality before interpreting charts.",
        "feature_id": "FR-13",
        "br_id": "BR-13",
        "oq_id": "All",
        "priority": "Could",
        "story_points": 2,
        "given": "the user opens the '09_EDA_Summary' page,",
        "when": "the automated data audit runs,",
        "then": "a pandas summary table, null-count summary, and statistical outlier badges are displayed."
    },
    # Epic 2: Core Dashboard & Global Filtering
    {
        "us_id": "US-03",
        "epic": "EP-02: Executive Dashboard & Navigation",
        "persona": "HR Manager (SH-03)",
        "statement": "As an HR Manager, I want high-level KPI cards showing Total Postings, Average Salary, Top City, and Top Role so that I can grasp macro talent market health at a glance.",
        "feature_id": "FR-15",
        "br_id": "BR-01, BR-05, BR-06, BR-07",
        "oq_id": "OQ-01, OQ-02, OQ-04, OQ-05",
        "priority": "Must",
        "story_points": 3,
        "given": "the user lands on the '01_Overview' dashboard,",
        "when": "the initial load completes,",
        "then": "four styled KPI metric cards display aggregate summary figures with clean Tailwind card styling."
    },
    {
        "us_id": "US-04",
        "epic": "EP-02: Executive Dashboard & Navigation",
        "persona": "HR Recruiter (SH-04)",
        "statement": "As an HR Recruiter, I want global sidebar filters for date range, industry sector, city, and experience level so that all downstream charts dynamically synchronize to my hiring criteria.",
        "feature_id": "FR-14",
        "br_id": "BR-01 to BR-12",
        "oq_id": "All",
        "priority": "Must",
        "story_points": 5,
        "given": "the sidebar filter widgets are modified by the user,",
        "when": "a specific industry or city is selected,",
        "then": "st.session_state updates and all active page visualizations re-render based strictly on the filtered cohort."
    },
    # Epic 3: Role & Skills Demand Analysis
    {
        "us_id": "US-05",
        "epic": "EP-03: Demand & Skills Intelligence",
        "persona": "HR Manager (SH-03)",
        "statement": "As an HR Manager, I want a ranked horizontal bar chart of job roles by posting volume so that I can identify high-demand talent segments.",
        "feature_id": "FR-03",
        "br_id": "BR-01",
        "oq_id": "OQ-01",
        "priority": "Must",
        "story_points": 3,
        "given": "the user is on the '02_Role_Demand' page,",
        "when": "the date or industry filter is applied,",
        "then": "the top 10 most demanded job roles are rendered in descending order with exact posting counts."
    },
    {
        "us_id": "US-06",
        "epic": "EP-03: Demand & Skills Intelligence",
        "persona": "Technical Recruiter (SH-04)",
        "statement": "As a Technical Recruiter, I want to see the top-N most requested skills within a selected industry so that I can draft competitive job descriptions.",
        "feature_id": "FR-05",
        "br_id": "BR-02",
        "oq_id": "OQ-03",
        "priority": "Must",
        "story_points": 3,
        "given": "an industry sector is selected on the '03_Skills' page,",
        "when": "the skill extraction pipeline runs,",
        "then": "the top 10 required skills are visualized on a frequency bar chart alongside a tabular percentage breakdown."
    },
    {
        "us_id": "US-07",
        "epic": "EP-03: Demand & Skills Intelligence",
        "persona": "Placement Officer (SH-06)",
        "statement": "As a Placement Officer, I want to see the proportion of experience levels (Fresher vs Mid vs Senior) per role so that I know which student cohorts match current demand.",
        "feature_id": "FR-10",
        "br_id": "BR-08",
        "oq_id": "OQ-08",
        "priority": "Must",
        "story_points": 3,
        "given": "the '10_Job_Type_ExpLevel' page is opened,",
        "when": "the experience distribution view is selected,",
        "then": "a 100% stacked bar chart shows the exact percentage of Fresher, Mid, and Senior postings per role."
    },
    {
        "us_id": "US-08",
        "epic": "EP-03: Demand & Skills Intelligence",
        "persona": "Career Counselor (SH-06)",
        "statement": "As a Career Counselor, I want to rank industries by demand for a specific skill keyword (e.g., 'Python') so that I can advise students on emerging sectors.",
        "feature_id": "FR-09",
        "br_id": "BR-11",
        "oq_id": "OQ-07",
        "priority": "Should",
        "story_points": 2,
        "given": "a skill keyword is entered in the search filter,",
        "when": "the industry query executes,",
        "then": "a ranked chart displays industries hiring most heavily for that skill."
    },
    # Epic 4: Compensation & Market Economics
    {
        "us_id": "US-09",
        "epic": "EP-04: Compensation & Geographic Benchmarks",
        "persona": "Job Seeker (SH-05)",
        "statement": "As a Job Seeker, I want to view average, minimum, and maximum salaries filtered by experience level and city so that I can benchmark fair compensation expectations.",
        "feature_id": "FR-04",
        "br_id": "BR-03",
        "oq_id": "OQ-02",
        "priority": "Must",
        "story_points": 3,
        "given": "the user filters by 'Fresher' and 'Bangalore' on '04_Salary_Benchmark',",
        "when": "salary metrics are calculated,",
        "then": "grouped bar charts and box plots reveal the min, mean, and max CTC offered."
    },
    {
        "us_id": "US-10",
        "epic": "EP-04: Compensation & Geographic Benchmarks",
        "persona": "HR Manager (SH-03)",
        "statement": "As an HR Manager, I want a single cross-industry salary comparison view so that I can budget company compensation packages competitively.",
        "feature_id": "FR-04",
        "br_id": "BR-06",
        "oq_id": "OQ-02",
        "priority": "Must",
        "story_points": 2,
        "given": "the user views the industry comparison tab on '04_Salary_Benchmark',",
        "when": "the chart renders,",
        "then": "all 8 industries are displayed sorted by average salary in descending order."
    },
    {
        "us_id": "US-11",
        "epic": "EP-04: Compensation & Geographic Benchmarks",
        "persona": "Technical Recruiter (SH-04)",
        "statement": "As a Technical Recruiter, I want an OLS scatter plot showing skill count versus offered salary with Pearson r so that I can justify higher budgets for complex multi-skilled roles.",
        "feature_id": "FR-11",
        "br_id": "BR-09",
        "oq_id": "OQ-09",
        "priority": "Should",
        "story_points": 5,
        "given": "the user navigates to the '07_Correlation' page,",
        "when": "the regression model evaluates,",
        "then": "a scatter plot with an OLS linear trendline is rendered alongside the Pearson r correlation and p-value."
    },
    {
        "us_id": "US-12",
        "epic": "EP-04: Compensation & Geographic Benchmarks",
        "persona": "Job Seeker (SH-05)",
        "statement": "As a Job Seeker, I want to view job posting concentration ranked by city so that I know which urban centers offer the greatest density of opportunities.",
        "feature_id": "FR-07",
        "br_id": "BR-07",
        "oq_id": "OQ-05",
        "priority": "Must",
        "story_points": 3,
        "given": "the '05_Geography' page is accessed,",
        "when": "city aggregates are computed,",
        "then": "a ranked geographic bar chart displays postings across all 15 Indian cities."
    },
    # Epic 5: Longitudinal Trends & Employer Insights
    {
        "us_id": "US-13",
        "epic": "EP-05: Longitudinal Trends & Employer Intelligence",
        "persona": "HR Manager (SH-03)",
        "statement": "As an HR Manager, I want an interactive monthly hiring timeline so that I can anticipate seasonal surges and time recruitment drives effectively.",
        "feature_id": "FR-06",
        "br_id": "BR-05",
        "oq_id": "OQ-04",
        "priority": "Must",
        "story_points": 3,
        "given": "the '06_Hiring_Trends' page is open,",
        "when": "the time-series query runs across 2022-2024,",
        "then": "a chronological line chart displays monthly hiring volumes with clear identification of hiring peaks."
    },
    {
        "us_id": "US-14",
        "epic": "EP-05: Longitudinal Trends & Employer Intelligence",
        "persona": "Job Seeker (SH-05)",
        "statement": "As a Job Seeker, I want to track monthly hiring trends for one specific role so that I can submit applications during historical peak recruitment windows.",
        "feature_id": "FR-06",
        "br_id": "BR-12",
        "oq_id": "OQ-04",
        "priority": "Should",
        "story_points": 2,
        "given": "the user selects a single role (e.g. 'Data Analyst') on '06_Hiring_Trends',",
        "when": "the time series filters to that role,",
        "then": "a single-trend line depicts monthly fluctuations for that specific job title."
    },
    {
        "us_id": "US-15",
        "epic": "EP-05: Longitudinal Trends & Employer Intelligence",
        "persona": "Technical Recruiter (SH-04)",
        "statement": "As a Technical Recruiter, I want to see the distribution of employment types (Full-time, Remote, Hybrid) per industry so that I can benchmark our flexible-work policy.",
        "feature_id": "FR-08",
        "br_id": "BR-10",
        "oq_id": "OQ-06",
        "priority": "Should",
        "story_points": 2,
        "given": "an industry is selected on '10_Job_Type_ExpLevel',",
        "when": "the job type distribution is computed,",
        "then": "a donut chart displays proportions of Full-time, Remote, Hybrid, Part-time, and Contract listings."
    },
    {
        "us_id": "US-16",
        "epic": "EP-05: Longitudinal Trends & Employer Intelligence",
        "persona": "Placement Officer (SH-06)",
        "statement": "As a Placement Officer, I want a ranked company hiring leaderboard within a selected domain so that I can prioritize institutional campus drive invitations.",
        "feature_id": "FR-12",
        "br_id": "BR-04",
        "oq_id": "OQ-10",
        "priority": "Must",
        "story_points": 3,
        "given": "a domain is selected on the '08_Companies' page,",
        "when": "the company leaderboard generates,",
        "then": "the top 10 recruiting firms appear in a ranked horizontal bar chart and an interactive sortable table."
    }
]

# ── 2. CREATE WORD DELIVERABLE ───────────────────────────────────────────────

doc = Document()
for s in doc.sections:
    s.top_margin = Cm(2.0); s.bottom_margin = Cm(2.0)
    s.left_margin = Cm(2.5); s.right_margin = Cm(2.5)

DARK  = RGBColor(0x1F,0x35,0x64); MID   = RGBColor(0x2E,0x74,0xB5)
STEEL = RGBColor(0x4A,0x86,0xC8); WHITE = RGBColor(0xFF,0xFF,0xFF)
BLACK = RGBColor(0x00,0x00,0x00); GREEN = RGBColor(0x1F,0x5C,0x2E)

def set_cell_bg(cell, hex_c):
    tc = cell._tc; pr = tc.get_or_add_tcPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'), 'clear'); s.set(qn('w:color'), 'auto')
    s.set(qn('w:fill'), hex_c); pr.append(s)

def set_tbl_borders(table, color='2E74B5'):
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc; pr = tc.get_or_add_tcPr()
            tb = OxmlElement('w:tcBorders')
            for e in ('top','left','bottom','right','insideH','insideV'):
                x = OxmlElement(f'w:{e}'); x.set(qn('w:val'), 'single')
                x.set(qn('w:sz'), '4'); x.set(qn('w:space'), '0')
                x.set(qn('w:color'), color); tb.append(x)
            pr.append(tb)

def add_heading(text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level==1 else 10)
    p.paragraph_format.space_after  = Pt(4)
    r = p.add_run(text); r.bold = True
    r.font.size = Pt(14 if level==1 else (12 if level==2 else 11))
    r.font.color.rgb = DARK if level==1 else (MID if level==2 else STEEL)

def add_para(text, bold=False, italic=False, size=10, color=BLACK):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text); r.bold = bold; r.italic = italic
    r.font.size = Pt(size); r.font.color.rgb = color

# Cover Banner
banner = doc.add_table(rows=1, cols=1); banner.alignment = WD_TABLE_ALIGNMENT.CENTER
bc = banner.cell(0,0); set_cell_bg(bc, '1F3564')
bp = bc.paragraphs[0]; bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_before = Pt(12); bp.paragraph_format.space_after = Pt(12)
r1 = bp.add_run('P01 – HIRING TREND ANALYSIS\n'); r1.bold = True; r1.font.size = Pt(18); r1.font.color.rgb = WHITE
r2 = bp.add_run('Step 11 · Agile User Stories & Sprint Plan Specification\n'); r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
r3 = bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA'); r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

add_heading('1. User Stories Overview & Methodology', 1)
add_para('This document translates business requirements into 16 granular, Agile User Stories structured across 5 functional Epics. Each user story is scoped to be buildable and testable within a half-day increment (1 to 5 story points), adhering to the standard template:', size=10)
add_para('"As a <persona>, I want <goal> so that <benefit>."', bold=True, italic=True, size=10.5, color=MID)
add_para('Acceptance criteria follow the industry-standard Given-When-Then BDD (Behavior-Driven Development) structure, ensuring unambiguous verification during unit and end-to-end testing in Steps 13 and 15.', size=10)
doc.add_paragraph()

# User Stories Master Table
add_heading('2. Master User Stories Table', 1)
us_tbl = doc.add_table(rows=len(USER_STORIES)+1, cols=8)
us_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

headers = ['US ID', 'Epic', 'Persona', 'User Story Statement', 'Feature ID', 'BR ID', 'MoSCoW', 'Pts']
widths = [Inches(0.6), Inches(1.1), Inches(0.9), Inches(2.2), Inches(0.7), Inches(0.6), Inches(0.6), Inches(0.4)]

for ci, h in enumerate(headers):
    c = us_tbl.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(h); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(8.5)

for ri, u in enumerate(USER_STORIES, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    vals = [u['us_id'], u['epic'].split(':')[0], u['persona'].split('(')[0].strip(), u['statement'], u['feature_id'], u['br_id'], u['priority'], str(u['story_points'])]
    for ci, val in enumerate(vals):
        c = us_tbl.cell(ri, ci); set_cell_bg(c, fill_c)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci in (0,1,2,4,5,6,7) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(val); r.font.size = Pt(8)
        if ci in (0, 6): r.bold = True

set_tbl_borders(us_tbl)
doc.add_paragraph()

# User Story Detail Cards (Given / When / Then)
add_heading('3. Acceptance Criteria (Given-When-Then Specifications)', 1)
for u in USER_STORIES:
    card = doc.add_table(rows=4, cols=2); card.alignment = WD_TABLE_ALIGNMENT.CENTER
    card.cell(0, 0).merge(card.cell(0, 1))
    set_cell_bg(card.cell(0, 0), '1F3564')
    hp = card.cell(0, 0).paragraphs[0]
    hr = hp.add_run(f"  {u['us_id']}: {u['statement'][:85]}…  [{u['priority']} | {u['story_points']} Pts]")
    hr.bold = True; hr.font.color.rgb = WHITE; hr.font.size = Pt(9.5)
    
    details = [
        ("User Story Narrative", u['statement']),
        ("Traceability Links", f"Epic: {u['epic']} | Features: {u['feature_id']} | Requirements: {u['br_id']} | OQ: {u['oq_id']}"),
        ("Acceptance Criteria (BDD)", f"GIVEN {u['given']}\nWHEN {u['when']}\nTHEN {u['then']}")
    ]
    fills = ['EEF4FB', 'FFFFFF', 'E5FFE5']
    for ri, (lbl, val) in enumerate(details, 1):
        set_cell_bg(card.cell(ri, 0), fills[ri-1]); set_cell_bg(card.cell(ri, 1), fills[ri-1])
        card.cell(ri, 0).width = Inches(1.8); card.cell(ri, 1).width = Inches(4.7)
        lp = card.cell(ri, 0).paragraphs[0]; lr = lp.add_run(lbl); lr.bold = True; lr.font.size = Pt(8.5)
        vp = card.cell(ri, 1).paragraphs[0]; vr = vp.add_run(val); vr.font.size = Pt(8.5)
        if lbl.startswith("Acceptance"): vr.font.color.rgb = GREEN; vr.bold = True
        
    set_tbl_borders(card)
    doc.add_paragraph()

# 4. Sprint Schedule
add_heading('4. 5-Day Sprint Execution Plan', 1)
add_para('The 16 user stories total 47 story points, structured across a 5-day delivery sprint aligned with Step 15 build milestones:', size=9.5)

sprint_plan = [
    ("Day 1: Foundations", "US-01, US-02", 5, "Database connector, validator, caching layer, and baseline EDA panel."),
    ("Day 2: Global UI & Demand", "US-03, US-04, US-05", 11, "Overview KPI cards, sidebar filter state synchronization, and role demand bar charts."),
    ("Day 3: Skills & Compensation", "US-06, US-07, US-08, US-09, US-10", 13, "Skill frequency extraction, experience breakdown per role, and salary benchmarking views."),
    ("Day 4: Analytics & Hotspots", "US-11, US-12, US-13, US-14", 13, "Skill-salary OLS correlation model, city geographic distribution, and monthly time-series charts."),
    ("Day 5: Synthesis & Hardening", "US-15, US-16", 5, "Job-type distribution, top hiring company leaderboard, and cross-browser smoke verification.")
]

sp_tbl = doc.add_table(rows=len(sprint_plan)+1, cols=4)
sp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Sprint Day', 'Assigned Stories', 'Total Points', 'Milestone Objective']):
    c = sp_tbl.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)

for ri, (day, stories, pts, obj) in enumerate(sprint_plan, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    for ci, val in enumerate([day, stories, str(pts), obj]):
        c = sp_tbl.cell(ri, ci); set_cell_bg(c, fill_c)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci in (0,1,2) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(val); r.font.size = Pt(9)
        if ci in (0, 2): r.bold = True

set_tbl_borders(sp_tbl)
doc.add_paragraph()

# Document Control
add_heading('Document Control', 2)
dc = doc.add_table(rows=2, cols=4); dc.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Version', 'Date', 'Author', 'Summary']):
    c = dc.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)
for ci, val in enumerate(['1.0', datetime.date.today().strftime('%d %b %Y'), 'Abhayjeet Singh', 'Initial User Stories & Sprint Plan deliverable']):
    c = dc.cell(1, ci); set_cell_bg(c, 'EEF4FB')
    p = c.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)
set_tbl_borders(dc)

doc.save(OUT_DOCX)
print(f"SAVED USER STORIES DOCX: {OUT_DOCX}")

# ── 3. UPDATE RTM EXCEL FILE (COLUMN K: USER STORY ID) ───────────────────────

wb = openpyxl.load_workbook(RTM_FILE)
ws = wb["RTM"]

# Mapping from BR_ID to User Story IDs
br_to_us_map = {
    "BR-01": "US-03, US-04, US-05",
    "BR-02": "US-04, US-06",
    "BR-03": "US-04, US-09",
    "BR-04": "US-04, US-16",
    "BR-05": "US-03, US-04, US-13",
    "BR-06": "US-03, US-04, US-10",
    "BR-07": "US-03, US-04, US-12",
    "BR-08": "US-04, US-07",
    "BR-09": "US-04, US-11",
    "BR-10": "US-04, US-15",
    "BR-11": "US-04, US-08",
    "BR-12": "US-04, US-14",
    "BR-13": "US-01, US-02"
}

# Update Column K (index 11) for BR rows
for row_idx in range(4, 17):
    br_id = ws.cell(row=row_idx, column=3).value
    if br_id in br_to_us_map:
        c = ws.cell(row=row_idx, column=11)
        c.value = br_to_us_map[br_id]
        c.fill = PatternFill('solid', fgColor='E2EFDA')
        c.font = Font(name='Calibri', size=9, bold=True, color='1F5C2E')
        c.alignment = Alignment(horizontal='center', vertical='center')

# Update Status in Column H (index 8) to 'Stories Mapped'
for row_idx in range(4, 17):
    c = ws.cell(row=row_idx, column=8)
    c.value = 'Stories Mapped'
    c.fill = PatternFill('solid', fgColor='E5FFE5')
    c.font = Font(name='Calibri', size=9, bold=True, color='1F5C2E')
    c.alignment = Alignment(horizontal='center', vertical='center')

# Update Changelog
ws_ch = wb["Changelog"]
for r in range(3, 10):
    if ws_ch.cell(r, 1).value and "Step 11" in str(ws_ch.cell(r, 1).value):
        ws_ch.cell(r, 2).value = datetime.date.today().strftime('%d %b %Y')
        ws_ch.cell(r, 5).value = "Populated Column K with User Story IDs (US-01..16). Status -> Stories Mapped."
        break

wb.save(RTM_FILE)
print(f"UPDATED RTM: {RTM_FILE}")
