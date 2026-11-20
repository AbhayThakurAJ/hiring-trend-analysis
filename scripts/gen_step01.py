"""
Step 01 – Project Understanding Document Generator
Project : P01 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES, MCA 2nd Yr DA
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os, datetime

# ── output path ──────────────────────────────────────────────────────────────
OUT_DIR  = r"C:\Hiring Trend Analysis\docs"
OUT_FILE = os.path.join(OUT_DIR, "P01_HTA_Step01_ProjectUnderstanding.docx")
os.makedirs(OUT_DIR, exist_ok=True)

# ── colour palette ────────────────────────────────────────────────────────────
DARK_BLUE  = RGBColor(0x1F, 0x35, 0x64)   # header / H1
MID_BLUE   = RGBColor(0x2E, 0x74, 0xB5)   # H2
STEEL      = RGBColor(0x4A, 0x86, 0xC8)   # H3 / table header bg
LIGHT_FILL = RGBColor(0xDA, 0xE8, 0xF5)   # alternate table row
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
BLACK      = RGBColor(0x00, 0x00, 0x00)
GREEN_OK   = RGBColor(0x37, 0x86, 0x36)

# ── helpers ───────────────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def set_cell_borders(table):
    for row in table.rows:
        for cell in row.cells:
            tc   = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for edge in ('top','left','bottom','right','insideH','insideV'):
                e = OxmlElement(f'w:{edge}')
                e.set(qn('w:val'),   'single')
                e.set(qn('w:sz'),    '4')
                e.set(qn('w:space'), '0')
                e.set(qn('w:color'), '2E74B5')
                tcBorders.append(e)
            tcPr.append(tcBorders)

def header_row(table, *texts, bg='1F3564'):
    row = table.rows[0]
    for i, (cell, text) in enumerate(zip(row.cells, texts)):
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.color.rgb = WHITE
        run.font.size = Pt(10)

def data_row(table, row_idx, *texts, alt=False):
    row = table.rows[row_idx]
    fill = 'DAE8F5' if alt else 'FFFFFF'
    for i, (cell, text) in enumerate(zip(row.cells, texts)):
        set_cell_bg(cell, fill)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(str(text))
        run.font.size = Pt(9.5)

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size  = Pt(14)
        run.font.color.rgb = DARK_BLUE
    elif level == 2:
        run.font.size  = Pt(12)
        run.font.color.rgb = MID_BLUE
    else:
        run.font.size  = Pt(11)
        run.font.color.rgb = STEEL
    return p

def add_para(doc, text, bold=False, italic=False, size=10, color=BLACK, indent=0):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent   = Cm(indent)
    p.paragraph_format.space_before  = Pt(2)
    p.paragraph_format.space_after   = Pt(2)
    run = p.add_run(text)
    run.bold   = bold
    run.italic = italic
    run.font.size      = Pt(size)
    run.font.color.rgb = color
    return p

def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent  = Cm(0.5 + level * 0.5)
    p.paragraph_format.space_after  = Pt(1)
    run = p.add_run(text)
    run.font.size = Pt(9.5)
    return p

def thin_rule(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot  = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    '6')
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), '2E74B5')
    pBdr.append(bot)
    pPr.append(pBdr)

# ═════════════════════════════════════════════════════════════════════════════
# BUILD DOCUMENT
# ═════════════════════════════════════════════════════════════════════════════
doc = Document()

# page margins
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── COVER BANNER ─────────────────────────────────────────────────────────────
banner = doc.add_table(rows=1, cols=1)
banner.alignment = WD_TABLE_ALIGNMENT.CENTER
bc = banner.cell(0,0)
set_cell_bg(bc, '1F3564')
bp = bc.paragraphs[0]
bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_before = Pt(10)
bp.paragraph_format.space_after  = Pt(10)
r1 = bp.add_run('P01 – HIRING TREND ANALYSIS\n')
r1.bold = True; r1.font.size = Pt(18); r1.font.color.rgb = WHITE
r2 = bp.add_run('Step 01 · Project Understanding Document\n')
r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
r3 = bp.add_run(f'Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA')
r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

meta = doc.add_table(rows=3, cols=4)
meta.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project ID','P01'),('Version','1.0'),
    ('Date', datetime.date.today().strftime('%d %b %Y')),('Status','Draft'),
    ('Domain','Data Analytics'),('Platform','Python + Streamlit + SQLite'),
    ('Hosting','Streamlit Community Cloud'),('Step','01 of 18'),
    ('Prepared by','Abhayjeet Singh'),('Roll No','2500320140004'),
    ('Institution','ABES Engineering College'),('Mentor','Dinesh Sereno'),
]
for i, row in enumerate(meta.rows):
    fill = 'EEF4FB' if i % 2 == 0 else 'FFFFFF'
    for j, cell in enumerate(row.cells):
        set_cell_bg(cell, fill)
        p   = cell.paragraphs[0]
        txt = meta_data[i*4 + j]
        run = p.add_run(f'{txt[0]}: ')
        run.bold = True; run.font.size = Pt(9)
        p.add_run(txt[1]).font.size = Pt(9)
set_cell_borders(meta)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — PROJECT OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '1. Project Overview', 1)
add_para(doc,
    'The Hiring Trend Analysis application is an end-to-end Python + Streamlit data analytics '
    'product that ingests structured recruitment data stored in a SQLite database and presents '
    'interactive exploratory data analysis (EDA), statistical insights, and Plotly visualisations '
    'to help HR managers, recruiters, placement officers, and job seekers understand evolving '
    'patterns in the job market.',
    size=10)

# ── 1.1 Problem Statement ────────────────────────────────────────────────────
add_heading(doc, '1.1  Problem Statement', 2)
ps_tbl = doc.add_table(rows=1, cols=1)
set_cell_bg(ps_tbl.cell(0,0), 'EEF4FB')
set_cell_borders(ps_tbl)
p = ps_tbl.cell(0,0).paragraphs[0]
p.paragraph_format.left_indent = Cm(0.3)
run = p.add_run(
    'HR professionals and job seekers currently lack a single, consolidated, and visually '
    'interactive tool to analyse historical hiring trends, benchmark salaries, identify '
    'in-demand skills, and track regional recruitment activity — forcing them to rely on '
    'fragmented spreadsheets and manual reports that are time-consuming, error-prone, and '
    'incapable of surfacing actionable insights quickly.'
)
run.font.size = Pt(10); run.italic = True
doc.add_paragraph()

# ── 1.2 Intended Audience ────────────────────────────────────────────────────
add_heading(doc, '1.2  Intended Audience & Usage Frequency', 2)
aud_tbl = doc.add_table(rows=6, cols=4)
aud_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(aud_tbl, 'Audience / User', 'Role in Context', 'Usage Frequency', 'Primary Need')
aud_data = [
    ('HR Manager / TA Lead',       'Primary decision-maker for hiring campaigns',    'Weekly',           'Identify demand spikes, role gaps'),
    ('Recruiter',                   'Operational hiring agent posting JDs',           'Daily',            'Skill benchmarks, top locations'),
    ('Job Seeker / Graduate',       'Candidate researching market trends',            'Occasional',       'In-demand skills, salary ranges'),
    ('College Placement Officer',   'Facilitates campus placements at ABES',         'Semester-wise',    'Industry demand, top recruiters'),
    ('Data Analyst / Researcher',   'Analyses workforce and labour market data',     'Monthly',          'Trend charts, correlation insights'),
]
for idx, row in enumerate(aud_data):
    data_row(aud_tbl, idx+1, *row, alt=(idx%2==0))
set_cell_borders(aud_tbl)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — STAKEHOLDERS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '2. Stakeholder Register', 1)
sh_tbl = doc.add_table(rows=7, cols=6)
sh_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(sh_tbl, 'SH ID', 'Stakeholder', 'Role', 'Interest', 'Influence', 'Key Decision via App')
sh_data = [
    ('SH-01', 'Dinesh Sereno',         'Academic Mentor',           'Project quality & methodology',           'H', 'Approve deliverables at each step'),
    ('SH-02', 'Abhayjeet Singh',       'Student Developer',         'Learning, grades & portfolio value',      'H', 'Build, test & deploy the app'),
    ('SH-03', 'HR Manager (Persona)',  'Primary App User',          'Faster, data-driven hiring decisions',    'H', 'Decide which roles/regions to target'),
    ('SH-04', 'Recruiter (Persona)',   'Operational App User',      'Skill & salary benchmarking',             'M', 'Set JD requirements & salary bands'),
    ('SH-05', 'Job Seeker (Persona)',  'End Consumer',              'Find in-demand skills & salary ranges',   'L', 'Choose upskilling path & job location'),
    ('SH-06', 'Placement Officer',     'Institutional Stakeholder', 'Campus placement success rate',           'M', 'Target companies & domains for drives'),
]
for idx, row in enumerate(sh_data):
    data_row(sh_tbl, idx+1, *row, alt=(idx%2==0))
set_cell_borders(sh_tbl)
add_para(doc, 'Influence scale — H: High · M: Medium · L: Low', italic=True, size=8.5)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — OPERATIONAL QUESTIONS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '3. Operational Questions (OQ)', 1)
add_para(doc,
    'The following 10 operational questions define EXACTLY what the app must answer. '
    'Each OQ is traceable to dataset fields (see Check column) and will map to at least '
    'one Functional Requirement in Step 5 (BRD).',
    size=9.5)
doc.add_paragraph()

oq_tbl = doc.add_table(rows=11, cols=5)
oq_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(oq_tbl, 'OQ ID', 'Operational Question', 'Business Value', 'Dataset Fields Required', 'Answerable?')
oq_data = [
    ('OQ-01','Which job roles have the highest posting volume in a selected time period?',
     'Identify high-demand positions for resource planning',
     'job_title, posted_date',         '✔ Yes'),
    ('OQ-02','What is the average, minimum, and maximum salary offered per industry/domain?',
     'Salary benchmarking for JDs and negotiations',
     'industry, salary_min, salary_max','✔ Yes'),
    ('OQ-03','Which skills are most frequently listed across all job postings?',
     'Guide upskilling programs and JD creation',
     'required_skills',                '✔ Yes'),
    ('OQ-04','How has monthly/quarterly hiring volume changed over time?',
     'Spot seasonal demand cycles for workforce planning',
     'posted_date, job_id',            '✔ Yes'),
    ('OQ-05','Which cities and states show the highest job posting concentration?',
     'Prioritise relocation and remote-work policies',
     'city, state',                    '✔ Yes'),
    ('OQ-06','What is the split of job types (Full-time / Part-time / Remote / Hybrid / Contract)?',
     'Understand flexible-work adoption trends',
     'job_type',                       '✔ Yes'),
    ('OQ-07','Which industries are the top hirers in a given domain/skill set?',
     'Focus talent pipelines on high-activity sectors',
     'industry, job_title',            '✔ Yes'),
    ('OQ-08','What experience level (Fresher / Mid / Senior) is most demanded overall and per role?',
     'Align recruitment campaigns with available talent pools',
     'experience_level',               '✔ Yes'),
    ('OQ-09','Is there a measurable correlation between required skills count and offered salary?',
     'Price complex roles competitively',
     'required_skills, salary_max',    '✔ Yes'),
    ('OQ-10','Which companies are the top recruiters in a selected domain or city?',
     'Build target employer lists for placement drives',
     'company_name, industry, city',   '✔ Yes'),
]
for idx, row in enumerate(oq_data):
    data_row(oq_tbl, idx+1, *row, alt=(idx%2==0))
set_cell_borders(oq_tbl)
add_para(doc, '✔ All 10 OQs are fully answerable from the planned synthetic dataset fields.', bold=True, color=GREEN_OK)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — SMART OBJECTIVES
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '4. Project Objectives (SMART)', 1)
obj_tbl = doc.add_table(rows=6, cols=7)
obj_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(obj_tbl, 'Obj ID', 'Objective', 'Specific', 'Measurable', 'Achievable', 'Relevant', 'Time-bound')
obj_data = [
    ('O-01',
     'Build a Streamlit app that answers all 10 OQs with interactive Plotly charts',
     'All 10 OQ screens present',
     '10/10 OQs functional',
     'Solo dev, 8 weeks',
     'Core deliverable',
     'By Step 15 completion'),
    ('O-02',
     'Generate & load ≥500 synthetic hiring records into SQLite',
     'Script + DB in /data',
     '≥500 rows verified',
     'Python Faker library',
     'Enables all analysis',
     'By Step 10'),
    ('O-03',
     'Achieve ≥90% pytest pass rate across all functional test cases',
     'Test file covers all FRs',
     'Pass rate ≥90%',
     '20–30 test cases',
     'Ensures reliability',
     'By Step 15'),
    ('O-04',
     'Deploy app on Streamlit Community Cloud with <3 s initial load',
     'Live URL accessible',
     'Load time <3 s',
     'Free-tier sufficient',
     'Portfolio ready',
     'By Step 16'),
    ('O-05',
     'Submit ABES Word project document and slide deck on time',
     '17 + 18 file names locked',
     'Both files submitted',
     'Templates available',
     'Academic requirement',
     'By Step 18'),
]
for idx, row in enumerate(obj_data):
    data_row(obj_tbl, idx+1, *row, alt=(idx%2==0))
set_cell_borders(obj_tbl)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — SCOPE
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '5. Project Scope', 1)

scope_tbl = doc.add_table(rows=1, cols=2)
scope_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_cell_bg(scope_tbl.cell(0,0), '1F3564')
set_cell_bg(scope_tbl.cell(0,1), '7B2D2D')
for label, col in [('✅  IN SCOPE',0),('❌  OUT OF SCOPE',1)]:
    p = scope_tbl.cell(0,col).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(label); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(11)
set_cell_borders(scope_tbl)

in_scope = [
    'EDA dashboard — distribution, missing value analysis, outlier detection',
    'Salary analysis — avg/min/max per industry, role, location',
    'Skills demand — top-N skill frequency, word-cloud equivalent',
    'Hiring trend over time — monthly/quarterly bar & line charts',
    'Geographic distribution — city/state heatmap / bar chart',
    'Job type breakdown — pie/donut chart (Full-time, Remote, Hybrid, etc.)',
    'Industry-wise hiring volume — horizontal bar chart',
    'Experience level demand — stacked bar per role',
    'Skill-salary correlation — scatter plot with trendline',
    'Top recruiters leaderboard — filterable table',
    'SQLite database with ≥500 synthetic records',
    'Pytest test suite (≥90% pass rate)',
    'Streamlit Community Cloud deployment',
    'ABES project document (Word) + slide deck (PPT)',
]
out_scope = [
    'Real-time web scraping from LinkedIn / Naukri / Indeed',
    'User authentication / login system',
    'Resume upload or parsing features',
    'Payment gateway or premium tiers',
    'Native mobile application (Android / iOS)',
    'Email or notification system',
    'Multi-language (non-English) data',
    'Machine-learning predictions or forecasting',
    'Admin panel or CMS',
    'Integration with HRMS or ATS systems',
]

in_tbl = doc.add_table(rows=len(in_scope)+1, cols=2)
in_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(in_tbl, '#', 'In-Scope Item', bg='1F5C2E')
for i, item in enumerate(in_scope):
    data_row(in_tbl, i+1, str(i+1), item, alt=(i%2==0))
set_cell_borders(in_tbl)
doc.add_paragraph()

out_tbl = doc.add_table(rows=len(out_scope)+1, cols=2)
out_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(out_tbl, '#', 'Out-of-Scope Item', bg='7B2D2D')
for i, item in enumerate(out_scope):
    data_row(out_tbl, i+1, str(i+1), item, alt=(i%2==0))
set_cell_borders(out_tbl)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — ASSUMPTIONS & CONSTRAINTS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '6. Assumptions & Constraints', 1)

add_heading(doc, '6.1  Assumptions', 2)
assumptions = [
    ('A-01', 'Synthetic data generated with Python Faker realistically mirrors Indian job-market distributions for roles, cities, and salaries.'),
    ('A-02', 'SQLite is sufficient as the data store for a college-scale project (≤10 000 rows).'),
    ('A-03', 'Streamlit Community Cloud free tier provides adequate uptime and performance for academic demonstration.'),
    ('A-04', 'All charts and EDA outputs are based on a single snapshot dataset; no live data feed is assumed.'),
    ('A-05', 'The app will be used on desktop/laptop browsers; mobile responsiveness is a nice-to-have, not a must-have.'),
    ('A-06', 'English is the only language for all labels, field values, and documentation.'),
    ('A-07', 'Mentor Dinesh Sereno will review and provide feedback at each major milestone step.'),
]
a_tbl = doc.add_table(rows=len(assumptions)+1, cols=3)
a_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(a_tbl, 'ID', 'Assumption', 'Impact if Wrong')
impacts = [
    'Data insights may not match real-world patterns',
    'Performance issues if dataset grows beyond project scope',
    'App may be unavailable during mentor demo — use local fallback',
    'Dashboard will show stale data — clearly documented',
    'Some charts may render poorly on small screens',
    'Non-English records will cause encoding errors',
    'Deliverable quality may not meet academic expectations',
]
for i, (aid, atext) in enumerate(assumptions):
    data_row(a_tbl, i+1, aid, atext, impacts[i], alt=(i%2==0))
set_cell_borders(a_tbl)
doc.add_paragraph()

add_heading(doc, '6.2  Constraints', 2)
constraints = [
    ('C-01', 'Budget',    'Zero — all tools, libraries, and hosting must be free-tier / open-source.'),
    ('C-02', 'Time',      'Academic semester timeline; target completion within 8 weeks from Step 1.'),
    ('C-03', 'Team',      'Solo developer — Abhayjeet Singh; no additional engineering resources.'),
    ('C-04', 'Data',      'Synthetic data only — no real personal or company data may be used (privacy compliance).'),
    ('C-05', 'Tech Stack','Python 3.11+, Streamlit, Plotly, SQLite, pytest — no paid frameworks.'),
    ('C-06', 'Hosting',   'Streamlit Community Cloud (free tier); no paid cloud (AWS/GCP/Azure) allowed.'),
    ('C-07', 'Doc Format','Final document must follow ABES Word template; slides must follow ABES PPT template.'),
]
c_tbl = doc.add_table(rows=len(constraints)+1, cols=4)
c_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(c_tbl, 'ID', 'Category', 'Constraint', 'Mitigation')
mitigations = [
    'Use Faker, Streamlit free, SQLite — all zero-cost',
    'Weekly milestone checkpoints with mentor',
    'Modular code design; reuse templates across steps',
    'Document data-generation script for reproducibility',
    'Well-documented, widely-supported open-source stack',
    'Maintain a local run script as backup demo option',
    'Use provided ABES templates from Step 17 onward',
]
for i, (cid, cat, ctext) in enumerate(constraints):
    data_row(c_tbl, i+1, cid, cat, ctext, mitigations[i], alt=(i%2==0))
set_cell_borders(c_tbl)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — SUCCESS CRITERIA
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '7. Success Criteria (Measurable)', 1)
sc_tbl = doc.add_table(rows=9, cols=4)
sc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(sc_tbl, 'SC ID', 'Success Criterion', 'Measurement Method', 'Target')
sc_data = [
    ('SC-01','All 10 OQs answered within the app',                    'Manual walkthrough against OQ list',     '10 / 10 screens present & functional'),
    ('SC-02','Dataset contains ≥500 synthetic hiring records',         'Row count query on SQLite DB',           '≥ 500 rows'),
    ('SC-03','App initial page load time < 3 seconds',                 'Browser DevTools Network tab',           '< 3 s on broadband'),
    ('SC-04','pytest suite passes ≥90% of test cases',                 'pytest -v output pass/fail ratio',       '≥ 90%'),
    ('SC-05','Zero critical bugs in deployed app at submission',        'Manual exploratory testing checklist',   '0 P1/P2 defects open'),
    ('SC-06','All 18 step deliverables produced and named correctly',  'File checklist reviewed by mentor',      '18 / 18 files present'),
    ('SC-07','ABES Word project document complete (all sections)',      'Document section checklist',             '100% sections filled'),
    ('SC-08','Presentation slide deck covers all major steps',         'Slide count & section coverage',         '12–18 slides, all major steps'),
]
for idx, row in enumerate(sc_data):
    data_row(sc_tbl, idx+1, *row, alt=(idx%2==0))
set_cell_borders(sc_tbl)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — OQ DATA-ANSWERABILITY CHECK
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '8. OQ ↔ Dataset Answerability Check', 1)
add_para(doc,
    'Each OQ has been cross-verified against the planned dataset schema. '
    'All fields listed below will be present in the synthetic dataset generated at Step 10. '
    'No OQ is left unanswerable.',
    size=9.5)
doc.add_paragraph()

chk_tbl = doc.add_table(rows=11, cols=4)
chk_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(chk_tbl, 'OQ ID', 'Key Fields Required', 'Available in Schema?', 'Chart / Analysis Type')
chk_data = [
    ('OQ-01','job_title, posted_date',               '✔ Yes','Bar chart – Top 10 roles by count'),
    ('OQ-02','industry, salary_min, salary_max',     '✔ Yes','Box plot / grouped bar – salary stats'),
    ('OQ-03','required_skills (pipe-delimited)',      '✔ Yes','Horizontal bar – top-N skill frequency'),
    ('OQ-04','posted_date → month/quarter',          '✔ Yes','Line chart – volume over time'),
    ('OQ-05','city, state',                          '✔ Yes','Choropleth / bar – geographic spread'),
    ('OQ-06','job_type',                             '✔ Yes','Pie / donut – work-mode distribution'),
    ('OQ-07','industry, job_title',                  '✔ Yes','Stacked bar – industry hiring share'),
    ('OQ-08','experience_level',                     '✔ Yes','Grouped bar – experience demand'),
    ('OQ-09','required_skills count, salary_max',    '✔ Yes','Scatter plot + trendline (OLS)'),
    ('OQ-10','company_name, industry, city',         '✔ Yes','Leaderboard table + bar chart'),
]
for idx, row in enumerate(chk_data):
    data_row(chk_tbl, idx+1, *row, alt=(idx%2==0))
set_cell_borders(chk_tbl)
add_para(doc, '✔ All 10 Operational Questions are confirmed answerable from the planned dataset.', bold=True, color=GREEN_OK, size=10)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# FOOTER NOTE
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, 'Document Control', 2)
dc_tbl = doc.add_table(rows=4, cols=4)
dc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(dc_tbl, 'Version', 'Date', 'Author', 'Change Summary')
dc_data = [
    ('1.0', datetime.date.today().strftime('%d %b %Y'), 'Abhayjeet Singh', 'Initial draft – Step 01 complete'),
]
data_row(dc_tbl, 1, *dc_data[0])
set_cell_borders(dc_tbl)

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('P01 – Hiring Trend Analysis  |  Step 01 Project Understanding  |  ABES MCA 2nd Yr DA  |  Mentor: Dinesh Sereno')
run.font.size = Pt(8); run.font.color.rgb = RGBColor(0x88,0x88,0x88); run.italic = True

# ── SAVE ─────────────────────────────────────────────────────────────────────
doc.save(OUT_FILE)
print(f"SAVED: {OUT_FILE}")
