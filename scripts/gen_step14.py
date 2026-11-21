"""
Step 14 – Build Plan Word Document Generator
Project : P071 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES MCA 2nd Yr DA
"""

import os, datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOCS_DIR = r"C:\Hiring Trend Analysis\docs"
OUT_DOCX = os.path.join(DOCS_DIR, "P01_HTA_Step14_BuildPlan.docx")

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

def code_box(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text); r.font.name = 'Courier New'; r.font.size = Pt(8.5)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), 'F0F4FF')
    pPr.append(shd)

# Banner
banner = doc.add_table(rows=1, cols=1); banner.alignment = WD_TABLE_ALIGNMENT.CENTER
bc = banner.cell(0,0); set_cell_bg(bc, '1F3564')
bp = bc.paragraphs[0]; bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_before = Pt(12); bp.paragraph_format.space_after = Pt(12)
r1 = bp.add_run('P071 – HIRING TREND ANALYSIS\n'); r1.bold = True; r1.font.size = Pt(18); r1.font.color.rgb = WHITE
r2 = bp.add_run('Step 14 · Comprehensive Implementation & Build Plan Specification\n'); r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
r3 = bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA'); r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

add_heading('1. Executive Build Summary & Deliverables Traceability', 1)
summary_text = (
    "The Hiring Trend Analysis (P071) platform is an end-to-end data analytics application constructed on a "
    "relational SQLite database of 652 Indian recruitment postings across 15 cities and 8 industry sectors. "
    "Designed to eliminate fragmented spreadsheets and subjective hiring decisions, the product provides real-time "
    "insights into role demand, salary benchmarks, skill frequencies, geographic concentration, and longitudinal "
    "seasonality trends. Structured as a modular Streamlit multipage application styled with Tailwind CSS, the platform "
    "resolves all 10 defined Operational Questions (OQ-01..10) with complete bidirectional traceability across all SDLC artifacts."
)
add_para(summary_text, size=10)
doc.add_paragraph()

# 2. Build Tasks Table
add_heading('2. Sequenced Build Tasks (BUILD-01 to BUILD-10)', 1)
add_para('Step 15 build execution is broken down into 10 structured, atomic tasks:', size=9.5)

build_tasks = [
    ("BUILD-01", "Environment & Dependencies Setup", "requirements.txt, runtime.txt, config.toml", "EP-01 / US-01", "Freeze pinned dependencies and custom theme settings", "TC-T-13"),
    ("BUILD-02", "Data Ingestion, Cleaning & Validation", "src/data_loader.py, src/cleaner.py, src/validator.py", "EP-01 / US-01, US-02", "Implement SQLite loader, imputation logic, and schema validation", "TC-T-01..06"),
    ("BUILD-03", "Analytical Calculation Engine", "src/analyser.py", "EP-03..05 / US-05..16", "Build 13 pure calculation functions for all OQs without UI dependencies", "TC-T-07..12"),
    ("BUILD-04", "Reusable Tailwind UI Components", "src/ui.py", "EP-02 / US-03", "Provide hero banners, KPI cards, and insight callouts via Tailwind CDN", "TC-T-14"),
    ("BUILD-05", "Streamlit Entrypoint & Global Filters", "app.py", "EP-02 / US-04", "Construct app entrypoint, session_state, and sidebar filter synchronization", "TC-F-14, 15"),
    ("BUILD-06", "Multipage Implementation (Part 1)", "pages/01_Overview.py to 05_Geography.py", "EP-02..04 / US-03, 05, 06, 09, 12", "Implement first 5 core pages for overview, demand, skills, salary, geography", "TC-F-01..03, 06, 07"),
    ("BUILD-07", "Multipage Implementation (Part 2)", "pages/06_Trends.py to 10_Job_Types.py", "EP-01..05 / US-07, 08, 11, 13..16", "Implement remaining 5 pages for trends, correlation, companies, EDA, job types", "TC-F-04, 05, 08..13"),
    ("BUILD-08", "Data Export & Error Hardening", "All pages/*.py", "FR-16 / US-16", "Add CSV download buttons and empty-state warning notifications", "TC-F-15, 16"),
    ("BUILD-09", "Automated Pytest Suite Verification", "tests/test_*.py", "All Epics / SC-04", "Execute 15 unit tests to achieve 100% pass rate in <10s", "All TC-T-01..15"),
    ("BUILD-10", "Documentation & Deployment Prep", "README.md, .gitignore", "DR-01..08", "Author repository documentation and configure for Streamlit Cloud", "TC-T-15")
]

bt_tbl = doc.add_table(rows=len(build_tasks)+1, cols=6)
bt_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Task ID', 'Task Objective', 'Impacted Files', 'Linked Stories', 'Execution Focus', 'Passing Test']):
    c = bt_tbl.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(8.5)

for ri, (bid, obj, files, stories, focus, test) in enumerate(build_tasks, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    for ci, val in enumerate([bid, obj, files, stories, focus, test]):
        c = bt_tbl.cell(ri, ci); set_cell_bg(c, fill_c)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci in (0, 3, 5) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(val); r.font.size = Pt(8)
        if ci in (0, 5): r.bold = True

set_tbl_borders(bt_tbl)
doc.add_paragraph()

# 3. Definition of Done
add_heading('3. Definition of Done (DoD)', 1)
dod_items = [
    ("Functional Completeness", "All 13 Business Requirements (BR-01..13) and 10 OQs implemented with zero unhandled exceptions."),
    ("Automated Testing Pass", "15/15 technical unit tests pass cleanly in pytest within 10 seconds execution latency."),
    ("Performance Benchmark", "Full analytical dataset query evaluates in under 1.0 second; initial page load < 3.0 seconds."),
    ("Bidirectional Traceability", "100% of RTM matrix columns populated (BR -> Features -> DB -> Stories -> Screens -> Tests)."),
    ("Clean Modular Architecture", "Strict separation of concerns: zero Streamlit imports inside src/analyser.py calculation engine."),
    ("Deployment Verification", "Live deployment functioning on Streamlit Community Cloud with public GitHub repository documentation.")
]
for title, desc in dod_items:
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(f"{title}: "); r1.bold = True; r1.font.size = Pt(9.5)
    r2 = p.add_run(desc); r2.font.size = Pt(9.5)

doc.add_paragraph()

# 4. Standard Commands
add_heading('4. Standard CLI Commands', 1)
code_box("""# 1. Install Dependencies
pip install -r requirements.txt

# 2. Run Synthetic Dataset Generator
python src/generate_data.py

# 3. Execute Pytest Test Suite
python -m pytest tests/ -v

# 4. Launch Streamlit Application Locally
streamlit run app.py""")

doc.add_paragraph()

# 5. Top 10 Viva Questions
add_heading('5. Top 10 Viva Questions & Model Answers', 1)
viva_qa = [
    ("Q1: What business problem does your project solve?", "Consolidates fragmented recruitment listings and salary guesswork into a real-time analytics platform for data-backed talent decisions."),
    ("Q2: Why choose SQLite over MySQL or PostgreSQL?", "SQLite is serverless, zero-configuration, and fully self-contained in data/app.db, making it ideal for academic demos and Streamlit Community Cloud."),
    ("Q3: How is data normalized in your database?", "Follows a Star Schema with 3 dimension tables (dim_industry, dim_city, dim_job_role), 1 fact table (fact_job_postings), and 1 bridge table (job_skills) in 1NF."),
    ("Q4: What is the purpose of vw_hiring_analysis?", "A pre-joined analytical view combining dimensions, facts, and derived attributes (e.g. salary_mid, quarter) to enable cached single-query dashboard reads."),
    ("Q5: How do you prevent Streamlit from reloading data repeatedly?", "By decorating load_data() with @st.cache_data(ttl=3600), storing the DataFrame in memory for sub-second UI filter response."),
    ("Q6: Why deliberately inject missing values and outliers?", "Real recruitment data is messy; 12 nulls and 3 outliers permit demonstrating robust EDA, imputation, and IQR outlier detection on Page 09."),
    ("Q7: What statistical method was used for OQ-09 (Skills vs Salary)?", "Ordinary Least Squares (OLS) linear regression combined with Pearson correlation (r) and p-value evaluation."),
    ("Q8: How did you implement Tailwind CSS without Node/npm?", "Injected the Tailwind CSS CDN stylesheet directly inside isolated iframes using st.components.v1.html in src/ui.py."),
    ("Q9: What is the difference between TC-F and TC-T tests?", "TC-F are functional user-facing test cases verifying UI behavior; TC-T are automated pytest unit tests verifying schema, math invariants, and latency."),
    ("Q10: How does the RTM safeguard project quality?", "Guarantees complete bidirectional traceability from business needs down to code and test cases, ensuring zero orphan requirements.")
]

for q, a in viva_qa:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(2)
    rq = p.add_run(q); rq.bold = True; rq.font.size = Pt(9.5); rq.font.color.rgb = DARK
    p2 = doc.add_paragraph()
    p2.paragraph_format.left_indent = Cm(0.5); p2.paragraph_format.space_after = Pt(4)
    ra = p2.add_run(f"➔ {a}"); ra.font.size = Pt(9); ra.italic = True

doc.add_paragraph()

# Document Control
add_heading('Document Control', 2)
dc = doc.add_table(rows=2, cols=4); dc.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Version', 'Date', 'Author', 'Summary']):
    c = dc.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)
for ci, val in enumerate(['1.0', datetime.date.today().strftime('%d %b %Y'), 'Abhayjeet Singh', 'Initial Build Plan deliverable and BUILD_PLAN.md generation']):
    c = dc.cell(1, ci); set_cell_bg(c, 'EEF4FB')
    p = c.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)
set_tbl_borders(dc)

doc.save(OUT_DOCX)
print(f"SAVED BUILD PLAN DOCX: {OUT_DOCX}")
