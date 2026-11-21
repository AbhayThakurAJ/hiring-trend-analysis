"""
Step 16 – Deployment Guide Document Generator
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
OUT_DOCX = os.path.join(DOCS_DIR, "P01_HTA_Step16_DeploymentGuide.docx")

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
r2 = bp.add_run('Step 16 · Production Deployment Guide & Cloud Publishing\n'); r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
r3 = bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA'); r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

add_heading('1. Production Readiness Audit & Pre-Deployment Checklist', 1)
add_para('Prior to initiating cloud deployment on Streamlit Community Cloud, the repository was audited against all deployment constraints:', size=9.5)

audit_checks = [
    ("Entrypoint & App Structure", "app.py is located at the repository root and initializes global sidebar filters correctly.", "PASS"),
    ("Pinned Dependencies", "requirements.txt pins exact runtime versions (streamlit, pandas, numpy, plotly, etc.).", "PASS"),
    ("Python Runtime Specified", "runtime.txt contains 'python-3.12' to ensure exact container version matching.", "PASS"),
    ("Database Packaging", "data/app.db is committed directly into the Git repository (size: 304 KB, well below GitHub limit).", "PASS"),
    ("Path Resolution Safety", "All file loads in src/data_loader.py use os.path.dirname relative resolution (no hardcoded absolute paths).", "PASS"),
    ("Credential & Secret Hygiene", "No API keys, passwords, or tokens are committed; .gitignore ignores local environments (.venv/).", "PASS"),
    ("Automated Testing Pass", "15/15 technical unit tests pass cleanly in pytest within 4.5 seconds.", "PASS")
]

aud_tbl = doc.add_table(rows=len(audit_checks)+1, cols=3)
aud_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Verification Item', 'Observed Implementation Details', 'Audit Status']):
    c = aud_tbl.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)

for ri, (item, details, status) in enumerate(audit_checks, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    for ci, val in enumerate([item, details, status]):
        c = aud_tbl.cell(ri, ci); set_cell_bg(c, fill_c)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(val); r.font.size = Pt(8.5)
        if ci in (0, 2): r.bold = True
        if ci == 2: r.font.color.rgb = GREEN

set_tbl_borders(aud_tbl)
doc.add_paragraph()

# 2. Step-by-Step GitHub & Streamlit Cloud Instructions
add_heading('2. Step-by-Step Cloud Deployment Walkthrough', 1)
add_para('Follow these exact steps to publish the repository to Streamlit Community Cloud:', size=9.5)

steps = [
    ("Step 1: Initialize Git & Push", "Initialize local git repository, commit all project files, and push to GitHub:\n"
     "git init\n"
     "git add .\n"
     "git commit -m 'Release v1.0: Complete Hiring Trend Analysis Platform'\n"
     "git branch -M main\n"
     "git remote add origin https://github.com/<your-username>/2500320140004-hiringtrendanalysis.git\n"
     "git push -u origin main"),
    ("Step 2: Authenticate on Streamlit Cloud", "Navigate to https://share.streamlit.io/ and log in using your GitHub account credentials."),
    ("Step 3: Create New Application", "Click the 'Create app' button in the Streamlit Cloud dashboard and choose 'I already have an app'."),
    ("Step 4: Configure Repository Settings", "Fill in the exact deployment parameters:\n"
     "• Repository: <your-username>/2500320140004-hiringtrendanalysis\n"
     "• Branch: main\n"
     "• Main file path: app.py\n"
     "• Custom App URL: 2500320140004-hiringtrendanalysis"),
    ("Step 5: Deploy & Monitor Build Log", "Click 'Deploy!'. Streamlit Cloud will provision a Debian-based container, install requirements.txt via pip, and launch the server in 2-3 minutes. Watch the 'Manage app' console log for confirmation.")
]

for title, desc in steps:
    add_heading(title, 2)
    add_para(desc, size=9)
    doc.add_paragraph()

# 3. Troubleshooting Matrix
add_heading('3. Common Cloud Deployment Issues & Exact Resolutions', 1)
add_para('In the event of build or runtime errors on Streamlit Community Cloud, consult the diagnostic table below:', size=9.5)

issues = [
    ("ModuleNotFoundError: No module named 'xyz'", "A library is imported in src/ or pages/ but was omitted from requirements.txt.", "Add the missing library with its exact version to requirements.txt and push to GitHub."),
    ("FileNotFoundError: data/app.db", "Absolute Windows paths (e.g. C:\\...) were used, or data/ was excluded in .gitignore.", "Use relative path resolution: os.path.join(os.path.dirname(__file__), '..', 'data', 'app.db') and ensure data/app.db is committed."),
    ("Cold start container timeout (> 10 mins)", "Unpinned requirements caused pip to attempt complex dependency resolution or build heavy wheels.", "Pin exact wheel-compatible versions (e.g. pandas==2.2.2) in requirements.txt."),
    ("PermissionError: [WinError 10013] Socket forbidden", "Trying to invoke local network sockets or bind ports inside sandbox environments.", "Streamlit handles port binding automatically on cloud containers; never hardcode local port bindings in app scripts.")
]

iss_tbl = doc.add_table(rows=len(issues)+1, cols=3)
iss_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Observed Cloud Symptom', 'Root Cause Explanation', 'Exact Technical Fix']):
    c = iss_tbl.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(8.5)

for ri, (sym, cause, fix) in enumerate(issues, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    for ci, val in enumerate([sym, cause, fix]):
        c = iss_tbl.cell(ri, ci); set_cell_bg(c, fill_c)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci == 0 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(val); r.font.size = Pt(8)
        if ci == 0: r.bold = True

set_tbl_borders(iss_tbl)
doc.add_paragraph()

# Document Control
add_heading('Document Control', 2)
dc = doc.add_table(rows=2, cols=4); dc.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Version', 'Date', 'Author', 'Summary']):
    c = dc.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)
for ci, val in enumerate(['1.0', datetime.date.today().strftime('%d %b %Y'), 'Abhayjeet Singh', 'Initial Deployment Guide & Cloud Publishing Specification']):
    c = dc.cell(1, ci); set_cell_bg(c, 'EEF4FB')
    p = c.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)
set_tbl_borders(dc)

doc.save(OUT_DOCX)
print(f"SAVED DEPLOYMENT GUIDE DOCX: {OUT_DOCX}")
