"""
Step 12 – UI Framework Generator & RTM Updater
Project : P071 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES MCA 2nd Yr DA
"""

import os, datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOCS_DIR = r"C:\Hiring Trend Analysis\docs"
SRC_DIR  = r"C:\Hiring Trend Analysis\src"
CONFIG_DIR = r"C:\Hiring Trend Analysis\.streamlit"
os.makedirs(DOCS_DIR, exist_ok=True)
os.makedirs(SRC_DIR, exist_ok=True)
os.makedirs(CONFIG_DIR, exist_ok=True)

OUT_DOCX    = os.path.join(DOCS_DIR, "P01_HTA_Step12_UIFramework.docx")
RTM_FILE    = os.path.join(DOCS_DIR, "P01_HTA_Step04_RTM.xlsx")
UI_PY_FILE  = os.path.join(SRC_DIR, "ui.py")
CONFIG_TOML = os.path.join(CONFIG_DIR, "config.toml")

# ── 1. WRITE .streamlit/config.toml ──────────────────────────────────────────
config_toml_content = """# Streamlit Configuration & Custom Theme for Hiring Trend Analysis
[theme]
primaryColor       = "#2E74B5"
backgroundColor    = "#FFFFFF"
secondaryBackgroundColor = "#F8FAFC"
textColor          = "#0F172A"
font               = "sans serif"

[server]
headless = true
enableCORS = false
enableXsrfProtection = true
maxUploadSize = 50

[browser]
gatherUsageStats = false
"""
with open(CONFIG_TOML, "w", encoding="utf-8") as f:
    f.write(config_toml_content)
print(f"SAVED: {CONFIG_TOML}")

# ── 2. WRITE src/ui.py (TAILWIND COMPONENTS VIA st.components.v1.html) ─────────
ui_py_content = '''"""
src/ui.py — Reusable UI Components styled with Tailwind CSS CDN
Provides professional hero headers, metric KPI cards, and insight banners
via Streamlit components iframe without requiring Node/npm.
"""
import streamlit as st
import streamlit.components.v1 as components

TAILWIND_CDN = '<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">'

def render_hero_header(title: str, subtitle: str, metadata: str = "P071 | MCA DA ABES"):
    """Render an executive Hero Banner styled with Tailwind CSS."""
    html_code = f"""
    {TAILWIND_CDN}
    <div class="bg-gradient-to-r from-blue-900 via-indigo-900 to-blue-800 text-white p-6 rounded-xl shadow-lg border border-blue-700 font-sans">
        <div class="flex items-center justify-between">
            <div>
                <h1 class="text-2xl font-bold tracking-tight text-white mb-1">📊 {title}</h1>
                <p class="text-blue-200 text-sm">{subtitle}</p>
            </div>
            <div class="hidden sm:block text-right">
                <span class="inline-block bg-blue-700 bg-opacity-60 text-blue-100 text-xs px-3 py-1 rounded-full border border-blue-500 font-mono">
                    {metadata}
                </span>
            </div>
        </div>
    </div>
    """
    components.html(html_code, height=110)

def render_kpi_cards(total_jobs: int, avg_salary_inr: float, top_city: str, top_role: str):
    """Render 4 responsive executive KPI Metric Cards in a single Tailwind grid."""
    avg_sal_lakhs = avg_salary_inr / 100000.0
    html_code = f"""
    {TAILWIND_CDN}
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 font-sans">
        <!-- Card 1 -->
        <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 hover:shadow-md transition">
            <div class="text-xs font-semibold uppercase tracking-wider text-gray-500">Total Postings</div>
            <div class="text-2xl font-extrabold text-blue-900 mt-1">{total_jobs:,}</div>
            <div class="text-xs text-green-600 mt-1 font-medium">↑ Verified Dataset</div>
        </div>
        <!-- Card 2 -->
        <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 hover:shadow-md transition">
            <div class="text-xs font-semibold uppercase tracking-wider text-gray-500">Average Salary</div>
            <div class="text-2xl font-extrabold text-indigo-900 mt-1">₹{avg_sal_lakhs:.1f} LPA</div>
            <div class="text-xs text-blue-600 mt-1 font-medium">Benchmark Avg</div>
        </div>
        <!-- Card 3 -->
        <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 hover:shadow-md transition">
            <div class="text-xs font-semibold uppercase tracking-wider text-gray-500">Top Hiring City</div>
            <div class="text-2xl font-extrabold text-blue-950 mt-1 truncate">{top_city}</div>
            <div class="text-xs text-purple-600 mt-1 font-medium">Prime Hotspot</div>
        </div>
        <!-- Card 4 -->
        <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 hover:shadow-md transition">
            <div class="text-xs font-semibold uppercase tracking-wider text-gray-500">Most Demanded Role</div>
            <div class="text-2xl font-extrabold text-blue-950 mt-1 truncate">{top_role}</div>
            <div class="text-xs text-amber-600 mt-1 font-medium">Highest Volume</div>
        </div>
    </div>
    """
    components.html(html_code, height=115)

def render_insight_box(insight_text: str, title: str = "Key Business Insight"):
    """Render a highlighted contextual insight callout banner."""
    html_code = f"""
    {TAILWIND_CDN}
    <div class="bg-blue-50 border-l-4 border-blue-600 p-4 rounded-r-lg shadow-xs font-sans">
        <div class="flex items-start">
            <span class="text-blue-600 font-bold mr-2 text-base">💡 {title}:</span>
            <p class="text-blue-900 text-sm leading-relaxed">{insight_text}</p>
        </div>
    </div>
    """
    components.html(html_code, height=75)
'''
with open(UI_PY_FILE, "w", encoding="utf-8") as f:
    f.write(ui_py_content)
print(f"SAVED: {UI_PY_FILE}")

# ── 3. BUILD 12_UI_Framework.docx ─────────────────────────────────────────────

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
r2 = bp.add_run('Step 12 · Streamlit UI Framework & Tailwind Wireframe Specification\n'); r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
r3 = bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA'); r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

# 1. Screen Registry Table
add_heading('1. Application Screen Registry (UI-01 to UI-10)', 1)
add_para('The Streamlit web application is structured as a 10-screen multipage dashboard. Every screen maps to specific business requirements, operational questions, and user stories.', size=9.5)

screens_data = [
    ("UI-01", "Overview & Executive KPIs", "pages/01_Overview.py", "Macro hiring snapshot, 4 Tailwind KPI cards, recent dataset highlights", "Date range, Industry, City", "US-03, US-04", "OQ-01, OQ-02, OQ-04"),
    ("UI-02", "Role Demand Analysis", "pages/02_Role_Demand.py", "Ranked horizontal bar charts of job posting volume by role", "Industry, Date range", "US-05", "OQ-01"),
    ("UI-03", "Skills Demand Analysis", "pages/03_Skills.py", "Top 10 skills by frequency with industry-level decomposition", "Industry multiselect", "US-06", "OQ-03"),
    ("UI-04", "Salary Benchmark & Stats", "pages/04_Salary.py", "Min, Mean, and Max salary comparisons per industry and city", "Experience level, City", "US-09, US-10", "OQ-02, OQ-07"),
    ("UI-05", "Geographic Hotspots", "pages/05_Geography.py", "City-level job volume distribution and regional concentration", "Region, Job type", "US-12", "OQ-05"),
    ("UI-06", "Hiring Volume Trends", "pages/06_Trends.py", "Monthly chronological timeline across 2022-2024 and role trends", "Role multiselect, Year", "US-13, US-14", "OQ-04"),
    ("UI-07", "Skill-Salary Correlation", "pages/07_Correlation.py", "Scatter plot with OLS trendline, Pearson r, and p-value readout", "Experience tier", "US-11", "OQ-09"),
    ("UI-08", "Company Leaderboard", "pages/08_Companies.py", "Top 10 hiring companies ranked by volume with tabular export", "Domain/Sector", "US-16", "OQ-10"),
    ("UI-09", "EDA & Data Quality Audit", "pages/09_EDA.py", "Pandas describe() table, null counts, and IQR outlier detection", "Table selector", "US-01, US-02", "BR-13 (EDA)"),
    ("UI-10", "Job Types & Experience", "pages/10_Job_Types.py", "Donut chart for work modes and 100% stacked bar for experience", "Industry", "US-07, US-08, US-15", "OQ-06, OQ-08, OQ-11")
]

scr_tbl = doc.add_table(rows=len(screens_data)+1, cols=7)
scr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['UI ID', 'Screen Title', 'File Path', 'Primary Purpose', 'Sidebar Filters', 'User Stories', 'OQ']):
    c = scr_tbl.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(8.5)

for ri, row in enumerate(screens_data, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    for ci, val in enumerate(row):
        c = scr_tbl.cell(ri, ci); set_cell_bg(c, fill_c)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci in (0, 4, 5, 6) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(val); r.font.size = Pt(8)
        if ci == 0: r.bold = True

set_tbl_borders(scr_tbl)
doc.add_paragraph()

# 2. ASCII Wireframes for Primary Screens
add_heading('2. ASCII Screen Wireframes', 1)
add_para('Below are textual wireframes demonstrating the visual organization, filter positioning, and chart container distribution for primary screens.', size=9.5)

wireframes = [
("UI-01: Overview & Executive KPIs", """+---------------------------------------------------------------------------------------------------+
| SIDEBAR FILTERS          | [ HERO BANNER: 📊 Hiring Trend Analysis — Executive Overview ]        |
|--------------------------+------------------------------------------------------------------------|
| [Date Slider: 2022-2024] | [ TOTAL POSTINGS ] [ AVERAGE SALARY ] [ TOP HIRING CITY ] [ TOP ROLE ] |
| [Industry: [All v]     ] | [     652        ] [   ₹12.4 LPA    ] [    Bangalore   ] [Data Analyst]|
| [City:     [All v]     ] |------------------------------------------------------------------------|
| [Exp Level:[All v]     ] | [ CHART: Monthly Hiring Trend Line ]    [ CHART: Top 5 Roles Bar ]     |
|                          | | Volume (Postings)                |    | Role 1: ======= (120)  |     |
| [Reset Filters]          | |  __/\__/\_                       |    | Role 2: =====   (85)   |     |
| [Download CSV]           | +----------------------------------+----+------------------------+-----+
|                          | [ TABULAR DATA PREVIEW (Last 5 Postings) ]                             |
+--------------------------+------------------------------------------------------------------------+"""),

("UI-04: Salary Benchmark & Stats", """+---------------------------------------------------------------------------------------------------+
| SIDEBAR FILTERS          | [ HERO BANNER: 💰 Salary Benchmark & Industry Compensation Analytics ]|
|--------------------------+------------------------------------------------------------------------|
| [Industry: Select v   ] | [ Metric: Min Salary ]   [ Metric: Avg Salary ]   [ Metric: Max Salary ] |
| [Experience: [Fresher] ] | [     ₹3.50 LPA      ]   [     ₹8.70 LPA      ]   [     ₹24.0 LPA      ] |
| [City:     [All v]     ] |------------------------------------------------------------------------|
|                          | [ CHART: Grouped Bar Chart — Salary by Industry (Min vs Avg vs Max) ]  |
|                          | | Industry A: [Min==][Avg====][Max========]                            |
|                          | | Industry B: [Min= ][Avg=== ][Max======= ]                            |
|                          | +----------------------------------------------------------------------+
|                          | [ CHART: Box Plot — Salary Distribution & Spread Across Exp Levels ]   |
|                          | [ DATA TABLE: Filtered Salary Statistics Table with CSV Export Button] |
+--------------------------+------------------------------------------------------------------------+"""),

("UI-07: Skill-Salary Correlation", """+---------------------------------------------------------------------------------------------------+
| SIDEBAR FILTERS          | [ HERO BANNER: 🔬 Statistical Correlation: Required Skills vs Offered Salary ]|
|--------------------------+------------------------------------------------------------------------|
| [Role Family: Analytics] | [ Pearson r: +0.48 ]   [ p-value: < 0.001 ]   [ R-Squared: 0.23 ]      |
| [Experience: [All v]   ] | [ Significant Positive Correlation: Multi-skilled roles earn higher CTC]|
|                          |------------------------------------------------------------------------|
|                          | [ SCATTER PLOT WITH OLS TRENDLINE ]                                    |
|                          | | Salary (₹)                                                           |
|                          | | 30L |         *    *  / (OLS Fit Line)                               |
|                          | | 20L |      *  *   *  /                                               |
|                          | | 10L |   *  *   *    /                                                |
|                          | |  0L +--*---*--------/-------------------                             |
|                          | |       1   2   3   4   5   6   7   8 (Skills Count)                   |
|                          | +----------------------------------------------------------------------+
|                          | [ INSIGHT BOX: Each additional technical skill adds ~₹1.2L to max CTC] |
+--------------------------+------------------------------------------------------------------------+""")
]

for title, wf in wireframes:
    add_heading(title, 2)
    code_box(wf)
    doc.add_paragraph()

# 3. Best Chart for each OQ
add_heading('3. Chart Selection Rationale for Operational Questions', 1)
add_para('Each chart type was chosen specifically to match the data grain, dimensional cardinality, and visual readability of its operational question:', size=9.5)

chart_rationale = [
    ("OQ-01 (Role Demand)", "Horizontal Bar Chart (`px.bar`)", "Horizontal orientation accommodates lengthy role names without label overlap; natural descending sort reveals highest demand instantly."),
    ("OQ-02 (Salary Benchmarking)", "Grouped Bar + Box Plot (`px.box`)", "Grouped bars show exact Min/Mean/Max metrics side-by-side; box plots illustrate variance, quartiles, and compensation outliers."),
    ("OQ-03 (Skills Frequency)", "Horizontal Bar Chart (`px.bar`)", "Ranked list of decomposed skills; top-10 view highlights core market requirements without visual clutter."),
    ("OQ-04 (Hiring Trends)", "Chronological Line Chart (`px.line`)", "Continuous time dimension (monthly granularity) is best interpreted as a continuous trendline showing trajectory and seasonality."),
    ("OQ-05 (Geographic Hotspots)", "Ranked Bar Chart (`px.bar`)", "Directly compares job counts across 15 distinct cities; horizontal layout provides clean reading order."),
    ("OQ-06 (Work Mode Split)", "Donut Chart (`px.pie` with hole=0.4)", "Represents part-to-whole proportions of exactly 5 mutually exclusive job types; center hole displays total count."),
    ("OQ-07 (Industries by Skill)", "Horizontal Bar Chart (`px.bar`)", "Enables recruiters and counselors to see which industries hire most heavily for a selected skill keyword."),
    ("OQ-08 (Experience Breakdown)", "100% Stacked Bar (`barmode='stack'`)", "Normalizes each role to 100%, allowing instant visual comparison of seniority composition across roles of varying total size."),
    ("OQ-09 (Skill-Salary Correlation)", "Scatter Plot + OLS (`trendline='ols'`)", "Visualizes bivariate continuous data (`skills_count` vs `salary_max`) with an ordinary least squares regression line and Pearson r."),
    ("OQ-10 (Company Leaderboard)", "Ranked Horizontal Bar + Interactive Table", "Quickly highlights top 10 hiring firms while providing full details in an accompanying sortable table.")
]

ch_tbl = doc.add_table(rows=len(chart_rationale)+1, cols=3)
ch_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Operational Question', 'Chosen Chart Type', 'Visual & Statistical Rationale']):
    c = ch_tbl.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)

for ri, (oq, ctype, rat) in enumerate(chart_rationale, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    for ci, val in enumerate([oq, ctype, rat]):
        c = ch_tbl.cell(ri, ci); set_cell_bg(c, fill_c)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci in (0, 1) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(val); r.font.size = Pt(8.5)
        if ci in (0, 1): r.bold = True

set_tbl_borders(ch_tbl)
doc.add_paragraph()

# 4. Global Styling & Tailwind Component Code
add_heading('4. Reusable Tailwind CSS Component Implementations', 1)
add_para('Below is the starter code implemented in `src/ui.py` providing the Tailwind Hero Banner, responsive KPI cards, and custom insight boxes:', size=9.5)
code_box(ui_py_content)
doc.add_paragraph()

# Document Control
add_heading('Document Control', 2)
dc = doc.add_table(rows=2, cols=4); dc.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Version', 'Date', 'Author', 'Summary']):
    c = dc.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)
for ci, val in enumerate(['1.0', datetime.date.today().strftime('%d %b %Y'), 'Abhayjeet Singh', 'Initial UI Framework, Wireframes & Tailwind Component deliverable']):
    c = dc.cell(1, ci); set_cell_bg(c, 'EEF4FB')
    p = c.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)
set_tbl_borders(dc)

doc.save(OUT_DOCX)
print(f"SAVED UI FRAMEWORK DOCX: {OUT_DOCX}")

# ── 4. UPDATE RTM EXCEL FILE (COLUMN L: UI SCREEN ID) ─────────────────────────

wb = openpyxl.load_workbook(RTM_FILE)
ws = wb["RTM"]

# Mapping from BR_ID to UI Screen IDs
br_to_ui_map = {
    "BR-01": "UI-01, UI-02",
    "BR-02": "UI-03",
    "BR-03": "UI-01, UI-04",
    "BR-04": "UI-08",
    "BR-05": "UI-01, UI-06",
    "BR-06": "UI-04",
    "BR-07": "UI-01, UI-05",
    "BR-08": "UI-10",
    "BR-09": "UI-07",
    "BR-10": "UI-10",
    "BR-11": "UI-10",
    "BR-12": "UI-06",
    "BR-13": "UI-09"
}

# Update Column L (index 12) for BR rows
for row_idx in range(4, 17):
    br_id = ws.cell(row=row_idx, column=3).value
    if br_id in br_to_ui_map:
        c = ws.cell(row=row_idx, column=12)
        c.value = br_to_ui_map[br_id]
        c.fill = PatternFill('solid', fgColor='E2EFDA')
        c.font = Font(name='Calibri', size=9, bold=True, color='1F5C2E')
        c.alignment = Alignment(horizontal='center', vertical='center')

# Update Status in Column H (index 8) to 'UI Mapped'
for row_idx in range(4, 17):
    c = ws.cell(row=row_idx, column=8)
    c.value = 'UI Mapped'
    c.fill = PatternFill('solid', fgColor='E5FFE5')
    c.font = Font(name='Calibri', size=9, bold=True, color='1F5C2E')
    c.alignment = Alignment(horizontal='center', vertical='center')

# Update Changelog
ws_ch = wb["Changelog"]
for r in range(3, 10):
    if ws_ch.cell(r, 1).value and "Step 12" in str(ws_ch.cell(r, 1).value):
        ws_ch.cell(r, 2).value = datetime.date.today().strftime('%d %b %Y')
        ws_ch.cell(r, 5).value = "Populated Column L with UI Screen IDs (UI-01..10). Status -> UI Mapped."
        break

wb.save(RTM_FILE)
print(f"UPDATED RTM: {RTM_FILE}")
