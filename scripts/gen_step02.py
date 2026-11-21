"""
Step 02 – How Might We (HMW) Statements Document Generator
Project : P071 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES, MCA 2nd Yr DA
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os, datetime

OUT_DIR  = r"C:\Hiring Trend Analysis\docs"
OUT_FILE = os.path.join(OUT_DIR, "P01_HTA_Step02_HowMightWe.docx")
os.makedirs(OUT_DIR, exist_ok=True)

# ── colours ───────────────────────────────────────────────────────────────────
DARK_BLUE  = RGBColor(0x1F, 0x35, 0x64)
MID_BLUE   = RGBColor(0x2E, 0x74, 0xB5)
STEEL      = RGBColor(0x4A, 0x86, 0xC8)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
BLACK      = RGBColor(0x00, 0x00, 0x00)
GREEN_OK   = RGBColor(0x37, 0x86, 0x36)
GOLD       = RGBColor(0xFF, 0xC0, 0x00)
PRIORITY_BG= 'FFF3CD'   # amber tint for priority rows
ALT_BG     = 'DAE8F5'
WHITE_HEX  = 'FFFFFF'

# ── helpers ───────────────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def set_borders(table, color='2E74B5'):
    for row in table.rows:
        for cell in row.cells:
            tc   = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcB  = OxmlElement('w:tcBorders')
            for edge in ('top','left','bottom','right','insideH','insideV'):
                e = OxmlElement(f'w:{edge}')
                e.set(qn('w:val'),   'single')
                e.set(qn('w:sz'),    '4')
                e.set(qn('w:space'), '0')
                e.set(qn('w:color'), color)
                tcB.append(e)
            tcPr.append(tcB)

def header_row(table, *texts, bg='1F3564'):
    row = table.rows[0]
    for cell, text in zip(row.cells, texts):
        set_cell_bg(cell, bg)
        p   = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.color.rgb = WHITE
        run.font.size = Pt(9.5)

def fill_row(table, row_idx, texts, bg=WHITE_HEX, bold_col=None):
    row = table.rows[row_idx]
    for i, (cell, text) in enumerate(zip(row.cells, texts)):
        set_cell_bg(cell, bg)
        p   = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(str(text))
        run.font.size = Pt(9.5)
        if bold_col is not None and i == bold_col:
            run.bold = True

def add_heading(doc, text, level=1):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size = Pt(14); run.font.color.rgb = DARK_BLUE
    elif level == 2:
        run.font.size = Pt(12); run.font.color.rgb = MID_BLUE
    else:
        run.font.size = Pt(11); run.font.color.rgb = STEEL

def add_para(doc, text, bold=False, italic=False, size=10, color=BLACK, indent=0):
    p   = doc.add_paragraph()
    p.paragraph_format.left_indent  = Cm(indent)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.bold = bold; run.italic = italic
    run.font.size = Pt(size); run.font.color.rgb = color
    return p

def thin_rule(doc):
    p   = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr= OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '6')
    bot.set(qn('w:space'), '1');    bot.set(qn('w:color'), '2E74B5')
    pBdr.append(bot); pPr.append(pBdr)

# ═════════════════════════════════════════════════════════════════════════════
doc = Document()
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
r1 = bp.add_run('P071 – HIRING TREND ANALYSIS\n')
r1.bold = True; r1.font.size = Pt(18); r1.font.color.rgb = WHITE
r2 = bp.add_run('Step 02 · How Might We (HMW) Statements\n')
r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
r3 = bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA')
r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

meta = doc.add_table(rows=2, cols=4)
meta.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project ID','P071'), ('Version','1.0'),
    ('Date', datetime.date.today().strftime('%d %b %Y')), ('Status','Draft'),
    ('Step','02 of 18'), ('Prev Step','01 – Project Understanding'),
    ('Next Step','03 – Needs Analysis (BR)'), ('Method','Design Thinking – HMW'),
]
for i, row in enumerate(meta.rows):
    fill = 'EEF4FB' if i % 2 == 0 else WHITE_HEX
    for j, cell in enumerate(row.cells):
        set_cell_bg(cell, fill)
        p   = cell.paragraphs[0]
        txt = meta_data[i*4 + j]
        run = p.add_run(f'{txt[0]}: '); run.bold = True; run.font.size = Pt(9)
        p.add_run(txt[1]).font.size = Pt(9)
set_borders(meta)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — INTRODUCTION
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '1. Introduction — Design Thinking & HMW', 1)
add_para(doc,
    'The "How Might We …?" (HMW) framework is a core design-thinking tool that reframes pain '
    'points as open-ended opportunity questions. A well-formed HMW is narrow enough to be '
    'actionable but broad enough to allow creative solutions. In this step we:',
    size=10)
for bullet in [
    'Identify 2–3 concrete pain points per stakeholder (derived from Step 01).',
    'Convert each pain point into a targeted HMW statement tied to this project.',
    'Cross-link every HMW to at least one Operational Question (OQ) from Step 01.',
    'Prioritise the top 5 HMWs that have the highest impact on the app\'s core value.',
]:
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_after = Pt(1)
    p.add_run(bullet).font.size = Pt(9.5)

add_para(doc,
    'Rule check — every HMW must be: specific to Hiring Trend Analysis · stakeholder-named · '
    'not a disguised solution · linked to at least one OQ.',
    italic=True, size=9, color=MID_BLUE)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — STAKEHOLDER PAIN POINTS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '2. Stakeholder Pain Points', 1)
add_para(doc,
    'Each stakeholder from the Step 01 register is mapped to 2–3 root pain points '
    'that the Hiring Trend Analysis app is positioned to address.',
    size=9.5)
doc.add_paragraph()

pain_data = {
    'SH-03 — HR Manager': [
        ('PP-01', 'Cannot quickly identify which job roles are surging or declining in demand without manually compiling multiple spreadsheets.'),
        ('PP-02', 'No single view to compare salary benchmarks across industries — budgeting is based on guesswork or outdated surveys.'),
        ('PP-03', 'Lacks visibility into seasonal hiring volume cycles, making it hard to plan recruitment campaign timelines.'),
    ],
    'SH-04 — Recruiter': [
        ('PP-04', 'JD skill lists are written from memory or copied from old postings, not driven by current market demand data.'),
        ('PP-05', 'Salary bands offered are set intuitively; no easy way to know if the offer is competitive for a given role + city + experience level.'),
        ('PP-06', 'No rapid way to identify which competitor companies are recruiting heavily in the same domain or city.'),
    ],
    'SH-05 — Job Seeker': [
        ('PP-07', 'Unsure which skills to learn next — no data on which skills actually appear most in real job postings.'),
        ('PP-08', 'Cannot verify whether their salary expectation aligns with market reality for their experience level and target location.'),
        ('PP-09', 'Does not know the best city to target for maximum job opportunities in their chosen domain.'),
    ],
    'SH-06 — Placement Officer': [
        ('PP-10', 'No data-backed method to decide which companies to invite for campus placement drives.'),
        ('PP-11', 'Cannot match student batch profiles (fresher, specific skills) to the actual experience levels and skills demanded by industry.'),
        ('PP-12', 'Unaware of seasonal hiring peaks — drives are scheduled based on tradition, not market timing.'),
    ],
    'SH-02 — Developer (Abhayjeet)': [
        ('PP-13', 'Unclear which chart type best answers each operational question without user research or domain knowledge.'),
        ('PP-14', 'Generating realistic synthetic hiring data that reflects Indian job-market distributions is non-trivial.'),
    ],
}

for sh_label, pains in pain_data.items():
    add_heading(doc, sh_label, 2)
    tbl = doc.add_table(rows=len(pains)+1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_row(tbl, 'PP ID', 'Pain Point', 'Root Cause')
    root_causes = {
        'PP-01': 'Data scattered across multiple systems; no aggregated view',
        'PP-02': 'No real-time salary benchmarking tool integrated into hiring workflow',
        'PP-03': 'No historical trend visibility; decisions made on intuition',
        'PP-04': 'No live market-skills data feed available to recruiters',
        'PP-05': 'Salary data is siloed; no cross-role, cross-city comparison tool',
        'PP-06': 'Competitor intelligence requires expensive tools or manual search',
        'PP-07': 'Job portals show listings but no aggregated skill-frequency analytics',
        'PP-08': 'Published salary surveys are infrequent and lack granularity',
        'PP-09': 'No geographic demand heatmap accessible to individual job seekers',
        'PP-10': 'Company hiring volume data is not publicly aggregated',
        'PP-11': 'Student profile data and market demand data live in separate silos',
        'PP-12': 'No time-series hiring trend data used in drive scheduling decisions',
        'PP-13': 'No prior EDA project experience; chart choice is a design skill gap',
        'PP-14': 'Real hiring datasets are proprietary; realism requires domain research',
    }
    for idx, (pid, ptext) in enumerate(pains):
        fill_row(tbl, idx+1, [pid, ptext, root_causes[pid]],
                 bg=ALT_BG if idx % 2 == 0 else WHITE_HEX, bold_col=0)
    set_borders(tbl)
    doc.add_paragraph()

thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — HMW STATEMENTS (MASTER TABLE)
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '3. How Might We Statements — Master Table', 1)
add_para(doc,
    'The 12 HMW statements below were derived directly from the pain points above. '
    'Each statement is specific to the Hiring Trend Analysis project, names a stakeholder, '
    'and links to at least one OQ. Priority HMWs are highlighted in amber.',
    size=9.5)
doc.add_paragraph()

hmw_rows = [
    # (HMW_ID, Statement, SH, PP_IDs, OQ_IDs, priority_reason or None)
    ('HMW-01',
     'How might we show an HR manager which specific job roles have spiked or dropped in posting volume over a chosen 3-month window, without opening a single spreadsheet?',
     'SH-03', 'PP-01', 'OQ-01, OQ-04',
     '★ PRIORITY — Directly delivers the core "trend" promise of the app title; highest HR manager value.'),
    ('HMW-02',
     'How might we display the exact top-10 skills demanded across a selected industry so a recruiter can write a data-driven job description in under 2 minutes?',
     'SH-04', 'PP-04', 'OQ-03',
     '★ PRIORITY — Addresses the most common recruiter pain point; high daily usage potential.'),
    ('HMW-03',
     'How might we let a job seeker instantly compare their salary expectation against the actual average offered for their experience level and target city?',
     'SH-05', 'PP-08', 'OQ-02, OQ-08',
     '★ PRIORITY — Salary benchmarking is the #1 information job seekers search for; broad audience appeal.'),
    ('HMW-04',
     'How might we give a placement officer a ranked list of the top-10 hiring companies in a selected domain so they can prioritise campus drive invitations with data?',
     'SH-06', 'PP-10', 'OQ-10',
     '★ PRIORITY — Directly replaces guesswork in company selection with an objective, ranked leaderboard.'),
    ('HMW-05',
     'How might we reveal the month-over-month change in total hiring volume on a single interactive timeline so an HR manager can align campaign launches with demand peaks?',
     'SH-03', 'PP-03', 'OQ-04',
     '★ PRIORITY — Seasonal planning is a recurring decision; a time-series chart provides reusable insight.'),
    ('HMW-06',
     'How might we surface the statistical correlation between the number of skills a role requires and its maximum offered salary, so a recruiter can price complex roles objectively?',
     'SH-04', 'PP-05', 'OQ-09',
     None),
    ('HMW-07',
     'How might we display a city-wise hiring concentration chart so that a job seeker relocating to a new city can compare opportunity density across the top 10 locations?',
     'SH-05', 'PP-09', 'OQ-05',
     None),
    ('HMW-08',
     'How might we show the job-type distribution (remote / hybrid / full-time / contract) for each industry so a recruiter can benchmark their flexible-work policy against market norms?',
     'SH-04', 'PP-05', 'OQ-06',
     None),
    ('HMW-09',
     'How might we break down experience-level demand (Fresher / Mid / Senior) by job role so a placement officer can match the right student batch to the right recruiter?',
     'SH-06', 'PP-11', 'OQ-08',
     None),
    ('HMW-10',
     'How might we provide HR managers a single-screen industry salary comparison — showing average, min, and max — so budget approvals no longer depend on third-party surveys?',
     'SH-03', 'PP-02', 'OQ-02, OQ-07',
     None),
    ('HMW-11',
     'How might we identify which industries show the fastest growth in hiring volume for a selected skill set, so a career counsellor can redirect students toward emerging sectors?',
     'SH-06', 'PP-12', 'OQ-07, OQ-04',
     None),
    ('HMW-12',
     'How might we show a time-series view of posting frequency for a specific job role so a job seeker can time their applications during peak hiring windows?',
     'SH-05', 'PP-07', 'OQ-01, OQ-04',
     None),
]

# build master table: 7 columns
col_heads = ['HMW ID','How Might We …? Statement','Stakeholder','Pain Point(s)','Linked OQ(s)','Priority']
tbl = doc.add_table(rows=len(hmw_rows)+1, cols=6)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

# set column widths
widths = [Cm(1.5), Cm(8.5), Cm(2.0), Cm(1.8), Cm(2.0), Cm(2.5)]
for row in tbl.rows:
    for i, cell in enumerate(row.cells):
        cell.width = widths[i]

header_row(tbl, *col_heads)

for idx, (hmw_id, stmt, sh, pp, oq, prio) in enumerate(hmw_rows):
    is_prio = prio is not None
    bg = PRIORITY_BG if is_prio else (ALT_BG if idx % 2 == 0 else WHITE_HEX)
    prio_text = '⭐ YES' if is_prio else '—'
    fill_row(tbl, idx+1, [hmw_id, stmt, sh, pp, oq, prio_text], bg=bg, bold_col=0)

set_borders(tbl)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — TOP 5 PRIORITY HMWs (detailed)
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '4. Top 5 Priority HMW Statements — Detailed Rationale', 1)
add_para(doc,
    'Priority was determined by: (a) direct link to the app\'s stated value proposition, '
    '(b) frequency of use by primary stakeholders, and (c) coverage of the most critical OQs.',
    size=9.5, italic=True)
doc.add_paragraph()

priority_detail = [
    ('HMW-01',
     'How might we show an HR manager which specific job roles have spiked or dropped in posting volume over a chosen 3-month window, without opening a single spreadsheet?',
     'SH-03 — HR Manager', 'OQ-01, OQ-04',
     'This directly delivers the core "Trend" promise in the project title. If the app cannot show role-level demand changes over time, it fails its primary purpose. High daily business impact.'),
    ('HMW-02',
     'How might we display the exact top-10 skills demanded across a selected industry so a recruiter can write a data-driven job description in under 2 minutes?',
     'SH-04 — Recruiter', 'OQ-03',
     'Skill-frequency analysis is the most actionable insight for recruiters. It transforms a passive chart into a decision-support tool used every time a new JD is opened.'),
    ('HMW-03',
     'How might we let a job seeker instantly compare their salary expectation against the actual average offered for their experience level and target city?',
     'SH-05 — Job Seeker', 'OQ-02, OQ-08',
     'Salary benchmarking is universally the #1 search query for job seekers. Solving this OQ makes the app relevant and shareable far beyond the classroom.'),
    ('HMW-04',
     'How might we give a placement officer a ranked list of the top-10 hiring companies in a selected domain so they can prioritise campus drive invitations with data?',
     'SH-06 — Placement Officer', 'OQ-10',
     'Placement drives have fixed time slots; a data-backed company leaderboard replaces subjective choices with an objective ranking that directly improves placement outcomes.'),
    ('HMW-05',
     'How might we reveal the month-over-month change in total hiring volume on a single interactive timeline so an HR manager can align campaign launches with demand peaks?',
     'SH-03 — HR Manager', 'OQ-04',
     'Seasonality insight is repeatable and recurring — it is consulted every quarter. A time-series trend chart is the highest-reuse feature in the entire app.'),
]

for i, (hid, stmt, sh, oq, rationale) in enumerate(priority_detail):
    p_tbl = doc.add_table(rows=5, cols=2)
    p_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_bg(p_tbl.cell(0,0), '1F3564')
    set_cell_bg(p_tbl.cell(0,1), '1F3564')
    # merge header
    p_tbl.cell(0,0).merge(p_tbl.cell(0,1))
    hdr_p = p_tbl.cell(0,0).paragraphs[0]
    hdr_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hdr_run = hdr_p.add_run(f'  {hid}  —  Priority #{i+1}')
    hdr_run.bold = True; hdr_run.font.color.rgb = WHITE; hdr_run.font.size = Pt(11)

    rows_data = [
        ('HMW Statement', stmt),
        ('Stakeholder',   sh),
        ('Linked OQs',    oq),
        ('Priority Reason', rationale),
    ]
    fills = ['FFF3CD', WHITE_HEX, 'FFF3CD', WHITE_HEX]
    for r_idx, (label, val) in enumerate(rows_data):
        set_cell_bg(p_tbl.cell(r_idx+1, 0), fills[r_idx])
        set_cell_bg(p_tbl.cell(r_idx+1, 1), fills[r_idx])
        p_tbl.cell(r_idx+1,0).width = Cm(3.5)
        p_tbl.cell(r_idx+1,1).width = Cm(14.0)
        lbl_run = p_tbl.cell(r_idx+1,0).paragraphs[0].add_run(label)
        lbl_run.bold = True; lbl_run.font.size = Pt(9.5)
        val_run = p_tbl.cell(r_idx+1,1).paragraphs[0].add_run(val)
        val_run.font.size = Pt(9.5)
        if r_idx == 3:
            val_run.italic = True; val_run.font.color.rgb = RGBColor(0x7B,0x36,0x06)
    set_borders(p_tbl)
    doc.add_paragraph()

thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — OQ COVERAGE CHECK
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '5. OQ Coverage Check — Every OQ ↔ ≥1 HMW', 1)
add_para(doc,
    'Per the project rules, every Operational Question from Step 01 must be covered '
    'by at least one HMW statement. The table below verifies this.',
    size=9.5)
doc.add_paragraph()

cov_data = [
    ('OQ-01','Which job roles have the highest posting volume?',         'HMW-01, HMW-12', '✔'),
    ('OQ-02','Average / min / max salary per industry?',                 'HMW-03, HMW-10', '✔'),
    ('OQ-03','Most frequently required skills across postings?',         'HMW-02',         '✔'),
    ('OQ-04','Month-over-month / quarterly hiring volume trend?',        'HMW-01, HMW-05, HMW-11, HMW-12', '✔'),
    ('OQ-05','Cities / states with highest job posting concentration?',  'HMW-07',         '✔'),
    ('OQ-06','Split of job types (remote / hybrid / full-time …)?',     'HMW-08',         '✔'),
    ('OQ-07','Which industries are the top hirers?',                     'HMW-10, HMW-11', '✔'),
    ('OQ-08','Experience level (Fresher / Mid / Senior) demand?',        'HMW-03, HMW-09', '✔'),
    ('OQ-09','Correlation: skill count ↔ offered salary?',              'HMW-06',         '✔'),
    ('OQ-10','Top recruiting companies in a domain / city?',             'HMW-04',         '✔'),
]

cov_tbl = doc.add_table(rows=len(cov_data)+1, cols=4)
cov_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(cov_tbl, 'OQ ID', 'Operational Question', 'Covered by HMW(s)', 'Status')
for idx, row in enumerate(cov_data):
    bg = ALT_BG if idx % 2 == 0 else WHITE_HEX
    fill_row(cov_tbl, idx+1, list(row), bg=bg, bold_col=0)
set_borders(cov_tbl)
add_para(doc, '✔ All 10 OQs are covered by at least one HMW statement. No gap identified.', bold=True, color=GREEN_OK)
doc.add_paragraph()
thin_rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — GOOD vs BAD HMW GUIDE (reference)
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '6. HMW Quality Reference — Good vs. Bad Examples', 1)
gb_tbl = doc.add_table(rows=5, cols=3)
gb_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(gb_tbl, 'Type', 'Example', 'Why It Fails / Works')
gb_data = [
    ('❌ Too Broad',   'How might we build a dashboard?',
     'Prescribes a solution (dashboard) instead of framing the problem.'),
    ('❌ Too Narrow',  'How might we add a blue bar chart for OQ-03?',
     'Locks design into one implementation; leaves no room for creative solutions.'),
    ('❌ Not specific','How might we improve HR processes?',
     'No link to Hiring Trend Analysis; could apply to any HR project.'),
    ('✔ Just Right',  'How might we show recruiters which 10 skills appear most in postings for their industry?',
     'Specific, stakeholder-named, problem-framed, links to OQ-03, not a solution.'),
]
for idx, row in enumerate(gb_data):
    bg = 'FFE5E5' if '❌' in row[0] else 'E5FFE5'
    fill_row(gb_tbl, idx+1, list(row), bg=bg)
set_borders(gb_tbl)
doc.add_paragraph()
thin_rule(doc)

# ── DOCUMENT CONTROL ──────────────────────────────────────────────────────────
add_heading(doc, 'Document Control', 2)
dc_tbl = doc.add_table(rows=2, cols=4)
dc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(dc_tbl, 'Version', 'Date', 'Author', 'Change Summary')
fill_row(dc_tbl, 1,
         ['1.0', datetime.date.today().strftime('%d %b %Y'),
          'Abhayjeet Singh', 'Initial draft – Step 02 complete'])
set_borders(dc_tbl)

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('P071 – Hiring Trend Analysis  |  Step 02 HMW Statements  |  ABES MCA 2nd Yr DA  |  Mentor: Dinesh Sereno')
run.font.size = Pt(8); run.font.color.rgb = RGBColor(0x88,0x88,0x88); run.italic = True

doc.save(OUT_FILE)
print(f"SAVED: {OUT_FILE}")
