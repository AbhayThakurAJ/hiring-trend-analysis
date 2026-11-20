"""
Step 17 – ABES Project Document Generator
Project: P01 – Hiring Trend Analysis
Student: Abhayjeet Singh | Roll: 2500320140004
Mentor: Dinesh Sereno | ABES Engineering College, MCA 2nd Yr DA
"""

import os, datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

OUT_DIR = r"C:\Hiring Trend Analysis\docs"
OUT_DOCX = os.path.join(OUT_DIR, "P01_HTA_Step17_ProjectDocument.docx")
os.makedirs(OUT_DIR, exist_ok=True)

# ── XML & Formatting Helpers ──────────────────────────────────────────────────

CLR_RED = RGBColor(0xA5, 0x1C, 0x30)    # #A51C30 (Primary ABES Burgundy/Red)
CLR_TEXT = RGBColor(0x15, 0x15, 0x14)   # #151514 (Base Dark Grey)
CLR_GRAY = RGBColor(0x7F, 0x7F, 0x7F)   # #7F7F7F (Muted Gray for notes/headers)

HEX_RED = "A51C30"
HEX_BORDER = "BFBFBF"

def set_cell_margins_and_border(cell, top=60, bottom=60, left=90, right=90, border_color=HEX_BORDER):
    """Set cell margins in dxa and thin border."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # Margins
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)
    
    # Borders (0.5 pt = sz 4)
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), border_color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def set_cell_shading(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def apply_table_styles(table, col_widths):
    """
    Apply ABES standard formatting to table:
    - fixed layout
    - no row splitting across pages
    - repeat header on every page
    - set column widths
    """
    tblPr = table._tbl.tblPr
    # Fixed layout
    layout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    tblPr.append(layout)
    
    # Left alignment matching page text margin
    tblCellMar = parse_xml(f'''
        <w:tblCellMar {nsdecls("w")}>
            <w:top w:w="60" w:type="dxa"/>
            <w:bottom w:w="60" w:type="dxa"/>
            <w:left w:w="90" w:type="dxa"/>
            <w:right w:w="90" w:type="dxa"/>
        </w:tblCellMar>
    ''')
    tblPr.append(tblCellMar)
    
    # Apply row rules and column widths
    for r_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        # Row cantSplit
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        # Header row rules
        if r_idx == 0:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths[c_idx]
            # Set explicit cell width in dxa
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{col_widths[c_idx]}" w:type="dxa"/>')
            tcPr.append(tcW)
            
            if r_idx == 0:
                set_cell_shading(cell, HEX_RED)
                set_cell_margins_and_border(cell, 80, 80, 90, 90)
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        r.font.name = 'Calibri'
                        r.font.size = Pt(9)
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            else:
                set_cell_margins_and_border(cell, 60, 60, 90, 90)
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(1)
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.line_spacing = 1.0
                    for r in p.runs:
                        r.font.name = 'Calibri'
                        r.font.size = Pt(9)
                        r.font.bold = False
                        r.font.italic = False
                        r.font.color.rgb = CLR_TEXT

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Georgia'
    run.font.size = Pt(15)
    run.font.bold = True
    run.font.color.rgb = CLR_RED
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Georgia'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = CLR_TEXT
    return p

def add_body_p(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(10.5)
    run.font.color.rgb = CLR_TEXT
    return p

def add_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(5)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(9.5)
    run.font.italic = True
    run.font.color.rgb = CLR_GRAY
    return p

def add_bullet(doc, text, bold_prefix=""):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.95)   # 540 dxa
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Calibri'
        r_pre.font.size = Pt(10.5)
        r_pre.font.bold = True
        r_pre.font.color.rgb = CLR_TEXT
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(10.5)
    run.font.color.rgb = CLR_TEXT
    return p

def add_label(doc, text, color=CLR_TEXT, before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.color.rgb = color
    return p

def add_code(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(text)
    run.font.name = 'Consolas'
    run.font.size = Pt(9)
    run.font.color.rgb = CLR_TEXT
    return p

def add_header_footer(doc):
    for s in doc.sections:
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(2.0)
        s.right_margin = Cm(2.0)
        s.header_distance = Cm(1.25)
        s.footer_distance = Cm(1.25)
        
        # Header: Right, 8 pt 7F7F7F
        header = s.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("ABES Engineering College | Project Document | P01")
        hrun.font.name = 'Calibri'
        hrun.font.size = Pt(8)
        hrun.font.color.rgb = CLR_GRAY
        
        # Footer: Center, 8 pt 7F7F7F
        footer = s.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun1 = fp.add_run("Abhayjeet Singh | 2500320140004     Page ")
        frun1.font.name = 'Calibri'
        frun1.font.size = Pt(8)
        frun1.font.color.rgb = CLR_GRAY
        
        # Add live page field
        fld_page = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        fp._p.append(fld_page)

# ═════════════════════════════════════════════════════════════════════════════
# BUILD DOCUMENT
# ═════════════════════════════════════════════════════════════════════════════

doc = Document()
add_header_footer(doc)

# ── PAGE 1: COVER ─────────────────────────────────────────────────────────────
# 1. HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad (bold 10 A51C30)
p_inst = doc.add_paragraph()
p_inst.paragraph_format.space_before = Pt(0)
p_inst.paragraph_format.space_after = Pt(5)
r = p_inst.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
r.font.name = 'Calibri'; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = CLR_RED

# 2. [Insert ABES logo here] (italic 9 pt 7F7F7F)
p_logo = doc.add_paragraph()
p_logo.paragraph_format.space_before = Pt(0)
p_logo.paragraph_format.space_after = Pt(5)
r = p_logo.add_run("[Insert ABES logo here]")
r.font.name = 'Calibri'; r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = CLR_GRAY

# 3. PROJECT DOCUMENT (Georgia 26 bold, before 70 pt, after 6 pt)
p_pdoc = doc.add_paragraph()
p_pdoc.paragraph_format.space_before = Pt(70)
p_pdoc.paragraph_format.space_after = Pt(6)
r = p_pdoc.add_run("PROJECT DOCUMENT")
r.font.name = 'Georgia'; r.font.size = Pt(26); r.font.bold = True; r.font.color.rgb = CLR_TEXT

# 4. <Project ID> - <Title> (Georgia 18 A51C30, after 12 pt, bottom border 1.5 pt A51C30)
p_title = doc.add_paragraph()
p_title.paragraph_format.space_before = Pt(0)
p_title.paragraph_format.space_after = Pt(12)
r = p_title.add_run("P01 - Hiring Trend Analysis")
r.font.name = 'Georgia'; r.font.size = Pt(18); r.font.color.rgb = CLR_RED

# Add bottom border 1.5 pt (sz 12) A51C30
pPr = p_title._p.get_or_add_pPr()
pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="{HEX_RED}"/></w:pBdr>')
pPr.append(pBdr)

# 5. Subtitle description
p_sub = doc.add_paragraph()
p_sub.paragraph_format.space_before = Pt(0)
p_sub.paragraph_format.space_after = Pt(24)
r = p_sub.add_run("Description · HMW · Stakeholders · Traceability · User Stories · BRD · TRD · Database Design · Test Cases · UI Framework")
r.font.name = 'Calibri'; r.font.size = Pt(10); r.font.color.rgb = CLR_GRAY

# 6. Cover Metadata Table (Field 2900 | Details 6738 = 9638 dxa)
cover_meta = [
    ("Project ID & Title", "P01 - Hiring Trend Analysis"),
    ("Domain", "Data Analytics"),
    ("Programme / Batch", "MCA 2nd Year Data Analyst - ABES_OND26_MCA Batch 01"),
    ("Learner(s) & Roll No(s)", "Abhayjeet Singh (Roll No: 2500320140004)"),
    ("Mentor", "Dinesh Sereno"),
    ("GitHub repository", "https://github.com/abhayjeet-singh/2500320140004-hiringtrendanalysis"),
    ("Live app (Streamlit)", "https://2500320140004-hiringtrendanalysis.streamlit.app/"),
    ("Version & date", f"v1.0 - {datetime.date.today().strftime('%d/%m/%Y')}")
]
tbl_cover = doc.add_table(rows=len(cover_meta)+1, cols=2)
tbl_cover.alignment = WD_TABLE_ALIGNMENT.LEFT
tbl_cover.rows[0].cells[0].paragraphs[0].text = "Field"
tbl_cover.rows[0].cells[1].paragraphs[0].text = "Details"
for i, (k, v) in enumerate(cover_meta, 1):
    tbl_cover.rows[i].cells[0].paragraphs[0].text = k
    tbl_cover.rows[i].cells[1].paragraphs[0].text = v
apply_table_styles(tbl_cover, [2900, 6738])

# Break to Page 2
doc.add_page_break()

# ── PAGE 2: TABLE OF CONTENTS ────────────────────────────────────────────────
# Word TOC Field: "TOC \h \o "1-1"" with updateFields on open
p_toc = doc.add_paragraph()
p_toc.paragraph_format.space_before = Pt(0)
p_toc.paragraph_format.space_after = Pt(0)
fld_toc = parse_xml(r'''
    <w:p %s>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \h \o "1-1" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
''' % nsdecls('w'))
p_toc._p.addnext(fld_toc)

# Instruct Word to update fields on document open
settings = doc.settings._element
settings.append(parse_xml(r'<w:updateFields %s w:val="true"/>' % nsdecls('w')))

# Break to Section 1
doc.add_page_break()

# ── 1. PROJECT DESCRIPTION ───────────────────────────────────────────────────
add_heading_1(doc, "1. Project Description")

desc_rows = [
    ("Overview", "The Hiring Trend Analysis platform is an end-to-end Python, Streamlit, and SQLite analytical application engineered to eliminate fragmented recruitment spreadsheets and intuition-driven decision-making. Operating on a structured database of 652 Indian recruitment postings, the application delivers real-time intelligence into role demand spikes, compensation benchmarks, skill penetration, urban talent concentration, and seasonal timeline cycles."),
    ("Problem statement", "HR professionals and job seekers lack a unified, data-driven analytical tool to evaluate hiring demand, benchmark salaries, identify in-demand skills, and track geographic concentration, forcing reliance on fragmented spreadsheets that are error-prone, slow, and incapable of surfacing actionable talent intelligence."),
    ("Intended users", "Talent Acquisition Managers (weekly resource planning), Technical Recruiters (daily JD authoring and salary benchmarking), Job Seekers (career upskilling and market compensation verification), and Academic Placement Officers (targeting recruiting employers for campus drives)."),
    ("In scope", "Interactive EDA dashboard, job role demand rankings (BR-01), skill demand analysis (BR-02), salary benchmarking by industry and location (BR-03, BR-06), company hiring leaderboards (BR-04), time-series hiring volume (BR-05, BR-12), city concentration analysis (BR-07), experience breakdown (BR-08), skill-salary OLS regression (BR-09), work mode distribution (BR-10), EDA data quality audit (BR-13), SQLite relational storage, automated pytest test suite, and cloud deployment."),
    ("Out of scope", "Real-time web scraping from live job portals (LinkedIn/Naukri), user login/authentication, resume parsing or NLP matching, candidate ATS workflow tracking, and mobile native app development."),
    ("Assumptions", "Synthetic dataset generated via Python Faker realistically mirrors Indian tech and business sectors; SQLite is fully performant for analytical workloads under 10,000 rows; Streamlit Community Cloud free tier provides sufficient uptime for academic demonstration."),
    ("Dataset", "652 synthetic job posting records stored across 5 normalized relational tables in SQLite (app.db) and an analytical view (vw_hiring_analysis) covering 2022 to 2024 across 15 Indian metropolitan cities and 8 industry sectors.")
]

tbl_desc = doc.add_table(rows=len(desc_rows)+1, cols=2)
tbl_desc.rows[0].cells[0].paragraphs[0].text = "Item"
tbl_desc.rows[0].cells[1].paragraphs[0].text = "Description"
for i, (item, val) in enumerate(desc_rows, 1):
    tbl_desc.rows[i].cells[0].paragraphs[0].text = item
    tbl_desc.rows[i].cells[1].paragraphs[0].text = val
apply_table_styles(tbl_desc, [2400, 7238])

# ── 2. HOW MIGHT WE STATEMENTS ───────────────────────────────────────────────
add_heading_1(doc, "2. How Might We (HMW) Statements")
add_note(doc, "Format: How might we <action> for <user> so that <outcome>?")

hmw_rows = [
    ("HMW-01", "How might we show an HR manager which specific job roles have spiked or dropped in posting volume over a chosen 3-month window, without opening a single spreadsheet?"),
    ("HMW-02", "How might we display the exact top-10 skills demanded across a selected industry so a recruiter can write a data-driven job description in under 2 minutes?"),
    ("HMW-03", "How might we let a job seeker instantly compare their salary expectation against the actual average offered for their experience level and target city?"),
    ("HMW-04", "How might we give a placement officer a ranked list of the top-10 hiring companies in a selected domain so they can prioritise campus drive invitations with data?"),
    ("HMW-05", "How might we reveal the month-over-month change in total hiring volume on a single interactive timeline so an HR manager can align campaign launches with demand peaks?"),
    ("HMW-06", "How might we surface the statistical correlation between the number of skills a role requires and its maximum offered salary, so a recruiter can price complex roles objectively?"),
    ("HMW-07", "How might we display a city-wise hiring concentration chart so that a job seeker relocating to a new city can compare opportunity density across the top 15 locations?"),
    ("HMW-08", "How might we show the job-type distribution (remote / hybrid / full-time / contract) for each industry so a recruiter can benchmark their flexible-work policy against market norms?"),
    ("HMW-09", "How might we break down experience-level demand (Fresher / Mid / Senior) by job role so a placement officer can match the right student batch to the right recruiter?"),
    ("HMW-10", "How might we provide HR managers a single-screen industry salary comparison — showing average, min, and max — so budget approvals no longer depend on third-party surveys?"),
    ("HMW-11", "How might we identify which industries show the fastest growth in hiring volume for a selected skill set, so a career counsellor can redirect students toward emerging sectors?"),
    ("HMW-12", "How might we show a time-series view of posting frequency for a specific job role so a job seeker can time their applications during peak hiring windows?")
]

tbl_hmw = doc.add_table(rows=len(hmw_rows)+1, cols=2)
tbl_hmw.rows[0].cells[0].paragraphs[0].text = "ID"
tbl_hmw.rows[0].cells[1].paragraphs[0].text = "HMW statement"
for i, (hid, hstmt) in enumerate(hmw_rows, 1):
    tbl_hmw.rows[i].cells[0].paragraphs[0].text = hid
    tbl_hmw.rows[i].cells[1].paragraphs[0].text = hstmt
apply_table_styles(tbl_hmw, [1300, 8338])

# ── 3. STAKEHOLDERS ──────────────────────────────────────────────────────────
add_heading_1(doc, "3. Stakeholders")

sh_rows = [
    ("SH-01", "Dinesh Sereno", "Academic Mentor & Reviewer", "Methodological rigor, project quality, and academic assessment", "High"),
    ("SH-02", "Abhayjeet Singh", "Student Developer & Data Analyst", "Learning, technical delivery, code quality, and end-to-end execution", "High"),
    ("SH-03", "HR Manager (Persona)", "Primary Business Decision-Maker", "Workforce planning, budgeting, identifying role demand spikes", "High"),
    ("SH-04", "Recruiter (Persona)", "Operational Talent Acquisition", "Writing job descriptions, benchmarking skill demand, competitive salary bands", "Medium"),
    ("SH-05", "Job Seeker (Persona)", "End User Candidate", "Salary benchmarking, in-demand technical skills, geographic job density", "Low"),
    ("SH-06", "Placement Officer (Persona)", "Institutional Placement Facilitator", "Matching student cohort skills to industry hiring demand, company drives", "Medium")
]

tbl_sh = doc.add_table(rows=len(sh_rows)+1, cols=5)
tbl_sh.rows[0].cells[0].paragraphs[0].text = "ID"
tbl_sh.rows[0].cells[1].paragraphs[0].text = "Stakeholder"
tbl_sh.rows[0].cells[2].paragraphs[0].text = "Role"
tbl_sh.rows[0].cells[3].paragraphs[0].text = "Interest"
tbl_sh.rows[0].cells[4].paragraphs[0].text = "Influence"
for i, row in enumerate(sh_rows, 1):
    for c_i, val in enumerate(row):
        tbl_sh.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_sh, [1000, 2400, 1900, 3138, 1200])

# ── 4. TRACEABILITY MATRIX ───────────────────────────────────────────────────
add_heading_1(doc, "4. Traceability Matrix")
add_note(doc, "One row per requirement. Priority uses MOSCOW: Must / Should / Could / Won't (this time).")

rtm_rows = [
    ("HMW-01", "BR-01", "Rank job roles by volume (date filter)", "SH-03", "FR-03", "Must", "US-05", "TC-F-01, TC-T-07", "UI-01, UI-02"),
    ("HMW-02", "BR-02", "Rank skills frequency by industry", "SH-04", "FR-05", "Must", "US-06", "TC-F-02, TC-T-02", "UI-03"),
    ("HMW-03", "BR-03", "Salary stats by experience and city", "SH-05", "FR-04", "Must", "US-09", "TC-F-03, TC-T-08", "UI-01, UI-04"),
    ("HMW-04", "BR-04", "Company hiring leaderboard by domain", "SH-06", "FR-12", "Must", "US-16", "TC-F-04, TC-T-02", "UI-08"),
    ("HMW-05", "BR-05", "Monthly hiring volume time series", "SH-03", "FR-06", "Must", "US-13", "TC-F-05, TC-T-10", "UI-01, UI-06"),
    ("HMW-10", "BR-06", "Cross-industry salary comparison", "SH-03", "FR-04", "Must", "US-10", "TC-F-06, TC-T-08", "UI-04"),
    ("HMW-07", "BR-07", "City posting concentration ranking", "SH-05", "FR-07", "Must", "US-12", "TC-F-07, TC-T-09", "UI-01, UI-05"),
    ("HMW-09", "BR-08", "Experience-level demand per role", "SH-06", "FR-10", "Must", "US-07", "TC-F-08, TC-T-02", "UI-10"),
    ("HMW-06", "BR-09", "Skill count vs salary correlation (OLS)", "SH-04", "FR-11", "Should", "US-11", "TC-F-09, TC-T-12", "UI-07"),
    ("HMW-08", "BR-10", "Job-type distribution per industry", "SH-04", "FR-08", "Should", "US-15", "TC-F-10, TC-T-11", "UI-10"),
    ("HMW-11", "BR-11", "Industry hiring volume by skill keyword", "SH-06", "FR-09", "Should", "US-08", "TC-F-11, TC-T-02", "UI-10"),
    ("HMW-12", "BR-12", "Role-level posting time series", "SH-05", "FR-06", "Should", "US-14", "TC-F-12, TC-T-10", "UI-06"),
    ("Assumed", "BR-13", "EDA summary statistics and quality audit", "SH-02", "FR-13", "Could", "US-01, 02", "TC-F-13, TC-T-01", "UI-09")
]

tbl_rtm = doc.add_table(rows=len(rtm_rows)+1, cols=9)
headers_rtm = ["HMW", "Req ID", "Requirement", "Stakeholder", "Feature", "Priority", "User Story", "Test Case", "UI Screen"]
for c_i, h in enumerate(headers_rtm):
    tbl_rtm.rows[0].cells[c_i].paragraphs[0].text = h
for i, row in enumerate(rtm_rows, 1):
    for c_i, val in enumerate(row):
        tbl_rtm.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_rtm, [894, 742, 1964, 1083, 1143, 798, 842, 1132, 1040])

# ── 5. USER STORIES ──────────────────────────────────────────────────────────
add_heading_1(doc, "5. User Stories")

us_rows = [
    ("US-01", "As a Data Analyst, I want to automatically load and validate the SQLite dataset with in-memory caching so that dashboard queries respond in under 2 seconds.",
     "Given the Streamlit app starts, When data loading is triggered, Then the dataset is read from data/app.db, validated, and cached.", "Must"),
    ("US-02", "As a Data Analyst, I want an EDA summary panel showing missing values and outliers so that I can audit data quality before interpreting charts.",
     "Given the user is on the EDA page, When audit calculations execute, Then descriptive statistics, missing counts, and IQR outlier metrics are displayed.", "Could"),
    ("US-03", "As an HR Manager, I want high-level KPI cards for Total Postings, Average Salary, Top City, and Top Role so that I can assess talent market health at a glance.",
     "Given the user opens the Overview dashboard, When the page renders, Then 4 responsive Tailwind KPI cards display macro summary metrics.", "Must"),
    ("US-04", "As an HR Recruiter, I want global sidebar filters for date range, industry, city, and experience level so that all charts dynamically synchronize.",
     "Given the user adjusts sidebar filters, When state updates, Then all active page visualizations re-render based strictly on the selected cohort.", "Must"),
    ("US-05", "As an HR Manager, I want a ranked horizontal bar chart of job roles by volume so that I can identify high-demand talent positions.",
     "Given the user views the Role Demand page, When filters are applied, Then top 10 roles are rendered in descending volume order.", "Must"),
    ("US-06", "As a Technical Recruiter, I want to see the top-10 in-demand skills per industry so that I can author market-competitive job descriptions.",
     "Given an industry is selected, When the skill breakdown executes, Then the top 10 skills appear in a frequency bar chart.", "Must"),
    ("US-07", "As a Placement Officer, I want to see experience level splits (Fresher vs Mid vs Senior) per role so that I know which student batches match demand.",
     "Given the Job Types & Exp page is active, When experience distribution renders, Then a 100% stacked bar chart displays seniority ratios per role.", "Must"),
    ("US-08", "As a Career Counselor, I want to rank industries by hiring for a specific skill keyword so that I can guide student upskilling.",
     "Given a skill keyword is selected, When industry aggregation runs, Then a ranked bar chart shows leading sectors requiring that competency.", "Should"),
    ("US-09", "As a Job Seeker, I want to view min, mean, and max salaries by experience tier and city so that I can benchmark fair compensation expectations.",
     "Given experience and city filters are chosen, When salary stats evaluate, Then grouped bar charts and box plots reveal the compensation spread.", "Must"),
    ("US-10", "As an HR Manager, I want a single cross-industry salary comparison view so that I can budget compensation packages competitively.",
     "Given the user views the industry salary tab, When the chart displays, Then all 8 industries are displayed sorted by average salary.", "Must"),
    ("US-11", "As a Technical Recruiter, I want an OLS regression scatter plot of skills count vs salary so that I can justify offers for multi-skilled roles.",
     "Given the user accesses the Correlation page, When the regression model runs, Then a scatter plot with OLS trendline, Pearson r, and p-value renders.", "Should"),
    ("US-12", "As a Job Seeker, I want to view job posting concentration ranked by city so that I know which metropolitan areas offer greatest opportunity density.",
     "Given the Geography screen is active, When city aggregates are computed, Then a ranked bar chart displays postings across all 15 Indian cities.", "Must"),
    ("US-13", "As an HR Manager, I want an interactive monthly hiring timeline so that I can anticipate seasonal hiring peaks and time drives effectively.",
     "Given the Hiring Trends page is open, When time-series aggregation runs, Then a chronological monthly line chart displays hiring trajectory.", "Must"),
    ("US-14", "As a Job Seeker, I want to track monthly hiring trends for a specific role so that I can apply during historical recruitment surges.",
     "Given a single job role is selected, When the timeline renders, Then an isolated monthly trendline shows seasonal hiring patterns for that role.", "Should"),
    ("US-15", "As a Technical Recruiter, I want to see work mode distributions (Remote, Hybrid, Full-time) so that I can benchmark flexible-work policies.",
     "Given the work modes page is active, When the proportion is evaluated, Then a donut chart displays percentages of remote, hybrid, and onsite jobs.", "Should"),
    ("US-16", "As a Placement Officer, I want a ranked company leaderboard in a domain so that I can prioritize campus recruitment invitations.",
     "Given an industry domain is selected, When company aggregates generate, Then the top 10 recruiting firms appear in a ranked chart and table.", "Must")
]

tbl_us = doc.add_table(rows=len(us_rows)+1, cols=4)
tbl_us.rows[0].cells[0].paragraphs[0].text = "US-ID"
tbl_us.rows[0].cells[1].paragraphs[0].text = "User story"
tbl_us.rows[0].cells[2].paragraphs[0].text = "Acceptance criteria (Given / When / Then)"
tbl_us.rows[0].cells[3].paragraphs[0].text = "Priority"
for i, row in enumerate(us_rows, 1):
    for c_i, val in enumerate(row):
        tbl_us.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_us, [900, 3400, 4138, 1200])

# ── 6. BUSINESS REQUIREMENTS DOCUMENT (BRD) ──────────────────────────────────
add_heading_1(doc, "6. Business Requirements Document (BRD)")

add_heading_2(doc, "6.1 Business objectives")
add_bullet(doc, "Deliver an intuitive, interactive data analytics dashboard that consolidates fragmented recruitment metrics into real-time visual insights.", "Centralize Talent Intelligence: ")
add_bullet(doc, "Enable HR managers and recruiters to benchmark compensation bands across industries, experience levels, and cities without purchasing expensive third-party surveys.", "Market Compensation Benchmarking: ")
add_bullet(doc, "Surface high-demand technical competencies from real job descriptions to guide candidate upskilling and university curriculum design.", "Skill Demand Forecasting: ")
add_bullet(doc, "Provide longitudinal seasonal hiring trend visibility (2022–2024) to optimize the scheduling of corporate talent acquisition drives and campus recruitment.", "Longitudinal Timeline Optimization: ")
add_bullet(doc, "Deliver a zero-cost, open-source academic capstone project meeting all ABES MCA curriculum requirements and deployable on free cloud infrastructure.", "Open Source & Academic Excellence: ")

add_heading_2(doc, "6.2 Success KPIs")
kpi_rows = [
    ("OQ Resolution Rate", "Fully Answered Operational Questions (10)", "Total Operational Questions (10)", "Project Completion", "100% (10/10 functional screens)"),
    ("Dataset Integrity", "Verified Synthetic Database Records (652)", "Minimum Prescribed Records (500)", "Database Generation", "130.4% (Target surpassed)"),
    ("Test Pass Rate", "Passing Technical Unit Tests (15)", "Total Executed Unit Tests (15)", "Continuous Integration", "100% (0 errors, 0 failures)"),
    ("Application Load Latency", "Measured Dashboard Load Time (seconds)", "Benchmark Threshold (3.0 seconds)", "Cloud Production", "< 3.0 seconds on broadband"),
    ("Query Execution Latency", "Analytical View Query Time (seconds)", "Latency Target (1.0 second)", "In-Memory Run", "0.012 s (< 1.0 s SLA)")
]
tbl_kpi = doc.add_table(rows=len(kpi_rows)+1, cols=5)
tbl_kpi.rows[0].cells[0].paragraphs[0].text = "KPI"
tbl_kpi.rows[0].cells[1].paragraphs[0].text = "Numerator"
tbl_kpi.rows[0].cells[2].paragraphs[0].text = "Denominator"
tbl_kpi.rows[0].cells[3].paragraphs[0].text = "Period"
tbl_kpi.rows[0].cells[4].paragraphs[0].text = "Target / note"
for i, row in enumerate(kpi_rows, 1):
    for c_i, val in enumerate(row):
        tbl_kpi.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_kpi, [1900, 2100, 2100, 1700, 1838])

add_heading_2(doc, "6.3 Business requirements")
br_rows = [
    ("BR-01", "The business needs to rank job roles by posting volume for any user-selected time period so that HR managers can identify surging or declining roles.", "Must", "SH-03 (HR Manager)"),
    ("BR-02", "The business needs to rank skills by frequency across job postings within a chosen industry so that recruiters can author data-backed job descriptions.", "Must", "SH-04 (Recruiter)"),
    ("BR-03", "The business needs to present average, min, and max salary figures segmented by experience tier and city so that job seekers can verify market expectations.", "Must", "SH-05 (Job Seeker)"),
    ("BR-04", "The business needs to rank enterprise companies by hiring volume within a selected domain so that placement officers can prioritize campus drive invitations.", "Must", "SH-06 (Placement Officer)"),
    ("BR-05", "The business needs to display aggregate job posting volume by calendar month so that HR managers can align campaign launches with seasonal demand peaks.", "Must", "SH-03 (HR Manager)"),
    ("BR-06", "The business needs to compare salary figures across all industries in a consolidated view so that HR managers can benchmark compensation without paid surveys.", "Must", "SH-03 (HR Manager)"),
    ("BR-07", "The business needs to present job posting counts ranked by city so that job seekers and recruiters can identify geographic hiring hotspots at a glance.", "Must", "SH-05 (Job Seeker)"),
    ("BR-08", "The business needs to show experience-level demand proportions (Fresher/Mid/Senior) per job role so that placement officers match student profiles to demand.", "Must", "SH-06 (Placement Officer)"),
    ("BR-09", "The business needs to quantify the correlation between skills count and salary so that recruiters can price complex, multi-skilled roles competitively.", "Should", "SH-04 (Recruiter)"),
    ("BR-10", "The business needs to present job-type distributions (Full-time, Remote, Hybrid) per industry so that recruiters can benchmark flexible-work policies.", "Should", "SH-04 (Recruiter)"),
    ("BR-11", "The business needs to rank industries by hiring volume for a user-selected skill keyword so that career counselors can direct students to emerging sectors.", "Should", "SH-06 (Placement Officer)"),
    ("BR-12", "The business needs to display a time-series view of posting frequency for a single selected role so that job seekers can time their applications effectively.", "Should", "SH-05 (Job Seeker)"),
    ("BR-13", "The business needs summary statistics, missing counts, and outlier flags for numeric fields so that analysts can assess data quality before interpreting charts.", "Could", "SH-02 (Data Analyst)")
]
tbl_br = doc.add_table(rows=len(br_rows)+1, cols=4)
tbl_br.rows[0].cells[0].paragraphs[0].text = "BR-ID"
tbl_br.rows[0].cells[1].paragraphs[0].text = "Requirement"
tbl_br.rows[0].cells[2].paragraphs[0].text = "Priority"
tbl_br.rows[0].cells[3].paragraphs[0].text = "Source (stakeholder)"
for i, row in enumerate(br_rows, 1):
    for c_i, val in enumerate(row):
        tbl_br.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_br, [1100, 5338, 1300, 1900])

add_heading_2(doc, "6.4 Constraints and out of scope")
add_bullet(doc, "Strict zero-budget constraint requiring all development frameworks, data storage, and cloud hosting to be 100% free-tier open source.", "Financial Budget Constraint: ")
add_bullet(doc, "Strict adherence to privacy compliance mandating 100% synthetic data generation with zero real PII or proprietary organizational data.", "Data Privacy & Governance: ")
add_bullet(doc, "Academic capstone timeline requiring completion across all 18 structured SDLC stages within a single semester.", "Timeline & Academic Deadlines: ")
add_bullet(doc, "Real-time web scraping from live commercial job boards (LinkedIn, Indeed, Naukri) is excluded to prevent bot detection and IP blocking.", "Out of Scope (Web Scraping): ")
add_bullet(doc, "User authentication, role-based access control, applicant resume parsing, and transactional application tracking (ATS) are excluded.", "Out of Scope (ATS & Auth): ")

# ── 7. TECHNICAL REQUIREMENTS DOCUMENT (TRD) ──────────────────────────────────
add_heading_1(doc, "7. Technical Requirements Document (TRD)")

add_heading_2(doc, "7.1 Technology stack (all free)")
stack_rows = [
    ("Language", "Python 3.12", "Modern, stable runtime with high performance and universal data analytics package support."),
    ("Data processing", "pandas 2.2.2 & numpy 1.26.4", "Vectorized in-memory data manipulation, datetime parsing, and statistical aggregations."),
    ("Storage", "SQLite 3 (app.db)", "Zero-configuration, serverless relational database engine stored directly within the repository."),
    ("Charts", "Plotly Express 5.22.0", "Interactive, client-side rendered HTML5 charts supporting zoom, hover tooltips, and exports."),
    ("App / UI", "Streamlit 1.35.0 & Tailwind CSS", "Rapid multipage web framework enhanced with modern responsive card styling via Tailwind CDN."),
    ("Testing", "pytest 8.2.2", "Industry standard automated unit testing framework supporting fixture injection and assertions."),
    ("Version control", "Git & GitHub", "Distributed source code management, commit traceability, and automated cloud CI/CD deployment."),
    ("Editor", "Visual Studio Code / Cursor", "Extensible development environment with integrated Python debugging and terminal tooling.")
]
tbl_stack = doc.add_table(rows=len(stack_rows)+1, cols=3)
tbl_stack.rows[0].cells[0].paragraphs[0].text = "Layer"
tbl_stack.rows[0].cells[1].paragraphs[0].text = "Tool"
tbl_stack.rows[0].cells[2].paragraphs[0].text = "Why"
for i, row in enumerate(stack_rows, 1):
    for c_i, val in enumerate(row):
        tbl_stack.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_stack, [2000, 3800, 3838])

add_heading_2(doc, "7.2 Functional requirements")
fr_rows = [
    ("FR-01", "Load dataset from SQLite into pandas DataFrame using @st.cache_data for instant retrieval.", "BR-01..13 / US-01"),
    ("FR-02", "Validate loaded dataset against business constraints; halt execution gracefully on schema breach.", "BR-R-01..10 / US-01"),
    ("FR-03", "Render interactive ranked horizontal bar chart of job posting volume by role for date range.", "BR-01 / US-05"),
    ("FR-04", "Render grouped bar chart and box plots of salary statistics across industries and cities.", "BR-03, 06 / US-09, 10"),
    ("FR-05", "Explode pipe-delimited skills and render horizontal frequency bar chart of top-N skills.", "BR-02 / US-06"),
    ("FR-06", "Render chronological time-series line chart of hiring volumes over monthly intervals.", "BR-05, 12 / US-13, 14"),
    ("FR-07", "Aggregate postings by city and render geographic concentration bar chart across metros.", "BR-07 / US-12"),
    ("FR-08", "Compute job-type proportions per industry and render interactive donut chart.", "BR-10 / US-15"),
    ("FR-09", "Filter postings by skill keyword and rank industries hiring most heavily for that skill.", "BR-11 / US-08"),
    ("FR-10", "Pivot experience levels per role and render a 100% stacked bar chart.", "BR-08 / US-07"),
    ("FR-11", "Derive skills count, evaluate Pearson r, and render OLS linear regression scatter plot.", "BR-09 / US-11"),
    ("FR-12", "Rank companies by posting volume in a selected domain and render leaderboard with table.", "BR-04 / US-16"),
    ("FR-13", "Render EDA summary statistics table, missing value counts, and IQR outlier identification.", "BR-13 / US-02"),
    ("FR-14", "Provide global sidebar filter widgets updating st.session_state across all pages.", "BR-01..13 / US-04"),
    ("FR-15", "Render responsive Tailwind hero banners and 4 KPI summary cards on the Overview screen.", "BR-01, 05, 07 / US-03"),
    ("FR-16", "Provide sanitized CSV download buttons on all analysis screens for filtered subsets.", "BR-01..13 / US-16")
]
tbl_fr = doc.add_table(rows=len(fr_rows)+1, cols=3)
tbl_fr.rows[0].cells[0].paragraphs[0].text = "FR-ID"
tbl_fr.rows[0].cells[1].paragraphs[0].text = "Requirement"
tbl_fr.rows[0].cells[2].paragraphs[0].text = "Linked BR / US"
for i, row in enumerate(fr_rows, 1):
    for c_i, val in enumerate(row):
        tbl_fr.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_fr, [1100, 6338, 2200])

add_heading_2(doc, "7.3 Non-functional requirements")
nfr_rows = [
    ("NFR-01", "Performance: Initial page load must complete in under 3.0 seconds; cached chart re-renders under 1.0 second."),
    ("NFR-02", "Reliability: Automated pytest test suite must achieve >= 90% pass rate across all data and logic modules."),
    ("NFR-03", "Security & Privacy: 100% synthetic dataset with zero real personal data or committed credentials."),
    ("NFR-04", "Usability: Clean responsive UI adhering to ABES typography and Tailwind CSS visual card hierarchy."),
    ("NFR-05", "Portability: Dynamic path resolution ensuring seamless cross-platform execution on Windows and Linux cloud hosts.")
]
tbl_nfr = doc.add_table(rows=len(nfr_rows)+1, cols=2)
tbl_nfr.rows[0].cells[0].paragraphs[0].text = "NFR-ID"
tbl_nfr.rows[0].cells[1].paragraphs[0].text = "Requirement"
for i, row in enumerate(nfr_rows, 1):
    for c_i, val in enumerate(row):
        tbl_nfr.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_nfr, [1100, 8538])

add_heading_2(doc, "7.4 Architecture and deployment")

p_arch_line = doc.add_paragraph()
p_arch_line.paragraph_format.space_before = Pt(2)
p_arch_line.paragraph_format.space_after = Pt(5)
r = p_arch_line.add_run("Raw data (data/raw) -> src/cleaner.py -> SQLite (data/app.db) -> Streamlit app (app.py) -> user")
r.font.name = 'Calibri'; r.font.size = Pt(10.5); r.font.bold = True; r.font.color.rgb = CLR_RED

stage_rows = [
    ("1. Local Development", "Run 'python src/generate_data.py' to seed SQLite database; execute 'streamlit run app.py' to launch local web server.", "App accessible at http://localhost:8501 with all 10 pages functioning."),
    ("2. GitHub Repository", "Initialize Git repository, commit all src/, data/app.db, and test files; push to main branch on GitHub.", "Repository public with zero committed secrets and clean .gitignore."),
    ("3. Streamlit Community Cloud", "Link GitHub repo at share.streamlit.io; set main file to app.py; deploy container on Python 3.12.", "Live public URL active with sub-second response times and 100% uptime.")
]
tbl_stage = doc.add_table(rows=len(stage_rows)+1, cols=3)
tbl_stage.rows[0].cells[0].paragraphs[0].text = "Stage"
tbl_stage.rows[0].cells[1].paragraphs[0].text = "Command / action"
tbl_stage.rows[0].cells[2].paragraphs[0].text = "Done when"
for i, row in enumerate(stage_rows, 1):
    for c_i, val in enumerate(row):
        tbl_stage.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_stage, [2200, 4838, 2600])

# ── 8. DATABASE DESIGN ───────────────────────────────────────────────────────
add_heading_1(doc, "8. Database Design")
add_note(doc, "Storage used: SQLite (data/app.db). Raw data stays unchanged; cleaned data is stored here.")

db_rows = [
    ("dim_industry", "industry_id", "INTEGER", "PK", "Unique industry identifier (1 to 8)"),
    ("dim_industry", "industry_name", "TEXT", "UNIQUE", "Industry sector name (e.g. IT & Software)"),
    ("dim_industry", "sector_category", "TEXT", "NOT NULL", "Broad sector classification (Services, Production)"),
    ("dim_city", "city_id", "INTEGER", "PK", "Unique geographic city identifier (1 to 15)"),
    ("dim_city", "city_name", "TEXT", "UNIQUE", "City name across major Indian metropolitan areas"),
    ("dim_city", "state_name", "TEXT", "NOT NULL", "Indian State or Union Territory name"),
    ("dim_city", "region", "TEXT", "NOT NULL", "Geographic zone (North, South, West, East, Central)"),
    ("dim_job_role", "role_id", "INTEGER", "PK", "Standardized job role identifier (1 to 12)"),
    ("dim_job_role", "role_title", "TEXT", "UNIQUE", "Standard job role title (e.g. Data Analyst)"),
    ("dim_job_role", "role_category", "TEXT", "NOT NULL", "Functional family (Analytics, Engineering, Design)"),
    ("dim_job_role", "seniority_band", "TEXT", "NOT NULL", "Default expected seniority (Fresher, Mid, Senior)"),
    ("fact_job_postings", "job_id", "INTEGER", "PK", "Unique job posting identifier (1 to 652)"),
    ("fact_job_postings", "role_id", "INTEGER", "FK", "Foreign key referencing dim_job_role(role_id)"),
    ("fact_job_postings", "company_name", "TEXT", "NOT NULL", "Hiring enterprise organization name"),
    ("fact_job_postings", "industry_id", "INTEGER", "FK", "Foreign key referencing dim_industry(industry_id)"),
    ("fact_job_postings", "city_id", "INTEGER", "FK", "Foreign key referencing dim_city(city_id)"),
    ("fact_job_postings", "job_type", "TEXT", "NOT NULL", "Work mode (Full-time, Hybrid, Remote, Contract)"),
    ("fact_job_postings", "experience_level", "TEXT", "NOT NULL", "Seniority tier (Fresher, Mid-level, Senior)"),
    ("fact_job_postings", "salary_min", "REAL", "NULLABLE", "Minimum compensation in INR (12 imputed nulls)"),
    ("fact_job_postings", "salary_max", "REAL", "NOT NULL", "Maximum compensation in INR"),
    ("fact_job_postings", "required_skills", "TEXT", "NOT NULL", "Pipe-delimited skills list (e.g. Python|SQL)"),
    ("fact_job_postings", "skills_count", "INTEGER", "NOT NULL", "Count of required technical competencies (1 to 8)"),
    ("fact_job_postings", "posted_date", "TEXT", "NOT NULL", "ISO 8601 posting date (YYYY-MM-DD)"),
    ("fact_job_postings", "posted_month", "TEXT", "NOT NULL", "Derived month key (YYYY-MM)"),
    ("fact_job_postings", "posted_year", "INTEGER", "NOT NULL", "Derived year (2022 to 2024)"),
    ("job_skills", "skill_id", "INTEGER", "PK", "Unique decomposed skill identifier (1 to 2601)"),
    ("job_skills", "job_id", "INTEGER", "FK", "Foreign key referencing fact_job_postings(job_id)"),
    ("job_skills", "skill_name", "TEXT", "NOT NULL", "Individual skill name extracted to 1NF"),
    ("job_skills", "is_primary", "INTEGER", "DEFAULT 0", "Flag indicating if skill is listed first (0 or 1)")
]

tbl_db = doc.add_table(rows=len(db_rows)+1, cols=5)
tbl_db.rows[0].cells[0].paragraphs[0].text = "Table / sheet"
tbl_db.rows[0].cells[1].paragraphs[0].text = "Column"
tbl_db.rows[0].cells[2].paragraphs[0].text = "Type"
tbl_db.rows[0].cells[3].paragraphs[0].text = "Key"
tbl_db.rows[0].cells[4].paragraphs[0].text = "Description"
for i, row in enumerate(db_rows, 1):
    for c_i, val in enumerate(row):
        tbl_db.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_db, [1700, 2300, 2100, 1400, 2138])

add_label(doc, "Relationship: one (1) dim_industry -- (many) fact_job_postings (classifies)")
add_label(doc, "Relationship: one (1) dim_city -- (many) fact_job_postings (locates)")
add_label(doc, "Relationship: one (1) dim_job_role -- (many) fact_job_postings (defines)")
add_label(doc, "Relationship: one (1) fact_job_postings -- (many) job_skills (requires)")

add_label(doc, "SQLite DDL:")
ddl_code = """CREATE TABLE dim_industry (
    industry_id INTEGER PRIMARY KEY AUTOINCREMENT,
    industry_name TEXT NOT NULL UNIQUE,
    sector_category TEXT NOT NULL
);

CREATE TABLE dim_city (
    city_id INTEGER PRIMARY KEY AUTOINCREMENT,
    city_name TEXT NOT NULL UNIQUE,
    state_name TEXT NOT NULL,
    region TEXT NOT NULL
);

CREATE TABLE dim_job_role (
    role_id INTEGER PRIMARY KEY AUTOINCREMENT,
    role_title TEXT NOT NULL UNIQUE,
    role_category TEXT NOT NULL,
    seniority_band TEXT NOT NULL
);

CREATE TABLE fact_job_postings (
    job_id INTEGER PRIMARY KEY AUTOINCREMENT,
    role_id INTEGER NOT NULL REFERENCES dim_job_role(role_id),
    company_name TEXT NOT NULL,
    industry_id INTEGER NOT NULL REFERENCES dim_industry(industry_id),
    city_id INTEGER NOT NULL REFERENCES dim_city(city_id),
    job_type TEXT NOT NULL,
    experience_level TEXT NOT NULL,
    salary_min REAL,
    salary_max REAL NOT NULL,
    required_skills TEXT NOT NULL,
    skills_count INTEGER NOT NULL,
    posted_date TEXT NOT NULL,
    posted_month TEXT NOT NULL,
    posted_year INTEGER NOT NULL
);

CREATE TABLE job_skills (
    skill_id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id INTEGER NOT NULL REFERENCES fact_job_postings(job_id) ON DELETE CASCADE,
    skill_name TEXT NOT NULL,
    is_primary INTEGER DEFAULT 0
);"""
add_code(doc, ddl_code)

# ── 9. TEST CASES ────────────────────────────────────────────────────────────
add_heading_1(doc, "9. Test Cases")

tc_rows = [
    ("TC-F-01", "BR-01 / US-05", "Select 2023-01 to 2023-12 on Date Slider in UI-02", "Top 10 job roles update to display exact 2023 counts", "Ranked chart displays 2023 role counts", "Pass"),
    ("TC-F-02", "BR-02 / US-06", "Select 'IT & Software' from Industry dropdown in UI-03", "Top 10 skills chart updates to Python, SQL, Java", "Top skills update to IT competencies", "Pass"),
    ("TC-F-03", "BR-03 / US-09", "Filter 'Fresher' and 'Bangalore' in UI-04", "Salary min, avg, max cards update to expected figures", "Salary cards display ₹3.5L to ₹6.5L ranges", "Pass"),
    ("TC-F-04", "BR-04 / US-16", "Select 'Banking & Financial Services' in UI-08", "Top 10 hiring banks appear in ranked order", "HDFC, ICICI, SBI lead banking leaderboard", "Pass"),
    ("TC-F-05", "BR-05 / US-13", "Toggle chronological monthly timeline in UI-06", "Line chart renders continuous monthly timeline", "36-month timeline displays Q1/Q3 peaks", "Pass"),
    ("TC-F-06", "BR-06 / US-10", "View cross-industry comparison tab in UI-04", "Horizontal bar chart displays all 8 industries", "8 industries displayed sorted by avg salary", "Pass"),
    ("TC-F-07", "BR-07 / US-12", "Load default geographic view in UI-05", "Ranked bar chart displays city job counts", "Bangalore, Hyderabad, Pune top chart", "Pass"),
    ("TC-F-08", "BR-08 / US-07", "Select Experience Breakdown per Role in UI-10", "100% stacked bar chart shows seniority ratios", "Fresher/Mid/Senior splits total 100%", "Pass"),
    ("TC-F-09", "BR-09 / US-11", "Load correlation scatter plot in UI-07", "Scatter plot renders with positive OLS trendline", "r = +0.48, p < 0.001 rendered with slope", "Pass"),
    ("TC-F-10", "BR-10 / US-15", "Select 'IT & Software' on work mode page in UI-10", "Donut chart shows Hybrid (30%) and Remote (18%)", "Work modes render with 48% flexible share", "Pass"),
    ("TC-F-11", "BR-11 / US-08", "Filter skill keyword 'Python' in UI-10", "Ranked chart displays industries hiring Python", "IT, Finance, Analytics lead Python demand", "Pass"),
    ("TC-F-12", "BR-12 / US-14", "Select single role 'Data Analyst' in UI-06", "Isolated monthly trendline renders for Data Analyst", "Data Analyst spikes visible in Jan and July", "Pass"),
    ("TC-F-13", "BR-13 / US-02", "Inspect EDA summary table in UI-09", "Descriptive statistics and null counts displayed", "12 nulls in salary_min correctly identified", "Pass"),
    ("TC-F-14", "BR-01..13 / US-04", "Click 'Reset Global Filters' button in UI-01", "All filters revert to defaults; KPI cards refresh", "Filters reset and cards show 652 postings", "Pass"),
    ("TC-F-15", "BR-01..13 / US-04", "Select filter combination yielding zero records", "Displays warning: 'No data matches selected filters'", "Clean warning notification displayed", "Pass"),
    ("TC-F-16", "BR-01..13 / US-16", "Click 'Download Filtered CSV' on analysis pages", "Browser downloads sanitized CSV matching filtered data", "CSV downloaded with expected row count", "Pass"),
    ("TC-T-01", "BR-13 / US-01", "Verify data/app.db SQLite file exists and > 10KB", "File exists and size > 10,000 bytes", "Database exists, size = 304 KB", "Pass"),
    ("TC-T-02", "BR-02, 04, 08 / US-01", "Verify all 5 tables and 1 analytical view exist", "Entities found in sqlite_master", "All 6 database entities verified", "Pass"),
    ("TC-T-03", "BR-13 / US-01", "Assert fact_job_postings row count >= 500", "Row count meets or exceeds 500", "Exact row count = 652 records", "Pass"),
    ("TC-T-04", "BR-13 / US-01", "Check master table row counts: 8, 15, 12", "Master tables match planned dimensions", "Counts match exactly: 8, 15, 12", "Pass"),
    ("TC-T-05", "BR-02 / US-01", "Execute PRAGMA foreign_key_check;", "0 foreign key violations returned", "0 violations (Perfect referential pass)", "Pass"),
    ("TC-T-06", "BR-13 / US-01", "Check vw_hiring_analysis view column count >= 25", "View contains at least 25 columns", "27 pre-joined columns present", "Pass"),
    ("TC-T-07", "BR-01 / US-05", "Verify role demand aggregation sums to total rows", "Sum of counts equals len(df)", "Sum equals 652 records exactly", "Pass"),
    ("TC-T-08", "BR-03, 06 / US-09", "Assert salary_min <= salary_max and avg > 200,000", "All rows satisfy salary logic", "All rows valid, mean > 200,000 INR", "Pass"),
    ("TC-T-09", "BR-07 / US-12", "Verify Bangalore or Hyderabad in top 3 cities", "Tech hubs lead geographic volume", "Bangalore & Hyderabad in top 3", "Pass"),
    ("TC-T-10", "BR-05, 12 / US-13", "Verify monthly dates cover >= 12 ascending months", "Months ordered chronologically", "36 distinct months sorted monotonically", "Pass"),
    ("TC-T-11", "BR-10 / US-15", "Check job type normalized value_counts sums to 1.0", "Proportions sum to exactly 1.0000", "Sum equals exactly 1.0000", "Pass"),
    ("TC-T-12", "BR-09 / US-11", "Assert skills_count is strictly between 1 and 8", "All rows satisfy 1 <= count <= 8", "All 652 rows satisfy condition", "Pass"),
    ("TC-T-13", "NFR-01 / US-01", "Check config.toml and ui.py exist in repository", "Required modules verified", "Both modules confirmed present", "Pass"),
    ("TC-T-14", "NFR-04 / US-03", "Verify src.ui contains Tailwind CDN stylesheet string", "Tailwind CDN imported cleanly", "Tailwind CDN string verified", "Pass"),
    ("TC-T-15", "NFR-01 / US-01", "Benchmark analytical view query latency < 1.0 second", "Query evaluates in < 1.0 second", "Query evaluated in 0.012 seconds", "Pass")
]

tbl_tc = doc.add_table(rows=len(tc_rows)+1, cols=6)
tbl_tc.rows[0].cells[0].paragraphs[0].text = "TC-ID"
tbl_tc.rows[0].cells[1].paragraphs[0].text = "Linked Req / US"
tbl_tc.rows[0].cells[2].paragraphs[0].text = "Steps"
tbl_tc.rows[0].cells[3].paragraphs[0].text = "Expected result"
tbl_tc.rows[0].cells[4].paragraphs[0].text = "Actual"
tbl_tc.rows[0].cells[5].paragraphs[0].text = "Pass / Fail"
for i, row in enumerate(tc_rows, 1):
    for c_i, val in enumerate(row):
        tbl_tc.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_tc, [900, 1500, 2400, 2638, 1100, 1100])

# ── 10. UI FRAMEWORK ─────────────────────────────────────────────────────────
add_heading_1(doc, "10. UI Framework")
add_note(doc, "The screen-by-screen design, built first as an HTML framework, then in Streamlit.")

ui_rows = [
    ("UI-01", "Overview & KPIs", "Macro talent market health snapshot", "4 Tailwind KPI metric cards, hiring trajectory line chart, top role bar chart", "Date window, Industry, Metro, Seniority", "US-03, US-04"),
    ("UI-02", "Role Demand", "Ranked job role posting volumes", "Horizontal bar chart with top-N slider, summary table, CSV export", "Industry sector, Date window", "US-05"),
    ("UI-03", "Skills Analysis", "Most requested technical competencies", "Decomposed skill frequency horizontal bar chart, tabular breakdown", "Industry multiselect, Top-N slider", "US-06"),
    ("UI-04", "Salary Benchmark", "Compensation benchmarks by domain & city", "Grouped bar chart (Min/Mean/Max), box plots, experience/metro slicers", "Seniority tier, Metro location", "US-09, US-10"),
    ("UI-05", "Geography", "Urban talent concentration across metros", "Horizontal bar chart across 15 Indian metropolitan areas, regional summary", "Geographic region, Work mode", "US-12"),
    ("UI-06", "Hiring Trends", "Longitudinal monthly recruitment cycles", "Chronological monthly line chart (2022–2024), single-role trajectory slicer", "Target role title, Calendar year", "US-13, US-14"),
    ("UI-07", "Correlation", "Skills count vs offered salary regression", "Scatter plot with OLS linear trendline, Pearson r, p-value, slope callouts", "Seniority tier, Functional family", "US-11"),
    ("UI-08", "Companies", "Top recruiting enterprise employers", "Ranked employer horizontal bar chart, sortable company table, CSV download", "Industry domain", "US-16"),
    ("UI-09", "EDA Quality", "Data cleaning, imputation & outlier audit", "Pandas describe() table, missing value counts, IQR outlier detection badge", "Entity selector", "US-01, US-02"),
    ("UI-10", "Job Types & Exp", "Workplace policies & seniority demand", "Work mode donut chart (Remote, Hybrid, Onsite), 100% stacked experience bar", "Industry sector", "US-07, US-08, US-15")
]

tbl_ui = doc.add_table(rows=len(ui_rows)+1, cols=6)
tbl_ui.rows[0].cells[0].paragraphs[0].text = "UI-ID"
tbl_ui.rows[0].cells[1].paragraphs[0].text = "Screen"
tbl_ui.rows[0].cells[2].paragraphs[0].text = "Purpose"
tbl_ui.rows[0].cells[3].paragraphs[0].text = "Components"
tbl_ui.rows[0].cells[4].paragraphs[0].text = "Filters"
tbl_ui.rows[0].cells[5].paragraphs[0].text = "User stories"
for i, row in enumerate(ui_rows, 1):
    for c_i, val in enumerate(row):
        tbl_ui.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_ui, [800, 1200, 1700, 2900, 1700, 1338])

# Break to Appendices
doc.add_page_break()

# ── APPENDIX A ───────────────────────────────────────────────────────────────
add_heading_1(doc, "Appendix A - Prompt to generate this document")

add_label(doc, "MY PREFERENCES", color=CLR_RED)
prefs_code = """PROJECT ID & TITLE: P01 - Hiring Trend Analysis
DOMAIN: Data Analytics
LEARNER(S) & ROLL NO(S): Abhayjeet Singh (Roll No: 2500320140004)
MENTOR: Dinesh Sereno
BATCH: MCA 2nd Year Data Analyst - ABES_OND26_MCA Batch 01
INTENDED USERS: HR Managers, Technical Recruiters, Job Seekers, Placement Officers
PROBLEM IN MY WORDS: Lack of a unified, real-time analytics tool to analyze role demand, salary benchmarks, skills, and geographic hiring concentration across India.
DATASET: 652 synthetic Indian job postings stored in SQLite across 5 tables with an analytical view.
KPIS I WANT: Total Postings, Average Maximum Salary, Top City, Top In-Demand Role.
FEATURES I WANT: Multipage Streamlit app, Tailwind KPI cards, Plotly charts, global sidebar filters, CSV downloads, EDA audit panel.
STORAGE: SQLite database (data/app.db).
TEAM SIZE & DEADLINE: Solo developer (Abhayjeet Singh); 8-week academic capstone timeline."""
add_code(doc, prefs_code)

add_label(doc, "INSTRUCTIONS FOR THE AI", color=CLR_RED, before=10)
kickoff_code = """You are my senior Business Analyst, Solution Architect, Python data analyst and QA mentor. We will build ONE Data Analyst project end to end, the way a professional IT team does, in 18 steps.
PROJECT: P01 - Hiring Trend Analysis | DOMAIN: Data Analytics
ME: Abhayjeet Singh, Roll 2500320140004, MCA 2nd yr DA, ABES | MENTOR: Dinesh Sereno
PRODUCT: Python + Streamlit app (EDA, statistics, Plotly charts, Tailwind UI), SQLite data, pytest tests, on Streamlit Community Cloud.
FINAL SUBMISSION: Project Document (ABES Word template) + ABES slide deck."""
add_code(doc, kickoff_code)

# ── APPENDIX B ───────────────────────────────────────────────────────────────
add_heading_1(doc, "Appendix B - AI tools disclosure")

ai_rows = [
    ("Claude 3.5 Sonnet / DeepSeek", "Framing How-Might-We statements, drafting agile user stories, and structuring acceptance criteria.", "Reviewed against project objectives and verified against all 10 Operational Questions."),
    ("Python-docx / Openpyxl Generator", "Automating generation of formatted Word deliverables and Excel Traceability Matrices.", "Verified by executing script, inspecting file sizes, and re-opening documents in Microsoft Word."),
    ("Pytest Automated Test Runner", "Executing automated unit tests to verify schema integrity, aggregation math, and latency.", "Executed directly via PowerShell terminal; 15/15 tests passed with 100% pass rate in 4.59s.")
]

tbl_ai = doc.add_table(rows=len(ai_rows)+1, cols=3)
tbl_ai.rows[0].cells[0].paragraphs[0].text = "Tool"
tbl_ai.rows[0].cells[1].paragraphs[0].text = "Used for"
tbl_ai.rows[0].cells[2].paragraphs[0].text = "How the output was checked"
for i, row in enumerate(ai_rows, 1):
    for c_i, val in enumerate(row):
        tbl_ai.rows[i].cells[c_i].paragraphs[0].text = val
apply_table_styles(tbl_ai, [2600, 3500, 3538])

# ── CONFIRMATION ITEMS CHECK ─────────────────────────────────────────────────
p_conf = doc.add_paragraph()
p_conf.paragraph_format.space_before = Pt(14)
r_conf = p_conf.add_run("ITEMS TO CONFIRM:")
r_conf.font.name = 'Calibri'; r_conf.font.size = Pt(9.5); r_conf.font.bold = True; r_conf.font.color.rgb = CLR_GRAY
add_note(doc, "• Live Streamlit App URL [To confirm]: https://2500320140004-hiringtrendanalysis.streamlit.app/ (Pending cloud deployment in Step 16)")
add_note(doc, "• GitHub Repository URL [To confirm]: https://github.com/abhayjeet-singh/2500320140004-hiringtrendanalysis (To be pushed to learner's personal GitHub account)")

doc.save(OUT_DOCX)
print(f"SAVED ABES PROJECT DOCUMENT DOCX: {OUT_DOCX}")
