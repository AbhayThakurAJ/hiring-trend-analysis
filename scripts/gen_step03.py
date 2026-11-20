"""
Step 03 – Needs Analysis (Business Requirements) Document Generator
Project : P01 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES, MCA 2nd Yr DA
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os, datetime

OUT_DIR  = r"C:\Hiring Trend Analysis\docs"
OUT_FILE = os.path.join(OUT_DIR, "P01_HTA_Step03_NeedsAnalysis_BR.docx")
os.makedirs(OUT_DIR, exist_ok=True)

# ── colours ───────────────────────────────────────────────────────────────────
DARK_BLUE = RGBColor(0x1F,0x35,0x64)
MID_BLUE  = RGBColor(0x2E,0x74,0xB5)
STEEL     = RGBColor(0x4A,0x86,0xC8)
WHITE     = RGBColor(0xFF,0xFF,0xFF)
BLACK     = RGBColor(0x00,0x00,0x00)
GREEN_OK  = RGBColor(0x37,0x86,0x36)
RED_ERR   = RGBColor(0xC0,0x00,0x00)
ORANGE    = RGBColor(0xC5,0x5A,0x11)

HEX = dict(
    dark='1F3564', mid='2E74B5', must='D6E4F0', should='E2EFDA',
    could='FFF3CD', wont='F4CCCC', alt='DAE8F5', white='FFFFFF',
    gap='FFE5E5', ok='E5FFE5', hdr_must='1F5C2E', hdr_should='2E5F1F',
    hdr_could='7B5700', banner='1F3564'
)

# ── helpers ───────────────────────────────────────────────────────────────────
def set_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_borders(table, color='2E74B5'):
    for row in table.rows:
        for cell in row.cells:
            tc   = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcB  = OxmlElement('w:tcBorders')
            for edge in ('top','left','bottom','right','insideH','insideV'):
                e = OxmlElement(f'w:{edge}')
                e.set(qn('w:val'),'single'); e.set(qn('w:sz'),'4')
                e.set(qn('w:space'),'0');    e.set(qn('w:color'),color)
                tcB.append(e)
            tcPr.append(tcB)

def hdr_row(table, *texts, bg='1F3564'):
    row = table.rows[0]
    for cell, text in zip(row.cells, texts):
        set_bg(cell, bg)
        p   = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True; run.font.color.rgb = WHITE; run.font.size = Pt(9.5)

def fill_cell(cell, text, bg=HEX['white'], bold=False, size=9.5,
              align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    set_bg(cell, bg)
    p   = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold; run.font.size = Pt(size)
    if color: run.font.color.rgb = color

def fill_row(table, row_idx, texts, bgs=None, bolds=None):
    row = table.rows[row_idx]
    for i, (cell, text) in enumerate(zip(row.cells, texts)):
        bg   = bgs[i]   if bgs   else HEX['white']
        bold = bolds[i] if bolds else False
        fill_cell(cell, text, bg=bg, bold=bold)

def add_heading(doc, text, level=1):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level==1 else 10)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size      = Pt(14 if level==1 else (12 if level==2 else 11))
    run.font.color.rgb = DARK_BLUE if level==1 else (MID_BLUE if level==2 else STEEL)

def add_para(doc, text, bold=False, italic=False, size=10, color=BLACK, indent=0):
    p   = doc.add_paragraph()
    p.paragraph_format.left_indent  = Cm(indent)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.bold   = bold; run.italic = italic
    run.font.size = Pt(size); run.font.color.rgb = color

def rule(doc):
    p   = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr= OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'6')
    bot.set(qn('w:space'),'1');    bot.set(qn('w:color'),'2E74B5')
    pBdr.append(bot); pPr.append(pBdr)

MOSCOW_BG = {'Must':'D6E4F0','Should':'E2EFDA','Could':'FFF3CD','Won\'t':'F4CCCC'}
MOSCOW_COLOR = {
    'Must':   RGBColor(0x1F,0x35,0x64),
    'Should': RGBColor(0x1F,0x5C,0x2E),
    'Could':  RGBColor(0x7B,0x57,0x00),
    "Won't":  RGBColor(0x7B,0x2D,0x2D),
}

# ═════════════════════════════════════════════════════════════════════════════
doc = Document()
for section in doc.sections:
    section.top_margin    = Cm(2.0); section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5); section.right_margin  = Cm(2.5)

# ── BANNER ────────────────────────────────────────────────────────────────────
banner = doc.add_table(rows=1, cols=1)
banner.alignment = WD_TABLE_ALIGNMENT.CENTER
bc = banner.cell(0,0); set_bg(bc,'1F3564')
bp = bc.paragraphs[0]
bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_before = Pt(10)
bp.paragraph_format.space_after  = Pt(10)
r1 = bp.add_run('P01 – HIRING TREND ANALYSIS\n')
r1.bold=True; r1.font.size=Pt(18); r1.font.color.rgb=WHITE
r2 = bp.add_run('Step 03 · Needs Analysis — Business Requirements (BR)\n')
r2.font.size=Pt(12); r2.font.color.rgb=RGBColor(0xBD,0xD5,0xEE)
r3 = bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA')
r3.font.size=Pt(9); r3.font.color.rgb=RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

meta = doc.add_table(rows=2,cols=4); meta.alignment=WD_TABLE_ALIGNMENT.CENTER
meta_data=[
    ('Project ID','P01'),('Version','1.0'),
    ('Date',datetime.date.today().strftime('%d %b %Y')),('Status','Draft'),
    ('Step','03 of 18'),('Method','MoSCoW Prioritisation'),
    ('Prev','02 – HMW'),('Next','04 – RTM (Traceability)'),
]
for i,row in enumerate(meta.rows):
    fill=HEX['alt'] if i%2==0 else HEX['white']
    for j,cell in enumerate(row.cells):
        set_bg(cell,fill)
        p=cell.paragraphs[0]; txt=meta_data[i*4+j]
        r=p.add_run(f'{txt[0]}: '); r.bold=True; r.font.size=Pt(9)
        p.add_run(txt[1]).font.size=Pt(9)
set_borders(meta); doc.add_paragraph(); rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §1  INTRODUCTION
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc,'1. Introduction — Needs Analysis Methodology',1)
add_para(doc,
    'Needs analysis translates the design-thinking opportunity questions from Step 02 into '
    'concrete, stakeholder-endorsed Business Requirements (BRs). Each BR answers the question: '
    '"What capability must exist for a stakeholder to realise the value promised by the HMW?" '
    'BRs describe WHAT the business needs — never HOW it will be built. '
    'Technology choices (Python, Streamlit, SQLite) appear only from Step 05 (BRD) onward.',
    size=10)
doc.add_paragraph()
for b in [
    'Each BR is written in the canonical form: "The business needs to <capability> so that <benefit>."',
    'Priority follows MoSCoW: Must · Should · Could · Won\'t (this release).',
    'Target: 10–15 BRs with ~60% Must-have.',
    'Every BR must be traceable to a source HMW and must be independently testable.',
    'Data requirements are listed per BR to feed the database design in Step 08.',
]:
    p=doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent=Cm(0.5)
    p.paragraph_format.space_after=Pt(1)
    p.add_run(b).font.size=Pt(9.5)
doc.add_paragraph(); rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §2  HMW → NEED DERIVATION
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc,'2. HMW → Stakeholder Need Derivation',1)
add_para(doc,
    'For each HMW statement the facilitator asked: "What does the stakeholder NEED to '
    'make this possible?" The raw need was then formalised into a BR sentence.',
    size=9.5)
doc.add_paragraph()

derivations = [
    ('HMW-01','SH-03 HR Mgr',
     'A ranked view of job roles by posting count for a chosen date range',
     'BR-01'),
    ('HMW-02','SH-04 Recruiter',
     'A ranked list of skills by frequency within a chosen industry',
     'BR-02'),
    ('HMW-03','SH-05 Job Seeker',
     'Salary statistics (avg/min/max) segmented by experience level AND city',
     'BR-03'),
    ('HMW-04','SH-06 Placement Officer',
     'A ranked leaderboard of companies by posting volume for a chosen domain',
     'BR-04'),
    ('HMW-05','SH-03 HR Mgr',
     'Aggregated monthly hiring volume as a navigable time series',
     'BR-05'),
    ('HMW-06','SH-04 Recruiter',
     'A quantified relationship between skills-count-per-posting and max salary',
     'BR-09'),
    ('HMW-07','SH-05 Job Seeker',
     'Posting counts grouped and ranked by city across the whole dataset',
     'BR-07'),
    ('HMW-08','SH-04 Recruiter',
     'A breakdown of job-type proportions per industry',
     'BR-10'),
    ('HMW-09','SH-06 Placement Officer',
     'Experience-level demand split (Fresher/Mid/Senior) per job role',
     'BR-08'),
    ('HMW-10','SH-03 HR Mgr',
     'A single cross-industry salary comparison (avg/min/max)',
     'BR-06'),
    ('HMW-11','SH-06 Placement Officer',
     'Industry hiring volume ranked for a user-selected skill set over time',
     'BR-11'),
    ('HMW-12','SH-05 Job Seeker',
     'A time-series view of posting frequency for one specific job role',
     'BR-12'),
    ('—  (Assumed)','SH-02 Developer',
     'Summary statistics for all numeric fields to validate dataset quality',
     'BR-13'),
]

drv_tbl = doc.add_table(rows=len(derivations)+1, cols=4)
drv_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr_row(drv_tbl,'Source HMW','Stakeholder','Raw Stakeholder Need','Maps to BR')
for i,(hmw,sh,need,br) in enumerate(derivations):
    bg = HEX['alt'] if i%2==0 else HEX['white']
    gap_bg = HEX['gap'] if hmw=='—  (Assumed)' else bg
    fill_row(drv_tbl, i+1,
             [hmw, sh, need, br],
             bgs=[gap_bg,gap_bg,gap_bg,gap_bg],
             bolds=[True,False,False,True])
set_borders(drv_tbl)
add_para(doc,'Row shaded pink = BR-13 has no direct HMW; it is an assumed baseline data-quality requirement (see §5 Gaps).',
         italic=True, size=8.5, color=ORANGE)
doc.add_paragraph(); rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §3  BUSINESS REQUIREMENTS TABLE (MASTER)
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc,'3. Business Requirements — Master Table',1)
add_para(doc,
    'All 13 BRs are listed below using MoSCoW priority. '
    'Must = 8 (62%) · Should = 4 (31%) · Could = 1 (8%) · Won\'t = 0.',
    size=9.5)
doc.add_paragraph()

# columns: BR ID | Requirement | Stakeholder | HMW | Priority | Measurement
brs = [
    # (id, requirement, stakeholder, hmw, priority, measurement)
    ('BR-01',
     'The business needs to rank job roles by posting volume for any user-selected time period '
     'so that HR managers can identify surging or declining roles without manual data aggregation.',
     'SH-03','HMW-01','Must',
     'All top-N roles displayed with correct descending count; date filter changes results correctly'),

    ('BR-02',
     'The business needs to rank skills by frequency of appearance across job postings within '
     'a chosen industry so that recruiters can build data-driven job descriptions.',
     'SH-04','HMW-02','Must',
     'Skill list updates when industry filter changes; frequency counts match raw data counts'),

    ('BR-03',
     'The business needs to present average, minimum, and maximum salary figures segmented by '
     'experience level and city so that job seekers can verify whether their expectations are market-aligned.',
     'SH-05','HMW-03','Must',
     'Three salary statistics correct for every experience–city combination; filters work independently'),

    ('BR-04',
     'The business needs to rank companies by total hiring volume within a user-selected domain '
     'so that placement officers can prioritise campus drive invitations with objective evidence.',
     'SH-06','HMW-04','Must',
     'Leaderboard rank order matches descending posting count per domain; ties broken consistently'),

    ('BR-05',
     'The business needs to aggregate and display total job posting volume by calendar month '
     'so that HR managers can align recruitment campaigns with seasonal demand peaks.',
     'SH-03','HMW-05','Must',
     'Monthly totals sum to overall dataset count; time axis is correctly ordered chronologically'),

    ('BR-06',
     'The business needs to compare average, minimum, and maximum offered salaries across all '
     'industries in a consolidated single view so that HR managers can benchmark compensation '
     'packages without consulting external salary surveys.',
     'SH-03','HMW-10','Must',
     'Stats per industry correct; at least 3 industries present; sorted by avg salary descending'),

    ('BR-07',
     'The business needs to present job posting counts ranked by city so that job seekers '
     'and recruiters can identify geographic hiring hotspots at a glance.',
     'SH-05','HMW-07','Must',
     'City counts correct and sorted; top-10 cities cover ≥60% of dataset records'),

    ('BR-08',
     'The business needs to show the proportion of job postings demanding each experience level '
     '(Fresher / Mid-level / Senior) broken down by job role so that placement officers can '
     'match the right student batch profile to the right recruiter.',
     'SH-06','HMW-09','Must',
     'Each role\'s experience split sums to 100%; filter by role updates chart correctly'),

    ('BR-09',
     'The business needs to quantify the statistical relationship between the number of skills '
     'required per posting and the maximum salary offered so that recruiters can price '
     'complex, multi-skill roles competitively.',
     'SH-04','HMW-06','Should',
     'Correlation coefficient displayed; scatter plot shows correct (skills_count, salary_max) pairs'),

    ('BR-10',
     'The business needs to present the distribution of job types (Full-time, Part-time, Remote, '
     'Hybrid, Contract) for each industry so that recruiters can benchmark their flexible-work '
     'policies against market norms.',
     'SH-04','HMW-08','Should',
     'Job-type proportions per industry sum to 100%; each type appears in correct slice'),

    ('BR-11',
     'The business needs to rank industries by total hiring volume for a user-selected skill '
     'keyword over time so that career counsellors can direct students toward emerging sectors.',
     'SH-06','HMW-11','Should',
     'Industry ranking changes when skill filter changes; volume counts match filtered records'),

    ('BR-12',
     'The business needs to display a time-series view of posting frequency for a specific '
     'user-selected job role so that job seekers can identify peak hiring windows and time '
     'their applications accordingly.',
     'SH-05','HMW-12','Should',
     'Monthly counts for selected role correct; zero-posting months shown as gaps, not omitted'),

    ('BR-13',
     'The business needs to present summary statistics (record count, mean, median, standard '
     'deviation, missing-value count, outlier flags) for all numeric fields so that analysts '
     'can rapidly assess dataset quality and distribution before interpreting any chart.',
     'SH-02','— (Assumed)','Could',
     'Stats match pandas .describe() output; missing-value count matches actual nulls in dataset'),
]

col_w = [Cm(1.5), Cm(8.0), Cm(2.0), Cm(1.8), Cm(1.8), Cm(3.5)]
master = doc.add_table(rows=len(brs)+1, cols=6)
master.alignment = WD_TABLE_ALIGNMENT.CENTER
for row in master.rows:
    for i,cell in enumerate(row.cells):
        cell.width = col_w[i]
hdr_row(master,'BR ID','Business Requirement','Stakeholder','Source HMW','Priority','Measurement / Acceptance Criterion')

for idx,(bid,req,sh,hmw,prio,measure) in enumerate(brs):
    bg = MOSCOW_BG.get(prio, HEX['white'])
    row = master.rows[idx+1]
    data = [bid, req, sh, hmw, prio, measure]
    for ci,(cell,val) in enumerate(zip(row.cells,data)):
        set_bg(cell, bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci in (0,2,3,4) else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(str(val))
        run.font.size = Pt(9)
        if ci == 0:
            run.bold = True
            run.font.color.rgb = MOSCOW_COLOR.get(prio, BLACK)
        if ci == 4:
            run.bold = True
            run.font.color.rgb = MOSCOW_COLOR.get(prio, BLACK)

set_borders(master)

# legend
leg = doc.add_table(rows=1, cols=4); leg.alignment=WD_TABLE_ALIGNMENT.CENTER
legend_items=[('Must — deliver in this release','D6E4F0','1F3564'),
              ('Should — high value, include if possible','E2EFDA','1F5C2E'),
              ('Could — nice-to-have if time allows','FFF3CD','7B5700'),
              ("Won't — out of scope (this release)",'F4CCCC','7B2D2D')]
for i,(label,bg,tc) in enumerate(legend_items):
    c=leg.cell(0,i); set_bg(c,bg)
    p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(label); r.bold=True; r.font.size=Pt(8)
    r.font.color.rgb=RGBColor(int(tc[:2],16),int(tc[2:4],16),int(tc[4:],16))
set_borders(leg,'AAAAAA')
doc.add_paragraph(); rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §4  DATA REQUIREMENTS PER BR
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc,'4. Data Requirements per Business Requirement',1)
add_para(doc,
    'The following table maps each BR to the specific dataset fields it needs. '
    'This feeds directly into the database design in Step 08. '
    'All fields are present in the planned synthetic dataset.',
    size=9.5)
doc.add_paragraph()

data_reqs = [
    ('BR-01','Must', 'job_title, posted_date',
     'Filter: date range → group by job_title → count job_id'),
    ('BR-02','Must', 'required_skills (pipe-delimited), industry',
     'Filter: industry → explode skills → count frequency per skill'),
    ('BR-03','Must', 'salary_min, salary_max, experience_level, city',
     'Filter: exp_level + city → compute avg/min/max of salary_max'),
    ('BR-04','Must', 'company_name, industry, job_id',
     'Filter: industry → group by company_name → count job_id → rank'),
    ('BR-05','Must', 'posted_date (→ month/year), job_id',
     'Group by month(posted_date) → count job_id → sort chronologically'),
    ('BR-06','Must', 'industry, salary_min, salary_max',
     'Group by industry → compute avg/min/max of salary_max → sort'),
    ('BR-07','Must', 'city, job_id',
     'Group by city → count job_id → sort descending'),
    ('BR-08','Must', 'job_title, experience_level',
     'Group by job_title + experience_level → count → pivot to %'),
    ('BR-09','Should','required_skills (count per row), salary_max',
     'Derive skills_count = len(split(required_skills)) → scatter vs salary_max → OLS'),
    ('BR-10','Should','job_type, industry',
     'Group by industry + job_type → count → pivot to proportion'),
    ('BR-11','Should','industry, required_skills, posted_date',
     'Filter: skill keyword in required_skills → group industry+month → count'),
    ('BR-12','Should','job_title, posted_date',
     'Filter: job_title == selected → group by month(posted_date) → count'),
    ('BR-13','Could', 'All numeric fields: salary_min, salary_max',
     'pandas describe() + null count + IQR-based outlier flag'),
]

dr_tbl = doc.add_table(rows=len(data_reqs)+1, cols=4)
dr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr_row(dr_tbl,'BR ID','Priority','Dataset Fields Required','Computation / Derivation Needed')
for i,(bid,prio,fields,comp) in enumerate(data_reqs):
    bg = MOSCOW_BG.get(prio, HEX['white'])
    row = dr_tbl.rows[i+1]
    for ci,(cell,val) in enumerate(zip(row.cells,[bid,prio,fields,comp])):
        set_bg(cell, bg)
        p=cell.paragraphs[0]
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER if ci<2 else WD_ALIGN_PARAGRAPH.LEFT
        run=p.add_run(str(val)); run.font.size=Pt(9)
        if ci==0: run.bold=True; run.font.color.rgb=MOSCOW_COLOR.get(prio,BLACK)
set_borders(dr_tbl)
doc.add_paragraph(); rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §5  GAP ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc,'5. Gap Analysis',1)
add_para(doc,
    'Two gap checks are performed: (A) every HMW must have at least one BR, '
    '(B) every BR must trace back to at least one HMW (or a documented assumption).',
    size=9.5)
doc.add_paragraph()

add_heading(doc,'5A — HMWs without a BR',2)
gap_a_tbl = doc.add_table(rows=13,cols=3)
gap_a_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr_row(gap_a_tbl,'HMW ID','Maps to BR','Status')
gap_a_data=[
    ('HMW-01','BR-01','✔ Covered'),('HMW-02','BR-02','✔ Covered'),
    ('HMW-03','BR-03','✔ Covered'),('HMW-04','BR-04','✔ Covered'),
    ('HMW-05','BR-05','✔ Covered'),('HMW-06','BR-09','✔ Covered'),
    ('HMW-07','BR-07','✔ Covered'),('HMW-08','BR-10','✔ Covered'),
    ('HMW-09','BR-08','✔ Covered'),('HMW-10','BR-06','✔ Covered'),
    ('HMW-11','BR-11','✔ Covered'),('HMW-12','BR-12','✔ Covered'),
]
for i,(hmw,br,status) in enumerate(gap_a_data):
    bg=HEX['ok'] if '✔' in status else HEX['gap']
    fill_row(gap_a_tbl,i+1,[hmw,br,status],bgs=[bg,bg,bg],bolds=[True,False,False])
set_borders(gap_a_tbl)
add_para(doc,'Result: All 12 HMWs have at least one BR. No HMW gap.',bold=True,color=GREEN_OK)
doc.add_paragraph()

add_heading(doc,'5B — BRs without a source HMW',2)
gap_b_tbl = doc.add_table(rows=14,cols=3)
gap_b_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr_row(gap_b_tbl,'BR ID','Source HMW','Status / Resolution')
gap_b_data=[
    ('BR-01','HMW-01','✔ Traced'),('BR-02','HMW-02','✔ Traced'),
    ('BR-03','HMW-03','✔ Traced'),('BR-04','HMW-04','✔ Traced'),
    ('BR-05','HMW-05','✔ Traced'),('BR-06','HMW-10','✔ Traced'),
    ('BR-07','HMW-07','✔ Traced'),('BR-08','HMW-09','✔ Traced'),
    ('BR-09','HMW-06','✔ Traced'),('BR-10','HMW-08','✔ Traced'),
    ('BR-11','HMW-11','✔ Traced'),('BR-12','HMW-12','✔ Traced'),
    ('BR-13','— (None)','⚠ Gap — Assumed baseline data-quality requirement. No HMW was written for this. Resolution: accepted as a standard EDA prerequisite; will be noted in BRD §Assumptions.'),
]
for i,(br,hmw,status) in enumerate(gap_b_data):
    bg = HEX['gap'] if '⚠' in status else HEX['ok']
    fill_row(gap_b_tbl,i+1,[br,hmw,status],bgs=[bg,bg,bg],bolds=[True,False,False])
set_borders(gap_b_tbl)
add_para(doc,
    '⚠ BR-13 has no source HMW — it is an assumed baseline EDA requirement accepted by convention.\n'
    '✔ All other 12 BRs are fully traced to a source HMW.',
    size=9.5, color=ORANGE)
doc.add_paragraph(); rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §6  TESTABILITY CHECK
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc,'6. Testability Check — Every BR Must Be Independently Testable',1)
add_para(doc,
    'A BR is testable if: (a) its measurement criterion has a clear pass/fail condition, '
    '(b) a tester can verify it without access to source code, '
    '(c) it does not depend on another BR to be evaluated.',
    size=9.5)
doc.add_paragraph()

test_tbl = doc.add_table(rows=14,cols=4)
test_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr_row(test_tbl,'BR ID','Key Testable Condition','Test Type','Testable?')
test_data = [
    ('BR-01','Posting count per role matches count(*) filtered by date range','Data validation','✔ Yes'),
    ('BR-02','Skill frequency matches manual count in raw data for given industry','Data validation','✔ Yes'),
    ('BR-03','Avg/min/max salary correct per exp_level + city combination','Statistical check','✔ Yes'),
    ('BR-04','Company rank order matches descending count(*) for selected domain','Data validation','✔ Yes'),
    ('BR-05','Monthly totals sum equals total dataset row count','Aggregation check','✔ Yes'),
    ('BR-06','Industry salary stats correct; ≥3 industries displayed','Statistical check','✔ Yes'),
    ('BR-07','City counts correct; sorted descending; top-10 covers ≥60% rows','Data validation','✔ Yes'),
    ('BR-08','Experience split per role sums to 100%; filter changes output','Percentage check','✔ Yes'),
    ('BR-09','Correlation coefficient value matches pandas .corr() result','Statistical check','✔ Yes'),
    ('BR-10','Job-type proportions per industry sum to 100%','Percentage check','✔ Yes'),
    ('BR-11','Industry ranking changes when skill keyword filter changes','Filter test','✔ Yes'),
    ('BR-12','Monthly counts for selected role match filtered row count','Data validation','✔ Yes'),
    ('BR-13','Summary stats match pandas .describe(); null count matches actual nulls','EDA validation','✔ Yes'),
]
for i,row in enumerate(test_data):
    bg = HEX['ok'] if '✔' in row[3] else HEX['gap']
    fill_row(test_tbl,i+1,list(row),bgs=[bg,bg,bg,bg],bolds=[True,False,False,False])
set_borders(test_tbl)
add_para(doc,'✔ All 13 BRs have clear, independently testable acceptance criteria. No BR is ambiguous.',
         bold=True,color=GREEN_OK)
doc.add_paragraph(); rule(doc)

# ══════════════════════════════════════════════════════════════════════════════
# §7  MOSCOW SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc,'7. MoSCoW Priority Summary',1)
ms_tbl = doc.add_table(rows=5,cols=4)
ms_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr_row(ms_tbl,'Priority','Count','BR IDs','% of Total')
ms_data = [
    ('Must',  '8','BR-01, BR-02, BR-03, BR-04, BR-05, BR-06, BR-07, BR-08','62%'),
    ('Should','4','BR-09, BR-10, BR-11, BR-12','31%'),
    ('Could', '1','BR-13','8%'),
    ("Won't", '0','—','0%'),
]
for i,(prio,cnt,ids,pct) in enumerate(ms_data):
    bg=MOSCOW_BG.get(prio,HEX['white'])
    fill_row(ms_tbl,i+1,[prio,cnt,ids,pct],
             bgs=[bg,bg,bg,bg],bolds=[True,True,False,False])
set_borders(ms_tbl)
doc.add_paragraph(); rule(doc)

# ── DOC CONTROL ───────────────────────────────────────────────────────────────
add_heading(doc,'Document Control',2)
dc=doc.add_table(rows=2,cols=4); dc.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr_row(dc,'Version','Date','Author','Change Summary')
fill_row(dc,1,['1.0',datetime.date.today().strftime('%d %b %Y'),
               'Abhayjeet Singh','Initial draft – Step 03 complete'],
         bgs=[HEX['alt']]*4)
set_borders(dc)
doc.add_paragraph()
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
run=p.add_run('P01 – Hiring Trend Analysis  |  Step 03 Needs Analysis  |  ABES MCA 2nd Yr DA  |  Mentor: Dinesh Sereno')
run.font.size=Pt(8); run.font.color.rgb=RGBColor(0x88,0x88,0x88); run.italic=True

doc.save(OUT_FILE)
print(f"SAVED: {OUT_FILE}")
