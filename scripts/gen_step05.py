"""
Step 05 – Business Requirements Document (BRD) Generator
Project : P071 – Hiring Trend Analysis
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
OUT_FILE = os.path.join(OUT_DIR, "P01_HTA_Step05_BRD.docx")
os.makedirs(OUT_DIR, exist_ok=True)

# ── palette ───────────────────────────────────────────────────────────────────
DARK  = RGBColor(0x1F,0x35,0x64)
MID   = RGBColor(0x2E,0x74,0xB5)
STEEL = RGBColor(0x4A,0x86,0xC8)
WHITE = RGBColor(0xFF,0xFF,0xFF)
BLACK = RGBColor(0x00,0x00,0x00)
GREEN = RGBColor(0x37,0x86,0x36)
AMBER = RGBColor(0xC5,0x5A,0x11)
RED   = RGBColor(0xC0,0x00,0x00)

HEX = dict(
    dark='1F3564', mid='2E74B5', steel='4A86C8',
    must='D6E4F0', should='E2EFDA', could='FFF3CD',
    alt='EEF4FB', white='FFFFFF', ok='E5FFE5',
    warn='FFF3CD', err='FFE5E5', grey='F2F2F2',
    hdr='FFFFFF', r_col='D6E4F0', a_col='E2EFDA',
    c_col='FFF8DC', i_col='F2F2F2',
    risk_h='C00000', risk_m='FF8C00', risk_l='1F5C2E',
)

MOSCOW_BG = {'Must':'D6E4F0','Should':'E2EFDA','Could':'FFF3CD'}
MOSCOW_FG = {'Must':'1F3564','Should':'1F5C2E','Could':'7B5700'}

# ── helpers ───────────────────────────────────────────────────────────────────
def bg(cell, hex_c):
    tc=cell._tc; pr=tc.get_or_add_tcPr()
    s=OxmlElement('w:shd')
    s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto')
    s.set(qn('w:fill'),hex_c); pr.append(s)

def borders(table, color='2E74B5'):
    for row in table.rows:
        for cell in row.cells:
            tc=cell._tc; pr=tc.get_or_add_tcPr()
            tb=OxmlElement('w:tcBorders')
            for e in ('top','left','bottom','right','insideH','insideV'):
                x=OxmlElement(f'w:{e}')
                x.set(qn('w:val'),'single'); x.set(qn('w:sz'),'4')
                x.set(qn('w:space'),'0');    x.set(qn('w:color'),color)
                tb.append(x)
            pr.append(tb)

def hdr(table, *texts, hbg='1F3564'):
    row=table.rows[0]
    for cell,text in zip(row.cells,texts):
        bg(cell,hbg)
        p=cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r=p.add_run(text); r.bold=True
        r.font.color.rgb=WHITE; r.font.size=Pt(9.5)

def frow(table, ri, vals, bgs=None, bolds=None, sizes=None, aligns=None):
    row=table.rows[ri]
    for i,(cell,val) in enumerate(zip(row.cells,vals)):
        b  = bgs[i]    if bgs    else HEX['white']
        bo = bolds[i]  if bolds  else False
        sz = sizes[i]  if sizes  else 9.5
        al = aligns[i] if aligns else WD_ALIGN_PARAGRAPH.LEFT
        bg(cell,b)
        p=cell.paragraphs[0]; p.alignment=al
        r=p.add_run(str(val)); r.bold=bo; r.font.size=Pt(sz)

def H(doc, text, lv=1):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(14 if lv==1 else 10)
    p.paragraph_format.space_after=Pt(4)
    r=p.add_run(text); r.bold=True
    r.font.size=Pt(14 if lv==1 else (12 if lv==2 else 11))
    r.font.color.rgb=DARK if lv==1 else (MID if lv==2 else STEEL)

def P(doc, text, bold=False, italic=False, size=10, color=BLACK, ind=0):
    p=doc.add_paragraph()
    p.paragraph_format.left_indent=Cm(ind)
    p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
    r=p.add_run(text); r.bold=bold; r.italic=italic
    r.font.size=Pt(size); r.font.color.rgb=color

def rule(doc):
    p=doc.add_paragraph()
    pr=p._p.get_or_add_pPr(); pb=OxmlElement('w:pBdr')
    bt=OxmlElement('w:bottom')
    bt.set(qn('w:val'),'single'); bt.set(qn('w:sz'),'6')
    bt.set(qn('w:space'),'1');    bt.set(qn('w:color'),'2E74B5')
    pb.append(bt); pr.append(pb)

def bullet(doc, text, lv=0):
    p=doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent=Cm(0.5+lv*0.5)
    p.paragraph_format.space_after=Pt(1)
    p.add_run(text).font.size=Pt(9.5)

# ═════════════════════════════════════════════════════════════════════════════
doc = Document()
for s in doc.sections:
    s.top_margin=Cm(2.0); s.bottom_margin=Cm(2.0)
    s.left_margin=Cm(2.5); s.right_margin=Cm(2.5)

# ══ COVER ════════════════════════════════════════════════════════════════════
ban=doc.add_table(rows=1,cols=1); ban.alignment=WD_TABLE_ALIGNMENT.CENTER
bc=ban.cell(0,0); bg(bc,'1F3564')
bp=bc.paragraphs[0]; bp.alignment=WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_before=Pt(14); bp.paragraph_format.space_after=Pt(14)
r1=bp.add_run('P071 – HIRING TREND ANALYSIS\n')
r1.bold=True; r1.font.size=Pt(20); r1.font.color.rgb=WHITE
r2=bp.add_run('Business Requirements Document (BRD)\n')
r2.font.size=Pt(14); r2.font.color.rgb=RGBColor(0xBD,0xD5,0xEE)
r3=bp.add_run('Step 05 of 18  |  Version 1.0  |  Status: Draft\n')
r3.font.size=Pt(10); r3.font.color.rgb=RGBColor(0xBD,0xD5,0xEE)
r4=bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | ABES MCA 2nd Yr DA | Mentor: Dinesh Sereno')
r4.font.size=Pt(9); r4.font.color.rgb=RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

# ══ §1 DOCUMENT CONTROL ══════════════════════════════════════════════════════
H(doc,'1. Document Control',1)
dc=doc.add_table(rows=8,cols=2); dc.alignment=WD_TABLE_ALIGNMENT.CENTER
dc_data=[
    ('Document Title','P071 – Hiring Trend Analysis: Business Requirements Document'),
    ('Document ID','P071-BRD-001'),
    ('Version','1.0'),
    ('Status','Draft — Pending Mentor Review'),
    ('Prepared By','Abhayjeet Singh | Roll: 2500320140004 | ABES MCA 2nd Yr DA'),
    ('Mentor / Reviewer','Dinesh Sereno | ABES Engineering College'),
    ('Date Created', datetime.date.today().strftime('%d %b %Y')),
    ('Next Review Date','On completion of Step 06 (PRD)'),
]
for i,(label,val) in enumerate(dc_data):
    b=HEX['alt'] if i%2==0 else HEX['white']
    bg(dc.cell(i,0),b); bg(dc.cell(i,1),b)
    r=dc.cell(i,0).paragraphs[0].add_run(label)
    r.bold=True; r.font.size=Pt(9.5)
    dc.cell(i,1).paragraphs[0].add_run(val).font.size=Pt(9.5)
borders(dc)

doc.add_paragraph()
rev=doc.add_table(rows=4,cols=4); rev.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(rev,'Version','Date','Author','Change Summary')
rev_data=[
    ('1.0',datetime.date.today().strftime('%d %b %Y'),'Abhayjeet Singh','Initial draft – Steps 01–04 compiled'),
    ('1.1','TBD (Step 06)','Abhayjeet Singh','Add Feature IDs after PRD'),
    ('1.2','TBD (Step 15)','Abhayjeet Singh','Update status to Verified after build'),
]
for i,row in enumerate(rev_data,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(rev,i,row,bgs=[b]*4)
borders(rev)
doc.add_paragraph(); rule(doc)

# ══ §2 EXECUTIVE SUMMARY ════════════════════════════════════════════════════
H(doc,'2. Executive Summary',1)
P(doc,
  'This Business Requirements Document (BRD) defines the complete set of business-level needs '
  'for the Hiring Trend Analysis (HTA) application — a Python + Streamlit data analytics product '
  'developed as an academic capstone project by Abhayjeet Singh, MCA 2nd Year, ABES Engineering '
  'College, under the mentorship of Dinesh Sereno.',
  size=10)
P(doc,
  'The HTA application will ingest structured hiring data from a SQLite database and provide '
  'HR managers, recruiters, job seekers, and placement officers with interactive exploratory '
  'data analysis (EDA), statistical insights, and Plotly visualisations that answer 10 defined '
  'Operational Questions (OQ-01 to OQ-10).',
  size=10)
P(doc,
  'This document consolidates findings from Steps 01–04: Project Understanding, HMW Statements, '
  'Needs Analysis, and the Requirements Traceability Matrix. It is the authoritative source of '
  'business requirements and governs all downstream design, development, and testing decisions. '
  'No technology implementation details appear in this document — those are addressed in the '
  'Product Requirements Document (PRD) in Step 06.',
  size=10)
doc.add_paragraph(); rule(doc)

# ══ §3 BUSINESS PROBLEM & OPPORTUNITY ═══════════════════════════════════════
H(doc,'3. Business Problem and Opportunity',1)
H(doc,'3.1  Current State — Pain Points',2)
P(doc,
  'HR professionals and job seekers currently operate without a consolidated analytics tool '
  'for the hiring domain. The key pain points identified during stakeholder analysis are:',
  size=10)
pains=[
    'HR Managers compile hiring data manually from spreadsheets, consuming 2–4 hours per report cycle with high risk of error.',
    'Recruiters write job descriptions from memory or outdated postings, resulting in skill lists that do not reflect current market demand.',
    'Job seekers have no data-backed tool to verify salary expectations against market reality for their experience level and city.',
    'Placement officers select companies for campus drives based on tradition and personal networks, not hiring volume data.',
    'No single open-source, freely-hosted tool exists for college-level hiring trend EDA in the Indian job market context.',
]
for p in pains: bullet(doc,p)

H(doc,'3.2  Opportunity Statement',2)
opp=doc.add_table(rows=1,cols=1); bg(opp.cell(0,0),'EEF4FB')
borders(opp,'2E74B5')
p_opp=opp.cell(0,0).paragraphs[0]
p_opp.paragraph_format.left_indent=Cm(0.3)
r_opp=p_opp.add_run(
    'By building a free, interactive, web-hosted data analytics application on a 500+ record '
    'synthetic dataset, the project demonstrates how data-driven decision-making can replace '
    'manual, intuition-driven HR processes — while simultaneously meeting all ABES academic '
    'capstone requirements for the MCA Data Analytics programme.')
r_opp.font.size=Pt(10); r_opp.italic=True

H(doc,'3.3  Desired Future State',2)
futures=[
    'HR managers open the HTA app and see role-demand trends, salary benchmarks, and geographic hotspots on one screen.',
    'Recruiters filter by industry and get the top-10 skills for any JD in under 2 minutes.',
    'Job seekers verify salary expectations and optimal city with real-time chart interactions.',
    'Placement officers rank top hiring companies by domain in seconds, without a single phone call.',
    'All insights are traceable to a reproducible SQLite dataset and verifiable by pytest.',
]
for f in futures: bullet(doc,f)
doc.add_paragraph(); rule(doc)

# ══ §4 OBJECTIVES & SUCCESS METRICS ═════════════════════════════════════════
H(doc,'4. Objectives and Success Metrics (KPIs)',1)
P(doc,'Five SMART objectives are defined. Each has at least one measurable KPI with a target value.',size=9.5)
doc.add_paragraph()

obj_tbl=doc.add_table(rows=6,cols=5); obj_tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(obj_tbl,'Obj ID','Objective','KPI','Target','Measurement Method')
obj_data=[
    ('O-01',
     'Build an app that answers all 10 OQs with interactive Plotly charts',
     'OQ screens functional','10 / 10','Manual walkthrough against OQ checklist'),
    ('O-02',
     'Generate and load ≥500 synthetic hiring records into SQLite',
     'Dataset row count','≥ 500 rows','SELECT COUNT(*) on hiring table'),
    ('O-03',
     'Achieve ≥90% pytest pass rate across all functional test cases',
     'pytest pass rate','≥ 90%','pytest -v output — pass / (pass + fail)'),
    ('O-04',
     'Deploy on Streamlit Community Cloud with <3 s initial load',
     'Page load time','< 3 seconds','Browser DevTools → Network tab'),
    ('O-05',
     'Submit ABES Word project document and slide deck on time',
     'Deliverables submitted','18 / 18 files','Mentor file-checklist sign-off'),
]
for i,row in enumerate(obj_data,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(obj_tbl,i,row,bgs=[b]*5,bolds=[True,False,False,True,False])
borders(obj_tbl)
doc.add_paragraph(); rule(doc)

# ══ §5 STAKEHOLDERS — RACI ════════════════════════════════════════════════════
H(doc,'5. Stakeholder Register & RACI Matrix',1)
H(doc,'5.1  Stakeholder Register',2)
sh_tbl=doc.add_table(rows=7,cols=5); sh_tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(sh_tbl,'SH ID','Name / Persona','Role','Interest','Influence')
sh_data=[
    ('SH-01','Dinesh Sereno','Academic Mentor & Reviewer','Project quality, methodology, grades','H'),
    ('SH-02','Abhayjeet Singh','Student Developer & BA','Build, test, deploy, document','H'),
    ('SH-03','HR Manager (Persona)','Primary App User','Data-driven hiring decisions','H'),
    ('SH-04','Recruiter (Persona)','Operational App User','Skill & salary benchmarking','M'),
    ('SH-05','Job Seeker (Persona)','End Consumer','Market insights, salary verification','L'),
    ('SH-06','Placement Officer (Persona)','Institutional User','Campus drive planning','M'),
]
for i,row in enumerate(sh_data,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(sh_tbl,i,row,bgs=[b]*5,bolds=[True,False,False,False,True])
borders(sh_tbl)
doc.add_paragraph()

H(doc,'5.2  RACI Matrix',2)
P(doc,'R=Responsible  A=Accountable  C=Consulted  I=Informed',italic=True,size=9,color=MID)
doc.add_paragraph()

activities=['Requirements\nDefinition','Database\nDesign','App\nDevelopment',
            'Unit\nTesting','Deployment','Documentation','Presentation\n& Viva']
raci_data={
    'SH-01 Mentor':      ['A','A','A','A','A','A','A'],
    'SH-02 Developer':   ['R','R','R','R','R','R','R'],
    'SH-03 HR Mgr':      ['C','I','C','I','I','I','I'],
    'SH-04 Recruiter':   ['C','I','I','I','I','I','I'],
    'SH-05 Job Seeker':  ['I','I','I','I','I','I','I'],
    'SH-06 Placement':   ['C','I','I','I','I','I','I'],
}
raci_colors={'R':'D6E4F0','A':'E2EFDA','C':'FFF3CD','I':'F2F2F2'}
raci_fg    ={'R':'1F3564','A':'1F5C2E','C':'7B5700','I':'555555'}

raci_tbl=doc.add_table(rows=len(raci_data)+1,cols=len(activities)+1)
raci_tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
# header
bg(raci_tbl.cell(0,0),'1F3564')
raci_tbl.cell(0,0).paragraphs[0].add_run('Stakeholder').font.color.rgb=WHITE
raci_tbl.cell(0,0).paragraphs[0].paragraphs[0] if False else None
for ci,act in enumerate(activities,1):
    c=raci_tbl.cell(0,ci); bg(c,'1F3564')
    p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(act); r.bold=True; r.font.color.rgb=WHITE; r.font.size=Pt(8.5)

for ri,(sh,vals) in enumerate(raci_data.items(),1):
    row=raci_tbl.rows[ri]
    # stakeholder cell
    bg(row.cells[0], HEX['alt'] if ri%2==1 else HEX['white'])
    row.cells[0].paragraphs[0].add_run(sh).font.size=Pt(9)
    for ci,val in enumerate(vals,1):
        c=row.cells[ci]; bg(c,raci_colors.get(val,HEX['white']))
        p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        rv=p.add_run(val); rv.bold=True; rv.font.size=Pt(10)
        rv.font.color.rgb=RGBColor(*[int(raci_fg.get(val,'000000')[i:i+2],16) for i in (0,2,4)])
borders(raci_tbl)
P(doc,'R=Responsible (does the work)  A=Accountable (owns the outcome)  C=Consulted (input sought)  I=Informed (kept updated)',
  italic=True,size=8.5,color=MID)
doc.add_paragraph(); rule(doc)

# ══ §6 SCOPE ═════════════════════════════════════════════════════════════════
H(doc,'6. Project Scope',1)
in_scope=[
    'EDA dashboard — data distribution, missing value analysis, outlier detection',
    'Job role demand analysis — ranking by posting volume with date filtering (BR-01)',
    'Skills frequency analysis — top-N skill ranking by industry (BR-02)',
    'Salary benchmarking — avg/min/max by experience level and city (BR-03)',
    'Company hiring leaderboard — ranked by domain (BR-04)',
    'Hiring volume time series — monthly aggregation (BR-05)',
    'Cross-industry salary comparison (BR-06)',
    'Geographic posting hotspot analysis — city-level ranking (BR-07)',
    'Experience-level demand breakdown by role (BR-08)',
    'Skill-count ↔ salary correlation — scatter + OLS trendline (BR-09)',
    'Job-type distribution by industry (BR-10)',
    'Industry hiring volume by skill keyword (BR-11)',
    'Role-level posting frequency time series (BR-12)',
    'EDA summary statistics and data quality panel (BR-13)',
    'SQLite database with ≥500 synthetic Indian job-market records',
    'pytest test suite with ≥90% pass rate',
    'Streamlit Community Cloud deployment',
    'ABES project document (Word) and slide deck (PPT)',
]
out_scope=[
    'Real-time web scraping from LinkedIn, Naukri, Indeed, or any live portal',
    'User authentication, login system, or role-based access control',
    'Resume upload, parsing, or matching features',
    'Payment gateway, subscription, or premium tier',
    'Native mobile application (Android / iOS)',
    'Email notifications or alert system',
    'Multi-language (non-English) data or UI',
    'Machine-learning predictions, forecasting, or recommendation engine',
    'Admin panel, CMS, or data entry interface',
    'Integration with external HRMS, ATS, or ERP systems',
]

scope_tbl=doc.add_table(rows=1,cols=2); scope_tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
bg(scope_tbl.cell(0,0),'1F5C2E'); bg(scope_tbl.cell(0,1),'7B2D2D')
for label,col in [('✅  IN SCOPE (14 items + infra)',0),('❌  OUT OF SCOPE (10 items)',1)]:
    p=scope_tbl.cell(0,col).paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(label); r.bold=True; r.font.color.rgb=WHITE; r.font.size=Pt(10)
borders(scope_tbl)

in_t=doc.add_table(rows=len(in_scope)+1,cols=2); in_t.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(in_t,'#','In-Scope Deliverable / Feature',hbg='1F5C2E')
for i,item in enumerate(in_scope):
    frow(in_t,i+1,[str(i+1),item],
         bgs=[HEX['ok'] if i%2==0 else HEX['white']]*2)
borders(in_t)
doc.add_paragraph()

out_t=doc.add_table(rows=len(out_scope)+1,cols=2); out_t.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(out_t,'#','Out-of-Scope Item',hbg='7B2D2D')
for i,item in enumerate(out_scope):
    frow(out_t,i+1,[str(i+1),item],
         bgs=[HEX['err'] if i%2==0 else HEX['white']]*2)
borders(out_t)
doc.add_paragraph(); rule(doc)

# ══ §7 BUSINESS REQUIREMENTS ══════════════════════════════════════════════════
H(doc,'7. Business Requirements',1)
P(doc,
  'Each requirement is stated in the canonical form: "The business needs to <capability> so that <benefit>." '
  'Priority follows MoSCoW. Source HMW traces back to Step 02. '
  'Acceptance criteria are independently testable.',
  size=9.5)
doc.add_paragraph()

brs_full=[
  ('BR-01','Must','HMW-01','SH-03',
   'Rank job roles by posting volume for a user-selected time period',
   'The business needs to rank job roles by posting volume for any user-selected time period '
   'so that HR managers can identify surging or declining roles without manual data aggregation.',
   'Without this capability, HR managers must compile posting data from multiple spreadsheets, '
   'consuming 2–4 hours per cycle. A ranked, filterable view eliminates this overhead.',
   'All top-N roles displayed in correct descending order; changing the date filter updates '
   'the ranking correctly; no roles missing from the chart.'),

  ('BR-02','Must','HMW-02','SH-04',
   'Rank skills by frequency within a chosen industry',
   'The business needs to rank skills by frequency of appearance across job postings within '
   'a chosen industry so that recruiters can build data-driven job descriptions.',
   'Recruiters currently write JD skill lists from memory or old postings. A live frequency '
   'ranking reduces JD preparation time and improves applicant-role fit.',
   'Skill frequency list updates when industry filter changes; counts match raw data counts; '
   'at least top-10 skills always visible.'),

  ('BR-03','Must','HMW-03','SH-05',
   'Present salary stats (avg/min/max) by experience level and city',
   'The business needs to present average, minimum, and maximum salary figures segmented by '
   'experience level and city so that job seekers can verify whether their salary expectations '
   'are market-aligned.',
   'Published salary surveys are infrequent, broad, and lack city-level or experience-level '
   'granularity. This requirement closes that gap for the Indian job market.',
   'Three salary statistics are correct for every experience–city combination; filters work '
   'independently; values match SQL aggregation on the same filters.'),

  ('BR-04','Must','HMW-04','SH-06',
   'Rank companies by hiring volume within a user-selected domain',
   'The business needs to rank companies by total hiring volume within a user-selected domain '
   'so that placement officers can prioritise campus drive invitations with objective evidence.',
   'Campus drive company selection is currently based on tradition and personal networks. '
   'A data-backed leaderboard replaces subjectivity with a reproducible ranking.',
   'Leaderboard rank order matches descending COUNT(*) for selected domain; ties are broken '
   'consistently (alphabetical); changing domain filter updates the list correctly.'),

  ('BR-05','Must','HMW-05','SH-03',
   'Display total hiring volume aggregated by calendar month',
   'The business needs to aggregate and display total job posting volume by calendar month '
   'so that HR managers can align recruitment campaign launches with seasonal demand peaks.',
   'Without trend visibility, campaign timing is intuition-driven. Monthly volume data '
   'enables data-informed campaign planning and budget allocation.',
   'Monthly totals across all months sum to total dataset row count; time axis is '
   'chronologically ordered; zero-posting months are shown, not omitted.'),

  ('BR-06','Must','HMW-10','SH-03',
   'Compare salary statistics across all industries in a single view',
   'The business needs to compare average, minimum, and maximum offered salaries across all '
   'industries in a consolidated single view so that HR managers can benchmark compensation '
   'packages without consulting external salary surveys.',
   'Third-party salary surveys are expensive or paywalled. A consolidated in-app comparison '
   'provides immediate, free, dataset-grounded benchmarking.',
   'Stats per industry are correct; at least 3 distinct industries visible; sorted by '
   'average salary descending by default.'),

  ('BR-07','Must','HMW-07','SH-05',
   'Present job posting counts ranked by city',
   'The business needs to present job posting counts ranked by city so that job seekers '
   'and recruiters can identify geographic hiring hotspots at a glance.',
   'Job seekers relocating to a new city have no quick way to compare opportunity density. '
   'A ranked city view makes the geographic decision data-driven.',
   'City counts are correct and sorted descending; top-10 cities visible by default; '
   'toggling a filter correctly updates the chart.'),

  ('BR-08','Must','HMW-09','SH-06',
   'Show experience-level demand split (Fresher/Mid/Senior) per job role',
   'The business needs to show the proportion of job postings demanding each experience '
   'level broken down by job role so that placement officers can match the right student '
   'batch profile to the right recruiter.',
   'Placement officers currently have no data to show recruiters why a specific student '
   'cohort is a good fit. An experience-demand breakdown provides that evidence.',
   "Each role's experience split sums to 100%; filtering by role updates the chart; "
   'all three experience levels appear in the legend even if one has 0%.'),

  ('BR-09','Should','HMW-06','SH-04',
   'Quantify correlation between skills count and maximum salary',
   'The business needs to quantify the statistical relationship between the number of '
   'skills required per posting and the maximum salary offered so that recruiters can '
   'price complex, multi-skill roles competitively.',
   'Salary bands are currently set by intuition. A correlation analysis provides an '
   'objective basis for pricing roles that demand broader skill sets.',
   'Correlation coefficient value matches pandas .corr() result; scatter plot shows '
   'correct (skills_count, salary_max) pairs; OLS trendline direction is correct.'),

  ('BR-10','Should','HMW-08','SH-04',
   'Present job-type distribution (Remote/Hybrid/FT/PT/Contract) per industry',
   'The business needs to present the distribution of job types for each industry so that '
   'recruiters can benchmark their flexible-work offerings against market norms.',
   'Flexible-work adoption varies widely by industry. Without a comparison view, '
   'recruiters cannot tell if their policies are competitive.',
   "Job-type proportions per industry sum to 100%; each type in correct slice; "
   'industry filter updates the chart correctly.'),

  ('BR-11','Should','HMW-11','SH-06',
   'Rank industries by hiring volume for a user-selected skill keyword',
   'The business needs to rank industries by total hiring volume for a user-selected '
   'skill keyword so that career counsellors can direct students toward emerging sectors.',
   'Students choosing a specialisation domain have no data on which industries are '
   'absorbing that skill most rapidly. This ranking fills that gap.',
   'Industry ranking changes when skill filter changes; volume counts match filtered '
   'row count; keyword matching is case-insensitive.'),

  ('BR-12','Should','HMW-12','SH-05',
   'Show posting frequency time series for a specific job role',
   'The business needs to display a time-series view of posting frequency for a specific '
   'user-selected job role so that job seekers can identify peak hiring windows and '
   'time their applications accordingly.',
   'Job application timing is currently random. A role-level time series reveals seasonal '
   'patterns, enabling job seekers to apply during high-activity windows.',
   'Monthly counts for the selected role match filtered row count; zero-posting months '
   'shown as gaps; changing role selection updates the chart correctly.'),

  ('BR-13','Could','— (Assumed)','SH-02',
   'Present EDA summary statistics and data-quality panel for all numeric fields',
   'The business needs to present summary statistics (count, mean, median, std deviation, '
   'missing-value count, outlier flags) for all numeric fields so that analysts can '
   'rapidly assess dataset quality and distribution before interpreting any chart.',
   'Without a data-quality panel, analysts may misinterpret charts built on fields '
   'with significant missing values or skewed distributions.',
   'Stats match pandas .describe() output; missing-value count matches actual nulls '
   'in dataset; outlier count uses IQR method consistently.'),
]

for bid,prio,hmw,sh,short,req,rationale,ac in brs_full:
    mbg=MOSCOW_BG.get(prio,HEX['white'])
    mfg=MOSCOW_FG.get(prio,'000000')
    # BR card
    card=doc.add_table(rows=6,cols=2); card.alignment=WD_TABLE_ALIGNMENT.CENTER
    # header spanning both columns
    card.cell(0,0).merge(card.cell(0,1))
    bg(card.cell(0,0),'1F3564')
    hp=card.cell(0,0).paragraphs[0]; hp.alignment=WD_ALIGN_PARAGRAPH.LEFT
    hr_=hp.add_run(f'  {bid}  —  {short}')
    hr_.bold=True; hr_.font.color.rgb=WHITE; hr_.font.size=Pt(11)

    card_data=[
        ('Requirement',    req),
        ('Rationale',      rationale),
        ('Stakeholder',    f'{sh}  |  Source HMW: {hmw}  |  Priority: {prio}'),
        ('Acceptance Criterion', ac),
        ('Status',         'Identified — Step 04 RTM'),
    ]
    fills=['EEF4FB',HEX['white'],mbg,'FFF8F0','E5FFE5']
    for ri,(label,val) in enumerate(card_data,1):
        bg(card.cell(ri,0),fills[ri-1]); bg(card.cell(ri,1),fills[ri-1])
        card.cell(ri,0).width=Cm(3.5); card.cell(ri,1).width=Cm(14)
        lr=card.cell(ri,0).paragraphs[0].add_run(label)
        lr.bold=True; lr.font.size=Pt(9.5)
        vr=card.cell(ri,1).paragraphs[0].add_run(val)
        vr.font.size=Pt(9.5)
        if label=='Stakeholder':
            vr.bold=True
            vr.font.color.rgb=RGBColor(*[int(mfg[i:i+2],16) for i in (0,2,4)])
    borders(card)
    doc.add_paragraph()

rule(doc)

# ══ §8 BUSINESS RULES ════════════════════════════════════════════════════════
H(doc,'8. Business Rules',1)
P(doc,
  'Business rules are constraints on data and behaviour that must be enforced regardless '
  'of the technology implementation. They are distinct from business requirements.',
  size=9.5)
doc.add_paragraph()

brules=[
    ('BR-R-01','Data Source',
     'All values displayed in the app must originate from the SQLite database. '
     'No hardcoded or mock values may appear in production charts.',
     'High'),
    ('BR-R-02','Currency',
     'All salary figures must be presented in Indian Rupees (₹ INR). '
     'No conversion to other currencies is permitted.',
     'High'),
    ('BR-R-03','Date Range',
     'All synthetic job postings must fall within 01-Jan-2022 to 31-Dec-2024. '
     'No posting date outside this range is valid.',
     'High'),
    ('BR-R-04','Experience Levels',
     'Experience level must be exactly one of three values: Fresher, Mid-level, Senior. '
     'No other values, synonyms, or null entries are permitted.',
     'High'),
    ('BR-R-05','Job Types',
     'Job type must be exactly one of five values: Full-time, Part-time, Remote, Hybrid, Contract. '
     'No other values are permitted.',
     'High'),
    ('BR-R-06','Chart Responsiveness',
     'All charts must visually update within 2 seconds of a filter change on a standard broadband connection.',
     'Medium'),
    ('BR-R-07','Data Privacy',
     'The dataset must contain zero real personal data. All names, companies, and locations '
     'must be synthetically generated. GDPR / Indian PDPB compliance is mandatory.',
     'High'),
    ('BR-R-08','Skills Per Posting',
     'Each job posting must have between 1 and 8 skills. Skills are stored as a '
     'pipe-delimited string (e.g., "Python|SQL|Tableau"). No posting may have 0 or >8 skills.',
     'Medium'),
    ('BR-R-09','Salary Validity',
     'salary_min must always be < salary_max for every record. '
     'No posting may have salary_min ≥ salary_max.',
     'High'),
    ('BR-R-10','Minimum Dataset Size',
     'The SQLite database must contain ≥500 hiring records before the app is launched. '
     'The app must refuse to start (show an error) if the count is below 500.',
     'Medium'),
]

br_tbl=doc.add_table(rows=len(brules)+1,cols=4); br_tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(br_tbl,'Rule ID','Category','Business Rule','Criticality')
crit_bg={'High':'FFE5E5','Medium':'FFF3CD','Low':'E5FFE5'}
for i,(rid,cat,rule_,crit) in enumerate(brules,1):
    b=crit_bg.get(crit,HEX['white'])
    frow(br_tbl,i,[rid,cat,rule_,crit],
         bgs=[b,b,b,b],bolds=[True,False,False,True],
         aligns=[WD_ALIGN_PARAGRAPH.CENTER,WD_ALIGN_PARAGRAPH.LEFT,
                 WD_ALIGN_PARAGRAPH.LEFT,WD_ALIGN_PARAGRAPH.CENTER])
borders(br_tbl)
doc.add_paragraph(); rule(doc)

# ══ §9 DATA REQUIREMENTS ═════════════════════════════════════════════════════
H(doc,'9. Data Requirements',1)
H(doc,'9.1  Dataset Overview',2)
dr_meta=[
    ('Dataset Name','hiring_trends'),
    ('Format','SQLite (.sqlite) + CSV (.csv) companion file'),
    ('Source','Synthetically generated — Python Faker + custom distribution rules'),
    ('Volume','≥500 records (target: 1 000 for richer analytics)'),
    ('Refresh Frequency','One-time static load; no live refresh in this release'),
    ('Date Range','01 Jan 2022 — 31 Dec 2024'),
    ('Geography','India — 15 major cities across 6 states'),
    ('Language','English only'),
]
dro=doc.add_table(rows=len(dr_meta),cols=2)
for i,(k,v) in enumerate(dr_meta):
    b=HEX['alt'] if i%2==0 else HEX['white']
    bg(dro.cell(i,0),b); bg(dro.cell(i,1),b)
    dro.cell(i,0).paragraphs[0].add_run(k).font.size=Pt(9.5)
    dro.cell(i,0).paragraphs[0].runs[0].bold=True
    dro.cell(i,1).paragraphs[0].add_run(v).font.size=Pt(9.5)
borders(dro)
doc.add_paragraph()

H(doc,'9.2  Field Dictionary',2)
fields=[
    ('job_id',       'INTEGER','PK','Unique job posting identifier','Auto-increment'),
    ('job_title',    'TEXT',   'NOT NULL','Job role name (e.g., Data Analyst)','12 distinct roles'),
    ('company_name', 'TEXT',   'NOT NULL','Hiring company name (synthetic)','50 synthetic companies'),
    ('industry',     'TEXT',   'NOT NULL','Industry sector','8 distinct sectors'),
    ('city',         'TEXT',   'NOT NULL','Posting city','15 Indian cities'),
    ('state',        'TEXT',   'NOT NULL','Posting state','6 states'),
    ('job_type',     'TEXT',   'NOT NULL','Employment type','Exactly 5 allowed values (BR-R-05)'),
    ('experience_level','TEXT','NOT NULL','Required experience band','Exactly 3 values (BR-R-04)'),
    ('required_skills','TEXT', 'NOT NULL','Pipe-delimited skill list','1–8 skills per row (BR-R-08)'),
    ('salary_min',   'REAL',   'NOT NULL','Minimum offered salary (₹ INR)','salary_min < salary_max (BR-R-09)'),
    ('salary_max',   'REAL',   'NOT NULL','Maximum offered salary (₹ INR)','Realistic Indian market ranges'),
    ('posted_date',  'TEXT',   'NOT NULL','Posting date (ISO 8601)','Within 2022-01-01 to 2024-12-31 (BR-R-03)'),
]
fd=doc.add_table(rows=len(fields)+1,cols=5); fd.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(fd,'Field','Type','Constraint','Description','Notes / Validation')
for i,row in enumerate(fields,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(fd,i,row,bgs=[b]*5,bolds=[True,False,False,False,False])
borders(fd)
doc.add_paragraph(); rule(doc)

# ══ §10 ASSUMPTIONS, CONSTRAINTS, DEPENDENCIES ═══════════════════════════════
H(doc,'10. Assumptions, Constraints and Dependencies',1)
H(doc,'10.1  Assumptions',2)
assm=[
    ('A-01','Synthetic data','Faker-generated data adequately represents Indian job-market distributions for roles, cities, and salaries.'),
    ('A-02','SQLite scale','SQLite is sufficient for ≤10 000 rows at college-project scale; no performance issues expected.'),
    ('A-03','Free hosting','Streamlit Community Cloud free tier provides adequate uptime for academic demo purposes.'),
    ('A-04','Static dataset','All charts are built on a snapshot dataset; no live data feed is assumed or required.'),
    ('A-05','Desktop browser','The app is designed for desktop/laptop browsers; mobile responsiveness is desirable but not required.'),
    ('A-06','English only','All data labels, field values, and UI copy are in English.'),
    ('A-07','Mentor availability','Dinesh Sereno will review and provide feedback at each major milestone step within 3 working days.'),
]
at=doc.add_table(rows=len(assm)+1,cols=3); at.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(at,'ID','Category','Assumption Statement')
for i,(aid,cat,stmt) in enumerate(assm,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(at,i,[aid,cat,stmt],bgs=[b]*3,bolds=[True,False,False])
borders(at)
doc.add_paragraph()

H(doc,'10.2  Constraints',2)
cons=[
    ('C-01','Budget',    'Zero — all tools, libraries, and hosting must be free-tier / open-source.',    'Use Faker, Streamlit, SQLite, pytest — all zero-cost'),
    ('C-02','Time',      'Academic semester; target completion within 8 weeks from Step 01.',              'Weekly milestone checkpoints with mentor'),
    ('C-03','Team',      'Solo developer — Abhayjeet Singh; no additional engineering resources.',         'Modular code design; reuse templates across steps'),
    ('C-04','Data',      'Synthetic data only — no real personal or company data (privacy compliance).', 'Document data-generation script for reproducibility'),
    ('C-05','Tech Stack','Python 3.11+, Streamlit, Plotly, SQLite, pytest — no paid frameworks.',         'Well-supported open-source stack'),
    ('C-06','Hosting',   'Streamlit Community Cloud free tier only; no paid cloud (AWS/GCP/Azure).',      'Maintain local run script as backup demo option'),
    ('C-07','Docs',      'Must follow ABES Word and PPT templates.',                                      'Templates obtained from ABES in Step 17'),
]
ct=doc.add_table(rows=len(cons)+1,cols=4); ct.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(ct,'ID','Category','Constraint','Mitigation')
for i,row in enumerate(cons,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(ct,i,row,bgs=[b]*4,bolds=[True,False,False,False])
borders(ct)
doc.add_paragraph()

H(doc,'10.3  Dependencies',2)
deps=[
    ('D-01','External','Python 3.11+','Required runtime; must be installed before development begins'),
    ('D-02','External','Streamlit ≥1.30','App framework; version pinned in requirements.txt'),
    ('D-03','External','Plotly ≥5.18','Charting library; version pinned in requirements.txt'),
    ('D-04','External','Faker ≥24.0','Synthetic data generation in Step 10'),
    ('D-05','External','pytest ≥8.0','Test runner for Step 13 and Step 15'),
    ('D-06','Internal','Step 10 (Dataset)','App cannot be built until ≥500 records exist in SQLite'),
    ('D-07','Internal','Step 08 (DB Design)','Dataset schema must be finalised before Step 10'),
    ('D-08','Platform','Streamlit Community Cloud','GitHub repo must be public for free-tier deployment'),
]
dt=doc.add_table(rows=len(deps)+1,cols=4); dt.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(dt,'ID','Type','Dependency','Impact if Unavailable')
for i,row in enumerate(deps,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(dt,i,row,bgs=[b]*4,bolds=[True,False,False,False])
borders(dt)
doc.add_paragraph(); rule(doc)

# ══ §11 RISKS AND MITIGATION ══════════════════════════════════════════════════
H(doc,'11. Risks and Mitigation',1)
risks=[
    ('R-01','Synthetic data lacks realism',   'M','H','M',
     'Use realistic Indian city names, salary ranges, and skill distributions; validate with mentor.'),
    ('R-02','Streamlit free-tier downtime during demo','L','H','M',
     'Maintain a local run script (streamlit run app.py) as a demo fallback.'),
    ('R-03','Scope creep beyond 18 steps',    'M','M','M',
     'Strict MoSCoW prioritisation; any new feature requires mentor approval and RTM update.'),
    ('R-04','pytest coverage gaps',           'M','M','M',
     'Write test cases in Step 13 before coding in Step 15 (test-first approach).'),
    ('R-05','Library version conflicts',      'L','M','L',
     'Pin all library versions in requirements.txt; test on clean virtual environment.'),
    ('R-06','Solo developer single point of failure','H','H','H',
     'Maintain detailed commit history and documentation so mentor can assess progress independently.'),
    ('R-07','Academic deadline slip',         'M','H','H',
     'Weekly step checkpoints; Step 15 (Build) must start no later than Step 10 completion + 1 week.'),
    ('R-08','GitHub public repo exposure',    'L','L','L',
     'No real data in repo; dataset is synthetic; no credentials committed.'),
]

risk_tbl=doc.add_table(rows=len(risks)+1,cols=7)
risk_tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(risk_tbl,'Risk ID','Risk Description','Probability\n(H/M/L)',
    'Impact\n(H/M/L)','Overall\nRating','Mitigation Strategy','Residual\nRisk')

risk_bg={'H':'FFE5E5','M':'FFF3CD','L':'E5FFE5'}
risk_fg={'H':'C00000','M':'7B5700','L':'1F5C2E'}
for i,(rid,desc,prob,imp,rating,mit) in enumerate(risks,1):
    row=risk_tbl.rows[i]
    base=HEX['alt'] if i%2==1 else HEX['white']
    vals=[rid,desc,prob,imp,rating,mit,'Low']
    for ci,val in enumerate(vals):
        c=row.cells[ci]
        is_rating=(ci in (2,3,4,6))
        b=risk_bg.get(val,base) if is_rating else base
        fg=risk_fg.get(val,'000000') if is_rating else '000000'
        bg(c,b)
        p=c.paragraphs[0]
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER if is_rating or ci==0 else WD_ALIGN_PARAGRAPH.LEFT
        r=p.add_run(str(val)); r.font.size=Pt(9)
        if is_rating or ci==0: r.bold=True
        if is_rating:
            r.font.color.rgb=RGBColor(*[int(fg[i2:i2+2],16) for i2 in (0,2,4)])
borders(risk_tbl)
doc.add_paragraph(); rule(doc)

# ══ §12 GLOSSARY ══════════════════════════════════════════════════════════════
H(doc,'12. Glossary',1)
glossary=[
    ('ATS',      'Applicant Tracking System — software for managing recruitment workflows.'),
    ('BA',       'Business Analyst — the role responsible for eliciting and documenting requirements.'),
    ('BRD',      'Business Requirements Document — this document; defines WHAT the business needs.'),
    ('BR',       'Business Requirement — a statement of a capability needed by a stakeholder.'),
    ('EDA',      'Exploratory Data Analysis — statistical summarisation and visualisation of a dataset.'),
    ('HMW',      'How Might We — a design-thinking question format that reframes pain points as opportunities.'),
    ('HTA',      'Hiring Trend Analysis — the name and shortcode of this project.'),
    ('INR / ₹',  'Indian Rupee — the currency used for all salary figures in the dataset.'),
    ('JD',       'Job Description — the document posted by a recruiter to attract candidates.'),
    ('KPI',      'Key Performance Indicator — a measurable value that tracks progress toward an objective.'),
    ('MoSCoW',   'Must / Should / Could / Won\'t — a priority framework for requirements.'),
    ('OLS',      'Ordinary Least Squares — a statistical method for fitting a linear trendline.'),
    ('OQ',       'Operational Question — a business question the app must answer (OQ-01 to OQ-10).'),
    ('PRD',      'Product Requirements Document — Step 06 document; defines HOW the product is built.'),
    ('RACI',     'Responsible / Accountable / Consulted / Informed — a stakeholder responsibility matrix.'),
    ('RTM',      'Requirements Traceability Matrix — the Excel file (Step 04) linking OQ→HMW→BR→Feature→Test.'),
    ('SQLite',   'A lightweight, file-based relational database; used as the data store for this project.'),
    ('Streamlit','An open-source Python framework for building interactive data web applications.'),
    ('SH',       'Stakeholder — a person or persona with an interest in the project outcome.'),
]
glo=doc.add_table(rows=len(glossary)+1,cols=2); glo.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(glo,'Term / Acronym','Definition')
for i,(term,defn) in enumerate(glossary,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    frow(glo,i,[term,defn],bgs=[b,b],bolds=[True,False])
borders(glo)
doc.add_paragraph(); rule(doc)

# ══ §13 SIGN-OFF ══════════════════════════════════════════════════════════════
H(doc,'13. Sign-off',1)
P(doc,
  'The following parties confirm that this BRD accurately and completely represents the '
  'business requirements for the Hiring Trend Analysis project as of the date signed.',
  size=10)
doc.add_paragraph()

sign=doc.add_table(rows=4,cols=5); sign.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr(sign,'Role','Name','Signature','Date','Status')
sign_data=[
    ('Prepared By','Abhayjeet Singh','_______________',
     datetime.date.today().strftime('%d %b %Y'),'Submitted'),
    ('Reviewed By','Dinesh Sereno (Mentor)','_______________','___________','Pending'),
    ('Approved By','Dinesh Sereno (Mentor)','_______________','___________','Pending'),
]
for i,row in enumerate(sign_data,1):
    b=HEX['alt'] if i%2==1 else HEX['white']
    st_bg=HEX['ok'] if row[4]=='Submitted' else HEX['warn']
    bgs=[b,b,b,b,st_bg]
    frow(sign,i,row,bgs=bgs,bolds=[True,False,False,False,True])
borders(sign)

doc.add_paragraph()
# footer
fp=doc.add_paragraph(); fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
fr=fp.add_run('P071 – Hiring Trend Analysis  |  BRD v1.0  |  Step 05 of 18  |  ABES MCA 2nd Yr DA  |  Mentor: Dinesh Sereno')
fr.font.size=Pt(8); fr.font.color.rgb=RGBColor(0x88,0x88,0x88); fr.italic=True

doc.save(OUT_FILE)
print(f"SAVED: {OUT_FILE}")
