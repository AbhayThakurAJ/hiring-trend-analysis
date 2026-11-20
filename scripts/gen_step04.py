"""
Step 04 – Requirements Traceability Matrix (RTM) Excel Generator
Project : P01 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES, MCA 2nd Yr DA
"""

import os, datetime
from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

OUT_DIR  = r"C:\Hiring Trend Analysis\docs"
OUT_FILE = os.path.join(OUT_DIR, "P01_HTA_Step04_RTM.xlsx")
os.makedirs(OUT_DIR, exist_ok=True)

# ── colour palette ────────────────────────────────────────────────────────────
C = dict(
    dark_blue ='1F3564', mid_blue  ='2E74B5', steel     ='4A86C8',
    must      ='D6E4F0', should    ='E2EFDA', could     ='FFF3CD',
    wont      ='F4CCCC', alt_row   ='EEF4FB', white     ='FFFFFF',
    ok_green  ='E5FFE5', warn_amber='FFF3CD', err_red   ='FFE5E5',
    hdr_txt   ='FFFFFF', empty_col ='F5F5F5', gold      ='FFD700',
    grey_light='F2F2F2', border_c  ='2E74B5',
)

MOSCOW_BG = {'Must':'D6E4F0','Should':'E2EFDA','Could':'FFF3CD',"Won't":'F4CCCC'}
MOSCOW_FG = {'Must':'1F3564','Should':'1F5C2E','Could':'7B5700',"Won't":'7B2D2D'}

def fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)

def font(bold=False, color='000000', size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic,
                name='Calibri')

def align(h='left', v='center', wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def thin_border(color='2E74B5'):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def med_border(color='1F3564'):
    s = Side(style='medium', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def style_cell(cell, bg=None, fg='000000', bold=False, size=10,
               h_align='left', v_align='center', wrap=True,
               italic=False, border_color='2E74B5'):
    if bg:
        cell.fill = fill(bg)
    cell.font      = font(bold=bold, color=fg, size=size, italic=italic)
    cell.alignment = align(h=h_align, v=v_align, wrap=wrap)
    cell.border    = thin_border(border_color)

# ═════════════════════════════════════════════════════════════════════════════
wb = Workbook()

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1  — RTM (main)
# ══════════════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "RTM"
ws.sheet_view.showGridLines = False
ws.freeze_panes = 'A4'          # freeze title rows

# ── Title block (rows 1-2) ────────────────────────────────────────────────────
ws.merge_cells('A1:N1')
t1 = ws['A1']
t1.value = 'P01 – HIRING TREND ANALYSIS  |  Requirements Traceability Matrix (RTM)  |  Step 04'
style_cell(t1, bg=C['dark_blue'], fg=C['hdr_txt'], bold=True, size=13,
           h_align='center', border_color=C['dark_blue'])
ws.row_dimensions[1].height = 28

ws.merge_cells('A2:N2')
t2 = ws['A2']
t2.value = (f"Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA"
            f"  |  Generated: {datetime.date.today().strftime('%d %b %Y')}  |  Status: Identified")
style_cell(t2, bg=C['steel'], fg=C['hdr_txt'], bold=False, size=9,
           h_align='center', border_color=C['dark_blue'])
ws.row_dimensions[2].height = 18

# ── Column headers (row 3) ────────────────────────────────────────────────────
HEADERS = [
    'OQ ID(s)',          # A
    'HMW ID',            # B
    'BR ID',             # C
    'Business Requirement (short)',  # D
    'Full BR Statement', # E
    'Stakeholder',       # F
    'Priority (MoSCoW)', # G
    'Status',            # H
    'Feature ID',        # I  — filled in Step 06
    'DB Table.Column',   # J  — filled in Step 08
    'User Story ID',     # K  — filled in Step 11
    'UI Screen ID',      # L  — filled in Step 12
    'TC-F (Functional)', # M  — filled in Step 13
    'TC-T (Technical)',  # N  — filled in Step 13
]
STEP_NOTES = {
    'I': 'Step 06', 'J': 'Step 08',
    'K': 'Step 11', 'L': 'Step 12',
    'M': 'Step 13', 'N': 'Step 13',
}

for col_idx, hdr in enumerate(HEADERS, start=1):
    cell = ws.cell(row=3, column=col_idx, value=hdr)
    col_letter = get_column_letter(col_idx)
    is_future = col_letter in STEP_NOTES
    bg  = C['grey_light'] if is_future else C['mid_blue']
    fg  = '888888'        if is_future else C['hdr_txt']
    style_cell(cell, bg=bg, fg=fg, bold=True, size=9,
               h_align='center', border_color=C['dark_blue'])
ws.row_dimensions[3].height = 32

# ── Column widths ─────────────────────────────────────────────────────────────
col_widths = {
    'A':12,'B':10,'C':9,'D':28,'E':52,'F':13,
    'G':13,'H':13,'I':11,'J':18,'K':12,'L':12,'M':14,'N':14
}
for col_letter, width in col_widths.items():
    ws.column_dimensions[col_letter].width = width

# ── RTM data rows ─────────────────────────────────────────────────────────────
rtm_data = [
    # (oq_ids, hmw, br_id, short_req, full_req, sh, priority)
    ('OQ-01; OQ-04','HMW-01','BR-01',
     'Rank job roles by posting volume (date filter)',
     'The business needs to rank job roles by posting volume for any user-selected time period '
     'so that HR managers can identify surging or declining roles without manual data aggregation.',
     'SH-03','Must'),

    ('OQ-03','HMW-02','BR-02',
     'Rank skills by frequency within a chosen industry',
     'The business needs to rank skills by frequency of appearance across job postings within '
     'a chosen industry so that recruiters can build data-driven job descriptions.',
     'SH-04','Must'),

    ('OQ-02; OQ-08','HMW-03','BR-03',
     'Salary stats by experience level and city',
     'The business needs to present average, minimum, and maximum salary figures segmented by '
     'experience level and city so that job seekers can verify whether their expectations are market-aligned.',
     'SH-05','Must'),

    ('OQ-10','HMW-04','BR-04',
     'Company hiring leaderboard by domain',
     'The business needs to rank companies by total hiring volume within a user-selected domain '
     'so that placement officers can prioritise campus drive invitations with objective evidence.',
     'SH-06','Must'),

    ('OQ-04','HMW-05','BR-05',
     'Monthly hiring volume time series',
     'The business needs to aggregate and display total job posting volume by calendar month '
     'so that HR managers can align recruitment campaigns with seasonal demand peaks.',
     'SH-03','Must'),

    ('OQ-02; OQ-07','HMW-10','BR-06',
     'Cross-industry salary comparison (avg/min/max)',
     'The business needs to compare average, minimum, and maximum offered salaries across all '
     'industries in a consolidated single view so that HR managers can benchmark compensation packages.',
     'SH-03','Must'),

    ('OQ-05','HMW-07','BR-07',
     'City-wise posting concentration ranking',
     'The business needs to present job posting counts ranked by city so that job seekers '
     'and recruiters can identify geographic hiring hotspots at a glance.',
     'SH-05','Must'),

    ('OQ-08','HMW-09','BR-08',
     'Experience-level demand split per job role',
     'The business needs to show the proportion of job postings demanding each experience level '
     '(Fresher / Mid-level / Senior) broken down by job role so that placement officers can '
     'match the right student batch profile to the right recruiter.',
     'SH-06','Must'),

    ('OQ-09','HMW-06','BR-09',
     'Skill count ↔ salary correlation (OLS)',
     'The business needs to quantify the statistical relationship between the number of skills '
     'required per posting and the maximum salary offered so that recruiters can price complex roles competitively.',
     'SH-04','Should'),

    ('OQ-06','HMW-08','BR-10',
     'Job-type distribution per industry',
     'The business needs to present the distribution of job types (Full-time, Part-time, Remote, '
     'Hybrid, Contract) for each industry so that recruiters can benchmark their flexible-work policies.',
     'SH-04','Should'),

    ('OQ-07; OQ-04','HMW-11','BR-11',
     'Industry hiring volume by skill keyword + time',
     'The business needs to rank industries by total hiring volume for a user-selected skill keyword '
     'over time so that career counsellors can direct students toward emerging sectors.',
     'SH-06','Should'),

    ('OQ-01; OQ-04','HMW-12','BR-12',
     'Role-level posting frequency time series',
     'The business needs to display a time-series view of posting frequency for a specific '
     'user-selected job role so that job seekers can identify peak hiring windows.',
     'SH-05','Should'),

    ('—','— (Assumed)','BR-13',
     'EDA summary statistics + data quality check',
     'The business needs to present summary statistics (count, mean, median, std dev, missing-value '
     'count, outlier flags) for all numeric fields so that analysts can assess dataset quality.',
     'SH-02','Could'),
]

for row_idx, (oq,hmw,br,short,full,sh,prio) in enumerate(rtm_data, start=4):
    is_alt = (row_idx % 2 == 0)
    prio_bg = MOSCOW_BG.get(prio, C['white'])
    prio_fg = MOSCOW_FG.get(prio, '000000')
    base_bg = C['alt_row'] if is_alt else C['white']
    # future columns
    future_bg = C['empty_col']

    vals = [oq, hmw, br, short, full, sh, prio, 'Identified',
            '—', '—', '—', '—', '—', '—']

    for ci, val in enumerate(vals, start=1):
        col_letter = get_column_letter(ci)
        cell = ws.cell(row=row_idx, column=ci, value=val)
        is_future_col = col_letter in STEP_NOTES

        if ci == 7:   # priority
            style_cell(cell, bg=prio_bg, fg=prio_fg, bold=True, size=9, h_align='center')
        elif ci == 8: # status
            style_cell(cell, bg='E5FFE5', fg='1F5C2E', bold=True, size=9, h_align='center')
        elif ci in (1,2,3,6): # IDs & SH
            style_cell(cell, bg=prio_bg if ci==3 else base_bg,
                       fg=prio_fg if ci==3 else '000000',
                       bold=(ci==3), size=9, h_align='center')
        elif is_future_col:
            style_cell(cell, bg=future_bg, fg='AAAAAA', size=9,
                       h_align='center', italic=True)
            # add comment note about which step fills it
            step_note = STEP_NOTES.get(col_letter,'')
            cell.value = f'← {step_note}'
        else:
            style_cell(cell, bg=base_bg, size=9)

    ws.row_dimensions[row_idx].height = 48

# ── Auto-filter on row 3 ──────────────────────────────────────────────────────
ws.auto_filter.ref = f'A3:N{3+len(rtm_data)}'

# ── Legend below data ─────────────────────────────────────────────────────────
legend_row = 4 + len(rtm_data) + 1
ws.merge_cells(f'A{legend_row}:N{legend_row}')
lc = ws[f'A{legend_row}']
lc.value = 'LEGEND  |  Priority: Must (blue)  Should (green)  Could (amber)  Won\'t (red)  |  Grey columns are populated in later Steps (06, 08, 11, 12, 13)  |  Status "Identified" = baseline; updated each step'
style_cell(lc, bg=C['grey_light'], fg='555555', size=8, italic=True, h_align='left')

# ── Freeze note row ───────────────────────────────────────────────────────────
note_row = legend_row + 1
ws.merge_cells(f'A{note_row}:N{note_row}')
nc = ws[f'A{note_row}']
nc.value = ('RTM UPDATE SCHEDULE: Step 06 → Feature ID  |  Step 08 → DB Table.Column  |  '
            'Step 11 → User Story ID  |  Step 12 → UI Screen ID  |  Step 13 → TC-F / TC-T')
style_cell(nc, bg=C['dark_blue'], fg=C['hdr_txt'], size=8, bold=True, h_align='center')

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2  — Coverage Check
# ══════════════════════════════════════════════════════════════════════════════
wc = wb.create_sheet("Coverage Check")
wc.sheet_view.showGridLines = False

# Title
wc.merge_cells('A1:F1')
style_cell(wc['A1'], bg=C['dark_blue'], fg=C['hdr_txt'], bold=True, size=13,
           h_align='center', border_color=C['dark_blue'])
wc['A1'].value = 'P01 – RTM Coverage Check  |  Step 04'
wc.row_dimensions[1].height = 26

# ── Section A: OQ Coverage ────────────────────────────────────────────────────
wc.merge_cells('A3:F3')
style_cell(wc['A3'], bg=C['mid_blue'], fg=C['hdr_txt'], bold=True, size=11, h_align='center')
wc['A3'].value = 'A — OQ Coverage: Every Operational Question covered by ≥1 HMW and ≥1 BR'
wc.row_dimensions[3].height = 22

oq_cov_hdrs = ['OQ ID','Operational Question (short)','Covered by HMW(s)','Covered by BR(s)','OQ ↔ HMW','OQ ↔ BR']
for ci, h in enumerate(oq_cov_hdrs, 1):
    c = wc.cell(row=4, column=ci, value=h)
    style_cell(c, bg=C['steel'], fg=C['hdr_txt'], bold=True, size=9, h_align='center')
wc.row_dimensions[4].height = 20

oq_cov = [
    ('OQ-01','Highest-posting-volume job roles','HMW-01, HMW-12','BR-01, BR-12','✔','✔'),
    ('OQ-02','Avg/min/max salary per industry', 'HMW-03, HMW-10','BR-03, BR-06','✔','✔'),
    ('OQ-03','Most-demanded skills',            'HMW-02',         'BR-02',       '✔','✔'),
    ('OQ-04','Month-on-month hiring trend',     'HMW-01, HMW-05, HMW-11, HMW-12','BR-01, BR-05, BR-11, BR-12','✔','✔'),
    ('OQ-05','City-wise posting hotspots',      'HMW-07',         'BR-07',       '✔','✔'),
    ('OQ-06','Job-type distribution',           'HMW-08',         'BR-10',       '✔','✔'),
    ('OQ-07','Top-hiring industries',           'HMW-10, HMW-11', 'BR-06, BR-11','✔','✔'),
    ('OQ-08','Experience-level demand',         'HMW-03, HMW-09', 'BR-03, BR-08','✔','✔'),
    ('OQ-09','Skill count ↔ salary correlation','HMW-06',         'BR-09',       '✔','✔'),
    ('OQ-10','Top recruiting companies',        'HMW-04',         'BR-04',       '✔','✔'),
]
for i, row_data in enumerate(oq_cov, start=5):
    for ci, val in enumerate(row_data, 1):
        c = wc.cell(row=i, column=ci, value=val)
        is_ok = val in ('✔','✔')
        bg = C['ok_green'] if (ci>=5 and '✔' in str(val)) else (C['alt_row'] if i%2==0 else C['white'])
        fg = '1F5C2E' if (ci>=5 and '✔' in str(val)) else '000000'
        style_cell(c, bg=bg, fg=fg, bold=(ci>=5), size=9, h_align='center' if ci!=2 else 'left')
    wc.row_dimensions[i].height = 18

# ── Section B: HMW Coverage ───────────────────────────────────────────────────
b_start = 5 + len(oq_cov) + 2
wc.merge_cells(f'A{b_start}:F{b_start}')
style_cell(wc[f'A{b_start}'], bg=C['mid_blue'], fg=C['hdr_txt'], bold=True, size=11, h_align='center')
wc[f'A{b_start}'].value = 'B — HMW Coverage: Every HMW maps to ≥1 BR'
wc.row_dimensions[b_start].height = 22

hmw_hdrs = ['HMW ID','HMW Statement (short)','Maps to BR','Status','Gap?','Resolution']
for ci, h in enumerate(hmw_hdrs, 1):
    c = wc.cell(row=b_start+1, column=ci, value=h)
    style_cell(c, bg=C['steel'], fg=C['hdr_txt'], bold=True, size=9, h_align='center')

hmw_cov = [
    ('HMW-01','Show HR which roles spiked/dropped in chosen window','BR-01','Covered','✔ No gap','—'),
    ('HMW-02','Show recruiter top-10 skills per industry',          'BR-02','Covered','✔ No gap','—'),
    ('HMW-03','Let job seeker compare salary vs market by exp+city','BR-03','Covered','✔ No gap','—'),
    ('HMW-04','Rank top-10 companies for placement officer',        'BR-04','Covered','✔ No gap','—'),
    ('HMW-05','Month-over-month hiring volume timeline for HR',     'BR-05','Covered','✔ No gap','—'),
    ('HMW-06','Skill count ↔ salary correlation for recruiter',     'BR-09','Covered','✔ No gap','—'),
    ('HMW-07','City hiring hotspots for job seeker',                'BR-07','Covered','✔ No gap','—'),
    ('HMW-08','Job-type distribution per industry for recruiter',   'BR-10','Covered','✔ No gap','—'),
    ('HMW-09','Exp-level demand per role for placement officer',    'BR-08','Covered','✔ No gap','—'),
    ('HMW-10','Cross-industry salary comparison for HR',            'BR-06','Covered','✔ No gap','—'),
    ('HMW-11','Fastest-growing industries by skill for counsellor', 'BR-11','Covered','✔ No gap','—'),
    ('HMW-12','Role-level posting trend for job seeker timing',     'BR-12','Covered','✔ No gap','—'),
]
for i, row_data in enumerate(hmw_cov, start=b_start+2):
    for ci, val in enumerate(row_data, 1):
        c = wc.cell(row=i, column=ci, value=val)
        bg = C['ok_green'] if (ci==5 and '✔' in str(val)) else (C['alt_row'] if i%2==0 else C['white'])
        fg = '1F5C2E' if (ci==5 and '✔' in str(val)) else '000000'
        style_cell(c, bg=bg, fg=fg, bold=(ci==5), size=9, h_align='center' if ci in (1,3,4,5) else 'left')
    wc.row_dimensions[i].height = 16

# ── Section C: BR Traceability ────────────────────────────────────────────────
c_start = b_start + 2 + len(hmw_cov) + 2
wc.merge_cells(f'A{c_start}:F{c_start}')
style_cell(wc[f'A{c_start}'], bg=C['mid_blue'], fg=C['hdr_txt'], bold=True, size=11, h_align='center')
wc[f'A{c_start}'].value = 'C — BR Traceability: Every BR traced to ≥1 HMW (or documented assumption)'
wc.row_dimensions[c_start].height = 22

br_hdrs = ['BR ID','BR Short Name','Source HMW','Priority','Traced?','Note']
for ci, h in enumerate(br_hdrs, 1):
    c = wc.cell(row=c_start+1, column=ci, value=h)
    style_cell(c, bg=C['steel'], fg=C['hdr_txt'], bold=True, size=9, h_align='center')

br_cov = [
    ('BR-01','Role demand ranking',           'HMW-01','Must',  '✔ Traced','—'),
    ('BR-02','Skill frequency ranking',       'HMW-02','Must',  '✔ Traced','—'),
    ('BR-03','Salary stats by exp + city',    'HMW-03','Must',  '✔ Traced','—'),
    ('BR-04','Company leaderboard',           'HMW-04','Must',  '✔ Traced','—'),
    ('BR-05','Monthly hiring volume',         'HMW-05','Must',  '✔ Traced','—'),
    ('BR-06','Cross-industry salary compare', 'HMW-10','Must',  '✔ Traced','—'),
    ('BR-07','City posting concentration',    'HMW-07','Must',  '✔ Traced','—'),
    ('BR-08','Exp-level demand per role',     'HMW-09','Must',  '✔ Traced','—'),
    ('BR-09','Skill-salary correlation',      'HMW-06','Should','✔ Traced','—'),
    ('BR-10','Job-type distribution',         'HMW-08','Should','✔ Traced','—'),
    ('BR-11','Industry hiring by skill',      'HMW-11','Should','✔ Traced','—'),
    ('BR-12','Role-level time series',        'HMW-12','Should','✔ Traced','—'),
    ('BR-13','EDA summary statistics',        '— (None)','Could','⚠ GAP',
     'No HMW; accepted as standard EDA baseline — documented in BRD §Assumptions'),
]
for i, row_data in enumerate(br_cov, start=c_start+2):
    for ci, val in enumerate(row_data, 1):
        c = wc.cell(row=i, column=ci, value=val)
        is_gap = '⚠' in str(row_data[4])
        is_ok  = '✔' in str(row_data[4])
        if ci == 5:
            bg = C['warn_amber'] if is_gap else C['ok_green']
            fg = '7B5700' if is_gap else '1F5C2E'
        elif ci == 4:
            prio = row_data[3]
            bg = MOSCOW_BG.get(prio, C['white'])
            fg = MOSCOW_FG.get(prio, '000000')
        else:
            bg = C['warn_amber'] if is_gap else (C['alt_row'] if i%2==0 else C['white'])
            fg = '7B5700' if is_gap else '000000'
        style_cell(c, bg=bg, fg=fg, bold=(ci in (1,5)), size=9,
                   h_align='center' if ci in (1,2,3,4,5) else 'left')
    wc.row_dimensions[i].height = 18

# Column widths for coverage sheet
for col_letter, width in [('A',10),('B',34),('C',22),('D',10),('E',12),('F',42)]:
    wc.column_dimensions[col_letter].width = width

# ── Summary box at top of coverage sheet ─────────────────────────────────────
# (insert after main table — append summary)
sum_row = c_start + 2 + len(br_cov) + 2
wc.merge_cells(f'A{sum_row}:F{sum_row}')
sc = wc[f'A{sum_row}']
sc.value = ('SUMMARY  |  OQ coverage: 10/10 ✔  |  HMW→BR coverage: 12/12 ✔  |  '
            'BR→HMW traceability: 12/13 ✔  (BR-13 gap documented)  |  '
            'Overall RTM status: COMPLETE with 1 documented gap')
style_cell(sc, bg='1F5C2E', fg=C['hdr_txt'], bold=True, size=9, h_align='center')
wc.row_dimensions[sum_row].height = 20

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3  — Changelog (for future step updates)
# ══════════════════════════════════════════════════════════════════════════════
wch = wb.create_sheet("Changelog")
wch.sheet_view.showGridLines = False
wch.merge_cells('A1:E1')
style_cell(wch['A1'], bg=C['dark_blue'], fg=C['hdr_txt'], bold=True,
           size=12, h_align='center', border_color=C['dark_blue'])
wch['A1'].value = 'RTM Changelog — P01 Hiring Trend Analysis'
wch.row_dimensions[1].height = 24

ch_hdrs = ['Step','Date Updated','Columns Updated','Updated by','Notes']
for ci, h in enumerate(ch_hdrs, 1):
    c = wch.cell(row=2, column=ci, value=h)
    style_cell(c, bg=C['mid_blue'], fg=C['hdr_txt'], bold=True, size=9, h_align='center')

changelog = [
    ('Step 04','  '+datetime.date.today().strftime('%d %b %Y'),
     'All columns created; Feature/DB/US/UI/TC left as placeholders',
     'Abhayjeet Singh','Initial RTM — baseline'),
    ('Step 06','—','Feature ID column','Abhayjeet Singh','To be filled in Step 06 (PRD)'),
    ('Step 08','—','DB Table.Column','Abhayjeet Singh','To be filled in Step 08 (Database Design)'),
    ('Step 11','—','User Story ID','Abhayjeet Singh','To be filled in Step 11 (User Stories)'),
    ('Step 12','—','UI Screen ID','Abhayjeet Singh','To be filled in Step 12 (UI Framework)'),
    ('Step 13','—','TC-F, TC-T','Abhayjeet Singh','To be filled in Step 13 (Test Cases)'),
]
for i, row_data in enumerate(changelog, start=3):
    for ci, val in enumerate(row_data, 1):
        c = wch.cell(row=i, column=ci, value=val)
        style_cell(c, bg=C['alt_row'] if i%2==0 else C['white'], size=9,
                   h_align='left' if ci>1 else 'center')
    wch.row_dimensions[i].height = 20

for col_letter, width in [('A',10),('B',16),('C',45),('D',18),('E',42)]:
    wch.column_dimensions[col_letter].width = width

# ── Tab colours ───────────────────────────────────────────────────────────────
ws.sheet_properties.tabColor  = '1F3564'
wc.sheet_properties.tabColor  = '1F5C2E'
wch.sheet_properties.tabColor = '7B5700'

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save(OUT_FILE)
print(f"SAVED: {OUT_FILE}")
