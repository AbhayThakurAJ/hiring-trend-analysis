"""
Step 09 – ER Diagram PNG Generator & Word Document Builder
Project : P01 – Hiring Trend Analysis
Student : Abhayjeet Singh | Roll: 2500320140004
Mentor  : Dinesh Sereno | ABES, MCA 2nd Yr DA
"""

import os, datetime
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_DIR   = r"C:\Hiring Trend Analysis\docs"
DOCX_FILE = os.path.join(OUT_DIR, "P01_HTA_Step09_ERDiagram.docx")
PNG_FILE  = os.path.join(OUT_DIR, "09_ER_Diagram.png")
os.makedirs(OUT_DIR, exist_ok=True)

# ── 1. GENERATE PROFESSIONAL ER DIAGRAM PNG USING PILLOW ─────────────────────
W, H = 1400, 950
img = Image.new("RGB", (W, H), color="#F8FAFC")
draw = ImageDraw.Draw(img)

# Try loading standard Windows fonts
try:
    font_title = ImageFont.truetype("arialbd.ttf", 24)
    font_sub   = ImageFont.truetype("arial.ttf", 14)
    font_tbl   = ImageFont.truetype("arialbd.ttf", 16)
    font_bold  = ImageFont.truetype("arialbd.ttf", 12)
    font_norm  = ImageFont.truetype("arial.ttf", 12)
    font_code  = ImageFont.truetype("consola.ttf", 11)
except Exception:
    font_title = font_sub = font_tbl = font_bold = font_norm = font_code = ImageFont.load_default()

# Header Banner
draw.rectangle([(0, 0), (W, 70)], fill="#1F3564")
draw.text((30, 15), "ENTITY-RELATIONSHIP (ER) DIAGRAM — P01 HIRING TREND ANALYSIS", fill="#FFFFFF", font=font_title)
draw.text((30, 45), "Relational Model: 3 Dimension Masters + 1 Core Fact + 1 Bridge Junction", fill="#BDD5EE", font=font_sub)

def draw_box(x, y, w, h, title, cols, header_color="#2E74B5"):
    # Card shadow
    draw.rounded_rectangle([(x+4, y+4), (x+w+4, y+h+4)], radius=8, fill="#E2E8F0")
    # Card body
    draw.rounded_rectangle([(x, y), (x+w, y+h)], radius=8, fill="#FFFFFF", outline="#94A3B8", width=2)
    # Header
    draw.rounded_rectangle([(x, y), (x+w, y+34)], radius=8, fill=header_color)
    draw.rectangle([(x, y+20), (x+w, y+34)], fill=header_color) # square bottom corners of header
    draw.text((x+15, y+8), title, fill="#FFFFFF", font=font_tbl)
    
    # Rows
    cur_y = y + 42
    for col_name, col_type, key_type in cols:
        badge_color = "#D6E4F0" if "PK" in key_type else ("#E2EFDA" if "FK" in key_type else "#F1F5F9")
        badge_text_color = "#1F3564" if "PK" in key_type else ("#1F5C2E" if "FK" in key_type else "#64748B")
        
        # Key Badge
        if key_type:
            draw.rounded_rectangle([(x+10, cur_y), (x+50, cur_y+18)], radius=4, fill=badge_color)
            draw.text((x+16, cur_y+2), key_type, fill=badge_text_color, font=font_code)
        
        # Col Name
        draw.text((x+58, cur_y+2), col_name, fill="#0F172A", font=font_bold if key_type else font_norm)
        # Col Type
        draw.text((x+w-100, cur_y+2), col_type, fill="#64748B", font=font_code)
        
        # Divider
        draw.line([(x+10, cur_y+22), (x+w-10, cur_y+22)], fill="#F1F5F9", width=1)
        cur_y += 24

# Tables data
dim_ind_cols = [
    ("industry_id", "INTEGER", "PK"),
    ("industry_name", "TEXT", ""),
    ("sector_category", "TEXT", ""),
]

dim_city_cols = [
    ("city_id", "INTEGER", "PK"),
    ("city_name", "TEXT", ""),
    ("state_name", "TEXT", ""),
    ("region", "TEXT", ""),
]

dim_role_cols = [
    ("role_id", "INTEGER", "PK"),
    ("role_title", "TEXT", ""),
    ("role_category", "TEXT", ""),
    ("seniority_band", "TEXT", ""),
]

fact_cols = [
    ("job_id", "INTEGER", "PK"),
    ("role_id", "INTEGER", "FK"),
    ("company_name", "TEXT", ""),
    ("industry_id", "INTEGER", "FK"),
    ("city_id", "INTEGER", "FK"),
    ("job_type", "TEXT", ""),
    ("experience_level", "TEXT", ""),
    ("salary_min", "REAL", ""),
    ("salary_max", "REAL", ""),
    ("required_skills", "TEXT", ""),
    ("skills_count", "INTEGER", ""),
    ("posted_date", "TEXT", ""),
    ("posted_month", "TEXT", ""),
    ("posted_year", "INTEGER", ""),
]

bridge_cols = [
    ("skill_id", "INTEGER", "PK"),
    ("job_id", "INTEGER", "FK"),
    ("skill_name", "TEXT", ""),
    ("is_primary", "INTEGER", ""),
]

# Coordinates layout
# Left column: Dimension Tables
draw_box(60, 110, 340, 130, "dim_industry", dim_ind_cols, header_color="#2E74B5")
draw_box(60, 300, 340, 155, "dim_city", dim_city_cols, header_color="#2E74B5")
draw_box(60, 520, 340, 155, "dim_job_role", dim_role_cols, header_color="#2E74B5")

# Center: Core Fact Table
draw_box(500, 110, 420, 410, "fact_job_postings (FACT)", fact_cols, header_color="#1F5C2E")

# Right column: Bridge Table
draw_box(1020, 110, 320, 155, "job_skills (BRIDGE)", bridge_cols, header_color="#7B5700")

# Relationship connectors & labels
def draw_conn(p1, p2, label, label_pos, card_left="||", card_right="o{"):
    draw.line([p1, p2], fill="#475569", width=2)
    # Background for label
    lx, ly = label_pos
    draw.rounded_rectangle([(lx-4, ly-2), (lx+110, ly+18)], radius=4, fill="#F1F5F9", outline="#CBD5E1", width=1)
    draw.text((lx+4, ly), label, fill="#1E293B", font=font_code)

# 1. dim_industry -> fact_job_postings
draw.line([(400, 175), (450, 175), (450, 210), (500, 210)], fill="#475569", width=2)
draw.text((410, 158), "1.0", fill="#1F3564", font=font_bold)
draw.text((475, 192), "1..N", fill="#1F3564", font=font_bold)
draw.rounded_rectangle([(415, 185), (485, 203)], radius=4, fill="#FEF3C7", outline="#F59E0B")
draw.text((420, 187), "classifies", fill="#92400E", font=font_code)

# 2. dim_city -> fact_job_postings
draw.line([(400, 375), (450, 375), (450, 240), (500, 240)], fill="#475569", width=2)
draw.text((410, 358), "1.0", fill="#1F3564", font=font_bold)
draw.text((475, 222), "1..N", fill="#1F3564", font=font_bold)
draw.rounded_rectangle([(425, 290), (475, 308)], radius=4, fill="#FEF3C7", outline="#F59E0B")
draw.text((430, 292), "locates", fill="#92400E", font=font_code)

# 3. dim_job_role -> fact_job_postings
draw.line([(400, 595), (460, 595), (460, 270), (500, 270)], fill="#475569", width=2)
draw.text((410, 578), "1.0", fill="#1F3564", font=font_bold)
draw.text((475, 252), "1..N", fill="#1F3564", font=font_bold)
draw.rounded_rectangle([(425, 420), (485, 438)], radius=4, fill="#FEF3C7", outline="#F59E0B")
draw.text((430, 422), "defines", fill="#92400E", font=font_code)

# 4. fact_job_postings -> job_skills
draw.line([(920, 180), (1020, 180)], fill="#475569", width=2)
draw.text((930, 162), "1.0", fill="#1F3564", font=font_bold)
draw.text((995, 162), "1..N", fill="#1F3564", font=font_bold)
draw.rounded_rectangle([(935, 185), (1005, 203)], radius=4, fill="#FEF3C7", outline="#F59E0B")
draw.text((940, 187), "requires", fill="#92400E", font=font_code)

# Legend Box
draw.rounded_rectangle([(500, 560), (1340, 680)], radius=8, fill="#FFFFFF", outline="#CBD5E1", width=2)
draw.text((520, 575), "DIAGRAM CONVENTIONS & CARDINALITY GUIDE", fill="#1F3564", font=font_bold)
draw.text((520, 605), "• 1.0  (One and only one): Strong parent master entity (dim_industry, dim_city, dim_job_role)", fill="#475569", font=font_norm)
draw.text((520, 628), "• 1..N (One to many): Child transactions in fact_job_postings and decomposed skills in job_skills", fill="#475569", font=font_norm)
draw.text((520, 651), "• PK = Primary Key (Identity) | FK = Foreign Key (Referential Integrity Check)", fill="#475569", font=font_norm)

# Footer banner
draw.rectangle([(0, H-40), (W, H)], fill="#0F172A")
draw.text((30, H-28), "Generated for ABES MCA Data Analytics Capstone | P01 Hiring Trend Analysis | Author: Abhayjeet Singh", fill="#94A3B8", font=font_code)

img.save(PNG_FILE)
print(f"SAVED PNG: {PNG_FILE}")


# ── 2. BUILD WORD DELIVERABLE (09_ER_Diagram.docx) ───────────────────────────
doc = Document()
for s in doc.sections:
    s.top_margin = Cm(2.0); s.bottom_margin = Cm(2.0)
    s.left_margin = Cm(2.5); s.right_margin = Cm(2.5)

DARK  = RGBColor(0x1F,0x35,0x64)
MID   = RGBColor(0x2E,0x74,0xB5)
STEEL = RGBColor(0x4A,0x86,0xC8)
WHITE = RGBColor(0xFF,0xFF,0xFF)
BLACK = RGBColor(0x00,0x00,0x00)
GREEN = RGBColor(0x1F,0x5C,0x2E)

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
r1 = bp.add_run('P01 – HIRING TREND ANALYSIS\n'); r1.bold = True; r1.font.size = Pt(18); r1.font.color.rgb = WHITE
r2 = bp.add_run('Step 09 · Entity-Relationship (ER) Diagram Specification\n'); r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
r3 = bp.add_run('Abhayjeet Singh | Roll: 2500320140004 | Mentor: Dinesh Sereno | ABES MCA 2nd Yr DA'); r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(0xBD,0xD5,0xEE)
doc.add_paragraph()

# 1. Embedded Image
add_heading('1. Visual Entity-Relationship Diagram', 1)
add_para('Below is the relational schema diagram rendered according to Crow\'s Foot notation, illustrating 3 Master Dimension tables, 1 Core Fact table, and 1 Junction Bridge table.', size=10)
p_img = doc.add_paragraph()
p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img.paragraph_format.space_before = Pt(8); p_img.paragraph_format.space_after = Pt(8)
p_img.add_run().add_picture(PNG_FILE, width=Inches(6.2))
add_para('Figure 9.1: Entity-Relationship Diagram for Hiring Trend Analysis schema.', italic=True, size=8.5, color=MID)
doc.add_paragraph()

# 2. Mermaid Code Block
add_heading('2. Mermaid "erDiagram" Code Specification', 1)
add_para('The following Mermaid syntax models all 5 relational entities, attribute datatypes, key qualifiers, and labelled cardinality relationships. This block can be directly rendered in markdown viewers or mermaid.live.', size=9.5)

mermaid_code = """erDiagram
    dim_industry ||--o{ fact_job_postings : "classifies"
    dim_city ||--o{ fact_job_postings : "locates"
    dim_job_role ||--o{ fact_job_postings : "defines"
    fact_job_postings ||--o{ job_skills : "requires"

    dim_industry {
        INTEGER industry_id PK
        TEXT industry_name
        TEXT sector_category
    }

    dim_city {
        INTEGER city_id PK
        TEXT city_name
        TEXT state_name
        TEXT region
    }

    dim_job_role {
        INTEGER role_id PK
        TEXT role_title
        TEXT role_category
        TEXT seniority_band
    }

    fact_job_postings {
        INTEGER job_id PK
        INTEGER role_id FK
        TEXT company_name
        INTEGER industry_id FK
        INTEGER city_id FK
        TEXT job_type
        TEXT experience_level
        REAL salary_min
        REAL salary_max
        TEXT required_skills
        INTEGER skills_count
        TEXT posted_date
        TEXT posted_month
        INTEGER posted_year
    }

    job_skills {
        INTEGER skill_id PK
        INTEGER job_id FK
        TEXT skill_name
        INTEGER is_primary
    }"""
code_box(mermaid_code)
doc.add_paragraph()

# 3. Plain English Relationship Narratives
add_heading('3. Relational Cardinality & Plain-English Narratives', 1)
add_para('Each foreign-key connection is governed by strict referential semantics described below:', size=9.5)

rel_tbl = doc.add_table(rows=5, cols=5)
rel_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['#', 'Relationship', 'Cardinality', 'Verb Phrase', 'Plain-English Business Description']):
    c = rel_tbl.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)

rel_data = [
    ('1', 'dim_industry ➔ fact_job_postings', '1 : N', 'classifies',
     'One industry sector classifies many job postings; every job posting must belong to exactly one industry sector.'),
    ('2', 'dim_city ➔ fact_job_postings', '1 : N', 'locates',
     'One geographic city locates many job postings; every job posting is stationed in exactly one city.'),
    ('3', 'dim_job_role ➔ fact_job_postings', '1 : N', 'defines',
     'One standardized job role defines many job postings; every posting references exactly one standard role template.'),
    ('4', 'fact_job_postings ➔ job_skills', '1 : N', 'requires',
     'One job posting requires many individual skills; each decomposed skill entry belongs to exactly one parent job posting.')
]
for ri, row in enumerate(rel_data, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    for ci, val in enumerate(row):
        c = rel_tbl.cell(ri, ci); set_cell_bg(c, fill_c)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci in (0,2) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(val); r.font.size = Pt(9)
        if ci in (0, 1): r.bold = True
set_tbl_borders(rel_tbl)
doc.add_paragraph()

# 4. Business Rules Enforced
add_heading('4. Business Rules Enforced in ER Design', 1)
rules_data = [
    ('Referential Integrity', 'Foreign key constraints prevent any job posting from referencing a non-existent role, city, or industry.'),
    ('Domain Integrity via CHECKs', 'Restricts categorical columns strictly to valid business values: job_type in (Full-time, Part-time, Remote, Hybrid, Contract); experience_level in (Fresher, Mid-level, Senior).'),
    ('Salary Logic Validation', 'Enforces salary_min > 0 AND salary_max > salary_min via table-level CHECK constraint (BR-R-09).'),
    ('Temporal Consistency', 'Restricts posted_year BETWEEN 2022 AND 2024; posted_month adheres to canonical YYYY-MM format.'),
    ('Cascading Deletions on Junction', 'Deleting a job posting automatically cascades to purge its corresponding skills in job_skills (ON DELETE CASCADE), preventing dangling skill pointers.'),
    ('Normalized Skill Searchability', 'Breaking pipe-delimited strings into job_skills resolves 1NF and allows high-performance GROUP BY counting without regex parsing overhead.')
]
for title, desc in rules_data:
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(f"{title}: "); r1.bold = True; r1.font.size = Pt(9.5)
    r2 = p.add_run(desc); r2.font.size = Pt(9.5)
doc.add_paragraph()

# 5. Design Problems & Solutions (Anticipated Issues & Mitigations)
add_heading('5. Architectural Critiques & Design Trade-offs', 1)
add_para('During schema refinement, three design challenges were evaluated and resolved:', size=9.5)

critiques = [
    ('1. Pipe-delimited vs Bridge Table Redundancy',
     'ISSUE: Both fact_job_postings.required_skills and job_skills table hold skill data (apparent duplication).\n'
     'RESOLUTION: Intentional hybrid design. The string column allows ultra-fast export and flat table displays in Streamlit without joins; the bridge table provides 1NF normalization for fast aggregation queries and skill frequency bar charts (OQ-03).'),
    ('2. Missing Company Master Table',
     'ISSUE: company_name is stored directly in fact_job_postings rather than as a separate dim_company.\n'
     'RESOLUTION: In synthetic recruitment datasets, companies are transient text entities (~50 arbitrary names). Creating a company dimension table would add unnecessary JOIN overhead for a read-heavy analytical dashboard with zero updates. Storing company_name directly in fact_job_postings avoids a superfluous 6th table.'),
    ('3. Temporal Granularity and Derived Keys',
     'ISSUE: Parsing ISO strings (YYYY-MM-DD) dynamically in SQL slows down time-series aggregations (OQ-04).\n'
     'RESOLUTION: Added pre-computed posted_month and posted_year directly in fact_job_postings, indexed via B-Tree index, ensuring sub-millisecond filtering across Streamlit date slider widgets.')
]

crit_tbl = doc.add_table(rows=len(critiques)+1, cols=2)
crit_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
c0 = crit_tbl.cell(0, 0); set_cell_bg(c0, '1F3564'); c0.paragraphs[0].add_run('Design Challenge').font.color.rgb = WHITE
c1 = crit_tbl.cell(0, 1); set_cell_bg(c1, '1F3564'); c1.paragraphs[0].add_run('Critique & Architectural Resolution').font.color.rgb = WHITE
c0.paragraphs[0].runs[0].bold = True; c1.paragraphs[0].runs[0].bold = True
crit_tbl.cell(0, 0).width = Inches(2.2); crit_tbl.cell(0, 1).width = Inches(4.3)

for ri, (chal, resol) in enumerate(critiques, 1):
    fill_c = 'EEF4FB' if ri%2==0 else 'FFFFFF'
    r_c0 = crit_tbl.cell(ri, 0); set_cell_bg(r_c0, fill_c)
    p0 = r_c0.paragraphs[0]; r0 = p0.add_run(chal); r0.bold = True; r0.font.size = Pt(9)
    
    r_c1 = crit_tbl.cell(ri, 1); set_cell_bg(r_c1, fill_c)
    p1 = r_c1.paragraphs[0]; r1 = p1.add_run(resol); r1.font.size = Pt(9)

set_tbl_borders(crit_tbl)
doc.add_paragraph()

# Document Control
add_heading('Document Control', 2)
dc = doc.add_table(rows=2, cols=4); dc.alignment = WD_TABLE_ALIGNMENT.CENTER
for ci, col in enumerate(['Version', 'Date', 'Author', 'Summary']):
    c = dc.cell(0, ci); set_cell_bg(c, '1F3564')
    p = c.paragraphs[0]; r = p.add_run(col); r.bold = True; r.font.color.rgb = WHITE; r.font.size = Pt(9)
for ci, val in enumerate(['1.0', datetime.date.today().strftime('%d %b %Y'), 'Abhayjeet Singh', 'Initial ER Diagram deliverable with Pillow PNG insertion']):
    c = dc.cell(1, ci); set_cell_bg(c, 'EEF4FB')
    p = c.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)
set_tbl_borders(dc)

doc.save(DOCX_FILE)
print(f"SAVED DOCX: {DOCX_FILE}")
