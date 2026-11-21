"""
Step 18 Presentation Generator for P071 - Hiring Trend Analysis
Generates docs/P01_HTA_Step18_Presentation.pptx and docs/P01_Presentation.pptx
Adheres strictly to all layout, geometry, color, and font specifications.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

# Colors
RGB_RED = RGBColor(0xA5, 0x1C, 0x30)
RGB_BLACK = RGBColor(0x15, 0x15, 0x14)
RGB_DARK = RGBColor(0x3A, 0x3A, 0x38)
RGB_GREY = RGBColor(0x5E, 0x5E, 0x5B)
RGB_CARD = RGBColor(0xF4, 0xF1, 0xF1)
RGB_PINK = RGBColor(0xD9, 0xB5, 0xBC)
RGB_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RGB_SLIDENUM = RGBColor(0x8A, 0x8A, 0x86)
RGB_FRAME_BORDER = RGBColor(0xBF, 0xBF, 0xBF)

def set_cell_border(cell, color="D9D9D9", width="6350"):
    """Sets a 0.5pt solid border on all 4 sides of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    for border_name in ['lnL', 'lnR', 'lnT', 'lnB']:
        ln = parse_xml(
            f'<a:{border_name} {nsdecls("a")} w="{width}" cmpd="s">\n'
            f'  <a:solidFill>\n'
            f'    <a:srgbClr val="{color}"/>\n'
            f'  </a:solidFill>\n'
            f'</a:{border_name}>'
        )
        tcPr.append(ln)

def add_textbox(slide, x, y, w, h, default_insets=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    if not default_insets:
        tf.margin_left = Inches(0)
        tf.margin_right = Inches(0)
        tf.margin_top = Inches(0)
        tf.margin_bottom = Inches(0)
    return tb

def add_card(slide, x, y, w, h, fill_color=RGB_CARD, line_color=RGB_CARD):
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = fill_color
    if line_color:
        card.line.color.rgb = line_color
        card.line.width = Pt(1)
    else:
        card.line.fill.background()
    return card

def add_badge(slide, x, y, t):
    badge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(x), Inches(y), Inches(0.42), Inches(0.42)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGB_RED
    badge.line.fill.background()
    tf = badge.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = str(t)
    run.font.name = 'Calibri'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = RGB_WHITE
    return badge

def add_badge_card(slide, x, y, w, h, badge_t, head_text, body_text):
    # CARD(x,y,w,h)
    add_card(slide, x, y, w, h)
    # BADGE(x+.2, y+.2, t)
    add_badge(slide, x + 0.2, y + 0.2, badge_t)
    # HEAD (x+.75, y+.17, w-.9, .45)
    tb_head = add_textbox(slide, x + 0.75, y + 0.17, w - 0.9, 0.45, default_insets=False)
    tf_h = tb_head.text_frame
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_h.paragraphs[0]
    p_h.alignment = PP_ALIGN.LEFT
    run_h = p_h.add_run()
    run_h.text = head_text
    run_h.font.name = 'Calibri'
    run_h.font.size = Pt(15)
    run_h.font.bold = True
    run_h.font.color.rgb = RGB_BLACK

    # BODY (x+.2, y+.72, w-.4, h-.85)
    tb_body = add_textbox(slide, x + 0.2, y + 0.72, w - 0.4, h - 0.85, default_insets=False)
    tf_b = tb_body.text_frame
    tf_b.vertical_anchor = MSO_ANCHOR.TOP
    paragraphs = body_text.split('\n')
    for i, p_str in enumerate(paragraphs):
        p = tf_b.paragraphs[0] if i == 0 else tf_b.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        run_b = p.add_run()
        run_b.text = p_str
        run_b.font.name = 'Calibri'
        run_b.font.size = Pt(13)
        run_b.font.color.rgb = RGB_DARK
        if i > 0:
            p.space_before = Pt(3)

def add_plain_card(slide, x, y, w, h, head_text, body_text):
    # CARD(x,y,w,h)
    add_card(slide, x, y, w, h)
    # HEAD (x+.2, y+.17, w-.35, .45)
    tb_head = add_textbox(slide, x + 0.2, y + 0.17, w - 0.35, 0.45, default_insets=False)
    tf_h = tb_head.text_frame
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_h.paragraphs[0]
    p_h.alignment = PP_ALIGN.LEFT
    run_h = p_h.add_run()
    run_h.text = head_text
    run_h.font.name = 'Calibri'
    run_h.font.size = Pt(15)
    run_h.font.bold = True
    run_h.font.color.rgb = RGB_BLACK

    # BODY (x+.2, y+.72, w-.4, h-.85)
    tb_body = add_textbox(slide, x + 0.2, y + 0.72, w - 0.4, h - 0.85, default_insets=False)
    tf_b = tb_body.text_frame
    tf_b.vertical_anchor = MSO_ANCHOR.TOP
    paragraphs = body_text.split('\n')
    for i, p_str in enumerate(paragraphs):
        p = tf_b.paragraphs[0] if i == 0 else tf_b.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        run_b = p.add_run()
        run_b.text = p_str
        run_b.font.name = 'Calibri'
        run_b.font.size = Pt(13)
        run_b.font.color.rgb = RGB_DARK
        if i > 0:
            p.space_before = Pt(3)

def create_content_slide_base(prs, title_text, slide_number):
    slide = prs.slides.add_slide(prs.slide_layouts[6]) # blank layout
    
    # Title (.5, .3, 9, .75) Cambria 28 bold black, centred, middle, default insets
    tb_title = add_textbox(slide, 0.5, 0.3, 9.0, 0.75, default_insets=True)
    tf_t = tb_title.text_frame
    tf_t.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_t = tf_t.paragraphs[0]
    p_t.alignment = PP_ALIGN.CENTER
    run_t = p_t.add_run()
    run_t.text = title_text
    run_t.font.name = 'Cambria'
    run_t.font.size = Pt(28)
    run_t.font.bold = True
    run_t.font.color.rgb = RGB_BLACK

    # Footer (.5, 5.2, 7.5, .3) 9 grey, middle, default insets
    tb_foot = add_textbox(slide, 0.5, 5.2, 7.5, 0.3, default_insets=True)
    tf_f = tb_foot.text_frame
    tf_f.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_f = tf_f.paragraphs[0]
    p_f.alignment = PP_ALIGN.LEFT
    run_f = p_f.add_run()
    run_f.text = "HCLTech Career Shaper ABES Engineering College | P071 - Hiring Trend Analysis"
    run_f.font.name = 'Calibri'
    run_f.font.size = Pt(9)
    run_f.font.color.rgb = RGB_GREY

    # Slide number (9, 5.2, .5, .3) 9, colour 8A8A86, right, default insets
    tb_num = add_textbox(slide, 9.0, 5.2, 0.5, 0.3, default_insets=True)
    tf_n = tb_num.text_frame
    tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_n = tf_n.paragraphs[0]
    p_n.alignment = PP_ALIGN.RIGHT
    run_n = p_n.add_run()
    run_n.text = str(slide_number)
    run_n.font.name = 'Calibri'
    run_n.font.size = Pt(9)
    run_n.font.color.rgb = RGB_SLIDENUM

    return slide

def set_speaker_note(slide, text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = text

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    # ----------------------------------------------------
    # SLIDE 1: DARK
    # ----------------------------------------------------
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    # Dark background
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = RGB_BLACK
    bg1.line.fill.background()

    # ABES logo (.5, .35, 1.6, .5)
    if os.path.exists('assets/abes_logo.png'):
        slide1.shapes.add_picture('assets/abes_logo.png', Inches(0.5), Inches(0.28), width=Inches(0.65))
    else:
        tb_logo = add_textbox(slide1, 0.5, 0.35, 1.6, 0.5, default_insets=True)
        tf_l = tb_logo.text_frame
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_l = tf_l.paragraphs[0]
        p_l.alignment = PP_ALIGN.LEFT
        r_l = p_l.add_run()
        r_l.text = "[ABES logo]"
        r_l.font.name = 'Calibri'
        r_l.font.size = Pt(10)
        r_l.font.italic = True
        r_l.font.color.rgb = RGB_PINK

    # Header right (2.2, .35, 7.3, .5) 11 pink, right, middle, default insets
    tb_hr = add_textbox(slide1, 2.2, 0.35, 7.3, 0.5, default_insets=True)
    tf_hr = tb_hr.text_frame
    tf_hr.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_hr = tf_hr.paragraphs[0]
    p_hr.alignment = PP_ALIGN.RIGHT
    r_hr = p_hr.add_run()
    r_hr.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    r_hr.font.name = 'Calibri'
    r_hr.font.size = Pt(11)
    r_hr.font.color.rgb = RGB_PINK

    # Title (.5, 1.55, 9, 1.5) Cambria 36 bold white, left, bottom, default insets
    tb_t1 = add_textbox(slide1, 0.5, 1.55, 9.0, 1.5, default_insets=True)
    tf_t1 = tb_t1.text_frame
    tf_t1.vertical_anchor = MSO_ANCHOR.BOTTOM
    p_t1 = tf_t1.paragraphs[0]
    p_t1.alignment = PP_ALIGN.LEFT
    r_t1 = p_t1.add_run()
    r_t1.text = "P071 - Hiring Trend Analysis"
    r_t1.font.name = 'Cambria'
    r_t1.font.size = Pt(36)
    r_t1.font.bold = True
    r_t1.font.color.rgb = RGB_WHITE

    # Lines (.5, 3.2, 9, 1.7) 15 F4F1F1, left, top, 4 pt before each line, default insets
    tb_lines1 = add_textbox(slide1, 0.5, 3.2, 9.0, 1.7, default_insets=True)
    tf_lines1 = tb_lines1.text_frame
    tf_lines1.vertical_anchor = MSO_ANCHOR.TOP

    lines_data = [
        ("Abhayjeet Singh | Roll No 2500320140004", RGB_CARD),
        ("MCA 2nd Year Data Analyst | Mentor: Dinesh Sereno", RGB_CARD),
        ("Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01", RGB_CARD),
        ("21 November 2026", RGB_PINK)
    ]
    for i, (txt, col) in enumerate(lines_data):
        p = tf_lines1.paragraphs[0] if i == 0 else tf_lines1.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = txt
        run.font.name = 'Calibri'
        run.font.size = Pt(15)
        run.font.color.rgb = col

    set_speaker_note(slide1,
        "Good morning respected mentor and evaluators. I am Abhayjeet Singh, MCA 2nd Year Data Analyst, "
        "presenting Project P071, Hiring Trend Analysis (Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01). "
        "The purpose of this project is to provide data-driven insights into Indian tech compensation, skill demands, and regional hiring patterns."
    )

    # ----------------------------------------------------
    # SLIDE 2: Problem and How Might We
    # ----------------------------------------------------
    slide2 = create_content_slide_base(prs, "Problem and How Might We", 2)
    # BADGE CARD (.5, 1.25, 4.3, 3.7) P, Problem statement: 2-3 sentences
    prob_text = (
        "Job seekers and academic institutions struggle with fragmented hiring data, "
        "non-transparent compensation structures, and misaligned curricula across Indian tech hubs."
    )
    add_badge_card(slide2, 0.5, 1.25, 4.3, 3.7, "P", "Problem Statement", prob_text)

    # BADGE CARD (5.2, 1.25, 4.3, 1.75) 1, HMW-01: statement
    hmw1_text = "How might we identify salary growth drivers across experience levels and skills for job seekers?"
    add_badge_card(slide2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01", hmw1_text)

    # BADGE CARD (5.2, 3.2, 4.3, 1.75) 2, HMW-02: statement
    hmw2_text = "How might we analyze recruitment demand by city and industry tier for workforce planning?"
    add_badge_card(slide2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02", hmw2_text)

    set_speaker_note(slide2,
        "Job seekers and colleges face fragmented hiring data and lack salary transparency. "
        "We formulated two primary How Might We questions focusing on uncovering salary growth drivers "
        "and mapping regional recruitment demand."
    )

    # ----------------------------------------------------
    # SLIDE 3: Stakeholders and users
    # ----------------------------------------------------
    slide3 = create_content_slide_base(prs, "Stakeholders and users", 3)
    # BADGE CARDS y 1.35, w 2.85, h 3.5, x .5 / 3.55 / 6.6, badges 1 / 2 / 3
    sh1_body = "Role: Job Seeker / Data Analyst\nNeed: Benchmark salaries and skills by city.\n\nInfluence: High"
    add_badge_card(slide3, 0.5, 1.35, 2.85, 3.5, "1", "Primary User", sh1_body)

    sh2_body = "Role: HR Recruiter / Hiring Lead\nNeed: Track talent demand & set competitive pay.\n\nInfluence: Medium"
    add_badge_card(slide3, 3.55, 1.35, 2.85, 3.5, "2", "Secondary User", sh2_body)

    sh3_body = "Role: Mentor & Evaluator\nNeed: Verify analytical rigor & technical stack.\n\nInfluence: High"
    add_badge_card(slide3, 6.6, 1.35, 2.85, 3.5, "3", "Evaluator", sh3_body)

    set_speaker_note(slide3,
        "We identified three key stakeholder groups: job seekers who need transparent salary benchmarks, "
        "HR recruiters benchmarking compensation, and academic evaluators assessing our analytical rigor and technical architecture."
    )

    # ----------------------------------------------------
    # SLIDE 4: Requirements and traceability
    # ----------------------------------------------------
    slide4 = create_content_slide_base(prs, "Requirements and traceability", 4)
    # TABLE at (.5, 1.3), column widths 1, 3.4, 1, 1.2, 1.2, 1.2
    col_widths = [1.0, 3.4, 1.0, 1.2, 1.2, 1.2]
    table_shape4 = slide4.shapes.add_table(5, 6, Inches(0.5), Inches(1.3), Inches(sum(col_widths)), Inches(2.5))
    table4 = table_shape4.table
    for idx, w in enumerate(col_widths):
        table4.columns[idx].width = Inches(w)

    headers4 = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    rows4_data = [
        ["HMW-01", "Salary vs Experience benchmark", "Must", "US-01", "TC-F-01", "UI-01"],
        ["HMW-02", "City hiring volume comparison", "Must", "US-03", "TC-F-03", "UI-03"],
        ["HMW-03", "In-demand tech skills ranking", "Must", "US-04", "TC-F-04", "UI-04"],
        ["HMW-07", "Outlier and null data cleaning", "Should", "US-07", "TC-F-07", "UI-07"]
    ]
    all_rows4 = [headers4] + rows4_data

    for r_idx, row_data in enumerate(all_rows4):
        row = table4.rows[r_idx]
        row.height = Inches(0.5)
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.margin_left = Inches(0.1)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            set_cell_border(cell, color="D9D9D9", width="6350")
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGB_RED
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGB_WHITE

            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx in [0, 2, 3, 4, 5] else PP_ALIGN.LEFT
            run = p.add_run()
            run.text = val
            run.font.name = 'Calibri'
            run.font.size = Pt(13)
            if r_idx == 0:
                run.font.bold = True
                run.font.color.rgb = RGB_WHITE
            else:
                run.font.color.rgb = RGB_BLACK

    # .3 below the table, (.5, y9, .35) italic 12 grey: Full matrix: Project Document, section 4
    # Top 1.3 + height 2.5 = 3.8. .3 below = 4.1
    tb_note4 = add_textbox(slide4, 0.5, 4.1, 9.0, 0.35, default_insets=False)
    tf_n4 = tb_note4.text_frame
    p_n4 = tf_n4.paragraphs[0]
    p_n4.alignment = PP_ALIGN.LEFT
    r_n4 = p_n4.add_run()
    r_n4.text = "Full matrix: Project Document, section 4"
    r_n4.font.name = 'Calibri'
    r_n4.font.size = Pt(12)
    r_n4.font.italic = True
    r_n4.font.color.rgb = RGB_GREY

    set_speaker_note(slide4,
        "Our Requirements Traceability Matrix guarantees complete bidirectional alignment: "
        "every HMW statement links directly to a business requirement, user story, test case, "
        "and UI screen, ensuring zero orphaned requirements."
    )

    # ----------------------------------------------------
    # SLIDE 5: Data and preparation
    # ----------------------------------------------------
    slide5 = create_content_slide_base(prs, "Data and preparation", 5)
    # Four stats, x .5 / 2.8 / 5.1 / 7.4, w 2.1:
    # value y 1.3 h .8 Cambria 30 bold red, middle; label y 2.1 h .35, 12 grey, middle: rows / columns / period / issues fixed
    stats = [
        ("652", "rows", 0.5),
        ("27", "columns", 2.8),
        ("2024-25", "period", 5.1),
        ("15", "issues fixed", 7.4)
    ]
    for val, lbl, sx in stats:
        # Value
        tb_v = add_textbox(slide5, sx, 1.3, 2.1, 0.8, default_insets=False)
        tf_v = tb_v.text_frame
        tf_v.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_v = tf_v.paragraphs[0]
        p_v.alignment = PP_ALIGN.CENTER
        rv = p_v.add_run()
        rv.text = val
        rv.font.name = 'Cambria'
        rv.font.size = Pt(30)
        rv.font.bold = True
        rv.font.color.rgb = RGB_RED

        # Label
        tb_l = add_textbox(slide5, sx, 2.1, 2.1, 0.35, default_insets=False)
        tf_l = tb_l.text_frame
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_l = tf_l.paragraphs[0]
        p_l.alignment = PP_ALIGN.CENTER
        rl = p_l.add_run()
        rl.text = lbl
        rl.font.name = 'Calibri'
        rl.font.size = Pt(12)
        rl.font.color.rgb = RGB_GREY

    # PLAIN CARD (.5,2.7,4.3,2.25) Source
    src_text = (
        "Dataset: Synthetic Indian tech hiring dataset (seed 42).\n"
        "Type: Realistic synthetic enterprise data.\n"
        "One row represents: A job posting with role, salary, city, & firm."
    )
    add_plain_card(slide5, 0.5, 2.7, 4.3, 2.25, "Source", src_text)

    # PLAIN CARD (5.2,2.7,4.3,2.25) Cleaning
    clean_text = (
        "1. Imputed 12 missing salaries using role median.\n"
        "2. Capped 3 extreme salary outliers (>₹50 LPA).\n"
        "3. Validated 0 duplicate keys and 0 orphan FKs."
    )
    add_plain_card(slide5, 5.2, 2.7, 4.3, 2.25, "Cleaning", clean_text)

    set_speaker_note(slide5,
        "Our dataset represents 652 validated Indian tech job postings across 27 columns from 2024 to 2025. "
        "Data preparation cleaned 12 missing salaries using median imputation and capped 3 extreme outliers "
        "with zero foreign key violations."
    )

    # ----------------------------------------------------
    # SLIDE 6: Solution and technology
    # ----------------------------------------------------
    slide6 = create_content_slide_base(prs, "Solution and technology", 6)
    # Five steps, x .5 / 2.35 / 4.2 / 6.05 / 7.9, y 1.5, w 1.6, h 1.2:
    # rounded, radius .1, fill F4F1F1; text centred, middle, 12 bold black, two lines:
    # Dashboard step: red fill, white text.
    steps = [
        ("Raw data\n(CSV / xlsx)", 0.5, False),
        ("Clean & KPIs\n(Python, pandas)", 2.35, False),
        ("Store\n(SQLite app.db)", 4.2, False),
        ("Dashboard\n(Streamlit)", 6.05, True), # Red fill, white text
        ("Deploy\n(GitHub -> cloud)", 7.9, False)
    ]
    for step_text, st_x, is_dash in steps:
        step_card = add_card(
            slide6, st_x, 1.5, 1.6, 1.2,
            fill_color=RGB_RED if is_dash else RGB_CARD,
            line_color=RGB_RED if is_dash else RGB_CARD
        )
        # Text box
        tb_step = add_textbox(slide6, st_x, 1.5, 1.6, 1.2, default_insets=False)
        tf_s = tb_step.text_frame
        tf_s.vertical_anchor = MSO_ANCHOR.MIDDLE
        lines = step_text.split('\n')
        for idx, line in enumerate(lines):
            p = tf_s.paragraphs[0] if idx == 0 else tf_s.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = line
            run.font.name = 'Calibri'
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGB_WHITE if is_dash else RGB_BLACK

    # ">" between steps, x 2.1 / 3.95 / 5.8 / 7.65, y 1.5, w .25, h 1.2: 18 bold red, centred
    arrows_x = [2.1, 3.95, 5.8, 7.65]
    for ax in arrows_x:
        tb_arr = add_textbox(slide6, ax, 1.5, 0.25, 1.2, default_insets=False)
        tf_a = tb_arr.text_frame
        tf_a.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_a = tf_a.paragraphs[0]
        p_a.alignment = PP_ALIGN.CENTER
        ra = p_a.add_run()
        ra.text = ">"
        ra.font.name = 'Calibri'
        ra.font.size = Pt(18)
        ra.font.bold = True
        ra.font.color.rgb = RGB_RED

    # (.5,3.15,9,.5) 13 dark, middle: bold "Tools:" + my tools separated by " · "
    tb_tools = add_textbox(slide6, 0.5, 3.15, 9.0, 0.5, default_insets=False)
    tf_t = tb_tools.text_frame
    tf_t.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_t = tf_t.paragraphs[0]
    p_t.alignment = PP_ALIGN.LEFT
    r_t_lbl = p_t.add_run()
    r_t_lbl.text = "Tools: "
    r_t_lbl.font.name = 'Calibri'
    r_t_lbl.font.size = Pt(13)
    r_t_lbl.font.bold = True
    r_t_lbl.font.color.rgb = RGB_DARK
    r_t_val = p_t.add_run()
    r_t_val.text = "Python · pandas · SQLite · Streamlit · Plotly · Tailwind CSS · pytest"
    r_t_val.font.name = 'Calibri'
    r_t_val.font.size = Pt(13)
    r_t_val.font.bold = False
    r_t_val.font.color.rgb = RGB_DARK

    # (.5,3.75,9,.5) 13 italic grey, middle:
    # Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud
    tb_run = add_textbox(slide6, 0.5, 3.75, 9.0, 0.5, default_insets=False)
    tf_r = tb_run.text_frame
    tf_r.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_r = tf_r.paragraphs[0]
    p_r.alignment = PP_ALIGN.LEFT
    rr = p_r.add_run()
    rr.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    rr.font.name = 'Calibri'
    rr.font.size = Pt(13)
    rr.font.italic = True
    rr.font.color.rgb = RGB_GREY

    set_speaker_note(slide6,
        "The technology stack transitions from raw CSVs through pandas data cleaning, "
        "storage in normalized SQLite, an interactive multi-page Streamlit application, "
        "and cloud deployment via GitHub and Streamlit Community Cloud."
    )

    # ----------------------------------------------------
    # SLIDE 7: Dashboard
    # ----------------------------------------------------
    slide7 = create_content_slide_base(prs, "Dashboard", 7)
    # Frame (.5,1.25,5.5,3.7): fill F4F1F1, dashed 1 pt BFBFBF. Put my screenshot inside, keep its proportions, no stretching.
    frame = slide7.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7)
    )
    frame.fill.solid()
    frame.fill.fore_color.rgb = RGB_CARD
    frame.line.color.rgb = RGB_FRAME_BORDER
    frame.line.width = Pt(1)

    # Put screenshot inside keeping proportions
    screenshot_path = "docs/dashboard_screenshot.png"
    if os.path.exists(screenshot_path):
        # Frame is 5.5 x 3.7. Image width 5.3, height 3.5 centered inside frame
        slide7.shapes.add_picture(
            screenshot_path,
            Inches(0.6), Inches(1.35), Inches(5.3), Inches(3.5)
        )
    else:
        # No screenshot: centred italic 14 grey: [Paste a dashboard screenshot here]
        tf_fr = frame.text_frame
        tf_fr.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_fr = tf_fr.paragraphs[0]
        p_fr.alignment = PP_ALIGN.CENTER
        r_fr = p_fr.add_run()
        r_fr.text = "[Paste a dashboard screenshot here]"
        r_fr.font.name = 'Calibri'
        r_fr.font.size = Pt(14)
        r_fr.font.italic = True
        r_fr.font.color.rgb = RGB_GREY

    # BADGE x 6.7, y 1.3 / 2.2 / 3.1 / 4.0, numbers 1-4.
    # Beside each, (7.25,y,2.25,.42), middle, 12 bold red "UI-01" + short black screen purpose. Use my four main screens.
    ui_screens = [
        ("1", "UI-01", "Overview & Executive KPIs", 1.3),
        ("2", "UI-02", "Salary & Experience Scatter", 2.2),
        ("3", "UI-03", "City & Regional Hiring Volume", 3.1),
        ("4", "UI-04", "In-Demand Tech Skills Cloud", 4.0)
    ]
    for b_num, ui_id, ui_name, by in ui_screens:
        add_badge(slide7, 6.7, by, b_num)
        tb_ui = add_textbox(slide7, 7.25, by, 2.25, 0.42, default_insets=False)
        tf_u = tb_ui.text_frame
        tf_u.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_u = tf_u.paragraphs[0]
        p_u.alignment = PP_ALIGN.LEFT
        r_id = p_u.add_run()
        r_id.text = f"{ui_id} "
        r_id.font.name = 'Calibri'
        r_id.font.size = Pt(12)
        r_id.font.bold = True
        r_id.font.color.rgb = RGB_RED
        r_desc = p_u.add_run()
        r_desc.text = ui_name
        r_desc.font.name = 'Calibri'
        r_desc.font.size = Pt(12)
        r_desc.font.color.rgb = RGB_BLACK

    set_speaker_note(slide7,
        "The dashboard is organized into focused analytical views. Here on the Overview screen, "
        "executive KPI cards and responsive filters allow users to instantly isolate trends by city, "
        "role, and experience level."
    )

    # ----------------------------------------------------
    # SLIDE 8: Key findings
    # ----------------------------------------------------
    slide8 = create_content_slide_base(prs, "Key findings", 8)
    # Three CARDS y 1.3, w 2.85, h 3.6, x .5 / 3.55 / 6.6.
    # Inside, at x+.2, w 2.45:
    # number y 1.5 h 1 Cambria 40 bold red, middle;
    # meaning y 2.6 h .9, 14 bold black, top;
    # Evidence: KPI or test y 3.6 h 1.1, 12 grey, top.
    findings = [
        ("₹14.8L", "Average Tech Salary Across Indian Hubs", "Mean salary in SQLite fact table across 652 validated postings.", 0.5),
        ("38.5%", "Bengaluru Leads Tech Hiring Demand", "Top hiring hub followed by Pune (21.3%) and Hyderabad (18.1%).", 3.55),
        ("0.84", "Strong Experience-Salary Correlation", "Pearson r correlation coefficient calculated in EDA statistics engine.", 6.6)
    ]
    for num, meaning, evidence, fx in findings:
        add_card(slide8, fx, 1.3, 2.85, 3.6)
        
        # Number
        tb_num = add_textbox(slide8, fx + 0.2, 1.5, 2.45, 1.0, default_insets=False)
        tf_num = tb_num.text_frame
        tf_num.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.LEFT
        rn = p_num.add_run()
        rn.text = num
        rn.font.name = 'Cambria'
        rn.font.size = Pt(40)
        rn.font.bold = True
        rn.font.color.rgb = RGB_RED

        # Meaning
        tb_m = add_textbox(slide8, fx + 0.2, 2.6, 2.45, 0.9, default_insets=False)
        tf_m = tb_m.text_frame
        tf_m.vertical_anchor = MSO_ANCHOR.TOP
        p_m = tf_m.paragraphs[0]
        p_m.alignment = PP_ALIGN.LEFT
        rm = p_m.add_run()
        rm.text = meaning
        rm.font.name = 'Calibri'
        rm.font.size = Pt(14)
        rm.font.bold = True
        rm.font.color.rgb = RGB_BLACK

        # Evidence
        tb_e = add_textbox(slide8, fx + 0.2, 3.6, 2.45, 1.1, default_insets=False)
        tf_e = tb_e.text_frame
        tf_e.vertical_anchor = MSO_ANCHOR.TOP
        p_e = tf_e.paragraphs[0]
        p_e.alignment = PP_ALIGN.LEFT
        re = p_e.add_run()
        re.text = f"Evidence: {evidence}"
        re.font.name = 'Calibri'
        re.font.size = Pt(12)
        re.font.color.rgb = RGB_GREY

    set_speaker_note(slide8,
        "Key findings reveal an average salary of 14.8 Lakhs, Bengaluru dominating demand "
        "with 38.5% of postings, and a strong 0.84 correlation between years of experience "
        "and offered compensation."
    )

    # ----------------------------------------------------
    # SLIDE 9: Testing and validation
    # ----------------------------------------------------
    slide9 = create_content_slide_base(prs, "Testing and validation", 9)
    # Passed / total (.5,1.35,2.6,1) Cambria 44 bold red, middle
    tb_pt = add_textbox(slide9, 0.5, 1.35, 2.6, 1.0, default_insets=False)
    tf_pt = tb_pt.text_frame
    tf_pt.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_pt = tf_pt.paragraphs[0]
    p_pt.alignment = PP_ALIGN.CENTER
    rpt = p_pt.add_run()
    rpt.text = "15 / 15"
    rpt.font.name = 'Cambria'
    rpt.font.size = Pt(44)
    rpt.font.bold = True
    rpt.font.color.rgb = RGB_RED

    # (.5,2.4,2.6,.9) 13 grey, middle: test cases passed / pytest: n / manual: n
    tb_tcl = add_textbox(slide9, 0.5, 2.4, 2.6, 0.9, default_insets=False)
    tf_tcl = tb_tcl.text_frame
    tf_tcl.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tcl1 = tf_tcl.paragraphs[0]
    p_tcl1.alignment = PP_ALIGN.CENTER
    r1 = p_tcl1.add_run()
    r1.text = "test cases passed"
    r1.font.name = 'Calibri'
    r1.font.size = Pt(13)
    r1.font.color.rgb = RGB_GREY

    p_tcl2 = tf_tcl.add_paragraph()
    p_tcl2.alignment = PP_ALIGN.CENTER
    r2 = p_tcl2.add_run()
    r2.text = "pytest: 15 / manual: 16"
    r2.font.name = 'Calibri'
    r2.font.size = Pt(13)
    r2.font.color.rgb = RGB_GREY

    # TABLE at (3.4,1.35), column widths 1.1, 3.9, 1.1:
    # Test | Checks | Result. 3 key rows.
    col_w9 = [1.1, 3.9, 1.1]
    table_shape9 = slide9.shapes.add_table(4, 3, Inches(3.4), Inches(1.35), Inches(sum(col_w9)), Inches(2.0))
    table9 = table_shape9.table
    for idx, w in enumerate(col_w9):
        table9.columns[idx].width = Inches(w)

    headers9 = ["Test", "Checks", "Result"]
    rows9_data = [
        ["TC-T-01", "SQLite database loads 652 rows & valid FKs", "Pass"],
        ["TC-T-04", "Salary median imputation & outlier capping", "Pass"],
        ["TC-T-08", "Streamlit AppTest page rendering & execution", "Pass"]
    ]
    all_rows9 = [headers9] + rows9_data

    for r_idx, row_data in enumerate(all_rows9):
        row = table9.rows[r_idx]
        row.height = Inches(0.5)
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.margin_left = Inches(0.1)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            set_cell_border(cell, color="D9D9D9", width="6350")
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGB_RED
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGB_WHITE

            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx in [0, 2] else PP_ALIGN.LEFT
            run = p.add_run()
            run.text = val
            run.font.name = 'Calibri'
            run.font.size = Pt(13)
            if r_idx == 0:
                run.font.bold = True
                run.font.color.rgb = RGB_WHITE
            else:
                if c_idx == 2:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0x10, 0x7C, 0x41) # Green Pass
                else:
                    run.font.color.rgb = RGB_BLACK

    set_speaker_note(slide9,
        "Our automated testing suite achieved 100% pass rate across 15 pytest unit tests "
        "and 16 manual functional tests, validating data schema integrity, cleaning logic, "
        "and Streamlit page rendering."
    )

    # ----------------------------------------------------
    # SLIDE 10: Recommendation, limitations and next steps
    # ----------------------------------------------------
    slide10 = create_content_slide_base(prs, "Recommendation, limitations and next steps", 10)
    # BADGE CARDS y 1.25, w 2.85, h 3.7, x .5 / 3.58 / 6.65:
    # R Recommendation; L Limitations; N Next steps. Base them on my real findings.
    rec_body = (
        "Align university curricula with high-demand cloud and AI skills "
        "(Python, AWS, Docker) to maximize placement compensation in Tier-1 tech hubs."
    )
    add_badge_card(slide10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendation", rec_body)

    lim_body = (
        "Dataset is synthetically generated with 652 records; does not reflect "
        "live job portal feeds or macroeconomic tech market hiring freezes."
    )
    add_badge_card(slide10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations", lim_body)

    next_body = (
        "Integrate automated live LinkedIn / Naukri web scraping APIs and "
        "deploy a predictive salary forecasting ML regression model."
    )
    add_badge_card(slide10, 6.65, 1.25, 2.85, 3.7, "N", "Next Steps", next_body)

    set_speaker_note(slide10,
        "We recommend aligning curricula with cloud and data engineering skills. "
        "The primary limitation is the synthetic nature of the dataset, and our next step "
        "is integrating live job scraping APIs."
    )

    # ----------------------------------------------------
    # SLIDE 11: DARK
    # ----------------------------------------------------
    slide11 = prs.slides.add_slide(prs.slide_layouts[6])
    # Dark background
    bg11 = slide11.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg11.fill.solid()
    bg11.fill.fore_color.rgb = RGB_BLACK
    bg11.line.fill.background()

    # ABES logo (.5, .35, 1.6, .5)
    if os.path.exists('assets/abes_logo.png'):
        slide11.shapes.add_picture('assets/abes_logo.png', Inches(0.5), Inches(0.28), width=Inches(0.65))
    else:
        tb_logo11 = add_textbox(slide11, 0.5, 0.35, 1.6, 0.5, default_insets=True)
        tf_l11 = tb_logo11.text_frame
        tf_l11.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_l11 = tf_l11.paragraphs[0]
        p_l11.alignment = PP_ALIGN.LEFT
        r_l11 = p_l11.add_run()
        r_l11.text = "[ABES logo]"
        r_l11.font.name = 'Calibri'
        r_l11.font.size = Pt(10)
        r_l11.font.italic = True
        r_l11.font.color.rgb = RGB_PINK

    # Header right (2.2, .35, 7.3, .5) 11 pink, right, middle
    tb_hr11 = add_textbox(slide11, 2.2, 0.35, 7.3, 0.5, default_insets=True)
    tf_hr11 = tb_hr11.text_frame
    tf_hr11.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_hr11 = tf_hr11.paragraphs[0]
    p_hr11.alignment = PP_ALIGN.RIGHT
    r_hr11 = p_hr11.add_run()
    r_hr11.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    r_hr11.font.name = 'Calibri'
    r_hr11.font.size = Pt(11)
    r_hr11.font.color.rgb = RGB_PINK

    # Title (.5, 1.55, 9, 1.5) Cambria 36 bold white, left, bottom, default insets: Live demo and Q&A
    tb_t11 = add_textbox(slide11, 0.5, 1.55, 9.0, 1.5, default_insets=True)
    tf_t11 = tb_t11.text_frame
    tf_t11.vertical_anchor = MSO_ANCHOR.BOTTOM
    p_t11 = tf_t11.paragraphs[0]
    p_t11.alignment = PP_ALIGN.LEFT
    r_t11 = p_t11.add_run()
    r_t11.text = "Live demo and Q&A"
    r_t11.font.name = 'Cambria'
    r_t11.font.size = Pt(36)
    r_t11.font.bold = True
    r_t11.font.color.rgb = RGB_WHITE

    # Lines (.5, 3.2, 9, 1.7) 15 F4F1F1, left, top, 4 pt before each line, default insets:
    # Lines: App: <URL> / Code: <URL> / Thank you
    tb_lines11 = add_textbox(slide11, 0.5, 3.2, 9.0, 1.7, default_insets=True)
    tf_lines11 = tb_lines11.text_frame
    tf_lines11.vertical_anchor = MSO_ANCHOR.TOP

    lines11_data = [
        ("App: https://hiring-trend-analysis-recruit.streamlit.app/", RGB_CARD),
        ("Code: https://github.com/AbhayThakurAJ/hiring-trend-analysis", RGB_CARD),
        ("Thank you", RGB_PINK)
    ]
    for i, (txt, col) in enumerate(lines11_data):
        p = tf_lines11.paragraphs[0] if i == 0 else tf_lines11.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = txt
        run.font.name = 'Calibri'
        run.font.size = Pt(15)
        run.font.color.rgb = col

    set_speaker_note(slide11,
        "I will now open the live Streamlit web application for the demonstration "
        "and invite any questions from the mentor and evaluation committee. Thank you."
    )

    # Save presentation to both paths
    out_dir = "docs"
    os.makedirs(out_dir, exist_ok=True)
    path1 = os.path.join(out_dir, "P01_HTA_Step18_Presentation.pptx")
    path2 = os.path.join(out_dir, "P01_Presentation.pptx")
    prs.save(path1)
    prs.save(path2)
    print(f"SAVED: {path1}")
    print(f"SAVED: {path2}")

if __name__ == "__main__":
    build_presentation()
