import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_document():
    doc = docx.Document()
    
    # Page setup - Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Header & Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.text = "HTA-01: Hiring Trend Analysis | Project Understanding (Step 1)"
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hp.style.font.name = "Calibri"
        hp.style.font.size = Pt(8.5)
        hp.style.font.color.rgb = RGBColor(120, 120, 120)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.text = "ABES Engineering College — MCA 2nd Year Data Analytics | Academic Project Deliverable"
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.style.font.name = "Calibri"
        fp.style.font.size = Pt(8.5)
        fp.style.font.color.rgb = RGBColor(120, 120, 120)

    # Styles setup
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = RGBColor(30, 41, 59) # Slate 800
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)

    # Document Title Block
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    run_sub = title_p.add_run("ENTERPRISE DATA ANALYTICS & RECRUITMENT INTELLIGENCE\n")
    run_sub.font.size = Pt(9.5)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(37, 99, 235) # Blue 600

    run_title = title_p.add_run("STEP 1: PROJECT UNDERSTANDING & SCOPE CHARTER")
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(15, 23, 42) # Slate 900

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(14)
    run_desc = p_meta.add_run("Project ID: HTA-01 | Domain: HR Analytics & Talent Acquisition | ABES Engineering College")
    run_desc.font.size = Pt(10)
    run_desc.font.italic = True
    run_desc.font.color.rgb = RGBColor(100, 116, 139)

    # Horizontal Divider rule
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    r_div = p_div.add_run("―" * 58)
    r_div.font.color.rgb = RGBColor(203, 213, 225)

    # Section 1: Academic & Project Metadata
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)
    r1 = h1.add_run("1. Document Control & Academic Profile")
    r1.font.size = Pt(14)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(30, 58, 138)

    t_meta = doc.add_table(rows=6, cols=2)
    t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        ("Project Identifier", "HTA-01: Hiring Trend Analysis & Recruitment Performance System"),
        ("Industry Domain", "Human Resources Analytics / Talent Acquisition Intelligence"),
        ("Student / Learner", "MCA 2nd Year Student (Data Analytics Specialization), ABES"),
        ("Roll Number", "2300320140001 (Default Assumption — Modifiable)"),
        ("Faculty Project Mentor", "Internal Faculty Guide, Department of MCA, ABES Engineering College"),
        ("Target Technology Stack", "Python 3.11+, Streamlit, SQLite 3, Plotly Express, Tailwind CSS, Pytest")
    ]
    for i, (k, v) in enumerate(meta_data):
        row = t_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        c0.text = k
        c1.text = v
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)
        c0.paragraphs[0].runs[0].font.bold = True
        c0.paragraphs[0].runs[0].font.size = Pt(9.5)
        c1.paragraphs[0].runs[0].font.size = Pt(9.5)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: Problem Statement
    h2 = doc.add_heading(level=1)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(6)
    r2 = h2.add_run("2. Problem Statement")
    r2.font.size = Pt(14)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(30, 58, 138)

    p_ps = doc.add_paragraph()
    p_ps.paragraph_format.space_after = Pt(10)
    r_ps_bold = p_ps.add_run("Core Problem Statement: ")
    r_ps_bold.bold = True
    r_ps = p_ps.add_run(
        "Talent acquisition teams and HR leadership lack centralized, quantitative visibility into recruitment velocity, departmental hiring demands, candidate attrition across funnel stages, and compensation benchmarks, leading to prolonged time-to-fill, escalating agency costs, and suboptimal workforce planning."
    )
    r_ps.font.color.rgb = RGBColor(15, 23, 42)

    p_ps_exp = doc.add_paragraph(
        "Recruitment pipelines generate large amounts of transactional data across job postings, applications, screening stages, interview schedules, compensation offers, and onboarding records. In the absence of an integrated analytical platform, hiring leaders rely on fragmented spreadsheets, resulting in data latency, lack of visibility into bottlenecks, and inability to forecast headcount capacity."
    )

    # Section 3: Intended Audience & User Personas
    h3 = doc.add_heading(level=1)
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(6)
    r3 = h3.add_run("3. Intended Audience & User Profiles")
    r3.font.size = Pt(14)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(30, 58, 138)

    t_users = doc.add_table(rows=5, cols=4)
    t_users.alignment = WD_TABLE_ALIGNMENT.CENTER
    user_headers = ["User Role / Persona", "Primary Objectives", "Access Frequency", "Key Operational Actions"]
    for j, text in enumerate(user_headers):
        cell = t_users.rows[0].cells[j]
        cell.text = text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 100, 100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9.5)

    user_rows = [
        ("Talent Acquisition Lead / Recruiting Manager", "Monitor daily pipeline velocity, recruiter workload, and stage conversion bottlenecks.", "Daily (Morning sync / Daily review)", "Reallocate recruiter loads, identify stalled candidate requisitions, evaluate stage drop-offs."),
        ("Department Hiring Managers (Tech, Sales, Product)", "Track department-specific requisitions, interview-to-offer ratios, and candidate caliber.", "Weekly (Prior to team pipeline reviews)", "Filter open roles by seniority/tech stack, inspect interview pass rates, review pending offer acceptances."),
        ("VP of HR / Chief People Officer", "Oversee strategic talent metrics, diversity ratios, time-to-fill trends, and source channel ROI.", "Bi-weekly / Monthly (Exec reporting)", "Formulate talent acquisition budgets, decide agency vendor renewals, calibrate compensation bands."),
        ("Finance Director / Head of FP&A", "Validate hiring actuals against budgeted headcount and monitor cost-per-hire economics.", "Monthly / Quarterly (Quarterly close)", "Audit total recruitment expenditure, compare accepted salaries against salary bands, approve requisition expansions.")
    ]
    for i, row in enumerate(user_rows):
        cells = t_users.rows[i+1].cells
        for j, val in enumerate(row):
            cells[j].text = val
            set_cell_background(cells[j], "F8FAFC" if i % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[j], 60, 60, 100, 100)
            p = cells[j].paragraphs[0]
            if j == 0:
                p.runs[0].font.bold = True
            for r in p.runs:
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(30, 41, 59)
    
    t_users.rows[0].cells[0].width = Inches(1.8)
    t_users.rows[0].cells[1].width = Inches(1.8)
    t_users.rows[0].cells[2].width = Inches(1.2)
    t_users.rows[0].cells[3].width = Inches(1.7)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: Stakeholders
    h4 = doc.add_heading(level=1)
    h4.paragraph_format.space_before = Pt(14)
    h4.paragraph_format.space_after = Pt(6)
    r4 = h4.add_run("4. Stakeholder Matrix")
    r4.font.size = Pt(14)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(30, 58, 138)

    t_sh = doc.add_table(rows=6, cols=6)
    t_sh.alignment = WD_TABLE_ALIGNMENT.CENTER
    sh_headers = ["SH ID", "Stakeholder", "Organizational Role", "Interest", "Influence", "Decision Made with the Application"]
    for j, text in enumerate(sh_headers):
        cell = t_sh.rows[0].cells[j]
        cell.text = text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9)

    sh_data = [
        ("SH-01", "Head of Talent Acquisition", "Business Process Owner & Primary Sponsor", "High", "High", "Decides recruitment channel spend allocations, recruiter capacity re-distribution, and SLA guidelines."),
        ("SH-02", "Senior Technical Recruiter", "Operational Power User", "High", "Low", "Prioritizes candidate screening queues, flags stage bottleneck delays, and escalates compensation deviations."),
        ("SH-03", "Engineering Hiring Manager", "Departmental Consumer", "Medium", "Medium", "Calibrates technical interview bar, adjusts role prerequisites, and initiates offer requests for candidates."),
        ("SH-04", "Chief Financial Officer (CFO)", "Budget & Executive Sponsor", "Medium", "High", "Approves annual hiring budget, validates cost-per-hire thresholds, and sanctions department expansion caps."),
        ("SH-05", "Academic Guide / Faculty Evaluator", "ABES Academic Mentor & Assessor", "High", "High", "Evaluates data architecture rigor, statistical validity, code quality, UI responsiveness, and viva defense.")
    ]
    for i, row in enumerate(sh_data):
        cells = t_sh.rows[i+1].cells
        for j, val in enumerate(row):
            cells[j].text = val
            set_cell_background(cells[j], "F8FAFC" if i % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[j], 60, 60, 80, 80)
            p = cells[j].paragraphs[0]
            if j in (0, 3, 4):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if j == 0:
                p.runs[0].font.bold = True
            for r in p.runs:
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 5: Operational Questions
    h5 = doc.add_heading(level=1)
    h5.paragraph_format.space_before = Pt(14)
    h5.paragraph_format.space_after = Pt(6)
    r5 = h5.add_run("5. Operational Questions & Data Feasibility Verification")
    r5.font.size = Pt(14)
    r5.font.bold = True
    r5.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "To ensure analytical relevance, ten operational questions (OQs) have been formalized. Each OQ is cross-checked against data availability in our planned SQLite schema to ensure 100% computational feasibility without synthetic gaps."
    )

    t_oq = doc.add_table(rows=11, cols=4)
    t_oq.alignment = WD_TABLE_ALIGNMENT.CENTER
    oq_headers = ["OQ ID", "Operational Question", "Required Data Fields / Metrics", "Feasibility & Schema Check"]
    for j, text in enumerate(oq_headers):
        cell = t_oq.rows[0].cells[j]
        cell.text = text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9)

    oq_data = [
        ("OQ-01", "What is the average and median Time-to-Hire (in calendar days) segmented across departments and seniority levels?", "applications.applied_date, applications.hire_date, jobs.department, jobs.experience_level", "Fully Feasible: Computed via Julian day delta: (julianday(hire_date) - julianday(applied_date))."),
        ("OQ-02", "Which candidate sourcing channels (LinkedIn, Referral, Direct, Agency) demonstrate the highest Offer Acceptance Rate (OAR)?", "candidates.source_channel, offers.status ('Accepted', 'Declined', 'Pending')", "Fully Feasible: Aggregate count(offers.accepted) / count(offers.issued) grouped by source_channel."),
        ("OQ-03", "At which stage of the recruitment funnel (Screening -> Technical -> Managerial -> HR -> Offer) do applicants experience the steepest drop-off?", "application_stages.stage_name, application_stages.status, application_stages.duration_days", "Fully Feasible: Funnel attrition waterfall using stage transition counts and pass percentages."),
        ("OQ-04", "How do accepted compensation offers deviate from the pre-approved salary band (min/max) across job families?", "offers.offered_salary, jobs.min_salary, jobs.max_salary, jobs.job_title", "Fully Feasible: Variance calculation: (offered_salary - midpoint(min_salary, max_salary))."),
        ("OQ-05", "What are the seasonal monthly patterns in application volume and job requisition openings over the trailing 24 months?", "applications.applied_date, jobs.posted_date", "Fully Feasible: Time-series monthly aggregation: strftime('%Y-%m', applied_date)."),
        ("OQ-06", "What is the demographic and geographical distribution (Gender, Tier 1/2/3 cities) of applicants compared to final hires?", "candidates.gender, candidates.location_tier, applications.status", "Fully Feasible: Proportion comparison between applicant pool and hired cohort."),
        ("OQ-07", "What are the predominant reasons cited by candidates when rejecting job offers, and what is the overall renege rate?", "offers.status, offers.decline_reason, offers.offered_date", "Fully Feasible: Pareto frequency analysis on decline_reason categories."),
        ("OQ-08", "How does recruiter throughput vary in terms of volume processed, active candidate load, and requisition closure rate?", "applications.recruiter_id, recruiters.name, applications.status, jobs.closed_date", "Fully Feasible: Recruiter KPI scorecard aggregating closed vs active requisitions."),
        ("OQ-09", "What is the interview-to-offer conversion efficiency ratio across different university and educational tiers?", "candidates.education_tier, applications.stage_reached, offers.offer_id", "Fully Feasible: Conversion ratio: count(offers) / count(interviewed_candidates) by education tier."),
        ("OQ-10", "What is the average cost-per-hire across sourcing channels and departments?", "applications.sourcing_cost, applications.agency_fee, applications.status='Hired'", "Fully Feasible: Cost sum (sourcing_cost + agency_fee) divided by total successful hires.")
    ]

    for i, row in enumerate(oq_data):
        cells = t_oq.rows[i+1].cells
        for j, val in enumerate(row):
            cells[j].text = val
            set_cell_background(cells[j], "F8FAFC" if i % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[j], 60, 60, 80, 80)
            p = cells[j].paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.bold = True
            for r in p.runs:
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 6: Objectives
    h6 = doc.add_heading(level=1)
    h6.paragraph_format.space_before = Pt(14)
    h6.paragraph_format.space_after = Pt(6)
    r6 = h6.add_run("6. Project Objectives (SMART)")
    r6.font.size = Pt(14)
    r6.font.bold = True
    r6.font.color.rgb = RGBColor(30, 58, 138)

    objectives = [
        ("OBJ-01", "Comprehensive KPI Pipeline Formulation", "Architect and automate 10 core talent acquisition metrics (including Time-to-Hire, Funnel Drop-off, Offer Acceptance Rate, Cost-per-Hire) derived from structured transactional tables, verifiable via SQL queries with 100% calculation accuracy."),
        ("OBJ-02", "Relational Schema & Synthetic Data Modeling", "Design a normalized 3NF SQLite database populated with 5,000 realistic, statistically coherent candidate records spanning 100+ job requisitions across 24 historical months, executing queries under 200 milliseconds."),
        ("OBJ-03", "Interactive Analytical Visualizations", "Construct a responsive Streamlit dashboard featuring 8+ Plotly visualizations with dynamic cross-filtering by department, role seniority, recruiter, and date range, ensuring cognitive clarity and visual accessibility."),
        ("OBJ-04", "Rigorous Automated Quality Assurance", "Develop an automated test suite containing 15+ pytest unit and integration test cases covering data integrity, boundary values, KPI aggregation logic, and UI filters, maintaining >85% code coverage."),
        ("OBJ-05", "Production Cloud Deployment & Documentation", "Package the application for zero-downtime deployment on Streamlit Community Cloud and produce the complete ABES-standard academic documentation and presentation deck for final institutional evaluation.")
    ]

    t_obj = doc.add_table(rows=6, cols=3)
    t_obj.alignment = WD_TABLE_ALIGNMENT.CENTER
    obj_headers = ["Objective ID", "SMART Goal Title", "Specific, Measurable, Achievable, Relevant, Time-bound Description"]
    for j, text in enumerate(obj_headers):
        cell = t_obj.rows[0].cells[j]
        cell.text = text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9)

    for i, row in enumerate(objectives):
        cells = t_obj.rows[i+1].cells
        for j, val in enumerate(row):
            cells[j].text = val
            set_cell_background(cells[j], "F8FAFC" if i % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[j], 60, 60, 80, 80)
            p = cells[j].paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.bold = True
            elif j == 1:
                p.runs[0].font.bold = True
            for r in p.runs:
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 7: Scope
    h7 = doc.add_heading(level=1)
    h7.paragraph_format.space_before = Pt(14)
    h7.paragraph_format.space_after = Pt(6)
    r7 = h7.add_run("7. Project Scope Specification")
    r7.font.size = Pt(14)
    r7.font.bold = True
    r7.font.color.rgb = RGBColor(30, 58, 138)

    t_scope = doc.add_table(rows=6, cols=3)
    t_scope.alignment = WD_TABLE_ALIGNMENT.CENTER
    scope_headers = ["Functional Area", "In-Scope (Deliverable in Phase 1)", "Out-of-Scope (Deferred / Excluded)"]
    for j, text in enumerate(scope_headers):
        cell = t_scope.rows[0].cells[j]
        cell.text = text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9)

    scope_data = [
        ("Data Ingestion & Modeling", "Synthesis of 5,000 realistic records with relational referential integrity in SQLite 3.", "Live web scraping of job portals or real-time ATS webhook integration (Workday/Greenhouse)."),
        ("Exploratory Data Analysis", "Interactive multi-dimensional filtering (dept, recruiter, seniority, date) and funnel drop-off calculation.", "Unsupervised clustering of unstructured candidate resumes or NLP text extraction."),
        ("Visualization & Dashboard", "Plotly charts (funnel, bar, scatter, boxplot, heatmap) embedded in modern Tailwind-inspired Streamlit UI.", "Native mobile app (iOS/Android) or multi-tenant user authentication (SSO/OAuth)."),
        ("Statistical Analysis", "Descriptive statistics, variance analysis, salary band deviation, correlation matrices.", "Deep learning predictive offer acceptance models or automated candidate ranking bots."),
        ("Deployment & Testing", "Streamlit Community Cloud deployment with continuous pytest validation suite.", "Kubernetes orchestration, multi-node clustering, or enterprise high-availability load balancers.")
    ]

    for i, row in enumerate(scope_data):
        cells = t_scope.rows[i+1].cells
        for j, val in enumerate(row):
            cells[j].text = val
            set_cell_background(cells[j], "F8FAFC" if i % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[j], 60, 60, 80, 80)
            p = cells[j].paragraphs[0]
            if j == 0:
                p.runs[0].font.bold = True
            for r in p.runs:
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 8: Assumptions & Constraints
    h8 = doc.add_heading(level=1)
    h8.paragraph_format.space_before = Pt(14)
    h8.paragraph_format.space_after = Pt(6)
    r8 = h8.add_run("8. Assumptions and Constraints")
    r8.font.size = Pt(14)
    r8.font.bold = True
    r8.font.color.rgb = RGBColor(30, 58, 138)

    t_ac = doc.add_table(rows=7, cols=4)
    t_ac.alignment = WD_TABLE_ALIGNMENT.CENTER
    ac_headers = ["Item ID", "Category", "Description", "Operational Impact / Mitigation"]
    for j, text in enumerate(ac_headers):
        cell = t_ac.rows[0].cells[j]
        cell.text = text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9)

    ac_data = [
        ("AS-01", "Assumption: Data Realism", "Synthetic data generated via Faker/NumPy replicates realistic corporate recruitment distributions.", "Ensures realistic variance in time-to-hire, funnel stage drop-offs, and salary ranges without privacy violations."),
        ("AS-02", "Assumption: Learner Identity", "Student assumed as MCA 2nd Year Data Analytics, ABES Engineering College (Roll 2300320140001).", "Can be updated at any project step if the student provides specific alternate credentials."),
        ("AS-03", "Assumption: Hosting Limits", "Streamlit Community Cloud 1GB RAM container is adequate for analytical query processing.", "SQLite queries are indexed and aggregated in memory with minimal memory footprint (<150MB)."),
        ("CO-01", "Constraint: Data Privacy", "Strict prohibition of real personally identifiable information (PII) under academic and corporate privacy policies.", "Full adherence to synthetic generation (Rule 5); candidate names, emails, and phone numbers are mock entities."),
        ("CO-02", "Constraint: Tech Stack", "Architecture restricted to open-source Python, SQLite, Streamlit, Plotly, and pytest.", "Ensures 100% reproducibility on any development machine without licensing or proprietary software hurdles."),
        ("CO-03", "Constraint: Academic Timeline", "Project deliverables must strictly align with the 18-step professional methodology and ABES submission format.", "Rigorous milestone tracking ensuring all artifacts (BRD, PRD, RTM, tests, slides) are maintained.")
    ]

    for i, row in enumerate(ac_data):
        cells = t_ac.rows[i+1].cells
        for j, val in enumerate(row):
            cells[j].text = val
            set_cell_background(cells[j], "F8FAFC" if i % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[j], 60, 60, 80, 80)
            p = cells[j].paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.bold = True
            elif j == 1:
                p.runs[0].font.bold = True
            for r in p.runs:
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 9: Success Criteria
    h9 = doc.add_heading(level=1)
    h9.paragraph_format.space_before = Pt(14)
    h9.paragraph_format.space_after = Pt(6)
    r9 = h9.add_run("9. Measurable Success Criteria")
    r9.font.size = Pt(14)
    r9.font.bold = True
    r9.font.color.rgb = RGBColor(30, 58, 138)

    t_sc = doc.add_table(rows=6, cols=4)
    t_sc.alignment = WD_TABLE_ALIGNMENT.CENTER
    sc_headers = ["Criteria ID", "Evaluation Dimension", "Target Threshold Metric", "Validation & Verification Method"]
    for j, text in enumerate(sc_headers):
        cell = t_sc.rows[0].cells[j]
        cell.text = text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9)

    sc_data = [
        ("SC-01", "Analytical Accuracy", "100% of OQ-01 through OQ-10 verifiable via deterministic SQL queries and test assertions.", "Automated verification scripts comparing SQL aggregations against pandas and pytest test cases."),
        ("SC-02", "Application Latency", "Dashboard initial page load < 2.0s; interactive slice-and-dice filter updates < 500ms.", "Execution profiling on standard Streamlit Community Cloud runner with 5,000 synthetic records."),
        ("SC-03", "Test Suite Coverage", "Minimum 85% statement coverage with zero failing unit or data validation tests.", "Automated pytest run with pytest-cov report generation."),
        ("SC-04", "UI/UX Usability", "Zero visual glitches, intuitive 3-click insight navigation, responsive on standard screen resolutions.", "Heuristic evaluation and cognitive walkthrough across all defined user personas."),
        ("SC-05", "Academic Rigor & Compliance", "Complete traceability across all 18 project steps, matching ABES thesis and presentation standards.", "Formal review of generated Word documentation, RTM matrix, and viva defense presentation.")
    ]

    for i, row in enumerate(sc_data):
        cells = t_sc.rows[i+1].cells
        for j, val in enumerate(row):
            cells[j].text = val
            set_cell_background(cells[j], "F8FAFC" if i % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[j], 60, 60, 80, 80)
            p = cells[j].paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.bold = True
            elif j == 1:
                p.runs[0].font.bold = True
            for r in p.runs:
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(30, 41, 59)

    # Save documents
    os.makedirs("docs", exist_ok=True)
    file_path = os.path.abspath("docs/01_Project_Understanding.docx")
    doc.save(file_path)
    print(f"Document successfully created at {file_path}")

create_document()
