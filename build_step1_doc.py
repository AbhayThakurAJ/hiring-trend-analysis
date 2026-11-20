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

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_document():
    doc = docx.Document()
    
    # Page setup - Margins (1 inch all around)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Header
        header = section.header
        hp = header.paragraphs[0]
        hp.text = "HTA-01: Hiring Trend Analysis | Step 1: Project Understanding & Scope Charter"
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hp.style.font.name = "Calibri"
        hp.style.font.size = Pt(8.5)
        hp.style.font.color.rgb = RGBColor(100, 116, 139)
        
        # Footer
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.text = "Student: Abhayjeet Singh (Roll: 2500320140004) | Mentor: Dinesh Sereno | MCA 2nd Yr DA, ABES"
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.style.font.name = "Calibri"
        fp.style.font.size = Pt(8.5)
        fp.style.font.color.rgb = RGBColor(100, 116, 139)

    # Base typography
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
    run_sub = title_p.add_run("MCA CAPSTONE PROJECT — DATA ANALYTICS TRACK\nDEPARTMENT OF COMPUTER APPLICATIONS, ABES ENGINEERING COLLEGE\n")
    run_sub.font.size = Pt(9.5)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(37, 99, 235) # Blue 600

    run_title = title_p.add_run("STEP 1: PROJECT UNDERSTANDING & SCOPE CHARTER")
    run_title.font.size = Pt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(15, 23, 42) # Slate 900

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(12)
    run_desc = p_meta.add_run("Project ID: HTA-01 | Title: Hiring Trend Analysis | Domain: Data Analytics")
    run_desc.font.size = Pt(10.5)
    run_desc.font.italic = True
    run_desc.font.color.rgb = RGBColor(71, 85, 105)

    # Divider line
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(14)
    r_div = p_div.add_run("―" * 58)
    r_div.font.color.rgb = RGBColor(203, 213, 225)

    # 1. Project Profile & Academic Details
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)
    r1 = h1.add_run("1. Academic Project Profile & Document Control")
    r1.font.size = Pt(13.5)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(30, 58, 138)

    t_profile = doc.add_table(rows=7, cols=2)
    t_profile.alignment = WD_TABLE_ALIGNMENT.CENTER
    profile_data = [
        ("Project Identifier & Title", "HTA-01: Hiring Trend Analysis"),
        ("Specialized Domain", "Data Analytics (Human Resources & Talent Acquisition Analytics)"),
        ("Student Name", "Abhayjeet Singh"),
        ("University / College Roll No.", "2500320140004"),
        ("Program & Academic Cohort", "Master of Computer Applications (MCA 2nd Year, Data Analytics Track)"),
        ("Faculty Project Mentor", "Dinesh Sereno, ABES Engineering College"),
        ("Technology & Tool Stack", "Python 3.11+, Streamlit, SQLite 3, Plotly Express, Tailwind-styled UI, Pytest")
    ]
    for i, (k, v) in enumerate(profile_data):
        row = t_profile.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.3)
        c1.width = Inches(4.2)
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

    # 2. Problem Statement
    h2 = doc.add_heading(level=1)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(6)
    r2 = h2.add_run("2. Problem Statement (One Sentence)")
    r2.font.size = Pt(13.5)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(30, 58, 138)

    p_ps = doc.add_paragraph()
    p_ps.paragraph_format.space_after = Pt(8)
    r_ps_bold = p_ps.add_run("Problem Statement: ")
    r_ps_bold.bold = True
    r_ps_bold.font.color.rgb = RGBColor(30, 58, 138)
    r_ps = p_ps.add_run(
        "Talent acquisition teams and HR leadership lack centralized, quantitative visibility into recruitment velocity, departmental hiring demands, candidate drop-offs across funnel stages, and compensation competitiveness, leading to extended time-to-hire, high agency expenditures, and suboptimal talent planning."
    )
    r_ps.font.size = Pt(11)
    r_ps.font.bold = True
    r_ps.font.color.rgb = RGBColor(15, 23, 42)

    p_context = doc.add_paragraph(
        "Context & Industry Challenge: Modern corporate talent pipelines generate vast volumes of transactional data across job postings, candidate applications, multi-round technical assessments, interview rounds, and salary offers. In most enterprises, this data resides across fragmented spreadsheets or siloed ATS exports. The absence of a unified analytics engine causes data latency, obscures pipeline bottlenecks, leads to offer reneges, and impedes data-driven headcount forecasting."
    )
    p_context.paragraph_format.space_after = Pt(8)

    # 3. Intended Audience & Users
    h3 = doc.add_heading(level=1)
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(6)
    r3 = h3.add_run("3. Intended Audience & System Users")
    r3.font.size = Pt(13.5)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(30, 58, 138)

    t_users = doc.add_table(rows=5, cols=4)
    t_users.alignment = WD_TABLE_ALIGNMENT.CENTER
    user_headers = ["User Role / Persona", "Primary Objectives & Use Case", "Access Frequency", "Key Operational Actions Taken in App"]
    for j, text in enumerate(user_headers):
        cell = t_users.rows[0].cells[j]
        cell.text = text
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 80, 80, 90, 90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9.5)

    user_rows = [
        ("Talent Acquisition Lead / Recruiting Manager", "Monitor daily applicant velocity, recruiter workload distribution, and funnel stage conversion rates.", "Daily (Morning standups & operational pipeline syncs)", "Rebalance candidate allocations across recruiters, spot stalled requisitions, escalate stage delays."),
        ("Department Hiring Managers (Tech, Sales, Product)", "Track department-specific open positions, interview pass rates, and candidate quality.", "Weekly (Prior to team hiring reviews)", "Filter open roles by tech stack/level, review interview-to-offer conversions, evaluate candidate salary requests."),
        ("VP of HR / Head of People", "Evaluate quarterly recruitment velocity, sourcing channel ROI, diversity ratios, and offer acceptance rates.", "Bi-weekly / Monthly (Executive business reviews)", "Calibrate recruitment agency budgets, adjust sourcing channel investments, revise organizational salary bands."),
        ("Finance Director / Head of FP&A", "Audit total recruitment expenditure against approved hiring budget and monitor cost-per-hire trends.", "Monthly / Quarterly (Budget reconciliation)", "Validate accepted offers against budgeted compensation ceilings, monitor external agency commission spend.")
    ]
    for i, row in enumerate(user_rows):
        cells = t_users.rows[i+1].cells
        for j, val in enumerate(row):
            cells[j].text = val
            set_cell_background(cells[j], "F8FAFC" if i % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[j], 60, 60, 90, 90)
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

    # 4. Stakeholder Matrix
    h4 = doc.add_heading(level=1)
    h4.paragraph_format.space_before = Pt(14)
    h4.paragraph_format.space_after = Pt(6)
    r4 = h4.add_run("4. Stakeholder Matrix (SH-01 to SH-05)")
    r4.font.size = Pt(13.5)
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
        ("SH-01", "Head of Talent Acquisition", "Business Process Owner & Primary Sponsor", "High", "High", "Decides recruitment channel spend allocations, recruiter capacity re-distribution, and SLA performance targets."),
        ("SH-02", "Senior Technical Recruiter", "Operational Power User", "High", "Low", "Prioritizes daily candidate screening queues, flags stage bottleneck delays, and initiates interview rounds."),
        ("SH-03", "Department Hiring Manager", "Departmental Hiring Consumer", "Medium", "Medium", "Calibrates interview assessment bar, adjusts role skill prerequisites, and approves final candidate selections."),
        ("SH-04", "Chief Financial Officer (CFO)", "Executive Budget Sponsor", "Medium", "High", "Approves annual hiring budget limits, validates cost-per-hire thresholds, and sanctions department expansion caps."),
        ("SH-05", "Dinesh Sereno", "ABES Faculty Project Mentor & Guide", "High", "High", "Evaluates analytical rigor, database architecture, statistical validity, Streamlit UI execution, and viva defense.")
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

    # 5. Operational Questions & Data Feasibility
    h5 = doc.add_heading(level=1)
    h5.paragraph_format.space_before = Pt(14)
    h5.paragraph_format.space_after = Pt(6)
    r5 = h5.add_run("5. Operational Questions (OQ-01 to OQ-10) & Data Feasibility")
    r5.font.size = Pt(13.5)
    r5.font.bold = True
    r5.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "To ensure business utility and analytical rigor, 10 operational questions have been formulated. Every question has been validated against our relational schema fields to guarantee that every single OQ can be answered directly from SQLite data with 100% mathematical integrity."
    )

    t_oq = doc.add_table(rows=11, cols=4)
    t_oq.alignment = WD_TABLE_ALIGNMENT.CENTER
    oq_headers = ["OQ ID", "Operational Question", "Target Schema Fields / Data Points", "Feasibility & Calculation Method"]
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
        ("OQ-01", "What is the average and median Time-to-Hire (in calendar days) across departments and experience levels?", "applications.applied_date, applications.hire_date, jobs.department, jobs.experience_level", "FEASIBLE: Computed via Julian day delta: CAST(julianday(hire_date) - julianday(applied_date) AS INTEGER)."),
        ("OQ-02", "Which candidate sourcing channels (LinkedIn, Referral, Career Portal, Agency) yield the highest Offer Acceptance Rate (OAR)?", "candidates.source_channel, offers.status ('Accepted', 'Declined', 'Pending')", "FEASIBLE: Computed as: COUNT(offers.status='Accepted') / COUNT(offers.status IN ('Accepted', 'Declined')) * 100."),
        ("OQ-03", "At which recruitment funnel stage (Screening -> Technical -> Managerial -> HR -> Offer) do applicants experience the steepest drop-off?", "application_stages.stage_name, application_stages.status, application_stages.stage_order", "FEASIBLE: Sequential stage funnel conversion waterfall comparing candidate counts passing each sequential stage."),
        ("OQ-04", "How do accepted compensation offers deviate from pre-approved salary bands (min/max) across job families?", "offers.offered_salary, jobs.min_salary, jobs.max_salary, jobs.job_title", "FEASIBLE: Band variance: (offered_salary - (min_salary + max_salary) / 2.0) and percentage deviation above/below midpoint."),
        ("OQ-05", "What are the seasonal monthly and quarterly trends in application volume and requisition openings over the 24-month horizon?", "applications.applied_date, jobs.posted_date", "FEASIBLE: Aggregated monthly timeseries using SQLite strftime('%Y-%m', applied_date) grouped by department."),
        ("OQ-06", "What is the demographic and geographical distribution (Gender, Tier 1/2/3 cities) of applicants versus hired candidates?", "candidates.gender, candidates.location_tier, applications.status", "FEASIBLE: Comparative percentage distribution between initial candidate pipeline and final hired cohort."),
        ("OQ-07", "What are the primary reasons cited by candidates when declining job offers, and what is the overall renege rate?", "offers.status, offers.decline_reason, offers.offered_date", "FEASIBLE: Frequency count and Pareto ranking of decline_reason categories for all declined offers."),
        ("OQ-08", "How does individual recruiter throughput compare in terms of total candidates screened, interviews scheduled, and requisitions closed?", "applications.recruiter_id, recruiters.name, applications.status, jobs.status", "FEASIBLE: Recruiter performance scorecard aggregating processed candidates, average cycle time, and successfully closed jobs."),
        ("OQ-09", "What is the interview-to-offer conversion efficiency ratio across different university tiers and candidate education levels?", "candidates.education_tier, applications.stage_reached, offers.offer_id", "FEASIBLE: Conversion ratio: COUNT(offers.offer_id) / COUNT(candidates reaching interview stage) by education tier."),
        ("OQ-10", "What is the average cost-per-hire across sourcing channels and departments?", "applications.sourcing_cost, applications.agency_fee, applications.status='Hired'", "FEASIBLE: Total acquisition cost SUM(sourcing_cost + agency_fee) divided by total successful hires.")
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

    # 6. Objectives (SMART)
    h6 = doc.add_heading(level=1)
    h6.paragraph_format.space_before = Pt(14)
    h6.paragraph_format.space_after = Pt(6)
    r6 = h6.add_run("6. Project Objectives (SMART Framework)")
    r6.font.size = Pt(13.5)
    r6.font.bold = True
    r6.font.color.rgb = RGBColor(30, 58, 138)

    objectives = [
        ("OBJ-01", "Core Recruitment KPI Formulation", "Specific & Measurable: Automate calculation of 10 essential recruitment analytics metrics (Time-to-Hire, Funnel Conversion, Offer Acceptance Rate, Cost-per-Hire) derived from structured transactional tables with 100% mathematical accuracy."),
        ("OBJ-02", "Relational Database & Data Engineering", "Achievable: Architect a normalized 3NF SQLite database populated with 5,000 synthetic candidate records spanning 100+ job requisitions across 24 historical months, executing analytical queries under 200 milliseconds."),
        ("OBJ-03", "Interactive Analytical Visualizations", "Relevant: Develop an intuitive Streamlit dashboard featuring 8+ interactive Plotly charts (funnel, boxplots, timeseries, heatmaps) with multi-dimensional filtering by department, seniority, recruiter, and date range."),
        ("OBJ-04", "Automated Quality Assurance Suite", "Time-bound: Implement a rigorous test suite using pytest with 15+ test cases verifying schema referential integrity, KPI calculation accuracy, boundary conditions, and filter logic, achieving >85% code coverage."),
        ("OBJ-05", "Cloud Deployment & Academic Deliverables", "Realistic & Complete: Deploy the production application on Streamlit Community Cloud with public accessibility and generate the complete ABES Word project report and viva presentation deck for final institutional evaluation.")
    ]

    t_obj = doc.add_table(rows=6, cols=3)
    t_obj.alignment = WD_TABLE_ALIGNMENT.CENTER
    obj_headers = ["Objective ID", "SMART Objective Title", "Specific, Measurable, Achievable, Relevant, Time-bound Description"]
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

    # 7. Project Scope: In-Scope vs Out-of-Scope
    h7 = doc.add_heading(level=1)
    h7.paragraph_format.space_before = Pt(14)
    h7.paragraph_format.space_after = Pt(6)
    r7 = h7.add_run("7. Project Scope Specification")
    r7.font.size = Pt(13.5)
    r7.font.bold = True
    r7.font.color.rgb = RGBColor(30, 58, 138)

    t_scope = doc.add_table(rows=6, cols=3)
    t_scope.alignment = WD_TABLE_ALIGNMENT.CENTER
    scope_headers = ["Functional Dimension", "In-Scope (Deliverable in Current Capstone Project)", "Out-of-Scope (Deferred to Future / Enterprise Phase)"]
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
        ("Data Management & Storage", "Synthetic generation of 5,000 realistic records with relational referential integrity in SQLite 3.", "Live ATS API integration (e.g. Workday, Greenhouse webhooks) and real candidate personal identifiable information (PII)."),
        ("Analytics & Metrics", "10 core recruitment KPIs, descriptive statistical distributions (median, IQR, standard deviations), and funnel attrition.", "Deep learning resume parsing, NLP text summarization of interview transcripts, or automated candidate sentiment analysis."),
        ("User Interface & Dashboard", "Multi-tab Streamlit dashboard with Tailwind CSS styling, responsive layout, and Plotly interactive charts.", "Native mobile application development (iOS/Android) and multi-tenant Role-Based Access Control (RBAC/OAuth)."),
        ("Quality Assurance", "Automated pytest suite testing database schemas, KPI queries, date delta calculations, and data boundary cases.", "Automated browser-based end-to-end Selenium/Cypress UI automation suites."),
        ("Deployment & Documentation", "Streamlit Community Cloud deployment, Git version control, ABES Word technical report, and viva slide deck.", "Multi-region Kubernetes clustering, enterprise high-availability load balancing, or paid cloud infrastructure.")
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

    # 8. Assumptions & Constraints
    h8 = doc.add_heading(level=1)
    h8.paragraph_format.space_before = Pt(14)
    h8.paragraph_format.space_after = Pt(6)
    r8 = h8.add_run("8. Assumptions and Constraints")
    r8.font.size = Pt(13.5)
    r8.font.bold = True
    r8.font.color.rgb = RGBColor(30, 58, 138)

    t_ac = doc.add_table(rows=7, cols=4)
    t_ac.alignment = WD_TABLE_ALIGNMENT.CENTER
    ac_headers = ["Item ID", "Classification", "Description", "Operational Impact & Project Mitigation"]
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
        ("AS-01", "Assumption: Data Realism", "Synthetic data generated via Faker/NumPy accurately simulates real-world recruitment distributions, stage drop-offs, and salary variances.", "Guarantees statistically meaningful analysis and visualization patterns without requiring sensitive corporate proprietary data."),
        ("AS-02", "Assumption: Infrastructure", "Streamlit Community Cloud 1GB RAM / 1 vCPU container is sufficient to host the SQLite database and perform in-memory caching.", "SQLite database design utilizes proper indices; queries aggregate data efficiently to maintain memory usage well below 150MB."),
        ("AS-03", "Assumption: Academic Standards", "The deliverables strictly adhere to the academic curriculum and guidelines set by the Department of MCA, ABES Engineering College.", "All reports, diagrams, matrices, and test suites are formatted in accordance with institutional project evaluation criteria."),
        ("CO-01", "Constraint: Data Privacy", "Strict prohibition against ingesting real personal identifiable information (PII) under corporate compliance and academic ethics (Rule 5).", "All candidate records, names, contact numbers, and emails are generated synthetically using randomized seed scripts."),
        ("CO-02", "Constraint: Technology Stack", "System implementation is restricted to open-source Python libraries (Streamlit, SQLite, Plotly, pytest) with zero proprietary licensing costs.", "Ensures 100% reproducibility and portability across any standard developer or examiner machine."),
        ("CO-03", "Constraint: Timeline & Scope", "Execution strictly follows the 18-step professional methodology with no code developed prior to Step 10.", "Ensures disciplined requirements engineering, architectural planning, and end-to-end traceability before implementation.")
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

    # 9. Measurable Success Criteria
    h9 = doc.add_heading(level=1)
    h9.paragraph_format.space_before = Pt(14)
    h9.paragraph_format.space_after = Pt(6)
    r9 = h9.add_run("9. Measurable Success Criteria")
    r9.font.size = Pt(13.5)
    r9.font.bold = True
    r9.font.color.rgb = RGBColor(30, 58, 138)

    t_sc = doc.add_table(rows=6, cols=4)
    t_sc.alignment = WD_TABLE_ALIGNMENT.CENTER
    sc_headers = ["Criteria ID", "Evaluation Dimension", "Target Quantitative Metric", "Verification & Assessment Method"]
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
        ("SC-01", "Analytical Accuracy", "100% of operational questions (OQ-01 to OQ-10) return deterministic, mathematically verified results.", "Automated verification scripts cross-checking SQL aggregations against pandas dataframe calculations."),
        ("SC-02", "Application Latency", "Initial dashboard load < 2.0 seconds; multi-filter dynamic slice-and-dice updates < 500 ms.", "Execution profiling on Streamlit Community Cloud hosting 5,000 synthetic records."),
        ("SC-03", "Test Suite Coverage", "Minimum 85% statement coverage with zero failing unit or data validation tests in pytest.", "Automated pytest run with pytest-cov HTML report generation."),
        ("SC-04", "UI/UX Usability", "Clean Tailwind-inspired visual hierarchy with intuitive 3-click insight navigation and zero visual clutter.", "Heuristic evaluation and cognitive walkthrough across all defined user personas."),
        ("SC-05", "Academic Rigor & Compliance", "Complete bidirectional traceability across all 18 project steps matching ABES evaluation standards.", "Formal verification of the Word technical report, Excel RTM, and viva presentation deck.")
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

    os.makedirs("docs", exist_ok=True)
    
    # Save multiple filename aliases so all links and conventions resolve perfectly
    primary_name = "docs/01_Abhayjeet_Hiring_Trend_Analysis_Project_Understanding.docx"
    doc.save(os.path.abspath(primary_name))
    doc.save(os.path.abspath("docs/01_Project_Understanding.docx"))
    doc.save(os.path.abspath("docs/01_Abhayjeet_project_understa.docx"))
    doc.save(os.path.abspath("docs/01_name_project_understa.docx"))
    print("All documents generated successfully.")

create_document()
