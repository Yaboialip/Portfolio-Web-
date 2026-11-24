# -*- coding: utf-8 -*-
"""
Generate comprehensive, professional CV document (.docx) for Alief Sukma Dewanta
containing all profile data, achievements, roles, projects, research, credentials, and links.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

import html

def create_element(name):
    return OxmlElement(name)

def add_hyperlink(paragraph, url, text, color="0078D7", underline=True):
    """Adds a clickable hyperlink to a paragraph in docx."""
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
    new_run = parse_xml(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    rPr = parse_xml(f'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    
    # color
    c = parse_xml(f'<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{color}"/>')
    rPr.append(c)
    
    # underline
    if underline:
        u = parse_xml(f'<w:u xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="single"/>')
        rPr.append(u)
    
    new_run.append(rPr)
    escaped_text = html.escape(str(text))
    new_run_text = parse_xml(f'<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xml:space="preserve">{escaped_text}</w:t>')
    new_run.append(new_run_text)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set cell padding (in twips, 1/20 of a pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_bg(cell, color_hex):
    """Set background color of a cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def add_heading_with_bottom_border(doc, title_text):
    """Creates a stylized executive section header with an accent bottom border."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    
    run = p.add_run(title_text.upper())
    run.font.name = "Arial"
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = RGBColor(16, 44, 87) # Deep Navy (#102C57)

    # XML bottom border on paragraph
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="008899"/></w:pBdr>')
    pPr.append(pBdr)
    return p

def generate_comprehensive_cv_doc(output_path="assets/Alief_Sukma_Dewanta_CV_Comprehensive.docx"):
    doc = docx.Document()

    # Set page margins (Standard 0.75 in / 1.9 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    # Base styling
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(30, 41, 59) # Slate Dark (#1E293B)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(3)

    # =========================================================================
    # HEADER / CONTACT BLOCK
    # =========================================================================
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2)
    run_name = p_name.add_run("ALIEF SUKMA DEWANTA")
    run_name.font.name = "Arial"
    run_name.font.size = Pt(22)
    run_name.font.bold = True
    run_name.font.color.rgb = RGBColor(11, 25, 44) # Deep Charcoal/Navy

    p_tagline = doc.add_paragraph()
    p_tagline.paragraph_format.space_before = Pt(0)
    p_tagline.paragraph_format.space_after = Pt(4)
    run_tag = p_tagline.add_run("IT Project Manager  |  IT Business Analyst  |  Systems Analyst  |  System Implementor")
    run_tag.font.name = "Arial"
    run_tag.font.size = Pt(11)
    run_tag.font.bold = True
    run_tag.font.color.rgb = RGBColor(0, 136, 153) # Teal / Cyan

    # Contact Info line
    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(6)
    
    p_contact.add_run("Surabaya, East Java, Indonesia (Open to Relocation)  •  ")
    p_contact.add_run("Email: ")
    add_hyperlink(p_contact, "mailto:aliefsd84@gmail.com", "aliefsd84@gmail.com", color="0078D7")
    p_contact.add_run("  •  LinkedIn: ")
    add_hyperlink(p_contact, "https://linkedin.com/in/aliefsd84", "linkedin.com/in/aliefsd84", color="0078D7")
    p_contact.add_run("  •  Portfolio Web: ")
    add_hyperlink(p_contact, "https://yaboialip.github.io/My-portfolio/", "yaboialip.github.io/My-portfolio", color="0078D7")
    p_contact.add_run("  •  GitHub: ")
    add_hyperlink(p_contact, "https://github.com/Yaboialip", "github.com/Yaboialip", color="0078D7")

    # Divider bar
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(0)
    p_div.paragraph_format.space_after = Pt(8)
    pPr = p_div._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="18" w:space="1" w:color="102C57"/></w:pBdr>')
    pPr.append(pBdr)

    # =========================================================================
    # 1. PROFESSIONAL EXECUTIVE SUMMARY
    # =========================================================================
    add_heading_with_bottom_border(doc, "Executive Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_after = Pt(4)
    p_sum.add_run(
        "Management Information Systems graduate (GPA 3.66/4.00, Telkom University, Convocation: Oct 26, 2026) "
        "complemented by international business coursework at Hanyang University, South Korea (GPA 4.50/4.50 • A+, A+). "
        "Bridges business stakeholders with engineering teams across the entire software development lifecycle (SDLC). "
        "Delivered 3 end-to-end web system applications, 2 ERP implementations (Odoo ERP), 1 enterprise big data lakehouse pipeline, "
        "and 1 published peer-reviewed enterprise architecture blueprint (TOGAF ADM 9.2). "
        "Formally certified in Google Project Management (Coursera), Google AI Essentials, Google Cloud Gen AI Academy APAC, "
        "IBM Systems Analyst, SAP SCM500 (Procurement), SAP HR050 (Human Capital Management), and EF SET C2 Advanced English (77/100). "
        "Proven leadership as Chairperson of a 40+ member Student Legislative Council (DPM) and Enterprise Architecture Laboratory Assistant."
    )

    # Key Highlights Table
    table_hl = doc.add_table(rows=1, cols=4)
    table_hl.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hl.autofit = False
    
    col_widths = [Inches(1.75), Inches(1.75), Inches(1.75), Inches(1.75)]
    metrics = [
        ("6+ Delivered Systems", "BRD, FSD, UAT & ERP"),
        ("GPA 3.66 / 4.00", "High Honors / Cum Laude"),
        ("6 Professional Certs", "Google PM, IBM, SAP, GCP"),
        ("C2 Advanced English", "EF SET 77/100 • EPrT 557")
    ]
    hdr_cells = table_hl.rows[0].cells
    for i, (m_top, m_sub) in enumerate(metrics):
        hdr_cells[i].width = col_widths[i]
        set_cell_bg(hdr_cells[i], "F1F5F9")
        set_cell_margins(hdr_cells[i], top=80, bottom=80, left=100, right=100)
        hp = hdr_cells[i].paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.paragraph_format.space_before = Pt(2)
        hp.paragraph_format.space_after = Pt(1)
        r1 = hp.add_run(m_top + "\n")
        r1.font.bold = True
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = RGBColor(16, 44, 87)
        r2 = hp.add_run(m_sub)
        r2.font.size = Pt(8)
        r2.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # =========================================================================
    # 2. EDUCATION & INTERNATIONAL STUDY
    # =========================================================================
    add_heading_with_bottom_border(doc, "Education & Academic Pedigree")

    # Telkom University
    p_edu1 = doc.add_paragraph()
    p_edu1.paragraph_format.space_before = Pt(2)
    p_edu1.paragraph_format.space_after = Pt(1)
    p_edu1.paragraph_format.keep_with_next = True
    r = p_edu1.add_run("Telkom University Surabaya")
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(15, 23, 42)
    p_edu1.add_run(" — Surabaya, East Java, Indonesia\n")
    r_deg = p_edu1.add_run("Bachelor of Science (B.S.) in Management Information Systems (Sistem Informasi)")
    r_deg.font.bold = True
    r_deg.font.color.rgb = RGBColor(0, 136, 153)
    p_edu1.add_run("  |  2022 – 2026 (Graduated Oct 2026 • Convocation: Oct 26, 2026)")

    bullets_edu1 = [
        ("Cumulative GPA: ", "3.66 / 4.00 (High Honors / Cum Laude). Completed 144 credit units."),
        ("Undergraduate Thesis (Grade A): ", "Stock Price Prediction for PT Pertamina Geothermal Energy (PGEO) Using ARIMA-GARCH, Implemented as a Web-Based System."),
        ("Thesis Document Link: ", "Available on Google Drive: "),
        ("Relevant Coursework: ", "Enterprise Architecture, IT/IS Governance, IT/IS Strategic Management, Information Systems Security, Data Warehouse & Business Intelligence, Big Data Management, Business Process Management, System Analysis & Design, System Testing & Implementation, Project Management for Information Systems."),
        ("Academic Governance: ", "Chairperson, Student Legislative Council (DPM Telkom University Surabaya, 2025–2026); Two-Semester Enterprise Architecture Laboratory Assistant (2024–2025).")
    ]
    for b_title, b_desc in bullets_edu1:
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.space_after = Pt(1.5)
        br = bp.add_run(b_title)
        br.font.bold = True
        bp.add_run(b_desc)
        if b_title == "Thesis Document Link: ":
            add_hyperlink(bp, "https://drive.google.com/file/d/1sFt-QrZIYCprjvEfZtW41hoVChWjGJXn/view?usp=drive_link", "Skripsi PGEO ARIMA-GARCH Manuscript [PDF]", color="0078D7")

    # Hanyang University
    p_edu2 = doc.add_paragraph()
    p_edu2.paragraph_format.space_before = Pt(4)
    p_edu2.paragraph_format.space_after = Pt(1)
    p_edu2.paragraph_format.keep_with_next = True
    r = p_edu2.add_run("Hanyang University")
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(15, 23, 42)
    p_edu2.add_run(" — Seoul, South Korea\n")
    r_deg = p_edu2.add_run("Winter International School / Pre-Exchange Program")
    r_deg.font.bold = True
    r_deg.font.color.rgb = RGBColor(0, 136, 153)
    p_edu2.add_run("  |  Dec 2025 – Feb 2026")

    bullets_edu2 = [
        ("Academic Standing: ", "Perfect 4.50 / 4.50 GPA (Grade A+, A+)."),
        ("Domain Curriculum: ", "Global Business Management, Strategic Decision Making, International Marketing and Support Services."),
        ("Multi-cultural Collaboration: ", "Led cross-cultural team presentations and analyzed multinational business case scenarios at a top-tier South Korean research institution.")
    ]
    for b_title, b_desc in bullets_edu2:
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.space_after = Pt(1.5)
        br = bp.add_run(b_title)
        br.font.bold = True
        bp.add_run(b_desc)

    # =========================================================================
    # 3. VERIFIED PROFESSIONAL CERTIFICATIONS
    # =========================================================================
    add_heading_with_bottom_border(doc, "Verified Professional Certifications")

    certs = [
        (
            "Google Project Management Professional Certificate",
            "Google Career Certificates (Coursera) • Issued 2026",
            "Validates competency across traditional Waterfall SDLC, Agile Scrum sprint ceremonies, Work Breakdown Structure (WBS), project charters, risk registers, and stakeholder governance.",
            "https://drive.google.com/file/d/1px_zULs9ktCtz4HS9y-mV5yoizBd3pu0/view"
        ),
        (
            "IBM Systems Analyst Professional Certificate",
            "IBM / Coursera • Issued 2025",
            "Covers software analysis & design, functional specification drafting, UML modeling (Use Case, Activity, Sequence), 3NF ERD schema modeling, and system testing.",
            "https://drive.google.com/file/d/1N1OpfzLb41qbAiJKFixpaQJyZxpbPYTs/view?usp=drive_link"
        ),
        (
            "Google AI Essentials",
            "Google Career Certificates • Issued 2026",
            "Covers foundational artificial intelligence paradigms, generative AI prompt engineering, workplace AI tool integration, and responsible AI safety frameworks.",
            "https://drive.google.com/file/d/1w9exZvruwxxjmYWl_AabMteworOG7CKH/view?usp=drive_link"
        ),
        (
            "Google Cloud Gen AI Academy APAC",
            "Google Cloud x Hack2skill • Cohort 1 Graduate • Issued 2026",
            "Trained in cloud generative AI architectures, foundation model tuning, Vertex AI deployment, and enterprise agent design via Model Context Protocol (MCP).",
            "https://drive.google.com/file/d/1NIVYBdtupU_G7StJX4JF7rXkCCHr8F6f/view?usp=drive_link"
        ),
        (
            "SAP Certified: SCM500 — Business Processes in Procurement",
            "SAP University Alliances • Issued 2026",
            "Validates proficiency in enterprise procurement lifecycle, vendor master management, purchase requisitions, inventory optimization, and goods receipt verification.",
            "https://drive.google.com/file/d/1Jld8wiO7VhbPhoA1k-4gh8whQ5Q0zFCP/view?usp=drive_link"
        ),
        (
            "SAP Certified: HR050 — Business Processes in Human Capital Management",
            "SAP University Alliances • Issued 2026",
            "Validates mastery of enterprise HCM workflows, organizational unit structures, employee master records, personnel administration, and corporate HR reporting.",
            "https://drive.google.com/file/d/1rbG-AEVhXtDvmtQMFCuqIaCu8rKrhB2F/view?usp=drive_link"
        ),
        (
            "EF SET English Certificate — C2 Advanced (Score 77/100)",
            "EF Standard English Test • Certified Bilingual Fluency",
            "CEFR C2 Advanced standard; verified English Proficiency Test (EPrT) score of 557 (B2 CEFR). Full professional fluency in international documentation and executive alignment.",
            "https://drive.google.com/file/d/1ZR8rK2Md03JX2uoQHXl3r09yCc6Qydpe/view?usp=drive_link"
        )
    ]

    for c_title, c_issuer, c_desc, c_link in certs:
        p_c = doc.add_paragraph(style='List Bullet')
        p_c.paragraph_format.space_after = Pt(2)
        r_ct = p_c.add_run(c_title)
        r_ct.font.bold = True
        p_c.add_run(f" | {c_issuer}\n")
        p_c.add_run(f"{c_desc}  [")
        add_hyperlink(p_c, c_link, "View Certificate Document", color="0078D7")
        p_c.add_run("]")

    # =========================================================================
    # 4. PROFESSIONAL & LEADERSHIP EXPERIENCE
    # =========================================================================
    add_heading_with_bottom_border(doc, "Professional & Leadership Experience")

    experiences = [
        (
            "Project Coordinator (Internship)",
            "CV. Netafarm Indolestari — Sidoarjo, East Java",
            "Jun 2025 – Aug 2025",
            [
                "Authored the formal Business Requirements Document (BRD) and Functional Specification Document (FSD / FRD), translating stakeholder needs across 4 operational divisions (Management, Sales, Warehouse, Field Ops) into structured functional scope, User Stories, Acceptance Criteria, and UAT scenarios.",
                "Coordinated a 3-person software engineering team over an 8-week Waterfall cycle using Jira, enforcing rigorous scope baseline compliance to deliver an OKR task tracking system for 50+ staff and 2,000+ agricultural SKUs.",
                "Conducted black-box and white-box testing, defect logging and triage, regression testing, and facilitated formal UAT sign-off with company leadership.",
                "Analyzed user survey data and task completion rates using Excel formulas (IF, VLOOKUP, Pivot Tables) to substantiate functional requirements before engineering handover.",
                "Awarded official Company Certificate of Appreciation for on-time and high-quality system delivery."
            ],
            "https://drive.google.com/drive/folders/1Ym-B6uoR0V2FTENJeCdkSvyU6m_mF0q4?usp=sharing",
            "Google Drive: Netafarm BRD & FSD Archive"
        ),
        (
            "Requirement Analyst (Academic Project)",
            "CV. Dimas Djaya Teknik — Surabaya, East Java",
            "Sep 2025 – Dec 2025",
            [
                "Authored comprehensive BRD and FSD for a multi-branch industrial cold-chain monitoring system covering 12 commercial refrigeration machines across distributed facilities.",
                "Engineered complete web QA test suite documented in '02-Workbook-Test-Case' (Excel), executing positive/negative test cases, boundary condition validations, and regression tracking.",
                "Designed application architecture (Laravel, PHP, MySQL, Python REST API) and established telemetry alert thresholds for temperature anomaly notifications.",
                "Used VLOOKUP and Pivot Tables in Excel to consolidate temperature sensor logs into monthly comparative reports, and designed the master data template for Excel data export.",
                "Validated features against Acceptance Criteria and obtained formal stakeholder sign-off."
            ],
            "https://drive.google.com/drive/folders/1yJeAjFuiJ__W35Yup2n8pDP-jdXFUZht?usp=sharing",
            "Google Drive: E-Temp BRD, FSD & 02-Workbook-Test-Case (Excel)"
        ),
        (
            "Laboratory Assistant, Enterprise Architecture",
            "Telkom University Surabaya — Department of Information Systems",
            "Sep 2024 – Dec 2025 (Two Consecutive Appointments)",
            [
                "Appointed for two consecutive academic semesters to mentor 50+ undergraduate students in Enterprise Architecture design and IT strategic planning.",
                "Authored practicum course modules and evaluated case studies applying the TOGAF ADM 9.2 framework modeled within Visual Paradigm for organizations with 50+ staff.",
                "Organized and served as Keynote Speaker for the virtual seminar 'DevOps x System Analyst: Modern IT Career Path' with 120+ student and professional attendees.",
                "Supervised student teams across Business, Data, Application, and Technology architectural domains."
            ],
            None, None
        ),
        (
            "Chairperson, Student Legislative Council (DPM)",
            "Telkom University Surabaya — Student Parliamentary Government",
            "Feb 2025 – Jan 2026",
            [
                "Elected to lead parliamentary governance across the university student body, directing 40+ council members and coordinating executive-legislative relations.",
                "Supervised student government budgeting, policy drafting, organizational compliance, and cross-departmental documentation practices.",
                "Established unified information management workflows, increasing transparency and institutional accountability."
            ],
            None, None
        ),
        (
            "Media & Information Commission Member",
            "Student Legislative Council (DPM) • Telkom University Surabaya",
            "Feb 2024 – Jan 2025",
            [
                "Managed public communications, official institutional releases, and cross-commission publication workflows.",
                "Maintained design consistency, publication timeliness, and stakeholder engagement metrics."
            ],
            None, None
        ),
        (
            "Vice Head, UI/UX Division",
            "CODER (Student Tech Organization) • Telkom University Surabaya",
            "Aug 2023 – Aug 2024",
            [
                "Led UI/UX division programs, designed hands-on Figma workshops, and mentored junior members in user-centered design methodologies.",
                "Collaborated on cross-functional software projects, conducting usability audits and prototyping user interfaces."
            ],
            None, None
        ),
        (
            "ERP Implementation Team Member (Course Project)",
            "Nuhsantara Merchandise — Surabaya, East Java",
            "Sep 2022 – Dec 2022",
            [
                "Implemented Odoo ERP across a 3-member team, configuring 3 core modules: Website/E-Commerce, Purchasing, and Inventory Management.",
                "Cleaned and migrated legacy SKU catalogs, vendor master lists, and opening balances; established automated reorder thresholds to prevent stockouts.",
                "Delivered operational SOP documentation and conducted end-user training sessions."
            ],
            None, None
        )
    ]

    for role_title, org, period, duties, drive_link, drive_text in experiences:
        p_r = doc.add_paragraph()
        p_r.paragraph_format.space_before = Pt(4)
        p_r.paragraph_format.space_after = Pt(1)
        p_r.paragraph_format.keep_with_next = True
        r_rt = p_r.add_run(role_title)
        r_rt.font.bold = True
        r_rt.font.size = Pt(10.5)
        r_rt.font.color.rgb = RGBColor(16, 44, 87)
        p_r.add_run(f" | {org}\n")
        r_per = p_r.add_run(period)
        r_per.font.italic = True
        r_per.font.size = Pt(9)
        r_per.font.color.rgb = RGBColor(100, 116, 139)

        for duty in duties:
            dp = doc.add_paragraph(style='List Bullet')
            dp.paragraph_format.space_after = Pt(1.5)
            dp.add_run(duty)

        if drive_link and drive_text:
            lp = doc.add_paragraph()
            lp.paragraph_format.space_before = Pt(0)
            lp.paragraph_format.space_after = Pt(3)
            lp.add_run("  📁 Artifact Archive: ")
            add_hyperlink(lp, drive_link, drive_text, color="0078D7")

    # =========================================================================
    # 5. FEATURED TECHNICAL PROJECTS & CASE STUDIES
    # =========================================================================
    add_heading_with_bottom_border(doc, "Featured Technical Projects & Case Studies")

    projects = [
        (
            "Order-to-Cash (O2C) & Inventory Automation: Business Analysis Case Study",
            "Lead Business Analyst & Systems Analyst  |  Nov 2026",
            "Simulated Client: PT Maju Distribusi Nusantara (B2B FMCG Distributor with 350+ Stores)",
            "Engineered complete Order-to-Cash automation eliminating manual WhatsApp and daily Excel dependencies using Odoo ERP.",
            [
                "BRD & FSD Specifications: Authored comprehensive BRD and FSD documents detailing 46 functional requirements (MoSCoW prioritized), 7 non-functional requirements, and 10 strict business rules.",
                "Process Modeling: Modeled AS-IS and TO-BE BPMN 2.0 workflows with automated managerial approvals for credit-limit breaches, discounts >10%, and orders >IDR 50M.",
                "Database Architecture: Designed 25-table relational schema with DBML documentation, featuring atomic stock reservation (Available = On Hand - Reserved), optimistic concurrency locking, and immutable audit logging.",
                "Agile Governance: Formulated 5-sprint hybrid Agile roadmap (100 story points, 17 user stories with Gherkin acceptance criteria, Jira import CSV, RACI matrix).",
                "Traceability & QA: Built 100% Requirements Traceability Matrix (RTM) closing 6 orphaned requirements, executed 62 system test cases, and designed 9 end-to-end UAT scenarios.",
                "Simulated KPI Impact (5,000 orders): Order processing time cut from 23.6h to 1.4h (94% faster); overselling reduced from 8.6% to 0.8%; invoice turnaround cut from 37.1h to 6.7 minutes."
            ],
            [
                ("Live Interactive KPI Dashboard", "https://yaboialip.github.io/o2c-automation-case-study/"),
                ("GitHub Open-Source Repository", "https://github.com/Yaboialip/o2c-automation-case-study")
            ]
        ),
        (
            "Enterprise E-Commerce Big Data Lakehouse & Analytics",
            "Data Architect & Business Analyst Lead  |  Telkom University Capstone • 2026",
            "Dataset: Brazilian E-Commerce (1.5M+ Rows, 100k Orders, R$20.4M GMV across 9 Relational Tables)",
            "Built an enterprise-scale medallion data architecture and interactive executive BI dashboard.",
            [
                "Architecture: Designed 3-layer medallion lakehouse warehouse on Hadoop HDFS (Raw landing, Staging cleansed, Cleaned analytics star schema).",
                "ETL/ELT Engineering: Engineered PySpark and Spark SQL pipelines joining 6 source tables into 118,107 unified fact records; simulated Kafka streaming ingestion.",
                "Business Intelligence: Deployed Apache Superset dashboard featuring 7 visualizations analyzing revenue velocity, payment method breakdown (73.7% credit card), and delivery turnaround vs customer review scores.",
                "Predictive Modeling: Trained Spark MLlib Random Forest and Linear Regression models to forecast customer satisfaction review scores and delivery delays."
            ],
            [
                ("Big Data Lakehouse Presentation Deck [PPT]", "https://drive.google.com/file/d/1-_KjfIORiWMeEm45z8R80D6LOb4IdRg4/view?usp=drive_link")
            ]
        ),
        (
            "Stock Price Volatility Prediction Web App (ARIMA-GARCH) — PGEO",
            "Systems Analyst & Quantitative Developer  |  Undergraduate Thesis (Grade A) • 2026",
            "Partner: PT Pertamina Geothermal Energy Tbk (IDX: PGEO)",
            "Developed econometric forecasting system modeling stock return volatility.",
            [
                "Econometric Modeling: Formulated hybrid ARIMA-GARCH statistical models to capture daily return volatility and asymmetric shocks in renewable energy equities.",
                "System Implementation: Built end-to-end web forecasting system using Python (pandas, statsmodels, arch, pmdarima) and Streamlit.",
                "Academic Validation: Defended before faculty board with Grade A evaluation."
            ],
            [
                ("Full Undergraduate Thesis / Skripsi Manuscript [PDF]", "https://drive.google.com/file/d/1sFt-QrZIYCprjvEfZtW41hoVChWjGJXn/view?usp=drive_link")
            ]
        ),
        (
            "Enterprise Architecture Blueprint (TOGAF ADM 9.2) — Infoglobal IT Division",
            "Enterprise Architect  |  Academic Project • 2024",
            "Partner: PT Infoglobal Teknologi Semesta (Surabaya)",
            "Developed an enterprise-wide IT modernization roadmap aligned with defense manufacturing business strategy.",
            [
                "Architectural Domains: Modeled full TOGAF architecture across Business Architecture, Information Systems Architecture (Data & Application), and Technology Architecture in Visual Paradigm.",
                "Governance Alignment: Formulated target migration roadmaps and baseline gap matrices to support core corporate applications."
            ],
            None
        ),
        (
            "IT Governance Assessment (COBIT 2019) — PT Infoglobal",
            "IT Governance Analyst  |  4-Member Research Team • 2024",
            "Partner: PT Infoglobal Teknologi Semesta",
            "Evaluated IT governance capability and organizational risk posture using COBIT 2019.",
            [
                "Design Factor Analysis: Evaluated 10 organizational design factors to determine enterprise governance profile.",
                "Priority Objectives: Identified and audited priority governance objectives: EDM03 (Risk Optimization), APO12 (Managed Risk), DSS05 (Managed Security Services), DSS04 (Managed Continuity), and MEA03 (Managed Compliance)."
            ],
            None
        ),
        (
            "Sales Performance Intelligence Dashboard — PT Gelatik",
            "Business Analyst  |  2025",
            "Partner: PT Gelatik Confectionery Manufacturer",
            "Engineered interactive executive Power BI dashboard analyzing regional sales velocity, customer buying frequencies, and product revenue performance.",
            [
                "Data Modeling: Built star-schema data models and crafted DAX measures for month-over-month sales velocity.",
                "Business Insights: Identified top 20% revenue-generating SKU tiers and regional distribution bottlenecks."
            ],
            None
        ),
        (
            "Enterprise AI Chatbot with Model Context Protocol (MCP)",
            "AI Solutions Architect  |  Google Cloud Gen AI Academy APAC • 2026",
            "Google Cloud Platform & Hack2skill Capstone",
            "Engineered an enterprise Gen AI chatbot utilizing Model Context Protocol (MCP) for dynamic context retrieval, integrating generative AI foundation models with enterprise cloud APIs.",
            None,
            None
        )
    ]

    for p_title, p_role, p_partner, p_summary, p_bullets, p_links in projects:
        p_proj = doc.add_paragraph()
        p_proj.paragraph_format.space_before = Pt(4)
        p_proj.paragraph_format.space_after = Pt(1)
        p_proj.paragraph_format.keep_with_next = True
        r_pt = p_proj.add_run(p_title)
        r_pt.font.bold = True
        r_pt.font.size = Pt(10.5)
        r_pt.font.color.rgb = RGBColor(16, 44, 87)
        p_proj.add_run(f"\n{p_role}")
        if p_partner:
            p_proj.add_run(f"  •  {p_partner}")

        if p_summary:
            p_ps = doc.add_paragraph()
            p_ps.paragraph_format.space_before = Pt(1)
            p_ps.paragraph_format.space_after = Pt(1.5)
            p_ps.add_run(p_summary).font.italic = True

        if p_bullets:
            for b in p_bullets:
                bp = doc.add_paragraph(style='List Bullet')
                bp.paragraph_format.space_after = Pt(1.5)
                # If bullet has a bold prefix
                if ":" in b and not b.startswith("http"):
                    parts = b.split(":", 1)
                    r_pre = bp.add_run(parts[0] + ":")
                    r_pre.font.bold = True
                    bp.add_run(parts[1])
                else:
                    bp.add_run(b)

        if p_links:
            lp = doc.add_paragraph()
            lp.paragraph_format.space_before = Pt(1)
            lp.paragraph_format.space_after = Pt(3)
            lp.add_run("  🔗 Project Artifacts: ")
            for idx, (l_title, l_url) in enumerate(p_links):
                if idx > 0:
                    lp.add_run("  |  ")
                add_hyperlink(lp, l_url, l_title, color="0078D7")

    # =========================================================================
    # 6. PEER-REVIEWED SCIENTIFIC RESEARCH & PUBLICATIONS
    # =========================================================================
    add_heading_with_bottom_border(doc, "Peer-Reviewed Scientific Publications")

    p_pub = doc.add_paragraph()
    p_pub.paragraph_format.space_before = Pt(2)
    p_pub.paragraph_format.space_after = Pt(2)
    r_pub_t = p_pub.add_run("Designing an IT Blueprint as a Strategy for Developing the Smart Village Concept\n")
    r_pub_t.font.bold = True
    r_pub_t.font.size = Pt(10.5)
    r_pub_t.font.color.rgb = RGBColor(16, 44, 87)
    p_pub.add_run("[Original: Perancangan IT Blueprint Sebagai Strategi Pengembangan Konsep Smart Village]\n")
    r_jrn = p_pub.add_run("Jurnal Masyarakat Mandiri (JMM)  •  Vol. 8, No. 6 (December 2024)  •  SINTA 2 Nationally Accredited Journal")
    r_jrn.font.bold = True
    r_jrn.font.color.rgb = RGBColor(0, 136, 153)

    pub_bullets = [
        "Framework & Scope: Applied the TOGAF ADM 9.2 framework to develop an IT Blueprint and Smart Village roadmap for Panjunan Village, Gresik, aligning village administrative processes with Sustainable Development Goals (SDGs).",
        "Phase Execution: Executed Phase A through Phase F (Architecture Vision, Business Architecture, Information Systems Architecture [Data & Application], Technology Architecture, and Migration Planning).",
        "Validation: Validated target blueprints through community stakeholder surveys with >73% approval and positive field readiness assessment.",
        "Accreditation & Indexing: Formally approved and published in SINTA S2 accredited national journal, indexed in Garuda and Google Scholar."
    ]
    for pb in pub_bullets:
        p_pb = doc.add_paragraph(style='List Bullet')
        p_pb.paragraph_format.space_after = Pt(1.5)
        if ":" in pb:
            parts = pb.split(":", 1)
            p_pb.add_run(parts[0] + ":").font.bold = True
            p_pb.add_run(parts[1])
        else:
            p_pb.add_run(pb)

    p_pub_links = doc.add_paragraph()
    p_pub_links.paragraph_format.space_before = Pt(1)
    p_pub_links.paragraph_format.space_after = Pt(4)
    p_pub_links.add_run("  📄 ResearchGate (Full PDF): ")
    add_hyperlink(p_pub_links, "https://www.researchgate.net/publication/392726934_PERANCANGAN_IT_BLUEPRINT_SEBAGAI_STRATEGI_PENGEMBANGAN_KONSEP_SMART_VILLAGE", "Read Article on ResearchGate", color="0078D7")
    p_pub_links.add_run("  •  DOI: ")
    add_hyperlink(p_pub_links, "https://doi.org/10.31764/jmm.v8i6.29063", "doi.org/10.31764/jmm.v8i6.29063", color="0078D7")

    # =========================================================================
    # 7. TECHNICAL & METHODOLOGICAL SKILL MATRIX
    # =========================================================================
    add_heading_with_bottom_border(doc, "Technical & Methodological Competency Matrix")

    skill_clusters = [
        ("Requirements & Business Analysis", "Requirements Gathering, BRD/FSD/FRD Authoring, MoSCoW Prioritization, User Stories (Gherkin Syntax), Acceptance Criteria, As-Is / To-Be BPMN 2.0 Process Modeling, Stakeholder Management, Fit-Gap Analysis"),
        ("Systems & Architecture Design", "TOGAF ADM 9.2 (Business/Data/App/Tech Domains), Visual Paradigm, UML (Use Case, Activity, Sequence), Relational Database Design (3NF ERD, DBML), API Contract Design, Role-Based Access Control (RBAC)"),
        ("IT Project Management & QA", "Waterfall SDLC, Agile Scrum Sprints, Work Breakdown Structure (WBS), Gantt Charts, Jira Software, Trello, Requirements Traceability Matrix (RTM), Black-box & White-box Testing, UAT Design & Facilitation, Defect Logging"),
        ("IT Governance & Compliance", "COBIT 2019 Design Factors & Governance Objectives, IT Risk & Compliance Assessment, IT Strategic Roadmaps, Sustainable Development Goals (SDGs) Alignment"),
        ("Enterprise ERP & Systems", "Odoo ERP (Inventory, Purchasing, Sales, E-Commerce, Multi-module Implementation), Master Data Cleansing & Migration, Automated Reorder Rules, SOP Documentation"),
        ("Data Analytics & Big Data", "Python (pandas, statsmodels, arch, pmdarima), SQL & MySQL, Apache Spark (PySpark), Spark SQL, Spark MLlib, Hadoop HDFS, Apache Kafka, Apache Superset, Power BI (DAX, Star Schema), ARIMA-GARCH Volatility Forecasting"),
        ("Software Tools & Office Suite", "Microsoft Excel (Advanced: IF, VLOOKUP, INDEX/MATCH, Pivot Tables, Data Modeling), Figma (UI/UX Wireframes & Prototypes), Google Cloud Platform (GCP), GitHub, Markdown, Microsoft Word & PowerPoint"),
        ("Languages & Soft Skills", "Indonesian (Native Fluency), English (EF SET 77/100 C2 Advanced, EPrT 557) • Cross-functional Leadership, Technical-to-Business Translation, Analytical Thinking, Complex Problem Solving, Public Presentation")
    ]

    table_skills = doc.add_table(rows=len(skill_clusters), cols=2)
    table_skills.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_skills.autofit = False

    t_widths = [Inches(2.2), Inches(4.8)]
    for r_idx, (cat, val) in enumerate(skill_clusters):
        row_cells = table_skills.rows[r_idx].cells
        row_cells[0].width = t_widths[0]
        row_cells[1].width = t_widths[1]

        bg_col = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
        set_cell_bg(row_cells[0], bg_col)
        set_cell_bg(row_cells[1], bg_col)
        set_cell_margins(row_cells[0], top=60, bottom=60, left=80, right=80)
        set_cell_margins(row_cells[1], top=60, bottom=60, left=80, right=80)

        p0 = row_cells[0].paragraphs[0]
        p0.paragraph_format.space_before = Pt(1)
        p0.paragraph_format.space_after = Pt(1)
        r0 = p0.add_run(cat)
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = RGBColor(16, 44, 87)

        p1 = row_cells[1].paragraphs[0]
        p1.paragraph_format.space_before = Pt(1)
        p1.paragraph_format.space_after = Pt(1)
        r1 = p1.add_run(val)
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # =========================================================================
    # 8. DIRECT CONTACT & VERIFICATION DIRECTORY
    # =========================================================================
    add_heading_with_bottom_border(doc, "Official Links & Credential Verification Directory")

    p_links_summary = doc.add_paragraph()
    p_links_summary.paragraph_format.space_after = Pt(2)
    p_links_summary.add_run("All documents, certificates, manuscripts, and project repositories referenced in this curriculum vitae can be verified at the direct addresses below:\n")

    dir_links = [
        ("Online Web Portfolio (Cyberpunk HUD):", "https://yaboialip.github.io/My-portfolio/"),
        ("O2C Automation Case Study (Live Dashboard):", "https://yaboialip.github.io/o2c-automation-case-study/"),
        ("O2C Automation Case Study (GitHub Code):", "https://github.com/Yaboialip/o2c-automation-case-study"),
        ("Netafarm OKR System BRD & FSD Archive:", "https://drive.google.com/drive/folders/1Ym-B6uoR0V2FTENJeCdkSvyU6m_mF0q4?usp=sharing"),
        ("E-Temp Monitoring BRD, FSD & Test Case Workbook:", "https://drive.google.com/drive/folders/1yJeAjFuiJ__W35Yup2n8pDP-jdXFUZht?usp=sharing"),
        ("Big Data Lakehouse Presentation Deck (PPT):", "https://drive.google.com/file/d/1-_KjfIORiWMeEm45z8R80D6LOb4IdRg4/view?usp=drive_link"),
        ("Undergraduate Thesis (Skripsi PGEO ARIMA-GARCH):", "https://drive.google.com/file/d/1sFt-QrZIYCprjvEfZtW41hoVChWjGJXn/view?usp=drive_link"),
        ("ResearchGate Smart Village IT Blueprint (PDF):", "https://www.researchgate.net/publication/392726934_PERANCANGAN_IT_BLUEPRINT_SEBAGAI_STRATEGI_PENGEMBANGAN_KONSEP_SMART_VILLAGE"),
        ("Journal Publication Official DOI (SINTA 2):", "https://doi.org/10.31764/jmm.v8i6.29063"),
        ("Google Project Management Certificate:", "https://drive.google.com/file/d/1px_zULs9ktCtz4HS9y-mV5yoizBd3pu0/view"),
        ("IBM Systems Analyst Certificate:", "https://drive.google.com/file/d/1N1OpfzLb41qbAiJKFixpaQJyZxpbPYTs/view?usp=drive_link"),
        ("Google AI Essentials Certificate:", "https://drive.google.com/file/d/1w9exZvruwxxjmYWl_AabMteworOG7CKH/view?usp=drive_link"),
        ("Google Cloud Gen AI Academy APAC Certificate:", "https://drive.google.com/file/d/1NIVYBdtupU_G7StJX4JF7rXkCCHr8F6f/view?usp=drive_link"),
        ("SAP Certified: SCM500 (Procurement) Certificate:", "https://drive.google.com/file/d/1Jld8wiO7VhbPhoA1k-4gh8whQ5Q0zFCP/view?usp=drive_link"),
        ("SAP Certified: HR050 (Human Capital) Certificate:", "https://drive.google.com/file/d/1rbG-AEVhXtDvmtQMFCuqIaCu8rKrhB2F/view?usp=drive_link"),
        ("EF SET English Certificate C2 Advanced:", "https://drive.google.com/file/d/1ZR8rK2Md03JX2uoQHXl3r09yCc6Qydpe/view?usp=drive_link")
    ]

    for label, url in dir_links:
        lp = doc.add_paragraph(style='List Bullet')
        lp.paragraph_format.space_after = Pt(1)
        lp.add_run(label + " ").font.bold = True
        add_hyperlink(lp, url, url, color="0078D7")

    # Save document
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    print(f"[SUCCESS] Comprehensive CV Word Document generated: {output_path}")

if __name__ == "__main__":
    generate_comprehensive_cv_doc()
