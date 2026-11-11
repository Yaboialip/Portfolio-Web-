import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ==============================================================================
# ALIEF SUKMA DEWANTA — EXECUTIVE PORTFOLIO PRESENTATION DECK GENERATOR
# Theme: Night City Tactical Minimalist (Cyberpunk 2077 HUD Aesthetics)
# High Contrast Palette: Obsidian Dark (#06080D, #0B0F17) + Neon Yellow (#FCEE0A) + Electric Cyan (#00F0FF)
# ==============================================================================

def create_portfolio_deck():
    prs = Presentation()
    # 16:9 Widescreen Layout
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette Definitions
    COLOR_BG_DARK = RGBColor(6, 8, 13)          # #06080D
    COLOR_BG_SURFACE = RGBColor(11, 15, 23)     # #0B0F17
    COLOR_BG_CARD = RGBColor(16, 22, 34)        # #101622
    COLOR_BORDER = RGBColor(26, 35, 50)         # #1A2332
    COLOR_BORDER_SUBTLE = RGBColor(36, 48, 68)  # #243044
    COLOR_YELLOW = RGBColor(252, 238, 10)       # #FCEE0A Neon Yellow
    COLOR_CYAN = RGBColor(0, 240, 255)          # #00F0FF Electric Cyan
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_TEXT_PRIMARY = RGBColor(236, 239, 244)
    COLOR_TEXT_SECONDARY = RGBColor(155, 168, 186)
    COLOR_TEXT_MUTED = RGBColor(93, 108, 130)

    FONT_DISPLAY = "Rajdhani"
    FONT_MONO = "JetBrains Mono"
    FONT_BODY = "Segoe UI"

    def add_slide_background(slide, slide_num, total_slides=13, category="OPERATIONAL DOSSIER"):
        # Dark Background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG_DARK
        bg.line.fill.background()

        # Top Diagnostic Rail (Header Bar)
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.35), Inches(12.133), Inches(0.4))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = COLOR_BG_SURFACE
        top_bar.line.color.rgb = COLOR_BORDER
        top_bar.line.width = Pt(1)

        tf = top_bar.text_frame
        tf.word_wrap = True
        tf.margin_top = Pt(3)
        tf.margin_bottom = Pt(3)
        tf.margin_left = Pt(10)
        tf.margin_right = Pt(10)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run1 = p.add_run()
        run1.text = "ALIEF SUKMA DEWANTA // "
        run1.font.name = FONT_DISPLAY
        run1.font.size = Pt(10)
        run1.font.bold = True
        run1.font.color.rgb = COLOR_YELLOW

        run2 = p.add_run()
        run2.text = f"{category} // TELKOM UNIV MIS GRADUATE (OCT 2026)"
        run2.font.name = FONT_MONO
        run2.font.size = Pt(9)
        run2.font.color.rgb = COLOR_CYAN

        # Right-side Slide Counter
        counter_box = slide.shapes.add_textbox(Inches(10.5), Inches(0.35), Inches(2.2), Inches(0.4))
        ctf = counter_box.text_frame
        ctf.margin_top = Pt(5)
        cp = ctf.paragraphs[0]
        cp.alignment = PP_ALIGN.RIGHT
        crun = cp.add_run()
        crun.text = f"SEC_REF: [{slide_num:02d} / {total_slides:02d}]"
        crun.font.name = FONT_MONO
        crun.font.size = Pt(9)
        crun.font.color.rgb = COLOR_TEXT_MUTED

        # Bottom Tactical Footer Rail
        btm_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(7.05), Inches(12.133), Pt(1))
        btm_line.fill.solid()
        btm_line.fill.fore_color.rgb = COLOR_BORDER
        btm_line.line.fill.background()

        btm_text = slide.shapes.add_textbox(Inches(0.6), Inches(7.08), Inches(12.133), Inches(0.35))
        btf = btm_text.text_frame
        btf.margin_top = Pt(2)
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.LEFT
        brun1 = bp.add_run()
        brun1.text = "TACTICAL PORTFOLIO & SPECIFICATION DOSSIER  •  GOOGLE PM & IBM SYSTEMS ANALYST CERTIFIED"
        brun1.font.name = FONT_MONO
        brun1.font.size = Pt(8)
        brun1.font.color.rgb = COLOR_TEXT_MUTED

        brun2 = bp.add_run()
        brun2.text = "  •  SURABAYA, ID (OPEN TO RELOCATION)  •  aliefsd84@gmail.com"
        brun2.font.name = FONT_MONO
        brun2.font.size = Pt(8)
        brun2.font.color.rgb = COLOR_CYAN

    def add_card(slide, left, top, width, height, border_color=COLOR_BORDER, fill_color=COLOR_BG_CARD):
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        box.fill.solid()
        box.fill.fore_color.rgb = fill_color
        box.line.color.rgb = border_color
        box.line.width = Pt(1.2)
        return box

    # --------------------------------------------------------------------------
    # SLIDE 1: COVER / TITLE (TACTICAL DOSSIER)
    # --------------------------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide1, 1, category="SYSTEM IDENTITY & VALUE MANIFESTO")

    # Left Column: Title, Subtitle, Key Telemetry
    title_box = slide1.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(7.5), Inches(2.2))
    ttf = title_box.text_frame
    ttf.word_wrap = True
    tp1 = ttf.paragraphs[0]
    trun_status = tp1.add_run()
    trun_status.text = "// TACTICAL DOSSIER: IT PROJECT MANAGEMENT & SYSTEMS ARCHITECTURE\n"
    trun_status.font.name = FONT_MONO
    trun_status.font.size = Pt(10)
    trun_status.font.bold = True
    trun_status.font.color.rgb = COLOR_CYAN

    trun_name = tp1.add_run()
    trun_name.text = "ALIEF SUKMA DEWANTA\n"
    trun_name.font.name = FONT_DISPLAY
    trun_name.font.size = Pt(36)
    trun_name.font.bold = True
    trun_name.font.color.rgb = COLOR_WHITE

    trun_role = tp1.add_run()
    trun_role.text = "IT Project Manager  •  IT Business Analyst  •  Systems Analyst"
    trun_role.font.name = FONT_DISPLAY
    trun_role.font.size = Pt(18)
    trun_role.font.bold = True
    trun_role.font.color.rgb = COLOR_YELLOW

    # Headline Summary Box
    summary_card = add_card(slide1, Inches(0.8), Inches(3.5), Inches(7.5), Inches(1.8), border_color=COLOR_BORDER_SUBTLE, fill_color=COLOR_BG_SURFACE)
    stf = summary_card.text_frame
    stf.word_wrap = True
    stf.margin_left = Pt(16)
    stf.margin_right = Pt(16)
    stf.margin_top = Pt(12)
    sp1 = stf.paragraphs[0]
    sp1_run1 = sp1.add_run()
    sp1_run1.text = "EXECUTIVE VALUE PROPOSITION\n"
    sp1_run1.font.name = FONT_MONO
    sp1_run1.font.size = Pt(10)
    sp1_run1.font.bold = True
    sp1_run1.font.color.rgb = COLOR_YELLOW

    sp1_run2 = sp1.add_run()
    sp1_run2.text = (
        "Fresh graduate in Management Information Systems (GPA 3.66/4.00, Telkom University, Convocation: Oct 26, 2026) "
        "bridging business stakeholders with technical execution. Delivered three system projects, one ERP deployment, and "
        "one enterprise big data pipeline using Agile and Waterfall methodologies. Formally certified in Google Project Management "
        "and IBM Systems Analysis, with proven mastery in BRD/FSD specifications, TOGAF EA, and UAT governance."
    )
    sp1_run2.font.name = FONT_BODY
    sp1_run2.font.size = Pt(10.5)
    sp1_run2.font.color.rgb = COLOR_TEXT_PRIMARY

    # Stats HUD Grid (4 Columns)
    stats_data = [
        ("6+", "Delivered Projects", COLOR_CYAN),
        ("3+", "Leadership Roles", COLOR_YELLOW),
        ("5", "Industry Certifications", COLOR_CYAN),
        ("3.66", "MIS GPA (High Honors)", COLOR_YELLOW)
    ]
    stat_w = Inches(1.75)
    for idx, (val, lbl, col) in enumerate(stats_data):
        sc = add_card(slide1, Inches(0.8) + idx * Inches(1.92), Inches(5.5), stat_w, Inches(1.2), border_color=COLOR_BORDER)
        sctf = sc.text_frame
        sctf.margin_top = Pt(10)
        sctf.margin_left = Pt(10)
        scp = sctf.paragraphs[0]
        scp.alignment = PP_ALIGN.CENTER
        vrun = scp.add_run()
        vrun.text = val + "\n"
        vrun.font.name = FONT_DISPLAY
        vrun.font.size = Pt(24)
        vrun.font.bold = True
        vrun.font.color.rgb = col
        lrun = scp.add_run()
        lrun.text = lbl
        lrun.font.name = FONT_MONO
        lrun.font.size = Pt(8.5)
        lrun.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right Column: Portrait Card
    portrait_path = "assets/images/alief-hero-portrait.png"
    if os.path.exists(portrait_path):
        p_card = add_card(slide1, Inches(8.7), Inches(1.2), Inches(3.9), Inches(5.5), border_color=COLOR_CYAN, fill_color=COLOR_BG_CARD)
        try:
            slide1.shapes.add_picture(portrait_path, Inches(8.9), Inches(1.4), width=Inches(3.5), height=Inches(4.6))
        except Exception as e:
            print(f"Notice: Portrait image failed to embed: {e}")
        
        # Caption below photo
        cap_box = slide1.shapes.add_textbox(Inches(8.9), Inches(6.05), Inches(3.5), Inches(0.55))
        ctf = cap_box.text_frame
        ctf.margin_top = Pt(2)
        cp = ctf.paragraphs[0]
        cp.alignment = PP_ALIGN.CENTER
        crun1 = cp.add_run()
        crun1.text = "ALIEF SUKMA DEWANTA  "
        crun1.font.name = FONT_MONO
        crun1.font.size = Pt(9.5)
        crun1.font.bold = True
        crun1.font.color.rgb = COLOR_WHITE
        crun2 = cp.add_run()
        crun2.text = "[STATUS: VERIFIED]"
        crun2.font.name = FONT_MONO
        crun2.font.size = Pt(9)
        crun2.font.color.rgb = COLOR_YELLOW

    # --------------------------------------------------------------------------
    # SLIDE 2: 4 MULTIDISCIPLINARY PROFESSIONAL ROLES / LENSES
    # --------------------------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide2, 2, category="OPERATIONAL LENSES & CORE COMPETENCIES")

    header_box = slide2.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf = header_box.text_frame
    hp = htf.paragraphs[0]
    hrun1 = hp.add_run()
    hrun1.text = "// MULTIDISCIPLINARY CAPABILITY MATRIX\n"
    hrun1.font.name = FONT_MONO
    hrun1.font.size = Pt(10)
    hrun1.font.color.rgb = COLOR_CYAN
    hrun2 = hp.add_run()
    hrun2.text = "Four Specialized Delivery Lenses Across the Systems Lifecycle"
    hrun2.font.name = FONT_DISPLAY
    hrun2.font.size = Pt(22)
    hrun2.font.bold = True
    hrun2.font.color.rgb = COLOR_WHITE

    roles_info = [
        (
            "IT PROJECT MANAGER",
            "Waterfall & Agile Sprint Governance",
            COLOR_YELLOW,
            [
                "Sprint planning, WBS dependency tracking, backlog grooming in Jira.",
                "Cross-functional developer coordination across 8-week delivery cycles.",
                "Comprehensive QA & UAT facilitation with defect tracking & sign-off.",
                "Risk mitigation, stakeholder status reporting & scope change governance."
            ]
        ),
        (
            "IT BUSINESS ANALYST",
            "Requirements Elicitation & BRD/FSD Authoring",
            COLOR_CYAN,
            [
                "Authored formal Business Requirements (BRD) & Functional Specs (FSD).",
                "Elicited multi-tier departmental requirements across 4 operational divisions.",
                "BPMN 2.0 As-Is / To-Be process mapping, User Stories & Acceptance Criteria.",
                "Data analysis & requirement validation in Excel (IF, VLOOKUP, Pivot Tables)."
            ]
        ),
        (
            "SYSTEMS ANALYST",
            "Enterprise Architecture & Data Modeling",
            COLOR_CYAN,
            [
                "Modeled UML Use Case, Activity, and Sequence diagrams in Visual Paradigm.",
                "Engineered 3NF-normalized relational schemas (MySQL) & time-series tables.",
                "Defined REST API data contracts, ingestion protocols & RBAC matrices.",
                "TOGAF ADM 9.2 Enterprise Architecture & COBIT 2019 governance frameworks."
            ]
        ),
        (
            "SYSTEM IMPLEMENTOR",
            "ERP Deployment, Data Cleansing & SOPs",
            COLOR_YELLOW,
            [
                "Configured multi-module Odoo ERP (Website/E-Commerce, Purchasing, Inventory).",
                "Master data auditing, SKU catalog standardization, and opening stock migration.",
                "Configured automated reordering rules & procurement requisition routing.",
                "Conducted end-user operational training and authored standard SOP manuals."
            ]
        )
    ]

    card_width = Inches(2.78)
    card_gap = Inches(0.20)
    for idx, (title, subtitle, accent_col, bullets) in enumerate(roles_info):
        left_pos = Inches(0.8) + idx * (card_width + card_gap)
        rcard = add_card(slide2, left_pos, Inches(1.9), card_width, Inches(4.8), border_color=accent_col, fill_color=COLOR_BG_CARD)
        rtf = rcard.text_frame
        rtf.word_wrap = True
        rtf.margin_left = Pt(12)
        rtf.margin_right = Pt(12)
        rtf.margin_top = Pt(14)

        p1 = rtf.paragraphs[0]
        r_title = p1.add_run()
        r_title.text = title + "\n"
        r_title.font.name = FONT_DISPLAY
        r_title.font.size = Pt(14)
        r_title.font.bold = True
        r_title.font.color.rgb = accent_col

        r_sub = p1.add_run()
        r_sub.text = subtitle + "\n\n"
        r_sub.font.name = FONT_MONO
        r_sub.font.size = Pt(8.5)
        r_sub.font.color.rgb = COLOR_TEXT_MUTED

        for b in bullets:
            bp = rtf.add_paragraph()
            bp.space_before = Pt(6)
            brun_icon = bp.add_run()
            brun_icon.text = "▸ "
            brun_icon.font.name = FONT_MONO
            brun_icon.font.size = Pt(9)
            brun_icon.font.color.rgb = accent_col
            brun_text = bp.add_run()
            brun_text.text = b
            brun_text.font.name = FONT_BODY
            brun_text.font.size = Pt(9)
            brun_text.font.color.rgb = COLOR_TEXT_PRIMARY

    # --------------------------------------------------------------------------
    # SLIDE 3: FEATURED PROJECT 01 — NETAFARM OKR SYSTEM (BRD & FSD)
    # --------------------------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide3, 3, category="FLAGSHIP SYSTEM PROJECT 01")

    # Header
    hb3 = slide3.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf3 = hb3.text_frame
    hp3 = htf3.paragraphs[0]
    hp3.add_run().text = "// FLAGSHIP PROJECT 01: CV. NETAFARM INDOLESTARI (SIDOARJO)\n"
    hp3.runs[0].font.name = FONT_MONO
    hp3.runs[0].font.size = Pt(10)
    hp3.runs[0].font.color.rgb = COLOR_CYAN

    hp3_2 = hp3.add_run()
    hp3_2.text = "OKR & Cross-Functional Task Management System"
    hp3_2.font.name = FONT_DISPLAY
    hp3_2.font.size = Pt(22)
    hp3_2.font.bold = True
    hp3_2.font.color.rgb = COLOR_WHITE

    # Left Column: Project Overview & Artifact Highlights
    p1_card = add_card(slide3, Inches(0.8), Inches(1.9), Inches(6.6), Inches(4.8), border_color=COLOR_BORDER, fill_color=COLOR_BG_CARD)
    p1tf = p1_card.text_frame
    p1tf.word_wrap = True
    p1tf.margin_left = Pt(14)
    p1tf.margin_right = Pt(14)
    p1tf.margin_top = Pt(12)

    p1_head = p1tf.paragraphs[0]
    p1_head.add_run().text = "CORE DELIVERABLES & GOVERNANCE ARTIFACTS\n"
    p1_head.runs[0].font.name = FONT_MONO
    p1_head.runs[0].font.size = Pt(10)
    p1_head.runs[0].font.bold = True
    p1_head.runs[0].font.color.rgb = COLOR_YELLOW

    bullets_p1 = [
        ("Business Problem: ", "Siloed communication across 4 departments (Management, Sales, Warehouse, Field Ops) resulting in untracked agricultural SKUs (2,000+) and missing delivery deadlines."),
        ("Formal BRD Deliverable: ", "Authored comprehensive Business Requirements Document aligning executive stakeholders, establishing business drivers, and defining As-Is vs. To-Be OKR workflows."),
        ("Formal FSD / FRD Deliverable: ", "Specified functional scope, User Stories, Acceptance Criteria (Given-When-Then), and granular Role-Based Access Control (RBAC) permission matrices."),
        ("Project Management (Jira): ", "Led 3 developers over an 8-week Waterfall delivery cycle, managing WBS, sprint backlog, black/white-box QA, and UAT sign-off with company leadership."),
        ("Data Validation (Excel): ", "Analyzed user surveys and task completion rates using Excel formulas (IF, VLOOKUP) to validate system requirements prior to engineering handover.")
    ]

    for label, desc in bullets_p1:
        bp = p1tf.add_paragraph()
        bp.space_before = Pt(6)
        r_lbl = bp.add_run()
        r_lbl.text = "• " + label
        r_lbl.font.name = FONT_BODY
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.color.rgb = COLOR_CYAN
        r_desc = bp.add_run()
        r_desc.text = desc
        r_desc.font.name = FONT_BODY
        r_desc.font.size = Pt(9.2)
        r_desc.font.color.rgb = COLOR_TEXT_PRIMARY

    # Google Drive Link Banner
    drive_p1 = p1tf.add_paragraph()
    drive_p1.space_before = Pt(10)
    dr1 = drive_p1.add_run()
    dr1.text = "📂 OFFICIAL GOOGLE DRIVE ARTIFACT ARCHIVE [BRD & FSD]:\n"
    dr1.font.name = FONT_MONO
    dr1.font.bold = True
    dr1.font.size = Pt(9)
    dr1.font.color.rgb = COLOR_YELLOW
    dr2 = drive_p1.add_run()
    dr2.text = "https://drive.google.com/drive/folders/1Ym-B6uoR0V2FTENJeCdkSvyU6m_mF0q4?usp=sharing"
    dr2.font.name = FONT_MONO
    dr2.font.size = Pt(8.5)
    dr2.font.color.rgb = COLOR_WHITE

    # Right Column: UI Screenshot Card
    neta_img = "assets/images/netafarm-admin-dashboard.png"
    if os.path.exists(neta_img):
        add_card(slide3, Inches(7.7), Inches(1.9), Inches(4.8), Inches(4.8), border_color=COLOR_CYAN, fill_color=COLOR_BG_CARD)
        try:
            slide3.shapes.add_picture(neta_img, Inches(7.85), Inches(2.1), width=Inches(4.5), height=Inches(3.3))
        except Exception as e:
            print(f"Notice: Netafarm image embed failed: {e}")
        
        cap3 = slide3.shapes.add_textbox(Inches(7.85), Inches(5.6), Inches(4.5), Inches(1.0))
        c3tf = cap3.text_frame
        c3tf.word_wrap = True
        c3tf.margin_top = Pt(2)
        c3p = c3tf.paragraphs[0]
        c3run = c3p.add_run()
        c3run.text = "TECH STACK & METRICS:\n"
        c3run.font.name = FONT_MONO
        c3run.font.bold = True
        c3run.font.size = Pt(9)
        c3run.font.color.rgb = COLOR_YELLOW
        c3run2 = c3p.add_run()
        c3run2.text = "Laravel 10, MySQL 8, Bootstrap 5, Chart.js  |  CV. Netafarm Indolestari  |  Awarded Official Company Certificate of Appreciation"
        c3run2.font.name = FONT_BODY
        c3run2.font.size = Pt(8.5)
        c3run2.font.color.rgb = COLOR_TEXT_SECONDARY

    # --------------------------------------------------------------------------
    # SLIDE 4: FEATURED PROJECT 02 — E-TEMP TELEMETRY (BRD, FSD & EXCEL QA)
    # --------------------------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide4, 4, category="FLAGSHIP SYSTEM PROJECT 02")

    # Header
    hb4 = slide4.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf4 = hb4.text_frame
    hp4 = htf4.paragraphs[0]
    hp4.add_run().text = "// FLAGSHIP PROJECT 02: CV. DIMAS DJAYA TEKNIK (SURABAYA)\n"
    hp4.runs[0].font.name = FONT_MONO
    hp4.runs[0].font.size = Pt(10)
    hp4.runs[0].font.color.rgb = COLOR_CYAN

    hp4_2 = hp4.add_run()
    hp4_2.text = "E-Temp Industrial Cold-Chain Monitoring System"
    hp4_2.font.name = FONT_DISPLAY
    hp4_2.font.size = Pt(22)
    hp4_2.font.bold = True
    hp4_2.font.color.rgb = COLOR_WHITE

    # Left Column: Project Overview & 3 Artifacts
    p2_card = add_card(slide4, Inches(0.8), Inches(1.9), Inches(6.6), Inches(4.8), border_color=COLOR_BORDER, fill_color=COLOR_BG_CARD)
    p2tf = p2_card.text_frame
    p2tf.word_wrap = True
    p2tf.margin_left = Pt(14)
    p2tf.margin_right = Pt(14)
    p2tf.margin_top = Pt(12)

    p2_head = p2tf.paragraphs[0]
    p2_head.add_run().text = "BRD, FSD, AND 02-WORKBOOK-TEST-CASE (EXCEL QA)\n"
    p2_head.runs[0].font.name = FONT_MONO
    p2_head.runs[0].font.size = Pt(10)
    p2_head.runs[0].font.bold = True
    p2_head.runs[0].font.color.rgb = COLOR_YELLOW

    bullets_p2 = [
        ("Business Problem: ", "Eliminated error-prone manual clipboard temperature checks across 12 industrial cold-storage machines, mitigating high-risk inventory spoilage."),
        ("Formal BRD Deliverable: ", "Authored Business Requirements Document establishing cold-chain compliance rules, multi-branch fleet operational constraints, and SLA targets."),
        ("Formal FSD Deliverable: ", "Specified real-time sensor polling intervals, automated excursion anomaly alert thresholds, time-series data schemas, and Python API + Laravel contracts."),
        ("02-Workbook-Test-Case (Excel QA): ", "Engineered and executed comprehensive web QA test suite in Excel: positive/negative test cases, boundary checks, defect tracking, and UAT sign-off matrix."),
        ("Data Visualization (Excel): ", "Consolidated multi-branch temperature logs via Excel VLOOKUP and Pivot Tables to generate monthly comparative machine performance reports.")
    ]

    for label, desc in bullets_p2:
        bp = p2tf.add_paragraph()
        bp.space_before = Pt(6)
        r_lbl = bp.add_run()
        r_lbl.text = "• " + label
        r_lbl.font.name = FONT_BODY
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.color.rgb = COLOR_CYAN
        r_desc = bp.add_run()
        r_desc.text = desc
        r_desc.font.name = FONT_BODY
        r_desc.font.size = Pt(9.2)
        r_desc.font.color.rgb = COLOR_TEXT_PRIMARY

    # Google Drive Link Banner
    drive_p2 = p2tf.add_paragraph()
    drive_p2.space_before = Pt(10)
    dr1 = drive_p2.add_run()
    dr1.text = "📂 OFFICIAL GOOGLE DRIVE ARTIFACT ARCHIVE [BRD, FSD & QA WORKBOOK]:\n"
    dr1.font.name = FONT_MONO
    dr1.font.bold = True
    dr1.font.size = Pt(9)
    dr1.font.color.rgb = COLOR_YELLOW
    dr2 = drive_p2.add_run()
    dr2.text = "https://drive.google.com/drive/folders/1yJeAjFuiJ__W35Yup2n8pDP-jdXFUZht?usp=sharing"
    dr2.font.name = FONT_MONO
    dr2.font.size = Pt(8.5)
    dr2.font.color.rgb = COLOR_WHITE

    # Right Column: Telemetry Screenshot Card
    etemp_img = "assets/images/etemp-dashboard.png"
    if os.path.exists(etemp_img):
        add_card(slide4, Inches(7.7), Inches(1.9), Inches(4.8), Inches(4.8), border_color=COLOR_YELLOW, fill_color=COLOR_BG_CARD)
        try:
            slide4.shapes.add_picture(etemp_img, Inches(7.85), Inches(2.1), width=Inches(4.5), height=Inches(3.3))
        except Exception as e:
            print(f"Notice: E-Temp image embed failed: {e}")
        
        cap4 = slide4.shapes.add_textbox(Inches(7.85), Inches(5.6), Inches(4.5), Inches(1.0))
        c4tf = cap4.text_frame
        c4tf.word_wrap = True
        c4tf.margin_top = Pt(2)
        c4p = c4tf.paragraphs[0]
        c4run = c4p.add_run()
        c4run.text = "TECH STACK & METRICS:\n"
        c4run.font.name = FONT_MONO
        c4run.font.bold = True
        c4run.font.size = Pt(9)
        c4run.font.color.rgb = COLOR_CYAN
        c4run2 = c4p.add_run()
        c4run2.text = "Laravel 11, PHP 8.2, Python API, MySQL, Chart.js  |  12 Industrial Cold Storage Machines  |  Client: CV. Dimas Djaya Teknik"
        c4run2.font.name = FONT_BODY
        c4run2.font.size = Pt(8.5)
        c4run2.font.color.rgb = COLOR_TEXT_SECONDARY

    # --------------------------------------------------------------------------
    # SLIDE 5: FEATURED PROJECT 03 — ENTERPRISE BIG DATA LAKEHOUSE (OLIST)
    # --------------------------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide5, 5, category="BIG DATA & ENTERPRISE ARCHITECTURE")

    hb5 = slide5.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf5 = hb5.text_frame
    hp5 = htf5.paragraphs[0]
    hp5.add_run().text = "// BIG DATA CAPSTONE PROJECT: E-COMMERCE ENTERPRISE LAKEHOUSE\n"
    hp5.runs[0].font.name = FONT_MONO
    hp5.runs[0].font.size = Pt(10)
    hp5.runs[0].font.color.rgb = COLOR_CYAN

    hp5_2 = hp5.add_run()
    hp5_2.text = "Enterprise E-Commerce Big Data Lakehouse & Executive Analytics"
    hp5_2.font.name = FONT_DISPLAY
    hp5_2.font.size = Pt(22)
    hp5_2.font.bold = True
    hp5_2.font.color.rgb = COLOR_WHITE

    p3_card = add_card(slide5, Inches(0.8), Inches(1.9), Inches(6.6), Inches(4.8), border_color=COLOR_BORDER, fill_color=COLOR_BG_CARD)
    p3tf = p3_card.text_frame
    p3tf.word_wrap = True
    p3tf.margin_left = Pt(14)
    p3tf.margin_right = Pt(14)
    p3tf.margin_top = Pt(12)

    p3_head = p3tf.paragraphs[0]
    p3_head.add_run().text = "DATA ARCHITECTURE & BUSINESS ANALYTICS\n"
    p3_head.runs[0].font.name = FONT_MONO
    p3_head.runs[0].font.size = Pt(10)
    p3_head.runs[0].font.bold = True
    p3_head.runs[0].font.color.rgb = COLOR_YELLOW

    bullets_p3 = [
        ("Scale & Dataset: ", "Architected pipeline processing 1.5M+ rows from 9 relational tables of Brazilian E-Commerce data (100k orders, R$20.4M total GMV)."),
        ("3-Layer Medallion Lakehouse: ", "Designed Hadoop HDFS data warehouse across Raw (Landing), Staging (Cleansed/Typed), and Cleaned (Analytics-ready Star Schema)."),
        ("ETL/ELT Batch & Streaming: ", "Engineered Spark SQL transformations joining 6 source tables into 118,107 unified fact records; simulated Kafka real-time ingestion topics."),
        ("Predictive Delivery Analytics: ", "Trained Spark MLlib Linear Regression model predicting order delivery duration based on freight value and geolocation distance."),
        ("Executive BI Dashboards: ", "Deployed Apache Superset dashboards tracking monthly revenue velocity, payment method breakdowns, freight cost impacts, and customer lifetime value.")
    ]

    for label, desc in bullets_p3:
        bp = p3tf.add_paragraph()
        bp.space_before = Pt(6)
        r_lbl = bp.add_run()
        r_lbl.text = "• " + label
        r_lbl.font.name = FONT_BODY
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.color.rgb = COLOR_CYAN
        r_desc = bp.add_run()
        r_desc.text = desc
        r_desc.font.name = FONT_BODY
        r_desc.font.size = Pt(9.2)
        r_desc.font.color.rgb = COLOR_TEXT_PRIMARY

    # Right Column: Superset Dashboard Screenshot
    olist_img = "assets/images/olist-superset-overview.png"
    if os.path.exists(olist_img):
        add_card(slide5, Inches(7.7), Inches(1.9), Inches(4.8), Inches(4.8), border_color=COLOR_CYAN, fill_color=COLOR_BG_CARD)
        try:
            slide5.shapes.add_picture(olist_img, Inches(7.85), Inches(2.1), width=Inches(4.5), height=Inches(3.3))
        except Exception as e:
            print(f"Notice: Olist image embed failed: {e}")
        
        cap5 = slide5.shapes.add_textbox(Inches(7.85), Inches(5.6), Inches(4.5), Inches(1.0))
        c5tf = cap5.text_frame
        c5tf.word_wrap = True
        c5tf.margin_top = Pt(2)
        c5p = c5tf.paragraphs[0]
        c5run = c5p.add_run()
        c5run.text = "TECH STACK & METRICS:\n"
        c5run.font.name = FONT_MONO
        c5run.font.bold = True
        c5run.font.size = Pt(8.5)
        c5run.font.color.rgb = COLOR_YELLOW
        c5run2 = c5p.add_run()
        c5run2.text = "Hadoop HDFS, PySpark, Kafka, MLlib, Superset  |  1.5M+ Rows\n"
        c5run2.font.name = FONT_BODY
        c5run2.font.size = Pt(8)
        c5run2.font.color.rgb = COLOR_TEXT_SECONDARY

        c5run3 = c5p.add_run()
        c5run3.text = "📊 SLIDE DECK: https://drive.google.com/file/d/1-_KjfIORiWMeEm45z8R80D6LOb4IdRg4/view?usp=drive_link"
        c5run3.font.name = FONT_MONO
        c5run3.font.size = Pt(7.8)
        c5run3.font.bold = True
        c5run3.font.color.rgb = COLOR_CYAN

    # --------------------------------------------------------------------------
    # SLIDE 6: FEATURED PROJECT 04 — ODOO ERP IMPLEMENTATION
    # --------------------------------------------------------------------------
    slide6 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide6, 6, category="ENTERPRISE SYSTEMS IMPLEMENTATION")

    hb6 = slide6.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf6 = hb6.text_frame
    hp6 = htf6.paragraphs[0]
    hp6.add_run().text = "// ENTERPRISE PROJECT 04: NUHSANTARA MERCHANDISE (SURABAYA)\n"
    hp6.runs[0].font.name = FONT_MONO
    hp6.runs[0].font.size = Pt(10)
    hp6.runs[0].font.color.rgb = COLOR_CYAN

    hp6_2 = hp6.add_run()
    hp6_2.text = "Odoo ERP Multi-Module Implementation & Business Process Reengineering"
    hp6_2.font.name = FONT_DISPLAY
    hp6_2.font.size = Pt(22)
    hp6_2.font.bold = True
    hp6_2.font.color.rgb = COLOR_WHITE

    # 3 Cards for the 3 ERP Domains
    erp_cards = [
        (
            "ERP IMPLEMENTOR LEAD",
            "Module Configuration & Master Data",
            COLOR_YELLOW,
            [
                "Configured 3 core Odoo ERP modules: Website/E-Commerce, Purchasing, and Inventory Management.",
                "Audited and cleaned legacy master catalogs (SKU numbering, barcode mapping, vendor records).",
                "Migrated opening stock balances and established automated reorder thresholds to prevent stockouts.",
                "Delivered comprehensive operational training and authored end-user SOP documentation."
            ]
        ),
        (
            "IT BUSINESS ANALYST",
            "Workflow Fit-Gap & Process Design",
            COLOR_CYAN,
            [
                "Conducted fit-gap analysis between partner's legacy manual procurement and standard Odoo ERP capabilities.",
                "Mapped As-Is retail inventory turnover and designed streamlined To-Be automated requisition approval flows.",
                "Established unified vendor price comparison lists and purchase order routing based on expenditure tiers.",
                "Defined KPI tracking for inventory holding costs, stock turnover ratio, and vendor lead times."
            ]
        ),
        (
            "PROJECT COORDINATOR",
            "Milestones & UAT Sign-off",
            COLOR_YELLOW,
            [
                "Coordinated cross-functional tasks across a 3-member team over a 4-month academic implementation.",
                "Created deployment roadmap, managed project timeline, and prioritized high-impact modules.",
                "Conducted rigorous User Acceptance Testing (UAT) sessions with retail leadership and store managers.",
                "Ensured seamless system handover with zero operational disruption to active merchandising sales."
            ]
        )
    ]

    ec_width = Inches(3.75)
    ec_gap = Inches(0.22)
    for idx, (title, subtitle, col, bullets) in enumerate(erp_cards):
        left_pos = Inches(0.8) + idx * (ec_width + ec_gap)
        ec = add_card(slide6, left_pos, Inches(1.9), ec_width, Inches(4.8), border_color=col, fill_color=COLOR_BG_CARD)
        etf = ec.text_frame
        etf.word_wrap = True
        etf.margin_left = Pt(14)
        etf.margin_right = Pt(14)
        etf.margin_top = Pt(14)

        ep = etf.paragraphs[0]
        ep.add_run().text = title + "\n"
        ep.runs[0].font.name = FONT_DISPLAY
        ep.runs[0].font.size = Pt(14)
        ep.runs[0].font.bold = True
        ep.runs[0].font.color.rgb = col

        ep.add_run().text = subtitle + "\n\n"
        ep.runs[1].font.name = FONT_MONO
        ep.runs[1].font.size = Pt(9)
        ep.runs[1].font.color.rgb = COLOR_TEXT_MUTED

        for b in bullets:
            bp = etf.add_paragraph()
            bp.space_before = Pt(6)
            bp.add_run().text = "▸ "
            bp.runs[0].font.name = FONT_MONO
            bp.runs[0].font.size = Pt(9)
            bp.runs[0].font.color.rgb = col
            bp.add_run().text = b
            bp.runs[1].font.name = FONT_BODY
            bp.runs[1].font.size = Pt(9.5)
            bp.runs[1].font.color.rgb = COLOR_TEXT_PRIMARY

    # --------------------------------------------------------------------------
    # SLIDE 7: FEATURED PROJECT 05 — ORDER-TO-CASH (O2C) & INVENTORY AUTOMATION
    # --------------------------------------------------------------------------
    slide7 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide7, 7, category="BUSINESS ANALYSIS & O2C AUTOMATION CASE STUDY")

    hb7 = slide7.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf7 = hb7.text_frame
    hp7 = htf7.paragraphs[0]
    hp7.add_run().text = "// FLAGSHIP CASE STUDY 05: PT MAJU DISTRIBUSI NUSANTARA (B2B FMCG)\n"
    hp7.runs[0].font.name = FONT_MONO
    hp7.runs[0].font.size = Pt(10)
    hp7.runs[0].font.color.rgb = COLOR_CYAN

    hp7_2 = hp7.add_run()
    hp7_2.text = "Order-to-Cash (O2C) & Inventory Automation: Business Analysis Case Study"
    hp7_2.font.name = FONT_DISPLAY
    hp7_2.font.size = Pt(22)
    hp7_2.font.bold = True
    hp7_2.font.color.rgb = COLOR_WHITE

    # Left Column: Detailed Specification & Engineering Card
    o2c_card = add_card(slide7, Inches(0.8), Inches(1.9), Inches(6.6), Inches(4.8), border_color=COLOR_BORDER, fill_color=COLOR_BG_CARD)
    o2ctf = o2c_card.text_frame
    o2ctf.word_wrap = True
    o2ctf.margin_left = Pt(14)
    o2ctf.margin_right = Pt(14)
    o2ctf.margin_top = Pt(12)

    o2c_head = o2ctf.paragraphs[0]
    o2c_head.add_run().text = "FULL-LIFECYCLE BA & SYSTEMS ARCHITECTURE (ODOO ERP)\n"
    o2c_head.runs[0].font.name = FONT_MONO
    o2c_head.runs[0].font.size = Pt(10)
    o2c_head.runs[0].font.bold = True
    o2c_head.runs[0].font.color.rgb = COLOR_YELLOW

    bullets_o2c = [
        ("Business Problem: ", "350+ B2B retail stores handled via WhatsApp/manual Excel with 6% mistyped orders, 9% overselling, and 1-2 day invoicing delays."),
        ("BRD & FSD Specifications: ", "Authored 46 functional requirements (MoSCoW prioritized), 7 non-functional requirements, and 10 strict business rules."),
        ("Architecture & Database Design: ", "Modeled AS-IS & TO-BE BPMN 2.0 workflows; designed 25-table relational ERD with atomic stock reservation (Available = On Hand - Reserved) & optimistic locking."),
        ("Agile & Test Management: ", "Planned 5 sprints (100 story points, 17 user stories with Gherkin syntax, Jira CSV, RACI); executed 62 test cases with 100% RTM traceability and 9 end-to-end UAT scenarios."),
        ("Validated Simulation Results: ", "5,000 simulated orders demonstrated order cycle reduced from 23.6h to 1.4h, overselling reduced from 8.6% to 0.8%, and invoice lead time cut from 37.1h to 6.7 min.")
    ]

    for label, desc in bullets_o2c:
        bp = o2ctf.add_paragraph()
        bp.space_before = Pt(5)
        r_lbl = bp.add_run()
        r_lbl.text = "• " + label
        r_lbl.font.name = FONT_BODY
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.color.rgb = COLOR_CYAN
        r_desc = bp.add_run()
        r_desc.text = desc
        r_desc.font.name = FONT_BODY
        r_desc.font.size = Pt(9.2)
        r_desc.font.color.rgb = COLOR_TEXT_PRIMARY

    # Right Column: KPI Comparison Diagram Card
    o2c_img = "assets/images/o2c-kpi-comparison.png"
    if os.path.exists(o2c_img):
        add_card(slide7, Inches(7.7), Inches(1.9), Inches(4.8), Inches(4.8), border_color=COLOR_YELLOW, fill_color=COLOR_BG_CARD)
        try:
            slide7.shapes.add_picture(o2c_img, Inches(7.85), Inches(2.1), width=Inches(4.5), height=Inches(3.3))
        except Exception as e:
            print(f"Notice: O2C image embed failed: {e}")
        
        cap7 = slide7.shapes.add_textbox(Inches(7.85), Inches(5.55), Inches(4.5), Inches(1.1))
        c7tf = cap7.text_frame
        c7tf.word_wrap = True
        c7tf.margin_top = Pt(2)
        c7p = c7tf.paragraphs[0]
        c7run = c7p.add_run()
        c7run.text = "LIVE DASHBOARD & GITHUB REPO:\n"
        c7run.font.name = FONT_MONO
        c7run.font.bold = True
        c7run.font.size = Pt(8.5)
        c7run.font.color.rgb = COLOR_YELLOW
        c7run2 = c7p.add_run()
        c7run2.text = "Interactive Dashboard: https://yaboialip.github.io/o2c-automation-case-study/\nGitHub Repository: https://github.com/Yaboialip/o2c-automation-case-study"
        c7run2.font.name = FONT_MONO
        c7run2.font.size = Pt(7.8)
        c7run2.font.color.rgb = COLOR_CYAN

    # --------------------------------------------------------------------------
    # SLIDE 8: SPECIALIZED TECHNICAL PROJECTS (GRID OF 5 PROJECTS)
    # --------------------------------------------------------------------------
    slide8 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide8, 8, category="TECHNICAL DOSSIER & SPECIALIZED PROJECTS")

    hb8 = slide8.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf8 = hb8.text_frame
    hp8 = htf8.paragraphs[0]
    hp8.add_run().text = "// ADDITIONAL SPECIALIZED ENGAGEMENTS & APPLIED RESEARCH\n"
    hp8.runs[0].font.name = FONT_MONO
    hp8.runs[0].font.size = Pt(10)
    hp8.runs[0].font.color.rgb = COLOR_CYAN

    hp8_2 = hp8.add_run()
    hp8_2.text = "Enterprise Architecture, Governance, BI & Applied Machine Learning"
    hp8_2.font.name = FONT_DISPLAY
    hp8_2.font.size = Pt(22)
    hp8_2.font.bold = True
    hp8_2.font.color.rgb = COLOR_WHITE

    spec_projects = [
        (
            "COBIT 2019 IT GOVERNANCE RESEARCH",
            "PT Infoglobal Teknologi Semesta (2024)",
            COLOR_CYAN,
            "Co-authored research evaluating 10 design factors and prioritizing EDM03, APO12, DSS05, DSS04, and MEA03 governance objectives.",
            "Tags: COBIT 2019 • Risk Management • IT Governance Audit"
        ),
        (
            "ENTERPRISE ARCHITECTURE BLUEPRINT",
            "Infoglobal IT Division (TOGAF ADM 9.2)",
            COLOR_YELLOW,
            "Modeled full TOGAF architecture in Visual Paradigm across Business, Data, Application, and Technology domains to align IT roadmap with corporate strategy.",
            "Tags: TOGAF 9.2 • Visual Paradigm • IT Blueprint"
        ),
        (
            "POWER BI SALES PERFORMANCE INTELLIGENCE",
            "PT Gelatik Confectionery (2025)",
            COLOR_CYAN,
            "Built interactive executive dashboard analyzing regional sales velocity, customer purchasing frequency, and revenue distribution trends.",
            "Tags: Power BI • DAX Data Modeling • Sales Analytics"
        ),
        (
            "ARIMA-GARCH STOCK VOLATILITY PREDICTION",
            "PT Pertamina Geothermal Energy Tbk (PGEO)",
            COLOR_YELLOW,
            "Undergraduate Thesis (Grade A, 2026): Python forecasting web app (Streamlit, arch, statsmodels) modeling PGEO return volatility on IDX.\nManuscript (PDF): drive.google.com/file/d/1sFt-QrZIYCprjvEfZtW41hoVChWjGJXn",
            "Tags: Python • Econometrics • Streamlit • Thesis (Grade A)"
        ),
        (
            "AI CHATBOT WITH MODEL CONTEXT PROTOCOL (MCP)",
            "Google Cloud Gen AI Academy APAC (2026)",
            COLOR_CYAN,
            "Built AI chatbot capstone leveraging Model Context Protocol (MCP) for dynamic context retrieval, integrating generative AI with enterprise cloud APIs.",
            "Tags: Google Cloud • Model Context Protocol • Gen AI"
        )
    ]

    # Render as 5 clean horizontal cards / two columns
    for idx, (title, org, col, desc, tags) in enumerate(spec_projects):
        if idx < 3:
            top_pos = Inches(1.9) + idx * Inches(1.6)
            left_pos = Inches(0.8)
            c_width = Inches(5.7)
        else:
            top_pos = Inches(1.9) + (idx - 3) * Inches(2.4)
            left_pos = Inches(6.8)
            c_width = Inches(5.7)

        sp_card = add_card(slide8, left_pos, top_pos, c_width, Inches(1.48) if idx < 3 else Inches(2.28), border_color=COLOR_BORDER, fill_color=COLOR_BG_CARD)
        sptf = sp_card.text_frame
        sptf.word_wrap = True
        sptf.margin_left = Pt(12)
        sptf.margin_right = Pt(12)
        sptf.margin_top = Pt(8)

        p = sptf.paragraphs[0]
        p.add_run().text = title + "\n"
        p.runs[0].font.name = FONT_DISPLAY
        p.runs[0].font.size = Pt(11.5)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = col

        p.add_run().text = org + "\n"
        p.runs[1].font.name = FONT_MONO
        p.runs[1].font.size = Pt(8)
        p.runs[1].font.color.rgb = COLOR_TEXT_MUTED

        p_desc = sptf.add_paragraph()
        p_desc.space_before = Pt(3)
        p_desc.add_run().text = desc
        p_desc.runs[0].font.name = FONT_BODY
        p_desc.runs[0].font.size = Pt(8.8)
        p_desc.runs[0].font.color.rgb = COLOR_TEXT_PRIMARY

        p_tags = sptf.add_paragraph()
        p_tags.space_before = Pt(4)
        p_tags.add_run().text = tags
        p_tags.runs[0].font.name = FONT_MONO
        p_tags.runs[0].font.size = Pt(7.5)
        p_tags.runs[0].font.color.rgb = COLOR_YELLOW

    # --------------------------------------------------------------------------
    # SLIDE 9: PEER-REVIEWED SCIENTIFIC RESEARCH PUBLICATION (TOGAF ADM 9.2)
    # --------------------------------------------------------------------------
    slide9 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide9, 9, category="DECLASSIFIED SCIENTIFIC RESEARCH")

    hb9 = slide9.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf9 = hb9.text_frame
    hp9 = htf9.paragraphs[0]
    hp9.add_run().text = "// PEER-REVIEWED SCIENTIFIC PUBLICATION: SINTA 2 ACCREDITED JOURNAL\n"
    hp9.runs[0].font.name = FONT_MONO
    hp9.runs[0].font.size = Pt(10)
    hp9.runs[0].font.color.rgb = COLOR_YELLOW

    hp9_2 = hp9.add_run()
    hp9_2.text = "Enterprise Architecture Design for Digital Village Public Services"
    hp9_2.font.name = FONT_DISPLAY
    hp9_2.font.size = Pt(22)
    hp9_2.font.bold = True
    hp9_2.font.color.rgb = COLOR_WHITE

    pub_card = add_card(slide9, Inches(0.8), Inches(1.9), Inches(11.7), Inches(4.8), border_color=COLOR_YELLOW, fill_color=COLOR_BG_CARD)
    pubtf = pub_card.text_frame
    pubtf.word_wrap = True
    pubtf.margin_left = Pt(20)
    pubtf.margin_right = Pt(20)
    pubtf.margin_top = Pt(16)

    p_meta = pubtf.paragraphs[0]
    p_meta.add_run().text = "JOURNAL RECORD: Jurnal Masyarakat Mandiri (JMM)  •  Vol. 8, No. 6 (December 2024)  •  ISSN 2614-4921 / E-ISSN 2598-8158\n"
    p_meta.runs[0].font.name = FONT_MONO
    p_meta.runs[0].font.size = Pt(9.5)
    p_meta.runs[0].font.bold = True
    p_meta.runs[0].font.color.rgb = COLOR_CYAN

    p_title = pubtf.add_paragraph()
    p_title.space_before = Pt(4)
    p_title.add_run().text = "Original Title: Perancangan Enterprise Architecture Menggunakan TOGAF ADM 9.2 pada Pelayanan Publik Desa Sukalaksana\n"
    p_title.runs[0].font.name = FONT_DISPLAY
    p_title.runs[0].font.size = Pt(13)
    p_title.runs[0].font.bold = True
    p_title.runs[0].font.color.rgb = COLOR_WHITE

    pub_points = [
        ("Research Objective: ", "Bridged fragmented administrative public services in Desa Sukalaksana into an integrated digital governance blueprint using the TOGAF ADM 9.2 framework."),
        ("Architectural Scope: ", "Executed Phase A through Phase F (Preliminary, Architecture Vision, Business Architecture, Information Systems Architecture [Data & Application], Technology Architecture, and Migration Planning)."),
        ("Gap Analysis & Resolution: ", "Identified and addressed 14 critical governance and operational gaps, developing target architectural blueprints modeled in Visual Paradigm."),
        ("Application & Tech Blueprint: ", "Specified an integrated public service platform, citizen portal, relational database schema, and mobile-friendly cloud deployment infrastructure."),
        ("Academic Validation: ", "Peer-reviewed, approved, and formally published in an accredited SINTA 2 national journal, demonstrating rigorous research execution.")
    ]

    for label, desc in pub_points:
        bp = pubtf.add_paragraph()
        bp.space_before = Pt(6)
        r_lbl = bp.add_run()
        r_lbl.text = "• " + label
        r_lbl.font.name = FONT_BODY
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(10)
        r_lbl.font.color.rgb = COLOR_YELLOW
        r_desc = bp.add_run()
        r_desc.text = desc
        r_desc.font.name = FONT_BODY
        r_desc.font.size = Pt(9.8)
        r_desc.font.color.rgb = COLOR_TEXT_PRIMARY

    doi_p = pubtf.add_paragraph()
    doi_p.space_before = Pt(10)
    doi_p.add_run().text = "🔗 DOI REFERENCE: https://doi.org/10.31764/jmm.v8i6.29744  |  INDEXED IN SINTA S2, GOOGLE SCHOLAR, GARUDA\n"
    doi_p.runs[0].font.name = FONT_MONO
    doi_p.runs[0].font.size = Pt(8.5)
    doi_p.runs[0].font.bold = True
    doi_p.runs[0].font.color.rgb = COLOR_CYAN

    rg_run = doi_p.add_run()
    rg_run.text = "📄 RESEARCHGATE FULL-TEXT (PDF): https://www.researchgate.net/publication/392726934_PERANCANGAN_IT_BLUEPRINT_SEBAGAI_STRATEGI_PENGEMBANGAN_KONSEP_SMART_VILLAGE"
    rg_run.font.name = FONT_MONO
    rg_run.font.size = Pt(8.5)
    rg_run.font.bold = True
    rg_run.font.color.rgb = COLOR_YELLOW

    # --------------------------------------------------------------------------
    # SLIDE 10: CAREER TRACK & LEADERSHIP RECORD
    # --------------------------------------------------------------------------
    slide10 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide10, 10, category="CAREER TIMELINE & LEADERSHIP")

    hb10 = slide10.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf10 = hb10.text_frame
    hp10 = htf10.paragraphs[0]
    hp10.add_run().text = "// PROFESSIONAL EXPERIENCE & LEADERSHIP ROLES\n"
    hp10.runs[0].font.name = FONT_MONO
    hp10.runs[0].font.size = Pt(10)
    hp10.runs[0].font.color.rgb = COLOR_CYAN

    hp10_2 = hp10.add_run()
    hp10_2.text = "Practical Industry Experience, Academic Mentorship & Organizational Leadership"
    hp10_2.font.name = FONT_DISPLAY
    hp10_2.font.size = Pt(22)
    hp10_2.font.bold = True
    hp10_2.font.color.rgb = COLOR_WHITE

    roles_timeline = [
        (
            "PROJECT COORDINATOR (INTERNSHIP)",
            "CV. Netafarm Indolestari (Jun – Aug 2025)",
            COLOR_CYAN,
            [
                "Authored formal Business Requirements Document (BRD) & Functional Specification Document (FSD / FRD).",
                "Coordinated a 3-person developer team over an 8-week Waterfall cycle in Jira to deliver an OKR task tracking system.",
                "Conducted black-box and white-box testing, managed defect resolution, and facilitated UAT sign-off with leadership.",
                "Awarded official Company Certificate of Appreciation for on-time and high-quality system handover."
            ]
        ),
        (
            "REQUIREMENT ANALYST (ACADEMIC PROJECT)",
            "CV. Dimas Djaya Teknik (Sep – Dec 2025)",
            COLOR_YELLOW,
            [
                "Authored formal BRD & FSD for multi-branch industrial cold-chain monitoring system (12 machines).",
                "Engineered complete web QA test suite in '02-Workbook-Test-Case' (Excel), validating features against Acceptance Criteria.",
                "Designed application architecture (Laravel, PHP, Python API, MySQL) and telemetry excursion alert logic.",
                "Consolidated temperature readings via Excel VLOOKUP and Pivot Tables to generate monthly comparative reports."
            ]
        ),
        (
            "LABORATORY ASSISTANT, ENTERPRISE ARCHITECTURE",
            "Telkom University Surabaya (Sep 2024 – Dec 2025)",
            COLOR_CYAN,
            [
                "Appointed for two consecutive academic semesters to mentor 50+ students in Enterprise Architecture design.",
                "Developed course modules and graded case studies applying the TOGAF ADM 9.2 framework in Visual Paradigm.",
                "Organized and served as Keynote Speaker for 'DevOps x System Analyst: Modern IT Career Path' seminar (120+ attendees).",
                "Supervised team modeling across Business, Data, Application, and Technology domains."
            ]
        )
    ]

    for idx, (title, org, col, bullets) in enumerate(roles_timeline):
        left_pos = Inches(0.8) + idx * Inches(3.97)
        t_card = add_card(slide10, left_pos, Inches(1.9), Inches(3.78), Inches(4.8), border_color=col, fill_color=COLOR_BG_CARD)
        ttf = t_card.text_frame
        ttf.word_wrap = True
        ttf.margin_left = Pt(12)
        ttf.margin_right = Pt(12)
        ttf.margin_top = Pt(12)

        tp = ttf.paragraphs[0]
        tp.add_run().text = title + "\n"
        tp.runs[0].font.name = FONT_DISPLAY
        tp.runs[0].font.size = Pt(12)
        tp.runs[0].font.bold = True
        tp.runs[0].font.color.rgb = col

        tp.add_run().text = org + "\n\n"
        tp.runs[1].font.name = FONT_MONO
        tp.runs[1].font.size = Pt(8.5)
        tp.runs[1].font.color.rgb = COLOR_TEXT_MUTED

        for b in bullets:
            bp = ttf.add_paragraph()
            bp.space_before = Pt(6)
            bp.add_run().text = "▸ "
            bp.runs[0].font.name = FONT_MONO
            bp.runs[0].font.size = Pt(8.5)
            bp.runs[0].font.color.rgb = col
            bp.add_run().text = b
            bp.runs[1].font.name = FONT_BODY
            bp.runs[1].font.size = Pt(9)
            bp.runs[1].font.color.rgb = COLOR_TEXT_PRIMARY

    # --------------------------------------------------------------------------
    # SLIDE 11: FORMAL CERTIFICATIONS & CREDENTIALS
    # --------------------------------------------------------------------------
    slide11 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide11, 11, category="VERIFIED CREDENTIALS & CERTIFICATIONS")

    hb11 = slide11.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf11 = hb11.text_frame
    hp11 = htf11.paragraphs[0]
    hp11.add_run().text = "// FORMAL INDUSTRY CERTIFICATIONS & ACADEMIC ACCREDITATION\n"
    hp11.runs[0].font.name = FONT_MONO
    hp11.runs[0].font.size = Pt(10)
    hp11.runs[0].font.color.rgb = COLOR_CYAN

    hp11_2 = hp11.add_run()
    hp11_2.text = "Verified Industry Credentials Across Project Management, Systems Analysis & SAP"
    hp11_2.font.name = FONT_DISPLAY
    hp11_2.font.size = Pt(22)
    hp11_2.font.bold = True
    hp11_2.font.color.rgb = COLOR_WHITE

    certs_data = [
        (
            "GOOGLE PROJECT MANAGEMENT PROFESSIONAL",
            "Google (Coursera)  •  Issued 2026",
            COLOR_YELLOW,
            "Mastered Agile & Scrum governance, Waterfall WBS planning, project charters, risk management registers, sprint estimation, and executive stakeholder communication across 6 comprehensive courses.\nDrive: drive.google.com/file/d/1px_zULs9ktCtz4HS9y-mV5yoizBd3pu0"
        ),
        (
            "GOOGLE AI ESSENTIALS & GCP GEN AI",
            "Google Career Certs & Cloud Academy  •  2026",
            COLOR_CYAN,
            "Certified in foundation AI prompt engineering, workplace AI productivity tooling, enterprise Generative AI APIs, and Model Context Protocol (MCP) multi-agent systems.\nDrive: drive.google.com/file/d/1w9exZvruwxxjmYWl_AabMteworOG7CKH"
        ),
        (
            "SAP CERTIFIED: SCM500 (PROCUREMENT)",
            "SAP SE  •  Issued 2026",
            COLOR_YELLOW,
            "Formal certification in enterprise procurement flows, purchase requisition lifecycle, master vendor data management, inventory valuation, and goods receipt verification.\nDrive: drive.google.com/file/d/1Jld8wiO7VhbPhoA1k-4gh8whQ5Q0zFCP"
        ),
        (
            "SAP CERTIFIED: HR050 (HUMAN CAPITAL)",
            "SAP SE  •  Issued 2026",
            COLOR_CYAN,
            "Certified in SAP enterprise human capital management workflows, organizational structure design, employee lifecycle data administration, and corporate personnel reporting.\nDrive: drive.google.com/file/d/1rbG-AEVhXtDvmtQMFCuqIaCu8rKrhB2F"
        ),
        (
            "EF SET ENGLISH: C2 ADVANCED",
            "EF Standard English Test  •  Score 77/100 (C2)",
            COLOR_YELLOW,
            "Demonstrated bilingual command of professional English at CEFR C2 Advanced standard; capable of cross-border executive alignment and international system documentation.\nDrive: drive.google.com/file/d/1ZR8rK2Md03JX2uoQHXl3r09yCc6Qydpe"
        ),
        (
            "BACHELOR OF SCIENCE IN INFORMATION SYSTEMS",
            "Telkom University  •  Graduated Oct 2026 (GPA 3.66/4.00)",
            COLOR_CYAN,
            "Completed 144 credits with High Honors. Specialization in Enterprise Systems, Big Data Analytics, Systems Architecture, and Business Process Management. Convocation: Oct 26, 2026."
        )
    ]

    for idx, (title, org, col, desc) in enumerate(certs_data):
        row = idx // 3
        col_idx = idx % 3
        c_left = Inches(0.8) + col_idx * Inches(3.97)
        c_top = Inches(1.9) + row * Inches(2.45)

        cc = add_card(slide11, c_left, c_top, Inches(3.78), Inches(2.25), border_color=col, fill_color=COLOR_BG_CARD)
        ctf = cc.text_frame
        ctf.word_wrap = True
        ctf.margin_left = Pt(12)
        ctf.margin_right = Pt(12)
        ctf.margin_top = Pt(10)

        cp = ctf.paragraphs[0]
        cp.add_run().text = title + "\n"
        cp.runs[0].font.name = FONT_DISPLAY
        cp.runs[0].font.size = Pt(11.5)
        cp.runs[0].font.bold = True
        cp.runs[0].font.color.rgb = col

        cp.add_run().text = org + "\n\n"
        cp.runs[1].font.name = FONT_MONO
        cp.runs[1].font.size = Pt(8.5)
        cp.runs[1].font.color.rgb = COLOR_TEXT_MUTED

        cp_desc = ctf.add_paragraph()
        cp_desc.space_before = Pt(2)
        cp_desc.add_run().text = desc
        cp_desc.runs[0].font.name = FONT_BODY
        cp_desc.runs[0].font.size = Pt(9)
        cp_desc.runs[0].font.color.rgb = COLOR_TEXT_PRIMARY

    # --------------------------------------------------------------------------
    # SLIDE 12: TECHNICAL & MANAGEMENT SKILLS MATRIX
    # --------------------------------------------------------------------------
    slide12 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide12, 12, category="TECHNICAL SKILLS & TOOLING MATRIX")

    hb12 = slide12.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf12 = hb12.text_frame
    hp12 = htf12.paragraphs[0]
    hp12.add_run().text = "// TOOLING & METHODOLOGY REPERTOIRE\n"
    hp12.runs[0].font.name = FONT_MONO
    hp12.runs[0].font.size = Pt(10)
    hp12.runs[0].font.color.rgb = COLOR_CYAN

    hp12_2 = hp12.add_run()
    hp12_2.text = "Comprehensive Competency Matrix Across Governance, Data, & Architecture"
    hp12_2.font.name = FONT_DISPLAY
    hp12_2.font.size = Pt(22)
    hp12_2.font.bold = True
    hp12_2.font.color.rgb = COLOR_WHITE

    skills_pillars = [
        (
            "PROJECT MANAGEMENT & DELIVERY",
            COLOR_YELLOW,
            [
                ("Methodologies:", "Waterfall (SDLC), Agile Scrum, Hybrid Governance"),
                ("Work Breakdown (WBS):", "Milestone scheduling, critical path, Gantt charts"),
                ("Tools & Tracking:", "Jira Software, Trello, Google Workspace"),
                ("Quality & Testing:", "UAT facilitation, black/white-box QA, defect logging"),
                ("Risk & Scope:", "Scope change control, risk registers, stakeholder reports")
            ]
        ),
        (
            "BUSINESS & SYSTEMS ANALYSIS",
            COLOR_CYAN,
            [
                ("Core Deliverables:", "BRD, FSD/FRD, Project Charters, User Stories"),
                ("Process Modeling:", "BPMN 2.0 (As-Is / To-Be), Flowcharts, Mind Mapping"),
                ("Data Analysis (Excel):", "VLOOKUP, INDEX/MATCH, Pivot Tables, IF conditions"),
                ("Systems Architecture:", "TOGAF ADM 9.2, COBIT 2019, Visual Paradigm"),
                ("UML & Data Schemas:", "Use Case, Activity, Sequence, 3NF Relational ERD")
            ]
        ),
        (
            "ENTERPRISE SYSTEMS & BIG DATA",
            COLOR_YELLOW,
            [
                ("ERP Implementation:", "Odoo ERP (Modules: E-Commerce, Inventory, Purchase)"),
                ("SAP Credentials:", "SAP SCM500 (Procurement), SAP HR050 (HCM)"),
                ("Big Data Architecture:", "Hadoop HDFS, Apache Spark (PySpark), Apache Kafka"),
                ("Databases & APIs:", "MySQL, PostgreSQL, RESTful JSON APIs, Postman"),
                ("Business Intelligence:", "Apache Superset, Power BI, Chart.js Dashboards")
            ]
        )
    ]

    p_width = Inches(3.78)
    p_gap = Inches(0.20)
    for idx, (pillar_name, col, items) in enumerate(skills_pillars):
        c_left = Inches(0.8) + idx * (p_width + p_gap)
        sk_card = add_card(slide11, c_left, Inches(1.9), p_width, Inches(4.8), border_color=col, fill_color=COLOR_BG_CARD)
        sktf = sk_card.text_frame
        sktf.word_wrap = True
        sktf.margin_left = Pt(14)
        sktf.margin_right = Pt(14)
        sktf.margin_top = Pt(14)

        sp = sktf.paragraphs[0]
        sp.add_run().text = pillar_name + "\n\n"
        sp.runs[0].font.name = FONT_DISPLAY
        sp.runs[0].font.size = Pt(13)
        sp.runs[0].font.bold = True
        sp.runs[0].font.color.rgb = col

        for key, val in items:
            ip = sktf.add_paragraph()
            ip.space_before = Pt(8)
            k_run = ip.add_run()
            k_run.text = "▸ " + key + "\n"
            k_run.font.name = FONT_MONO
            k_run.font.size = Pt(9.5)
            k_run.font.bold = True
            k_run.font.color.rgb = COLOR_WHITE

            v_run = ip.add_run()
            v_run.text = val
            v_run.font.name = FONT_BODY
            v_run.font.size = Pt(9.2)
            v_run.font.color.rgb = COLOR_TEXT_SECONDARY

    # --------------------------------------------------------------------------
    # SLIDE 13: CONTACT & DISPATCH (TACTICAL CALL TO ACTION)
    # --------------------------------------------------------------------------
    slide13 = prs.slides.add_slide(blank_layout)
    add_slide_background(slide13, 13, category="COMMUNICATION CHANNELS & DISPATCH")

    # Header
    hb13 = slide13.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8))
    htf13 = hb13.text_frame
    hp13 = htf13.paragraphs[0]
    hp13.add_run().text = "// INITIATE COMMUNICATION: OPEN FOR IMMEDIATE OPPORTUNITIES\n"
    hp13.runs[0].font.name = FONT_MONO
    hp13.runs[0].font.size = Pt(10)
    hp13.runs[0].font.color.rgb = COLOR_CYAN

    hp13_2 = hp13.add_run()
    hp13_2.text = "Let's Build Resilient, High-Impact IT Systems Together"
    hp13_2.font.name = FONT_DISPLAY
    hp13_2.font.size = Pt(22)
    hp13_2.font.bold = True
    hp13_2.font.color.rgb = COLOR_WHITE

    # Left Box: Roles of Interest
    c13_left = add_card(slide13, Inches(0.8), Inches(1.9), Inches(5.7), Inches(4.8), border_color=COLOR_YELLOW, fill_color=COLOR_BG_CARD)
    c13tf = c13_left.text_frame
    c13tf.word_wrap = True
    c13tf.margin_left = Pt(16)
    c13tf.margin_right = Pt(16)
    c13tf.margin_top = Pt(14)

    c13p = c13tf.paragraphs[0]
    c13p.add_run().text = "TARGET ROLES & AVAILABILITY\n"
    c13p.runs[0].font.name = FONT_MONO
    c13p.runs[0].font.size = Pt(10.5)
    c13p.runs[0].font.bold = True
    c13p.runs[0].font.color.rgb = COLOR_YELLOW

    c13_roles = [
        ("IT Project Manager / Project Coordinator", "Governing sprint schedules, team velocity, WBS dependencies, and UAT sign-off."),
        ("IT Business Analyst (IT BA)", "Conducting requirement elicitation, authoring BRD/FSD, BPMN workflows, and user stories."),
        ("Systems Analyst / Enterprise Architect", "Designing software architecture, UML diagrams, relational data schemas, and API contracts."),
        ("ERP Specialist / System Implementor", "Managing module deployment, master data migration, fit-gap analysis, and end-user training.")
    ]

    for r_title, r_desc in c13_roles:
        rp = c13tf.add_paragraph()
        rp.space_before = Pt(8)
        rt_run = rp.add_run()
        rt_run.text = "• " + r_title + "\n"
        rt_run.font.name = FONT_BODY
        rt_run.font.bold = True
        rt_run.font.size = Pt(10)
        rt_run.font.color.rgb = COLOR_CYAN
        rd_run = rp.add_run()
        rd_run.text = r_desc
        rd_run.font.name = FONT_BODY
        rd_run.font.size = Pt(9.2)
        rd_run.font.color.rgb = COLOR_TEXT_PRIMARY

    # Right Box: Direct Channels & Links
    c13_right = add_card(slide13, Inches(6.8), Inches(1.9), Inches(5.7), Inches(4.8), border_color=COLOR_CYAN, fill_color=COLOR_BG_CARD)
    rtf13 = c13_right.text_frame
    rtf13.word_wrap = True
    rtf13.margin_left = Pt(16)
    rtf13.margin_right = Pt(16)
    rtf13.margin_top = Pt(14)

    rp1 = rtf13.paragraphs[0]
    rp1.add_run().text = "DIRECT CONTACT FREQUENCIES\n"
    rp1.runs[0].font.name = FONT_MONO
    rp1.runs[0].font.size = Pt(10.5)
    rp1.runs[0].font.bold = True
    rp1.runs[0].font.color.rgb = COLOR_CYAN

    channels = [
        ("EMAIL DISPATCH:", "aliefsd84@gmail.com", "Primary contact channel for recruitment & discussions"),
        ("LINKEDIN PROFILE:", "linkedin.com/in/aliefsd84", "Connect for professional recommendations & endorsements"),
        ("ONLINE PORTFOLIO:", "https://yaboialip.github.io/My-portfolio/", "Interactive Web HUD version with live demo decrypters"),
        ("O2C AUTOMATION CASE STUDY:", "https://yaboialip.github.io/o2c-automation-case-study/", "Live simulated KPI dashboard & complete open-source repo")
    ]

    for ch_lbl, ch_val, ch_note in channels:
        cp = rtf13.add_paragraph()
        cp.space_before = Pt(8)
        c_lbl = cp.add_run()
        c_lbl.text = "▸ " + ch_lbl + "\n"
        c_lbl.font.name = FONT_MONO
        c_lbl.font.size = Pt(9)
        c_lbl.font.color.rgb = COLOR_TEXT_MUTED
        c_val = cp.add_run()
        c_val.text = ch_val + "\n"
        c_val.font.name = FONT_BODY
        c_val.font.bold = True
        c_val.font.size = Pt(9.5)
        c_val.font.color.rgb = COLOR_WHITE
        c_note = cp.add_run()
        c_note.text = ch_note
        c_note.font.name = FONT_BODY
        c_note.font.size = Pt(8.5)
        c_note.font.color.rgb = COLOR_TEXT_SECONDARY

    # Save to assets/Alief_Sukma_Dewanta_Portfolio.pptx
    output_path = "assets/Alief_Sukma_Dewanta_Portfolio.pptx"
    prs.save(output_path)
    print(f"[SUCCESS] Generated 13-slide executive portfolio presentation deck: {output_path}")

if __name__ == "__main__":
    create_portfolio_deck()
