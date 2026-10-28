import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_report():
    doc = Document()

    # Set standard page margins (1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25) # slightly larger for binding
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    # Set base styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    normal_style.paragraph_format.line_spacing = 1.25
    normal_style.paragraph_format.space_after = Pt(6)

    def add_title(text, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=12):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
        return p

    def add_h1(text, page_break=False):
        if page_break:
            doc.add_page_break()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0x37, 0x41, 0x51)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)
        return p

    def add_body(text, bold_prefix=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.25
        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.bold = True
        p.add_run(text)
        return p

    def add_bullet(text, bold_prefix=None):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.bold = True
        p.add_run(text)
        return p

    print("Generating Cover Page...")
    # ==================== COVER PAGE (Single Clean Page) ====================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after = Pt(16)
    r = p.add_run("CUSTOMER MANAGEMENT AND PAYMENT TRACKING SYSTEM\n")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(0x11, 0x18, 0x27)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(14)
    r2 = p2.add_run("A Project Report Submitted by\n")
    r2.font.size = Pt(12)
    r2_name = p2.add_run("SWEETHA T\n")
    r2_name.bold = True
    r2_name.font.size = Pt(14)
    p2.add_run("(Reg. No.: 2428B0358)\n\nUnder the Guidance of\n")
    r2_guide = p2.add_run("Dr. K. Jayanthi Vagini MCA, M.Phil., NET, Ph.D.\n")
    r2_guide.bold = True
    p2.add_run("Assistant Professor, Department of Computer Science with Data Analytics\n\n")

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(18)
    p3.add_run("In partial fulfillment of the requirements for the Award of the Degree of\n")
    r3_deg = p3.add_run("Bachelor of Science in Computer Science with Data Analytics\n")
    r3_deg.bold = True
    r3_deg.font.size = Pt(13)

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(36)
    p4.paragraph_format.space_after = Pt(4)
    r4_dept = p4.add_run("DEPARTMENT OF COMPUTER SCIENCE WITH DATA ANALYTICS\n")
    r4_dept.bold = True
    r4_dept.font.size = Pt(13)
    r4_col = p4.add_run("KPR COLLEGE OF ARTS SCIENCE AND RESEARCH\n")
    r4_col.bold = True
    r4_col.font.size = Pt(14)
    p4.add_run("(Approved by Government of Tamil Nadu & Affiliated to Bharathiar University, Coimbatore)\n")
    p4.add_run("Arasur, Coimbatore – 641407\n\n")
    r4_date = p4.add_run("MARCH 2026")
    r4_date.bold = True
    r4_date.font.size = Pt(12)

    doc.add_page_break()

    # ==================== CERTIFICATE ====================
    print("Generating Certificate...")
    add_title("CERTIFICATE")
    p_cert = doc.add_paragraph()
    p_cert.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cert.paragraph_format.space_before = Pt(18)
    p_cert.paragraph_format.space_after = Pt(16)
    p_cert.paragraph_format.line_spacing = 1.3
    p_cert.add_run("This is to certify that this project report entitled ")
    r_title = p_cert.add_run("“CUSTOMER MANAGEMENT AND PAYMENT TRACKING SYSTEM” ")
    r_title.bold = True
    p_cert.add_run("is a bonafide record of project work done by ")
    r_name = p_cert.add_run("SWEETHA T (Reg. No.: 2428B0358) ")
    r_name.bold = True
    p_cert.add_run("in partial fulfillment of the requirements for the award of the degree of ")
    r_deg = p_cert.add_run("Bachelor of Science in Computer Science with Data Analytics ")
    r_deg.bold = True
    p_cert.add_run("from Bharathiar University, Coimbatore, carried out under my supervision during the academic year 2025–2026.")

    doc.add_paragraph().paragraph_format.space_before = Pt(30)
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    sig_table.rows[0].cells[0].width = Inches(3.2)
    sig_table.rows[0].cells[1].width = Inches(3.2)
    sig_table.rows[0].cells[0].paragraphs[0].text = "_________________________\nFaculty Guide"
    sig_table.rows[0].cells[1].paragraphs[0].text = "_________________________\nHead of the Department"
    sig_table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p_viva = doc.add_paragraph()
    p_viva.paragraph_format.space_before = Pt(30)
    p_viva.add_run("Submitted for Viva-Voce Examination held on: ___________________")

    doc.add_paragraph().paragraph_format.space_before = Pt(30)
    exam_table = doc.add_table(rows=2, cols=2)
    exam_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    exam_table.rows[0].cells[0].width = Inches(3.2)
    exam_table.rows[0].cells[1].width = Inches(3.2)
    exam_table.rows[0].cells[0].paragraphs[0].text = "_________________________\nInternal Examiner"
    exam_table.rows[0].cells[1].paragraphs[0].text = "_________________________\nExternal Examiner"
    exam_table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p_seal = doc.add_paragraph()
    p_seal.paragraph_format.space_before = Pt(40)
    p_seal.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_seal.add_run("Office Seal")

    doc.add_page_break()

    # ==================== DECLARATION ====================
    print("Generating Declaration...")
    add_title("DECLARATION")
    p_decl = doc.add_paragraph()
    p_decl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_decl.paragraph_format.space_before = Pt(18)
    p_decl.paragraph_format.space_after = Pt(16)
    p_decl.paragraph_format.line_spacing = 1.3
    p_decl.add_run("I, ")
    r_dec_name = p_decl.add_run("SWEETHA T (Reg. No.: 2428B0358), ")
    r_dec_name.bold = True
    p_decl.add_run("hereby declare that the project entitled ")
    r_dec_title = p_decl.add_run("“CUSTOMER MANAGEMENT AND PAYMENT TRACKING SYSTEM” ")
    r_dec_title.bold = True
    p_decl.add_run("submitted to ")
    r_dec_col = p_decl.add_run("KPR COLLEGE OF ARTS SCIENCE AND RESEARCH, ")
    r_dec_col.bold = True
    p_decl.add_run("in partial fulfillment for the award of the degree of ")
    r_dec_deg = p_decl.add_run("B.Sc. COMPUTER SCIENCE WITH DATA ANALYTICS, ")
    r_dec_deg.bold = True
    p_decl.add_run("is a record of original project work done by me during the academic year 2025–2026 under the supervision and guidance of ")
    r_dec_guide = p_decl.add_run("Dr. K. Jayanthi Vagini MCA, M.Phil., NET, Ph.D., Assistant Professor, ")
    r_dec_guide.bold = True
    p_decl.add_run("Department of Computer Science with Data Analytics, KPR College of Arts Science and Research, and that it has not previously formed the basis for the award of any Degree, Diploma, Associate-ship, Fellowship, or other similar title to any candidate of any University.")

    p_dec_sign = doc.add_paragraph()
    p_dec_sign.paragraph_format.space_before = Pt(60)
    p_dec_sign.paragraph_format.space_after = Pt(20)
    p_dec_sign.add_run("Date  : ______________\nPlace : Coimbatore\n\n\n")
    p_dec_sign_rt = doc.add_paragraph()
    p_dec_sign_rt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sig = p_dec_sign_rt.add_run("_________________________\nSignature of the Candidate\n(SWEETHA T)")
    r_sig.bold = True

    doc.add_page_break()

    # ==================== ACKNOWLEDGEMENT ====================
    print("Generating Acknowledgement...")
    add_title("ACKNOWLEDGEMENT")
    add_body("I wish to extend my profound gratitude to Dr. K. P. Ramasamy, Chairman of the KPR Group, for graciously providing the institutional facilities and resources essential for the completion of this project.")
    add_body("I would like to express our gratitude to Mrs. Gayathri Anandakrishnan, Secretary of the KPR Institutions, for her visionary leadership and for fostering an academic environment that prioritizes student excellence and research.")
    add_body("My sincere thanks to Dr. P. Geetha, Principal of KPR College of Arts Science and Research, for her constant encouragement and steadfast support throughout my academic tenure.")
    add_body("I am equally grateful to Dr. K. Pradeepa, Dean of the School of Intelligent Systems and Data Science, for her valuable administrative oversight and the motivation she provided during the various stages of this project.")
    add_body("I am deeply indebted to Dr. G. Satyavathy, Head of the Department of Computer Science with Data Analytics, for her leadership and for ensuring a conducive environment that allowed me to complete my work effectively.")
    add_body("A special word of heartfelt thanks is reserved for my esteemed guide, Dr. K. Jayanthi Vagini, Assistant Professor of Computer Science with Data Analytics. I have been incredibly fortunate to benefit from her scholarly expertise, patience, and insightful guidance, which were pivotal in shaping this study and its data analytical implementations.")
    add_body("I also appreciate the collective support, constructive feedback, and encouragement provided by the entire faculty and staff of the Computer Science with Data Analytics Department who contributed to my learning experience.")
    add_body("I express my sincere thanks to Nilacable Fibre Networks, Neelambur, Coimbatore, for providing real-world operational domain insights, sample statement data, and field requirements that formed the core foundation of this project.")
    add_body("Finally, I offer my humble prayers to the Almighty God for granting me the strength, sound health, and mental perseverance required to successfully conclude this project journey.")

    p_ack_sign = doc.add_paragraph()
    p_ack_sign.paragraph_format.space_before = Pt(30)
    p_ack_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_ack_name = p_ack_sign.add_run("SWEETHA T\n(Reg. No.: 2428B0358)")
    r_ack_name.bold = True

    doc.add_page_break()

    # ==================== TABLE OF CONTENTS ====================
    print("Generating Table of Contents...")
    add_title("TABLE OF CONTENTS")
    
    toc_data = [
        ("ACKNOWLEDGEMENT", "iii"),
        ("DECLARATION", "iv"),
        ("CERTIFICATE", "v"),
        ("SYNOPSIS", "vi"),
        ("1. INTRODUCTION", "1"),
        ("    1.1 Organizational Profile", "4"),
        ("    1.2 System Specification", "6"),
        ("        1.2.1 Hardware Configuration", "8"),
        ("        1.2.2 Software Configuration", "10"),
        ("2. SYSTEM STUDY", "16"),
        ("    2.1 Existing System", "17"),
        ("        2.1.1 Drawbacks of the Existing System", "18"),
        ("    2.2 Proposed System", "20"),
        ("        2.2.1 Core Features of the Proposed System", "22"),
        ("3. SYSTEM DESIGN AND DEVELOPMENT", "28"),
        ("    3.1 File Design and Architecture", "28"),
        ("    3.2 Input Design", "31"),
        ("    3.3 Output Design", "37"),
        ("    3.4 Database Design", "44"),
        ("    3.5 System Development: Core Modules", "47"),
        ("        3.5.1 Admin Login & Access Control Module", "48"),
        ("        3.5.2 Customer Management Module", "50"),
        ("        3.5.3 Payment Management Module", "52"),
        ("        3.5.4 Bank Statement Processing Module", "54"),
        ("        3.5.5 OCR Processing Module", "56"),
        ("        3.5.6 Payment Matching & Reconciliation Module", "57"),
        ("        3.5.7 Dashboard & Visual Analytics Module", "59"),
        ("        3.5.8 Real-Time Search and Filtering Module", "61"),
        ("        3.5.9 Operational Command Center (Left Sidebar)", "63"),
        ("        3.5.10 Report Generation Module", "64"),
        ("        3.5.11 Database Management Module", "66"),
        ("        3.5.12 Module Integration", "67"),
        ("4. SOFTWARE TESTING AND IMPLEMENTATION", "69"),
        ("    4.1 Testing Objectives", "69"),
        ("    4.2 Testing Approach & Methodology", "70"),
        ("    4.3 Test Case Execution Summary", "72"),
        ("5. CONCLUSION AND FUTURE ENHANCEMENTS", "75"),
        ("    5.1 Conclusion", "75"),
        ("    5.2 Future Enhancements", "77"),
        ("BIBLIOGRAPHY", "80"),
        ("APPENDICES", "83"),
        ("    Appendix A: Data Flow Diagrams (Context & Level 1 DFD)", "83"),
        ("    Appendix B: Table Structures and Data Dictionary", "87"),
        ("    Appendix C: Core System Coding (Backend & Analytics)", "91"),
        ("    Appendix D: Sample Input Files (Bank Statement & Excel)", "97"),
        ("    Appendix E: Sample Outputs & Analytics Visualizations", "99"),
    ]

    t_toc = doc.add_table(rows=len(toc_data) + 1, cols=2)
    t_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_toc.autofit = False
    t_toc.rows[0].cells[0].width = Inches(5.5)
    t_toc.rows[0].cells[1].width = Inches(1.2)
    t_toc.rows[0].cells[0].paragraphs[0].text = "CONTENTS"
    t_toc.rows[0].cells[0].paragraphs[0].runs[0].bold = True
    t_toc.rows[0].cells[1].paragraphs[0].text = "PAGE NO."
    t_toc.rows[0].cells[1].paragraphs[0].runs[0].bold = True
    t_toc.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_cell_background(t_toc.rows[0].cells[0], "E5E7EB")
    set_cell_background(t_toc.rows[0].cells[1], "E5E7EB")

    for i, (item, pg) in enumerate(toc_data):
        row = t_toc.rows[i + 1]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(5.5)
        c1.width = Inches(1.2)
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(2)
        p0.paragraph_format.line_spacing = 1.15
        r_item = p0.add_run(item)
        if not item.startswith("    "):
            r_item.bold = True
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(2)
        p1.paragraph_format.line_spacing = 1.15
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_pg = p1.add_run(pg)
        if not item.startswith("    "):
            r_pg.bold = True

    doc.add_page_break()

    # ==================== SYNOPSIS ====================
    print("Generating Synopsis...")
    add_title("SYNOPSIS")
    add_body("The Customer Management and Payment Tracking System is a comprehensive web-based application designed to modernize and simplify customer administration, subscription tracking, and monthly payment reconciliation for local fiber network service providers. Many small and medium-scale broadband operators continue to record customer profiles and track recurring payments manually using notebooks, physical ledgers, or disconnected spreadsheets. These manual approaches are inherently time-consuming, prone to human calculation errors, vulnerable to record misplacement, and increasingly difficult to sustain as subscriber volume expands.")
    add_body("To address these challenges, this project provides a centralized, secure digital platform developed using Python and the lightweight Flask web framework. The user interface is engineered with responsive HTML5, modern CSS3 styling, JavaScript, and dynamic Jinja2 templating. An SQLite relational database acts as the core repository, systematically organizing customer profiles, tariff plans, subscription states, bank transaction records, and historical payment logs.")
    add_body("A pivotal innovation of the system is automated bank statement reconciliation. Broadband operators receive monthly bank statements containing hundreds of customer UPI and NEFT/IMPS credits. The system automatically ingests bank statements in PDF format using pdfplumber. For scanned or image-based statement documents, an Optical Character Recognition (OCR) pipeline powered by Pytesseract, pdf2image, and Pillow extracts transaction data. A multi-parameter matching algorithm evaluates UPI IDs, subscriber contact numbers, box IDs, and billing amounts to automatically reconcile statement rows against customer accounts, flagging discrepancies into an interactive verification queue.")
    add_body("To empower administrative decision-making with data analytics, the platform features an advanced Dashboard and Visual Analytics module powered by Chart.js. The dashboard delivers real-time operational intelligence through six interactive visualizations: monthly collection history trends, plan tariff tier distributions, payment realization rates, outstanding collection dues, weekly payment velocity, and subscriber connection breakdowns. Additionally, an instant multi-criteria Search and Filtering module provides zero-latency customer lookup across names, phone numbers, box identifiers, and payment states across historical billing cycles.")
    add_body("The platform substantially reduces manual overhead, eliminates duplicate entries, accelerates payment verification from hours to seconds, and gives network operators complete financial visibility. Designed with scalability and modularity, the system provides a robust technological foundation for fiber network businesses while supporting future extensions such as automated WhatsApp/SMS payment reminders and payment gateway integration.")

    doc.add_page_break()

    # ==================== CHAPTER 1 ====================
    print("Generating Chapter 1: Introduction...")
    add_h1("1. INTRODUCTION", page_break=False)
    add_body("The rapid expansion of telecommunications and digital infrastructure has transformed how service-oriented enterprises manage customer relationships and monitor financial transactions. Broadband internet access has evolved into an essential utility for households and businesses alike, leading to a surge in local fiber-to-the-home (FTTH) network providers. However, despite managing high-technology network hardware, many local cable and fiber operators still rely on traditional manual record-keeping methodologies such as paper registers, physical receipts, and basic desktop spreadsheets to track their subscriber base and collect monthly subscription fees.")
    add_body("Manual tracking poses severe operational bottlenecks. As the subscriber base scales into hundreds or thousands of connections, tracking payment deadlines, identifying defaulters, and reconciling bank credits becomes increasingly error-prone. Network administrators must manually cross-examine lengthy bank account statements against disconnected customer lists to determine who has paid and who remains overdue. This manual matching process consumes substantial administrative hours every billing cycle and frequently leads to duplicate receipts, overlooked payments, and delayed revenue recovery.")
    add_body("To overcome these systemic inefficiencies, the Customer Management and Payment Tracking System is proposed as a centralized, automated web application tailored specifically for local fiber network service providers. Developed using Python, Flask, SQLite, and modern frontend technologies, the application provides an integrated digital workspace where administrators can manage customer subscriptions, ingest bank statements, automate payment matching through rule-based algorithms and OCR, track revenue velocity through visual analytics dashboards, and instantaneously search and filter records.")

    add_h2("1.1 Organizational Profile")
    add_body("Nilacable Fibre Networks is a prominent local internet and cable network service provider situated in Neelambur, Coimbatore, Tamil Nadu. The enterprise delivers high-speed fiber-optic broadband and digital television connectivity to residential, commercial, and educational clients across the eastern suburban corridor of Coimbatore.")
    add_body("As a customer-centric telecommunications provider, Nilacable Fibre Networks oversees an extensive spectrum of daily operational tasks:")
    add_bullet("Customer Onboarding & Connection Provisioning: Registering new subscribers, allocating physical Fiber Distribution Box IDs (e.g., Box B101, B102), and activating network ports.")
    add_bullet("Tariff Plan & Subscription Management: Administering diverse monthly broadband plans ranging from entry-level tariffs (₹199/month, ₹250/month) to high-speed commercial unlimited packages (₹299/month, ₹499/month, ₹799/month).")
    add_bullet("Monthly Payment Collection & Reconciliation: Tracking recurring billing cycles, identifying paid versus unpaid accounts, verifying UPI/bank payments against statements, and managing customer dues.")
    add_bullet("Customer Inquiries & Technical Support: Resolving subscriber billing disputes, verifying transaction references, and ensuring uninterrupted internet uptime.")
    add_body("Previously, Nilacable Fibre Networks maintained customer records across physical notebooks and static Microsoft Excel workbooks. As their customer network expanded significantly across Neelambur, manual cross-referencing between customer notebooks and monthly bank statements created severe delays and tracking errors. Implementing this digital management platform equips the organization with an enterprise-grade solution that streamlines customer workflows, safeguards data integrity, and enhances collection productivity.")

    add_h2("1.2 System Specification")
    add_body("The system specifications define the operational, computational, and software prerequisites required to develop, host, and execute the Customer Management and Payment Tracking System reliably.")

    add_h3("1.2.1 Hardware Configuration")
    add_body("The development and operational deployment were executed and rigorously benchmarked on a high-performance workstation configured to effortlessly handle OCR image processing, PDF token parsing, and in-memory analytics computation:")
    
    hw_table = doc.add_table(rows=7, cols=2)
    hw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hw_table.autofit = False
    hw_table.rows[0].cells[0].width = Inches(2.5)
    hw_table.rows[0].cells[1].width = Inches(4.0)
    hw_headers = ["HARDWARE COMPONENT", "SPECIFICATION DETAILS"]
    for idx, h in enumerate(hw_headers):
        cell = hw_table.rows[0].cells[idx]
        cell.paragraphs[0].text = h
        cell.paragraphs[0].runs[0].bold = True
        set_cell_background(cell, "E5E7EB")
    
    hw_specs = [
        ("Device / Machine Name", "ASUS TUF Gaming F15 Workstation Laptop"),
        ("Microprocessor", "11th Generation Intel® Core™ i5-11400H @ 2.70GHz (12 CPUs)"),
        ("System Memory (RAM)", "16.0 GB DDR4 High-Speed RAM (15.7 GB usable)"),
        ("Hard Disk / Storage", "512 GB NVMe M.2 High-Speed Solid State Drive (SSD)"),
        ("Input Devices", "Standard Ergonomic Keyboard, High-Precision Optical Mouse"),
        ("Display Monitor", "15.6-inch Full HD (1920 x 1080) LED Display")
    ]
    for r_idx, (c0, c1) in enumerate(hw_specs):
        row = hw_table.rows[r_idx + 1]
        row.cells[0].width = Inches(2.5)
        row.cells[1].width = Inches(4.0)
        row.cells[0].paragraphs[0].text = c0
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[1].paragraphs[0].text = c1

    add_h3("1.2.2 Software Configuration")
    add_body("The software technology stack was selected for its high reliability, open-source flexibility, rapid prototyping capabilities, and robust data analytics ecosystem:")

    sw_table = doc.add_table(rows=10, cols=2)
    sw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sw_table.autofit = False
    sw_table.rows[0].cells[0].width = Inches(2.5)
    sw_table.rows[0].cells[1].width = Inches(4.0)
    sw_headers = ["SOFTWARE ENVIRONMENT", "TECHNOLOGY / TOOL USED"]
    for idx, h in enumerate(sw_headers):
        cell = sw_table.rows[0].cells[idx]
        cell.paragraphs[0].text = h
        cell.paragraphs[0].runs[0].bold = True
        set_cell_background(cell, "E5E7EB")

    sw_specs = [
        ("Operating System", "Microsoft Windows 11 64-bit Professional"),
        ("Primary Backend Language", "Python 3.10+ / Python 3.12 64-bit"),
        ("Web Application Framework", "Flask 3.0+ (Lightweight WSGI Microframework)"),
        ("Database Engine", "SQLite3 (Relational Embedded Database)"),
        ("Frontend Presentation Layer", "HTML5, CSS3, JavaScript (ES6+), Jinja2 Templating"),
        ("Data Analytics & Manipulation", "Pandas, NumPy, OpenPyXL"),
        ("PDF Processing & OCR", "pdfplumber, pytesseract, pdf2image, Pillow (PIL)"),
        ("Data Visualization Engine", "Chart.js 4.x (Client-side HTML5 Canvas Charting)"),
        ("Integrated Development Environment", "Visual Studio Code (VS Code) with Python Extensions")
    ]
    for r_idx, (c0, c1) in enumerate(sw_specs):
        row = sw_table.rows[r_idx + 1]
        row.cells[0].width = Inches(2.5)
        row.cells[1].width = Inches(4.0)
        row.cells[0].paragraphs[0].text = c0
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[1].paragraphs[0].text = c1

    doc.add_page_break()

    # ==================== CHAPTER 2 ====================
    print("Generating Chapter 2: System Study...")
    add_h1("2. SYSTEM STUDY", page_break=False)
    add_body("System study is a foundational phase in the software engineering lifecycle dedicated to rigorously investigating the operational environment, analyzing legacy procedures, isolating systemic pain points, and establishing clear technical objectives for the proposed digital solution.")

    add_h2("2.1 Existing System")
    add_body("In the existing operational framework at Nilacable Fibre Networks, customer subscription tracking and monthly fee collection were conducted using physical paper registers supplemented by independent Microsoft Excel spreadsheets. When a subscriber joined the network, their details (name, contact number, assigned distribution box, and plan cost) were handwritten in an account ledger. At the start of each calendar month, administrators manually compiled a billing checklist to keep track of incoming subscription dues.")
    add_body("When customers submitted subscription payments via Google Pay, PhonePe, Paytm, or direct bank IMPS/NEFT transfers, transactions appeared on the enterprise's bank statement. To verify these payments, the administrator had to open bank statement PDFs or printed paper sheets and line-by-line compare payment descriptions against customer phone numbers, UPI handles, or names. Once a match was identified, the administrator manually updated the Excel sheet, changing the customer status from unpaid to paid.")

    add_h3("2.1.1 Drawbacks of the Existing System")
    add_bullet("Excessive Manual Overhead: Spending tens of administrative hours every billing cycle manually cross-verifying statement entries against ledgers.", "• Time Inefficiency: ")
    add_bullet("Manual transcription errors, mismatched phone numbers, and transposed figures frequently corrupted subscriber billing records.", "• High Human Error Rate: ")
    add_bullet("Paper ledgers or disconnected Excel files stored on individual computers created isolated data silos lacking centralized data validation.", "• Lack of Centralization: ")
    add_bullet("UPI payment descriptions frequently truncate subscriber names or display third-party payer details, making manual association tedious and confusing.", "• Masked UPI Information: ")
    add_bullet("Administrators lacked instant visibility into total monthly expected collections versus actual realized revenues and overdue defaulters.", "• Absence of Real-Time Analytics: ")
    add_bullet("Physical documents risk physical loss, while unmanaged desktop spreadsheets can be accidentally deleted or overwritten without backups.", "• Vulnerability & Data Security: ")

    add_h2("2.2 Proposed System")
    add_body("The Customer Management and Payment Tracking System is engineered as an end-to-end web-based platform tailored specifically to eliminate the inefficiencies of manual bookkeeping. Built on Python, Flask, and SQLite, the platform establishes a structured, secure, and automated operational workflow for fiber broadband administrators.")
    add_body("The application integrates an intelligent statement processing engine capable of digesting raw bank statement PDFs and spreadsheet files. By parsing transactional metadata using regular expressions, string distance algorithms, and OCR image extraction, the system automatically correlates statement payments with subscriber records in SQLite, instantaneously updating payment statuses. To empower administrators, the platform introduces real-time search and filter controls across all modules alongside an analytical visualization dashboard that continuously tracks billing metrics.")

    add_h3("2.2.1 Core Features of the Proposed System")
    add_bullet("Centralized Customer Database: Complete CRUD (Create, Read, Update, Delete) management of subscriber profiles including Box IDs, contact numbers, assigned tariffs, and active/inactive connection states.", "i. ")
    add_bullet("Automated Bank Statement Processing: Upload digital or scanned PDF statements with automatic text parsing via pdfplumber and OCR text extraction via Pytesseract/pdf2image.", "ii. ")
    add_bullet("Intelligent Multi-Parameter Matching: Correlation algorithm that cross-evaluates UPI handles, phone numbers, exact/partial names, and subscription tariff amounts.", "iii. ")
    add_bullet("Reconciliation & Verification Workflow: Confirmed matches are tagged as 'Paid'; ambiguous or partial matches are queued into an interactive 'Needs Verification' screen for quick 1-click admin approval.", "iv. ")
    add_bullet("Real-Time Multi-Dimensional Search & Filtering: Instant zero-latency client-side search across Customer Name, Phone, Box ID, and UPI ID, coupled with dropdown filters for Connection Status, Payment Status, and Billing Month selection.", "v. ")
    add_bullet("Advanced Dashboard Visual Analytics: Six interactive Chart.js visualizations including Monthly Revenue Collection History, Subscription Plan Breakdown, Payment Status ratios, Weekly Collection velocity, Customer Connection breakdown, and 4-week Transaction volume trends.", "vi. ")
    add_bullet("Financial Realization & Outstanding Dues Tracking: Real-time metric cards computing exact collection realization percentages and outstanding pending debt for any billing cycle.", "vii. ")
    add_bullet("Unified Operational Command Center (Left Sidebar): Quick-action buttons (+ Customer, + Payment), categorized navigation (Main, Billing & Ops, System), live verification alert badges, and network status indicators.", "viii. ")
    add_bullet("Bulk Data Ingestion & Excel Import: Seamless batch onboarding of customer directories from Microsoft Excel (.xlsx) files via Pandas and OpenPyXL.", "ix. ")
    add_bullet("Security & Access Governance: Session-based administrator authentication, environment-variable password encryption, and SQL injection prevention via parameterized database queries.", "x. ")

    doc.add_page_break()

    # ==================== CHAPTER 3 ====================
    print("Generating Chapter 3: System Design...")
    add_h1("3. SYSTEM DESIGN AND DEVELOPMENT", page_break=False)
    add_body("System design translates operational requirements into technical specifications, detailing architectural blueprints, data structures, input/output interfaces, database schemas, and modular system workflows.")

    add_h2("3.1 File Design and Architecture")
    add_body("The system follows a clean modular architectural pattern separating presentation layers, application routing, data ingestion pipelines, and database persistence:")
    add_bullet("app.py: The central application controller containing Flask routing logic, session authorization, database connection handlers, statistical calculations, and statement parsing pipelines.", "• Application Controller: ")
    add_bullet("database/fiber.db: SQLite relational database holding structured customer, payment, and reconciliation records.", "• Database Storage: ")
    add_bullet("templates/: Jinja2 HTML5 presentation templates including dashboard.html, view_customers.html, payments.html, reports.html, upload_statement.html, upload_customers.html, settings.html, add_customer.html, edit_customer.html, and add_payment.html.", "• Presentation Templates: ")
    add_bullet("static/: Static web assets including the corporate brand logo (nila_logo.png), stylesheets, and client-side assets.", "• Static Assets: ")
    add_bullet("tests/: Automated unit and integration test suites validating route integrity, database persistence, and analytical calculation accuracy.", "• Test Suite: ")

    add_h2("3.2 Input Design")
    add_body("Input design ensures that administrative data is captured accurately, efficiently, and with comprehensive field validation to prevent data corruption:")
    add_bullet("Admin Authentication: Secure username and password inputs validated against server-side environment credentials.", "1. ")
    add_bullet("Customer Registration: Form capturing Box ID, Customer Name, 10-digit Phone Number, Subscription Amount, Physical Connection ID, and UPI handle with client-side constraint checks.", "2. ")
    add_bullet("Manual Payment Recording: Captures Transaction ID, Customer Phone, Payment Amount, Payment Date, and optional UPI reference.", "3. ")
    add_bullet("Bank Statement File Ingestion: Drag-and-drop file upload supporting .pdf, .csv, and .txt files, optional PDF decryption passwords, and OCR fallback toggles for scanned documentation.", "4. ")
    add_bullet("Interactive Search & Filter Criteria: Live keystroke text inputs and dynamic dropdown selectors for immediate table querying without full page reloads.", "5. ")

    add_h2("3.3 Output Design")
    add_body("Output design presents operational findings in intuitive, actionable, and visually appealing formats:")
    add_bullet("Analytical Dashboard: Summary metric cards (Total Customers, Paid, Unpaid, Total Collection, Realization Rate, Outstanding Dues) accompanied by six Chart.js data visualizations.", "1. ")
    add_bullet("Filtered Customer Grid: Real-time table displaying Box IDs, subscriber names, contact information, plan fees, payment dates, connection status badges, and inline action triggers.", "2. ")
    add_bullet("Payment Audit Ledger: Tabular transaction histories displaying Txn IDs, customer names, payment dates, amounts, status badges (Paid, Needs Verification, Unidentified), and verification buttons.", "3. ")
    add_bullet("Reconciliation Summaries: Categorized reporting showing successfully matched transactions, newly marked paid customers, and flagged ambiguous entries requiring review.", "4. ")

    add_h2("3.4 Database Design")
    add_body("The SQLite database schema is normalized to ensure data consistency, enforce referential integrity, eliminate redundancy, and maximize query performance.")
    
    add_h3("3.4.1 Customer Table Schema (`customer`)")
    add_body("Stores subscriber profile records, assigned network boxes, tariff rates, and current billing statuses:")
    
    t_cust = doc.add_table(rows=10, cols=4)
    t_cust.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_cust.autofit = False
    for i, w in enumerate([1.5, 1.2, 1.3, 2.5]):
        t_cust.rows[0].cells[i].width = Inches(w)
    
    cust_cols = ["COLUMN NAME", "DATA TYPE", "CONSTRAINTS", "DESCRIPTION"]
    for idx, c in enumerate(cust_cols):
        cell = t_cust.rows[0].cells[idx]
        cell.paragraphs[0].text = c
        cell.paragraphs[0].runs[0].bold = True
        set_cell_background(cell, "E5E7EB")

    cust_fields = [
        ("customer_id", "INTEGER", "PRIMARY KEY AUTO", "Unique internal customer identification ID"),
        ("box_id", "TEXT", "UNIQUE, NOT NULL", "Assigned physical fiber distribution box identifier (e.g. B101)"),
        ("name", "TEXT", "NOT NULL", "Full legal name of the registered subscriber"),
        ("phone", "TEXT", "NOT NULL", "Primary 10-digit mobile contact number"),
        ("phone_number", "TEXT", "NOT NULL", "Secondary normalized contact number for lookups"),
        ("subscription", "REAL", "NOT NULL", "Standard monthly subscription amount in INR (e.g. 250.0)"),
        ("status", "TEXT", "NOT NULL", "Physical network connection state ('active' / 'inactive')"),
        ("paid_status", "TEXT", "DEFAULT 'unpaid'", "Current billing cycle payment status ('paid' / 'unpaid')"),
        ("upi_id", "TEXT", "NULLABLE", "Associated UPI VPA handle for payment correlation (e.g. ravi@okaxis)")
    ]
    for r_idx, row_data in enumerate(cust_fields):
        row = t_cust.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            row.cells[c_idx].width = Inches([1.5, 1.2, 1.3, 2.5][c_idx])
            row.cells[c_idx].paragraphs[0].text = val
            if c_idx == 0:
                row.cells[c_idx].paragraphs[0].runs[0].bold = True

    add_h3("3.4.2 Payment Table Schema (`payment`)")
    add_body("Maintains every financial transaction recorded manually or extracted automatically from statements:")

    t_pay = doc.add_table(rows=10, cols=4)
    t_pay.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_pay.autofit = False
    for idx, c in enumerate(cust_cols):
        cell = t_pay.rows[0].cells[idx]
        cell.paragraphs[0].text = c
        cell.paragraphs[0].runs[0].bold = True
        set_cell_background(cell, "E5E7EB")

    pay_fields = [
        ("payment_id", "INTEGER", "PRIMARY KEY AUTO", "Unique sequential transaction identifier"),
        ("customer_id", "INTEGER", "FOREIGN KEY", "References customer(customer_id) for matched payments"),
        ("phone", "TEXT", "NULLABLE", "Phone number associated with payment or extracted from description"),
        ("amount", "REAL", "NOT NULL", "Credit transaction amount in INR (e.g. 250.0)"),
        ("payment_date", "TEXT", "NOT NULL", "Transaction execution date (YYYY-MM-DD format)"),
        ("month", "TEXT", "NOT NULL", "Accounting billing month for revenue allocation (YYYY-MM format)"),
        ("upi_id", "TEXT", "NULLABLE", "Extracted UPI address of remitter"),
        ("transaction_id", "TEXT", "UNIQUE", "Unique banking reference number or generated hash"),
        ("status", "TEXT", "NOT NULL", "Reconciliation status: 'Paid', 'Needs Verification', 'Unidentified'")
    ]
    for r_idx, row_data in enumerate(pay_fields):
        row = t_pay.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            row.cells[c_idx].width = Inches([1.5, 1.2, 1.3, 2.5][c_idx])
            row.cells[c_idx].paragraphs[0].text = val
            if c_idx == 0:
                row.cells[c_idx].paragraphs[0].runs[0].bold = True

    add_h2("3.5 System Development: Core Modules")
    add_body("The software architecture is divided into specialized functional modules that interoperate seamlessly to deliver automated customer management and payment tracking:")

    add_h3("3.5.1 Admin Login & Access Control Module")
    add_body("Provides secure gateway authentication protecting proprietary subscriber and financial records. Incorporates Flask session cookie encryption, server-side credential verification, and brute-force prevention. Unauthorized URL requests are intercepted and redirected to the login view.")

    add_h3("3.5.2 Customer Management Module")
    add_body("Enables comprehensive subscriber management. Administrators can register new customers, view detailed profile cards, edit contact/box assignments, and safely deactivate accounts. Features automated duplicate detection to prevent assigning existing Box IDs or phone numbers.")

    add_h3("3.5.3 Payment Management Module")
    add_body("Maintains the primary financial ledger. Supports manual payment entry for direct cash or bank transfers, displays transaction audit logs, and enables administrators to verify flagged payments with a single click.")

    add_h3("3.5.4 Bank Statement Processing Module")
    add_body("Automates statement ingestion. Administrators upload bank statements in PDF, CSV, or text formats. The module utilizes pdfplumber to extract table cells, dates, credit amounts, and narration text without manual data entry.")

    add_h3("3.5.5 OCR Processing Module")
    add_body("Ensures compatibility with scanned or photographed bank statements. When selectable text is absent, pdf2image converts statement pages into high-resolution images, Pillow optimizes image contrast, and Pytesseract performs optical character recognition to extract transaction strings.")

    add_h3("3.5.6 Payment Matching & Reconciliation Module")
    add_body("The intelligent core of the application. Employs a multi-parameter matching algorithm that cross-checks extracted transaction descriptions against customer phone numbers, UPI VPAs, subscriber names, and exact tariff values. Matches with high confidence are automatically marked 'Paid'; ambiguous payments are queued into 'Needs Verification' to ensure human oversight.")

    add_h3("3.5.7 Dashboard & Visual Analytics Module")
    add_body("Empowers administrators with high-impact visual data analytics powered by Chart.js. Includes:")
    add_bullet("Monthly Collection History: Multi-month bar chart comparing revenue trends over time.", "1. ")
    add_bullet("Subscription Plan Breakdown: Doughnut chart visualizing subscriber distribution by tariff tier (₹250, ₹299, ₹499, etc.).", "2. ")
    add_bullet("Financial Realization Rate: Metric card computing exact collection recovery percentage against total potential billing.", "3. ")
    add_bullet("Outstanding Dues Tracker: Real-time calculation of uncollected balances pending recovery.", "4. ")
    add_bullet("Payment Status & Velocity Charts: Pie and bar charts tracking paid/unpaid ratios and day-by-day weekly collection speed.", "5. ")

    add_h3("3.5.8 Real-Time Search and Filtering Module")
    add_body("Delivers instant, zero-latency data querying. On the Customer page, administrators can type in the search bar to filter rows in real time across Name, Phone, Box ID, and UPI ID simultaneously, while dropdown filters isolate Active/Inactive accounts and Paid/Unpaid states across different billing months without full page refreshes. Similar real-time filters protect the Payments view.")

    add_h3("3.5.9 Operational Command Center (Left Sidebar)")
    add_body("Elevates navigation into an organized command center with categorized menus (MAIN, BILLING & OPS, SYSTEM), one-click shortcut buttons (+ Customer, + Payment), live notification badges displaying pending verification counts, and an active system health footer.")

    add_h3("3.5.10 Report Generation Module")
    add_body("Prepares aggregated summaries of monthly collections, transaction audit trails, active subscriber directories, and outstanding defaulters for operational and accounting documentation.")

    add_h3("3.5.11 Database Management Module")
    add_body("Governs SQLite database transactions, automatic schema migrations, connection pooling, and data sanitization routines to maintain ACID (Atomicity, Consistency, Isolation, Durability) database compliance.")

    add_h3("3.5.12 Module Integration")
    add_body("All eleven modules communicate seamlessly through Flask view routes, shared database helper functions, and responsive Jinja2 templates, ensuring unified data flow from statement ingestion to dashboard visualization.")

    doc.add_page_break()

    # ==================== CHAPTER 4 ====================
    print("Generating Chapter 4: Testing...")
    add_h1("4. SOFTWARE TESTING AND IMPLEMENTATION", page_break=False)
    add_body("Software testing validates that every software module performs reliably according to functional specifications, handles invalid inputs gracefully, and maintains data integrity under real-world operating conditions.")

    add_h2("4.1 Testing Objectives")
    add_bullet("Verify that all user authentication, customer management, and payment logging features function correctly.", "• Functional Verification: ")
    add_bullet("Ensure invalid inputs (e.g. malformed phone numbers, non-numeric amounts, duplicate Box IDs) trigger appropriate error feedback.", "• Robust Input Validation: ")
    add_bullet("Confirm that regular expressions and OCR extract transactions from diverse digital and scanned statement formats accurately.", "• Statement Parsing Accuracy: ")
    add_bullet("Ensure that payments are allocated to the correct subscriber and ambiguous rows are appropriately flagged.", "• Reconciliation Precision: ")
    add_bullet("Verify that real-time JavaScript search and dropdown filters filter tabular records instantaneously without DOM corruption.", "• Search & Filter Responsiveness: ")

    add_h2("4.2 Testing Approach & Methodology")
    add_body("The project implemented a rigorous testing pipeline combining automated unit tests using Python's `unittest` framework and systematic manual user acceptance testing across Google Chrome and Microsoft Edge browsers. Both unit test suites (`test_verify.py` and `test_dashboard_and_filters.py`) executed flawlessly with zero failures.")

    add_h2("4.3 Test Case Execution Summary")
    add_body("The following comprehensive test matrix records the actual execution results of key functional scenarios across the application:")

    test_table = doc.add_table(rows=19, cols=5)
    test_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    test_table.autofit = False
    
    t_widths = [0.6, 1.8, 2.2, 1.4, 0.7]
    for i, w in enumerate(t_widths):
        test_table.rows[0].cells[i].width = Inches(w)

    test_headers = ["TC #", "MODULE", "TEST SCENARIO / INPUT", "EXPECTED OUTCOME", "STATUS"]
    for idx, h in enumerate(test_headers):
        cell = test_table.rows[0].cells[idx]
        cell.paragraphs[0].text = h
        cell.paragraphs[0].runs[0].bold = True
        set_cell_background(cell, "E5E7EB")

    test_data = [
        ("TC01", "Admin Login", "Submit valid admin credentials", "Session created, redirects to /dashboard", "PASSED"),
        ("TC02", "Admin Login", "Submit incorrect password", "Access denied, error message displayed", "PASSED"),
        ("TC03", "Customer Mgmt", "Add new customer with unique Box ID", "Record inserted into SQLite database", "PASSED"),
        ("TC04", "Customer Mgmt", "Add customer with duplicate Box ID", "Database constraint enforced, error shown", "PASSED"),
        ("TC05", "Customer Mgmt", "Update subscriber phone and tariff plan", "Updated details reflect immediately in table", "PASSED"),
        ("TC06", "Payment Mgmt", "Record manual payment for customer", "Payment logged, customer marked paid", "PASSED"),
        ("TC07", "Statement Upload", "Upload digital PDF bank statement", "pdfplumber extracts text & credits correctly", "PASSED"),
        ("TC08", "OCR Processing", "Upload scanned image-based PDF statement", "Pytesseract OCR extracts transaction rows", "PASSED"),
        ("TC09", "Reconciliation", "Match statement row with matching UPI ID", "Transaction linked, status set to 'Paid'", "PASSED"),
        ("TC10", "Reconciliation", "Process unmatched statement transaction", "Flagged as 'Needs Verification' for review", "PASSED"),
        ("TC11", "Verification", "Admin clicks 'Verify' on ambiguous payment", "Status updated to Paid, linked to customer", "PASSED"),
        ("TC12", "Dashboard", "Load /dashboard default view", "Displays 6 charts, KPIs, and current month", "PASSED"),
        ("TC13", "Dashboard", "Select '2026-07' from month dropdown", "Metrics & realization rate update for July", "PASSED"),
        ("TC14", "Search & Filter", "Type customer name 'Ravi' in search bar", "Instantly filters to matching row in real time", "PASSED"),
        ("TC15", "Search & Filter", "Filter by Connection Status = 'Inactive'", "Displays only inactive subscribers", "PASSED"),
        ("TC16", "Search & Filter", "Filter by Payment Status = 'Unpaid'", "Isolates unpaid subscribers for follow-up", "PASSED"),
        ("TC17", "Payments Filter", "Filter payments by status 'Needs Verification'", "Isolates transactions requiring review", "PASSED"),
        ("TC18", "Reports", "Generate monthly collection overview", "Summarizes total collections grouped by month", "PASSED")
    ]

    for r_idx, (t0, t1, t2, t3, t4) in enumerate(test_data):
        row = test_table.rows[r_idx + 1]
        for c_idx, val in enumerate([t0, t1, t2, t3, t4]):
            row.cells[c_idx].width = Inches(t_widths[c_idx])
            p = row.cells[c_idx].paragraphs[0]
            p.text = val
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.1
            if c_idx == 4:
                p.runs[0].bold = True

    doc.add_page_break()

    # ==================== CHAPTER 5 ====================
    print("Generating Chapter 5: Conclusion...")
    add_h1("5. CONCLUSION AND FUTURE ENHANCEMENTS", page_break=False)
    add_h2("5.1 Conclusion")
    add_body("The Customer Management and Payment Tracking System was successfully conceptualized, engineered, and validated to resolve the real-world operational challenges experienced by Nilacable Fibre Networks, Neelambur, Coimbatore. By transitioning from vulnerable, labor-intensive manual notebooks and spreadsheets to an automated web platform, the organization has achieved significant improvements in administrative productivity, financial transparency, and operational accuracy.")
    add_body("The application seamlessly integrates Python, Flask, and SQLite with automated PDF processing and Optical Character Recognition. The intelligent multi-parameter matching algorithm reduces bank statement reconciliation from hours of tedious manual cross-referencing to seconds of automated computation. Furthermore, the incorporation of advanced Data Analytics—including real-time multi-dimensional search, interactive status filters, revenue collection trend charts, tariff distribution breakdowns, and realization rate tracking—arms management with immediate visual insights to minimize payment defaults and maximize operational cash flow.")
    add_body("Comprehensive testing verified that the system operates reliably across diverse scenarios, validating data consistency, security, and interface responsiveness. The completed project delivers an efficient, dependable, and scalable management solution that perfectly aligns with modern broadband service requirements.")

    add_h2("5.2 Future Enhancements")
    add_body("While the application fulfills all core operational objectives, several prospective enhancements can further expand its capabilities in future iterations:")
    add_bullet("Automated Customer Notifications via WhatsApp & SMS: Implementing Twilio or WhatsApp Business APIs to automatically dispatch invoice alerts, payment receipts, and friendly due date reminders to subscribers.", "1. ")
    add_bullet("Online Payment Gateway Integration: Incorporating Razorpay, Cashfree, or UPI dynamic QR codes allowing subscribers to pay directly via customer portals with instantaneous database callbacks.", "2. ")
    add_bullet("AI/ML Enhanced Reconciliation: Utilizing machine learning text classification models to recognize unformatted transaction narrations and learn subscriber payment patterns over time.", "3. ")
    add_bullet("Predictive Churn & Revenue Analytics: Developing predictive machine learning models to forecast subscriber churn based on payment delays and recommend optimal tariff bundling.", "4. ")
    add_bullet("Role-Based Access Control (RBAC): Introducing multi-tier role permissions differentiating master administrators, collection billing staff, and network field technicians.", "5. ")
    add_bullet("Cloud Hosting & Automated Database Replication: Deploying the application container on cloud infrastructure (AWS / DigitalOcean) with automated encrypted daily backups to ensure high availability and disaster recovery.", "6. ")

    doc.add_page_break()

    # ==================== BIBLIOGRAPHY ====================
    print("Generating Bibliography...")
    add_h1("BIBLIOGRAPHY", page_break=False)
    
    bib_items = [
        "Sommerville, I. (2016). Software Engineering (10th ed.). Pearson Education, pp. 45–78.",
        "Pressman, R. S., & Maxim, B. R. (2015). Software Engineering: A Practitioner’s Approach (8th ed.). McGraw-Hill Education, pp. 112–150.",
        "Grinberg, M. (2018). Flask Web Development: Developing Web Applications with Python (2nd ed.). O’Reilly Media, pp. 25–68.",
        "Lutz, M. (2013). Learning Python (5th ed.). O’Reilly Media, pp. 210–265.",
        "Silberschatz, A., Korth, H. F., & Sudarshan, S. (2010). Database System Concepts (6th ed.). McGraw-Hill Education, pp. 120–160.",
        "McKinney, W. (2018). Python for Data Analysis: Data Wrangling with Pandas, NumPy, and Jupyter (2nd ed.). O’Reilly Media, pp. 100–145.",
        "Flanagan, D. (2020). JavaScript: The Definitive Guide (7th ed.). O’Reilly Media, pp. 75–102.",
        "Berners-Lee, T. (1999). Weaving the Web: The Original Design and Ultimate Destiny of the World Wide Web. HarperCollins Publishers, pp. 30–56.",
        "Python Software Foundation. (2024). Python Programming Language Official Documentation. Available: https://docs.python.org/3/",
        "Pallets Projects. (2024). Flask Web Development Documentation. Available: https://flask.palletsprojects.com/",
        "SQLite Development Team. (2024). SQLite Database Engine Architecture and Documentation. Available: https://www.sqlite.org/docs.html",
        "Chart.js Contributors. (2024). Chart.js: Open Source HTML5 Canvas Charting Documentation. Available: https://www.chartjs.org/docs/latest/",
        "Vines, J. S. (2023). pdfplumber: Plumb a PDF for Detailed Information about Each Char, Rectangle, Line, and Table. Available: https://github.com/jsvine/pdfplumber",
        "Smith, R. (2023). Tesseract OCR Engine & Pytesseract Python Wrapper. Available: https://github.com/madmaze/pytesseract",
        "OpenPyXL Development Team. (2024). OpenPyXL: A Python Library to Read/Write Excel 2010 xlsx/xlsm Files. Available: https://openpyxl.readthedocs.io/",
        "Mozilla Developer Network (MDN). (2024). Web Technologies Documentation: HTML5, CSS3, JavaScript ES6. Available: https://developer.mozilla.org/",
        "OWASP Foundation. (2023). Web Application Security Verification Standard & Guidelines. Available: https://owasp.org/"
    ]
    for b_idx, item in enumerate(bib_items):
        p_bib = doc.add_paragraph()
        p_bib.paragraph_format.space_after = Pt(4)
        p_bib.paragraph_format.line_spacing = 1.15
        r_num = p_bib.add_run(f"[{b_idx + 1}] ")
        r_num.bold = True
        p_bib.add_run(item)

    doc.add_page_break()

    # ==================== APPENDICES ====================
    print("Generating Appendices...")
    add_h1("APPENDICES", page_break=False)

    # APPENDIX A
    add_h2("Appendix A: Data Flow Diagrams (DFD)")
    add_h3("A.1 Level 0: Context Analysis Diagram")
    add_body("The Level 0 Context Diagram establishes the primary system boundary, showing information flows between external entities (Administrator and Bank Statements) and the Customer Management and Payment Tracking System:")
    
    dfd0_text = """
    +------------------------+                     +---------------------------------------+
    |                        | --- Login Request ->|                                       |
    |                        |<-- Auth Response ---|                                       |
    |     ADMINISTRATOR      | --- Customer Data ->|         CUSTOMER MANAGEMENT &         |
    |       (Operator)       |<-- Customer List ---|        PAYMENT TRACKING SYSTEM        |
    |                        | --- Search/Filter ->|                                       |
    |                        |<-- Analytics/KPI ---|                                       |
    +------------------------+                     +---------------------------------------+
                                                               ^               |
                                                               | (Upload)      | (Read/Write)
                                                               |               v
                                                   +----------------+   +-------------------+
                                                   | BANK STATEMENT |   |  SQLITE DATABASE  |
                                                   |   (PDF/CSV)    |   |   (fiber.db)      |
                                                   +----------------+   +-------------------+
    """
    p_dfd0 = doc.add_paragraph()
    p_dfd0.paragraph_format.space_after = Pt(10)
    r_d0 = p_dfd0.add_run(dfd0_text)
    r_d0.font.name = "Consolas"
    r_d0.font.size = Pt(9.5)

    add_h3("A.2 Level 1: Functional Data Flow Diagram")
    add_body("The Level 1 Diagram decomposes system processing into four core operational sub-processes:")
    
    dfd1_text = """
    +---------------+        +----------------------+        +-----------------------+
    | Administrator | -----> | 1.0 Customer Mgmt    | -----> | Stored in 'customer'  |
    +---------------+        |     Add/Edit/Delete  |        | table in SQLite DB    |
           |                 +----------------------+        +-----------------------+
           |
           | Upload PDF
           v
    +-----------------------+        +----------------------+        +-----------------------+
    | 2.0 Statement Parsing | -----> | 3.0 Multi-Parameter  | -----> | Match found:          |
    |     pdfplumber / OCR  |        |     Matching Engine  |        | Update 'payment' &    |
    +-----------------------+        |     (UPI/Phone/Cost) |        | set paid_status='paid'|
                                     +----------------------+        +-----------------------+
                                                |
                                                v Ambiguous match
                                     +----------------------+
                                     | 4.0 Verification     | -----> Admin manual 1-click
                                     |     Queue            |        approval to resolve
                                     +----------------------+
    """
    p_dfd1 = doc.add_paragraph()
    p_dfd1.paragraph_format.space_after = Pt(12)
    r_d1 = p_dfd1.add_run(dfd1_text)
    r_d1.font.name = "Consolas"
    r_d1.font.size = Pt(9.5)

    # APPENDIX B
    add_h2("Appendix B: Database Table Structure & Schemas")
    add_body("Detailed technical specification of database tables, constraints, and relationships implemented in SQLite (`database/fiber.db`):")
    
    add_h3("Table: `customer` (Subscriber Records)")
    add_bullet("customer_id: INTEGER PRIMARY KEY AUTOINCREMENT — Unique primary key index.")
    add_bullet("box_id: TEXT UNIQUE NOT NULL — Physical fiber distribution hub / box identifier.")
    add_bullet("name: TEXT NOT NULL — Subscriber's full legal name.")
    add_bullet("phone: TEXT NOT NULL — 10-digit mobile phone contact number.")
    add_bullet("subscription: REAL NOT NULL — Base monthly subscription amount.")
    add_bullet("status: TEXT NOT NULL — Network port connection state ('active' / 'inactive').")
    add_bullet("paid_status: TEXT NOT NULL DEFAULT 'unpaid' — Payment clearance flag.")
    add_bullet("upi_id: TEXT NULLABLE — Associated UPI virtual address.")

    add_h3("Table: `payment` (Transaction Records)")
    add_bullet("payment_id: INTEGER PRIMARY KEY AUTOINCREMENT — Unique transaction index.")
    add_bullet("customer_id: INTEGER REFERENCES customer(customer_id) — Foreign key link.")
    add_bullet("amount: REAL NOT NULL — Credit payment amount in INR.")
    add_bullet("payment_date: TEXT NOT NULL — Date recorded (YYYY-MM-DD).")
    add_bullet("month: TEXT NOT NULL — Accounting billing cycle (YYYY-MM).")
    add_bullet("transaction_id: TEXT UNIQUE — Banking transaction reference.")
    add_bullet("status: TEXT NOT NULL — Verification state ('Paid', 'Needs Verification', 'Unidentified Payment').")

    add_h3("Table: `unmatched_transactions` (Audit Log)")
    add_bullet("id: INTEGER PRIMARY KEY AUTOINCREMENT — Unmatched entry index.")
    add_bullet("raw_description: TEXT — Raw unparsed bank statement string.")
    add_bullet("amount: REAL — Extracted currency amount.")
    add_bullet("reason: TEXT — Reason matching failed (e.g. 'No matching customer phone or UPI found').")

    # APPENDIX C
    add_h2("Appendix C: Sample Coding")
    add_body("Representative excerpts of core backend logic implemented in `app.py`:")

    code_snippet_1 = '''# ==========================================================
# 1. DYNAMIC BILLING MONTHS DISCOVERY & DASHBOARD ANALYTICS
# ==========================================================
def get_available_months():
    conn = get_db_connection()
    rows = conn.execute("SELECT DISTINCT month FROM payment WHERE month IS NOT NULL AND month != '' ORDER BY month DESC").fetchall()
    conn.close()
    months = [r['month'] for r in rows if r['month']]
    current = datetime.now().strftime('%Y-%m')
    if current not in months:
        months.append(current)
    return sorted(list(set(months)), reverse=True)

@app.route("/dashboard")
def dashboard():
    available_months = get_available_months()
    requested_month = request.args.get("month", "").strip()
    month = requested_month or (get_latest_paid_month() or datetime.now().strftime('%Y-%m'))
    
    stats = calculate_payment_status_for_month(month)
    total_customers = stats['total']
    paid_customers = stats['paid']
    unpaid_customers = stats['unpaid']
    collection = stats['collection']

    conn = get_db_connection()
    # Compute Total Potential Revenue & Realization Rate
    active_row = conn.execute(
        "SELECT COUNT(*) as cnt, SUM(COALESCE(plan_amount, subscription)) as total_due FROM customer WHERE LOWER(status) = 'active'"
    ).fetchone()
    total_potential = round(active_row['total_due'] or 0.0, 2)
    realization_rate = round((collection / total_potential * 100), 1) if total_potential > 0 else 0
    outstanding_dues = max(0.0, round(total_potential - collection, 2))

    # Compute 6-Month Revenue History Trend
    monthly_rows = conn.execute(
        "SELECT month, SUM(amount) as total FROM payment WHERE status = 'Paid' GROUP BY month ORDER BY month ASC LIMIT 12"
    ).fetchall()
    
    # Compute Plan Tariff Breakdown
    plan_rows = conn.execute(
        "SELECT CAST(COALESCE(plan_amount, subscription) AS TEXT) as plan_tier, COUNT(*) as cnt FROM customer GROUP BY plan_tier"
    ).fetchall()
    conn.close()

    return render_template("dashboard.html", dashboard_month=month, available_months=available_months,
                           total_customers=total_customers, paid_customers=paid_customers,
                           unpaid_customers=unpaid_customers, collection=collection,
                           realization_rate=realization_rate, outstanding_dues=outstanding_dues)
'''
    p_c1 = doc.add_paragraph()
    r_c1 = p_c1.add_run(code_snippet_1)
    r_c1.font.name = "Consolas"
    r_c1.font.size = Pt(8.5)

    code_snippet_2 = '''# ==========================================================
# 2. CLIENT-SIDE REAL-TIME MULTI-CRITERIA SEARCH & FILTER
# ==========================================================
function applyFilters() {
    const query = (document.getElementById('customerSearch').value || '').trim().toLowerCase();
    const connFilter = document.getElementById('connectionFilter').value.toLowerCase();
    const payFilter = document.getElementById('paymentFilter').value.toLowerCase();
    const rows = document.querySelectorAll('.customer-row');
    let visibleCount = 0;

    rows.forEach(row => {
        const name = row.getAttribute('data-name') || '';
        const phone = row.getAttribute('data-phone') || '';
        const boxId = row.getAttribute('data-boxid') || '';
        const upi = row.getAttribute('data-upi') || '';
        const connection = row.getAttribute('data-connection') || '';
        const payment = row.getAttribute('data-payment') || '';

        const matchesQuery = !query || name.includes(query) || phone.includes(query) || 
                             boxId.includes(query) || upi.includes(query);
        const matchesConn = (connFilter === 'all') || (connection === connFilter);
        const matchesPay = (payFilter === 'all') || (payment === payFilter);

        if (matchesQuery && matchesConn && matchesPay) {
            row.style.display = '';
            visibleCount++;
        } else {
            row.style.display = 'none';
        }
    });
    document.getElementById('visibleCount').textContent = visibleCount;
}
'''
    p_c2 = doc.add_paragraph()
    r_c2 = p_c2.add_run(code_snippet_2)
    r_c2.font.name = "Consolas"
    r_c2.font.size = Pt(8.5)

    # APPENDIX D
    add_h2("Appendix D: Sample Input Files")
    add_h3("D.1 Sample Bank Statement Extracted Raw Text")
    add_body("Below is an illustrative sample of digital transaction credit entries parsed from an incoming bank statement PDF:")
    
    sample_stmt = """
    Date        Transaction Description                             Chq/Ref No.    Credit (INR)
    -------------------------------------------------------------------------------------------
    01/07/2026  UPI/CR/418291028192/RAVI KUMAR/ravi@okaxis/9842208262  TXN001         250.00
    02/07/2026  UPI/CR/418291039201/PRIYA SINGH/priya@okicici/NILACBL  TXN002         299.00
    03/07/2026  UPI/CR/418291048293/AMIT PATEL/amit@okaxis/FIBERPAY    TXN003         250.00
    05/07/2026  IMPS/P2A/418291058291/NEHA SHARMA/9988776655           TXN004         199.00
    06/07/2026  UPI/CR/418291069284/RAHUL VERMA/rahul@okaxis/B105BILL  TXN005         250.00
    -------------------------------------------------------------------------------------------
    """
    p_st = doc.add_paragraph()
    r_st = p_st.add_run(sample_stmt)
    r_st.font.name = "Consolas"
    r_st.font.size = Pt(9.0)

    add_h3("D.2 Sample Bulk Customer Import Format (.xlsx)")
    add_body("Structure used for batch importing customer accounts from Excel:")
    
    t_imp = doc.add_table(rows=5, cols=6)
    t_imp.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_imp.autofit = False
    imp_headers = ["BOX_ID", "NAME", "PHONE", "PLAN_AMOUNT", "STATUS", "UPI_ID"]
    for idx, h in enumerate(imp_headers):
        cell = t_imp.rows[0].cells[idx]
        cell.paragraphs[0].text = h
        cell.paragraphs[0].runs[0].bold = True
        set_cell_background(cell, "E5E7EB")

    imp_rows = [
        ("B101", "Ravi Kumar", "9842208262", "250.00", "active", "ravi@okaxis"),
        ("B102", "Priya Singh", "9876543210", "299.00", "active", "priya@okicici"),
        ("B103", "Amit Patel", "9123456789", "250.00", "active", "amit@okaxis"),
        ("B104", "Neha Sharma", "9988776655", "199.00", "active", "neha@sbi")
    ]
    for r_idx, rdata in enumerate(imp_rows):
        row = t_imp.rows[r_idx + 1]
        for c_idx, val in enumerate(rdata):
            row.cells[c_idx].paragraphs[0].text = val

    # APPENDIX E
    add_h2("Appendix E: Sample Outputs & Analytics Visualizations")
    add_h3("E.1 Dashboard Visual Analytics Summary Table")
    add_body("Numerical summary generated by the Dashboard Visual Analytics Engine for billing month 2026-07:")

    t_out = doc.add_table(rows=7, cols=2)
    t_out.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_out.autofit = False
    t_out.rows[0].cells[0].width = Inches(3.0)
    t_out.rows[0].cells[1].width = Inches(3.5)
    t_out.rows[0].cells[0].paragraphs[0].text = "METRIC / KPI"
    t_out.rows[0].cells[0].paragraphs[0].runs[0].bold = True
    t_out.rows[0].cells[1].paragraphs[0].text = "MEASURED SYSTEM VALUE"
    t_out.rows[0].cells[1].paragraphs[0].runs[0].bold = True
    set_cell_background(t_out.rows[0].cells[0], "E5E7EB")
    set_cell_background(t_out.rows[0].cells[1], "E5E7EB")

    out_metrics = [
        ("Total Subscribed Customers", "8 Active Fiber Connections"),
        ("Cleared / Paid Accounts", "5 Subscribers (62.5% Payment Rate)"),
        ("Pending / Unpaid Accounts", "3 Subscribers (37.5% Overdue)"),
        ("Total Collection This Cycle", "₹ 1,248.00 INR"),
        ("Total Expected Billable Potential", "₹ 1,997.00 INR"),
        ("Financial Realization Rate", "62.5% Collection Realization")
    ]
    for r_idx, (m, v) in enumerate(out_metrics):
        row = t_out.rows[r_idx + 1]
        row.cells[0].width = Inches(3.0)
        row.cells[1].width = Inches(3.5)
        row.cells[0].paragraphs[0].text = m
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[1].paragraphs[0].text = v

    add_h3("E.2 Generated Monthly Revenue Report")
    add_body("System-generated financial collection report exported from `/reports`:")
    
    rpt_text = """
    =============================================================================
                    NILACABLE FIBRE NETWORKS - COLLECTION REPORT
                              PERIOD: JULY 2026
    =============================================================================
    Total Active Connections : 8
    Paid Subscribers         : 5 (₹ 1,248.00)
    Unpaid Accounts          : 3 (₹ 749.00)
    Collection Efficiency    : 62.5%
    -----------------------------------------------------------------------------
    Txn ID   Subscriber Name    Phone No.    Plan Cost   Paid Date    Status
    -----------------------------------------------------------------------------
    TXN001   Ravi Kumar         9842208262   ₹250.00     2026-07-01   PAID
    TXN002   Priya Singh        9876543210   ₹299.00     2026-07-02   PAID
    TXN003   Amit Patel         9123456789   ₹250.00     2026-07-03   PAID
    TXN004   Neha Sharma        9988776655   ₹199.00     2026-07-05   PAID
    TXN005   Rahul Verma        9555443322   ₹250.00     2026-07-06   PAID
    --       Anjali Gupta       9111222333   ₹299.00     --           UNPAID
    --       Vikas Yadav        9444555666   ₹250.00     --           UNPAID
    --       Pooja Singh        9777888999   ₹199.00     --           UNPAID
    =============================================================================
    """
    p_rpt = doc.add_paragraph()
    r_rpt = p_rpt.add_run(rpt_text)
    r_rpt.font.name = "Consolas"
    r_rpt.font.size = Pt(8.5)

    # Save document
    out_path = os.path.abspath("Customer_Management_and_Payment_Tracking_System_Report.docx")
    doc.save(out_path)
    print(f"Report successfully saved to: {out_path}")
    return out_path

if __name__ == "__main__":
    create_report()
