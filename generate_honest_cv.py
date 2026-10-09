#!/usr/bin/env python3
"""
generate_honest_cv.py
Automated generator for Aaron Alex Mathew's honest, verifiable, ATS-friendly CV in .docx format.
Adheres strictly to single-column ATS parsing guidelines while offering clean, human-readable typography.
"""

import os
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def add_hyperlink(paragraph, url, text, color="1D4ED8", underline=True, bold=False):
    """
    Inserts a genuine OpenXML hyperlink element into a docx paragraph.
    Properly parsed by ATS systems and clickable in Word/PDF.
    """
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')

    if color:
        c = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
        rPr.append(c)

    if underline:
        u = parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>')
        rPr.append(u)

    if bold:
        b = parse_xml(f'<w:b {nsdecls("w")}/>')
        rPr.append(b)

    # Set font family to Calibri and size to 9.5pt
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Calibri" w:hAnsi="Calibri"/>')
    rPr.append(rFonts)
    sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="19"/>')  # 19 half-points = 9.5pt
    rPr.append(sz)

    new_run.append(rPr)
    text_node = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def add_section_header(doc, title):
    """Adds a bold section heading with a crisp horizontal bottom border."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    
    run = p.add_run(title.upper())
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)  # Deep Professional Navy

    # Add subtle bottom border line via XML
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="94A3B8"/></w:pBdr>')
    pPr.append(pBdr)
    return p

def add_bullet_point(doc, bold_prefix, text_content, space_after=2):
    """Adds a clean, standard list bullet with bold lead-in."""
    bp = doc.add_paragraph(style='List Bullet')
    bp.paragraph_format.space_before = Pt(0)
    bp.paragraph_format.space_after = Pt(space_after)
    bp.paragraph_format.left_indent = Inches(0.25)
    bp.paragraph_format.line_spacing = 1.05

    if bold_prefix:
        r_bold = bp.add_run(bold_prefix)
        r_bold.bold = True
        r_bold.font.name = 'Calibri'
        r_bold.font.size = Pt(9.5)
        r_bold.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    r_text = bp.add_run(text_content)
    r_text.font.name = 'Calibri'
    r_text.font.size = Pt(9.5)
    r_text.font.color.rgb = RGBColor(0x37, 0x41, 0x51)
    return bp

def build_cv():
    doc = Document()

    # ATS-Friendly Page Margins: 0.55" top/bottom, 0.6" left/right
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Configure Normal Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    # =========================================================================
    # HEADER
    # =========================================================================
    header_p = doc.add_paragraph()
    header_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    header_p.paragraph_format.space_before = Pt(0)
    header_p.paragraph_format.space_after = Pt(1)

    name_run = header_p.add_run("AARON ALEX MATHEW")
    name_run.bold = True
    name_run.font.name = 'Calibri'
    name_run.font.size = Pt(17)
    name_run.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)

    contact_p = doc.add_paragraph()
    contact_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    contact_p.paragraph_format.space_before = Pt(0)
    contact_p.paragraph_format.space_after = Pt(6)

    c_run = contact_p.add_run("Kerala, India  |  +91 80759 85482  |  ")
    c_run.font.name = 'Calibri'
    c_run.font.size = Pt(9.5)

    add_hyperlink(contact_p, "mailto:aaronalexmathew48@gmail.com", "aaronalexmathew48@gmail.com")
    sep1 = contact_p.add_run("  |  ")
    sep1.font.name = 'Calibri'
    sep1.font.size = Pt(9.5)

    add_hyperlink(contact_p, "https://github.com/RYuK006", "GitHub: RYuK006")
    sep2 = contact_p.add_run("  |  ")
    sep2.font.name = 'Calibri'
    sep2.font.size = Pt(9.5)

    add_hyperlink(contact_p, "https://www.linkedin.com/in/aaron-alex-mathew/", "LinkedIn: aaron-alex-mathew")

    # =========================================================================
    # PROFESSIONAL SUMMARY
    # =========================================================================
    add_section_header(doc, "Professional Summary")
    sum_p = doc.add_paragraph()
    sum_p.paragraph_format.space_before = Pt(2)
    sum_p.paragraph_format.space_after = Pt(6)
    sum_p.paragraph_format.line_spacing = 1.08

    sum_text = (
        "AI systems builder and Computer Science undergraduate specializing in Artificial Intelligence, with demonstrated "
        "experience architecting deep learning pipelines, cross-modal transformer frameworks, and full-stack platforms. "
        "Track record spans affective computing research at IIIT Kottayam (achieving 80.03% binary classification accuracy on CMU-MOSEI), "
        "materials informatics pipelines filtering theoretical candidate pools for superconductor discovery, and shipping "
        "production web platforms with 300+ active users. Passionate about solving high-complexity problems across machine learning "
        "research, data pipelines, and scalable platform engineering."
    )
    sum_run = sum_p.add_run(sum_text)
    sum_run.font.name = 'Calibri'
    sum_run.font.size = Pt(9.5)

    # =========================================================================
    # FORMAL WORK & RESEARCH EXPERIENCE (Placed BEFORE projects)
    # =========================================================================
    add_section_header(doc, "Work & Research Experience")

    # Entry: IIIT Kottayam Internship
    exp_p = doc.add_paragraph()
    exp_p.paragraph_format.space_before = Pt(3)
    exp_p.paragraph_format.space_after = Pt(1)
    exp_p.paragraph_format.keep_with_next = True

    role_run = exp_p.add_run("Research Intern — Affective Computing Lab")
    role_run.bold = True
    role_run.font.name = 'Calibri'
    role_run.font.size = Pt(10)
    role_run.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)

    inst_run = exp_p.add_run(" | Indian Institute of Information Technology (IIIT), Kottayam")
    inst_run.bold = True
    inst_run.font.name = 'Calibri'
    inst_run.font.size = Pt(10)

    date_run = exp_p.add_run("  [May 18, 2026 – July 16, 2026]")
    date_run.font.name = 'Calibri'
    date_run.font.size = Pt(9)
    date_run.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    sub_exp = doc.add_paragraph()
    sub_exp.paragraph_format.space_before = Pt(0)
    sub_exp.paragraph_format.space_after = Pt(3)
    sub_exp.paragraph_format.keep_with_next = True
    adv_run = sub_exp.add_run("Advisor: Dr. Anu Maria Sebastian, Associate Faculty | Focus: Multimodal Sentiment Analysis (MSA) via Cross-Modal Transformers")
    adv_run.italic = True
    adv_run.font.name = 'Calibri'
    adv_run.font.size = Pt(9)
    adv_run.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    add_bullet_point(
        doc,
        "Architecture Design: ",
        "Designed and implemented a text-anchored cross-attention transformer framework in PyTorch integrating textual, acoustic, and visual modalities, utilizing text tokens as the anchor query attending across aligned acoustic and visual feature streams."
    )
    add_bullet_point(
        doc,
        "Benchmark Performance: ",
        "Evaluated on the standard CMU-MOSEI benchmark dataset: achieved 80.03% binary classification accuracy and 0.8687 weighted F1 score, validating multi-modal synergy."
    )
    add_bullet_point(
        doc,
        "Baseline Outperformance: ",
        "Outperformed conventional early fusion (78.14%) and late fusion (77.23%) baselines through a novel parallel cross-attention layer architecture configured with 16 attention heads and a 128-dimensional embedding space."
    )
    add_bullet_point(
        doc,
        "Ablations & Interpretability: ",
        "Conducted 11 ablation studies confirming that cross-attention requires sufficient representational capacity (16 heads vs. 4 heads) to surpass simpler baselines; generated attention heatmaps confirming model anchoring on emotionally salient video frames rather than monotonic temporal matching."
    )
    add_bullet_point(
        doc,
        "Research Publication: ",
        'Authored preprint "Towards Holistic Emotion Understanding: A Text-Anchored Cross-Attention Transformer for Multimodal Sentiment Analysis" [Status: Preprint, available on request; formal target: IEEE Transactions on Affective Computing].',
        space_after=5
    )

    # =========================================================================
    # INDEPENDENT & RESEARCH PROJECTS
    # =========================================================================
    add_section_header(doc, "Independent & Research Projects")

    # Project 1: CrowdVotesAI
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(3)
    p1.paragraph_format.space_after = Pt(1)
    p1.paragraph_format.keep_with_next = True

    p1_title = p1.add_run("CrowdVotesAI — Crowdsourced Prediction Platform ")
    p1_title.bold = True
    p1_title.font.name = 'Calibri'
    p1_title.font.size = Pt(10)
    p1_title.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)

    p1_status = p1.add_run("[Shipped | Active Development]\n")
    p1_status.bold = True
    p1_status.font.name = 'Calibri'
    p1_status.font.size = Pt(9)
    p1_status.font.color.rgb = RGBColor(0x16, 0x65, 0x34)  # Forest Green

    p1_links = p1.add_run("Live: ")
    p1_links.font.name = 'Calibri'
    p1_links.font.size = Pt(9)
    add_hyperlink(p1, "https://crowdvotesai.in", "https://crowdvotesai.in")

    p1_sep = p1.add_run("  |  GitHub: ")
    p1_sep.font.name = 'Calibri'
    p1_sep.font.size = Pt(9)
    add_hyperlink(p1, "https://github.com/RYuK006/CrowdVote", "github.com/RYuK006/CrowdVote")

    p1_stack = p1.add_run("  |  Stack: React, Node.js, Firebase/Firestore, Vite, Vercel")
    p1_stack.italic = True
    p1_stack.font.name = 'Calibri'
    p1_stack.font.size = Pt(9)
    p1_stack.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    add_bullet_point(
        doc,
        "Production Scale & Traction: ",
        "Architected, deployed, and currently maintain a full-stack real-time polling platform serving 300+ active users for community predictions across competitive events."
    )
    add_bullet_point(
        doc,
        "Credibility Scoring Algorithm: ",
        "Engineered a proprietary credibility-based vote weighting algorithm that tracks user historical forecasting accuracy and dynamically adjusts crowd consensus weights based on predictive track records."
    )
    add_bullet_point(
        doc,
        "Zero-Downtime Data Migration: ",
        "Executed zero-downtime database migrations in Firebase NoSQL, successfully pivoting the architecture from regional electoral polling to a global 2026 FIFA World Cup predictor MVP."
    )
    add_bullet_point(
        doc,
        "Real-Time Backend Infrastructure: ",
        "Optimized Firestore query patterns and real-time listeners for sub-second synchronization across concurrent voting sessions; deployed on Vercel with automated CI/CD pipelines.",
        space_after=4
    )

    # Project 2: Q-MATIS
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(3)
    p2.paragraph_format.space_after = Pt(1)
    p2.paragraph_format.keep_with_next = True

    p2_title = p2.add_run("Q-MATIS: Superconducting Materials Discovery Pipeline ")
    p2_title.bold = True
    p2_title.font.name = 'Calibri'
    p2_title.font.size = Pt(10)
    p2_title.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)

    p2_status = p2.add_run("[Active Research Project | In Development]\n")
    p2_status.bold = True
    p2_status.font.name = 'Calibri'
    p2_status.font.size = Pt(9)
    p2_status.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)  # Cobalt Blue

    p2_links = p2.add_run("GitHub: ")
    p2_links.font.name = 'Calibri'
    p2_links.font.size = Pt(9)
    add_hyperlink(p2, "https://github.com/RYuK006/Q-MATIS", "github.com/RYuK006/Q-MATIS")

    p2_stack = p2.add_run("  |  Stack: Python, PyTorch Geometric, PyMatGen, Materials Project API, SciPy, NumPy")
    p2_stack.italic = True
    p2_stack.font.name = 'Calibri'
    p2_stack.font.size = Pt(9)
    p2_stack.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    add_bullet_point(
        doc,
        "Candidate Generation & Filtering: ",
        "Architected a scalable materials informatics pipeline designed to screen a theoretical pool of 1 billion+ candidate compound structures for thermodynamic stability and superconductivity."
    )
    add_bullet_point(
        doc,
        "Graph Neural Network Engineering: ",
        "Implemented GNN stability and critical temperature (Tc) predictors—including ModernCGCNN and native PyTorch Geometric ALIGNN (rewritten without DGL dependency)—achieving a 3.53% MAE improvement over baseline CGCNN on standardized randomized splits."
    )
    add_bullet_point(
        doc,
        "Data Ingestion & Featurization: ",
        "Constructed automated data pipelines querying the Materials Project API, extracting crystallographic data (CIF), and featurizing atomic coordinates and bond lengths into graph node/edge tensors."
    )
    add_bullet_point(
        doc,
        "Statistical Calibration & Architecture: ",
        "Formulated uncertainty quantification using Prediction Interval Coverage Probability (PICP) and embedding validation via Silhouette scores; authored 19-phase modular refactor with full PyTest suites.",
        space_after=4
    )

    # Project 3: μLearnPRC Community Website
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(3)
    p3.paragraph_format.space_after = Pt(1)
    p3.paragraph_format.keep_with_next = True

    p3_title = p3.add_run("μLearnPRC Community Website ")
    p3_title.bold = True
    p3_title.font.name = 'Calibri'
    p3_title.font.size = Pt(10)
    p3_title.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)

    p3_status = p3.add_run("[In Development | Launching This Week]\n")
    p3_status.bold = True
    p3_status.font.name = 'Calibri'
    p3_status.font.size = Pt(9)
    p3_status.font.color.rgb = RGBColor(0xB4, 0x53, 0x09)  # Amber

    p3_links = p3.add_run("Live Staging: ")
    p3_links.font.name = 'Calibri'
    p3_links.font.size = Pt(9)
    add_hyperlink(p3, "https://mu-learn-prc-website.vercel.app/", "mu-learn-prc-website.vercel.app")

    p3_sep = p3.add_run(" (Migrating to mulearnprc.com)  |  Stack: React, Next.js, TailwindCSS, Framer Motion, Vercel")
    p3_sep.italic = True
    p3_sep.font.name = 'Calibri'
    p3_sep.font.size = Pt(9)
    p3_sep.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    add_bullet_point(
        doc,
        "Full-Stack Platform Ownership: ",
        "Independently designed, architected, and built the complete official web platform for μLearn PRC (student tech community at Providence College of Engineering), serving as the community hub for member profiles, announcements, and events."
    )
    add_bullet_point(
        doc,
        "Measured Lighthouse Metrics: ",
        "Achieved measured Google Lighthouse production audit scores of 93/100 Accessibility, 88/100 Code Quality, and 72/100 Performance on initial Vercel deployment."
    )
    add_bullet_point(
        doc,
        "Performance & SEO Optimization: ",
        "Executed systematic optimizations: eliminated render-blocking resources, implemented next/image lazy loading, and built structured meta-tag frameworks to target 70+ SEO score ahead of public domain launch."
    )
    add_bullet_point(
        doc,
        "Modern Responsive UI: ",
        "Built cinematic, accessible user interfaces using Framer Motion micro-animations with strict WCAG compliance and seamless cross-device mobile/desktop parity.",
        space_after=4
    )

    # Project 4: Multimodal Emotion Recognition Transformer (Research System)
    p4 = doc.add_paragraph()
    p4.paragraph_format.space_before = Pt(3)
    p4.paragraph_format.space_after = Pt(1)
    p4.paragraph_format.keep_with_next = True

    p4_title = p4.add_run("Multimodal Emotion Recognition Transformer ")
    p4_title.bold = True
    p4_title.font.name = 'Calibri'
    p4_title.font.size = Pt(10)
    p4_title.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)

    p4_status = p4.add_run("[Research Project | Preprint Completed]\n")
    p4_status.bold = True
    p4_status.font.name = 'Calibri'
    p4_status.font.size = Pt(9)
    p4_status.font.color.rgb = RGBColor(0x43, 0x38, 0xCA)  # Indigo

    p4_details = p4.add_run("Scope: Text-Anchored Cross-Modal Attention Framework  |  Stack: PyTorch, Transformers, Librosa, OpenFace, CMU-MOSEI")
    p4_details.italic = True
    p4_details.font.name = 'Calibri'
    p4_details.font.size = Pt(9)
    p4_details.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    add_bullet_point(
        doc,
        "Cross-Modal Fusion Engine: ",
        "Formulated and coded the complete dual-stream cross-attention PyTorch framework treating text as the query anchor attending across extracted acoustic (Librosa/COVAREP) and visual (FACET/OpenFace) feature sequences."
    )
    add_bullet_point(
        doc,
        "Comprehensive Evaluation Suite: ",
        "Orchestrated 11 rigorous experimental evaluations covering modality ablations, attention head scaling (4 vs. 16 heads), and fusion strategies to validate architectural gains on CMU-MOSEI."
    )
    add_bullet_point(
        doc,
        "Qualitative Diagnostics: ",
        "Extracted attention weight matrices and computed t-SNE latent space projections, confirming cross-modal semantic clustering on emotionally salient frames.",
        space_after=5
    )

    # =========================================================================
    # TECHNICAL SKILLS
    # =========================================================================
    add_section_header(doc, "Technical Skills")

    skills_data = [
        ("AI & Machine Learning: ", "PyTorch, TensorFlow, HuggingFace Transformers, PyTorch Geometric, Scikit-Learn, SciPy, Matplotlib"),
        ("Agentic & LLM Systems: ", "Claude API, OpenAI API, Prompt Engineering, Evaluation Benchmarks, Local Models (Ollama)"),
        ("Programming Languages: ", "Python (Primary), JavaScript (ES6+), TypeScript, SQL, C++"),
        ("Backend & Databases: ", "Node.js, Express.js, FastAPI, PostgreSQL, Firebase / Firestore, RESTful APIs"),
        ("Frontend & UI Design: ", "React, Next.js, TailwindCSS, HTML5, CSS3, Framer Motion, Figma"),
        ("Developer Tools & OS: ", "Git, GitHub, Linux (Ubuntu/WSL), Vercel, Vite, Jupyter Notebook, VS Code")
    ]

    for label, items in skills_data:
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(1.5)
        sp.paragraph_format.line_spacing = 1.05

        lbl_run = sp.add_run(label)
        lbl_run.bold = True
        lbl_run.font.name = 'Calibri'
        lbl_run.font.size = Pt(9.5)
        lbl_run.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

        itm_run = sp.add_run(items)
        itm_run.font.name = 'Calibri'
        itm_run.font.size = Pt(9.5)
        itm_run.font.color.rgb = RGBColor(0x37, 0x41, 0x51)

    # =========================================================================
    # EDUCATION & COURSEWORK
    # =========================================================================
    add_section_header(doc, "Education & Relevant Coursework")

    edu_p = doc.add_paragraph()
    edu_p.paragraph_format.space_before = Pt(2)
    edu_p.paragraph_format.space_after = Pt(1)
    edu_p.paragraph_format.keep_with_next = True

    deg_run = edu_p.add_run("Bachelor of Technology (B.Tech) in Computer Science & Engineering (AI Specialization)")
    deg_run.bold = True
    deg_run.font.name = 'Calibri'
    deg_run.font.size = Pt(10)
    deg_run.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)

    grad_run = edu_p.add_run("  [Expected Graduation: 2028]")
    grad_run.font.name = 'Calibri'
    grad_run.font.size = Pt(9)
    grad_run.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    inst_p = doc.add_paragraph()
    inst_p.paragraph_format.space_before = Pt(0)
    inst_p.paragraph_format.space_after = Pt(2)
    inst_p.paragraph_format.keep_with_next = True

    clg_run = inst_p.add_run("Providence College of Engineering, Chengannur, Kerala (Affiliated with APJ Abdul Kalam Technological University - KTU)")
    clg_run.font.name = 'Calibri'
    clg_run.font.size = Pt(9.5)
    clg_run.font.color.rgb = RGBColor(0x37, 0x41, 0x51)

    cw_p = doc.add_paragraph()
    cw_p.paragraph_format.space_before = Pt(0)
    cw_p.paragraph_format.space_after = Pt(4)
    cw_lbl = cw_p.add_run("Relevant Coursework: ")
    cw_lbl.bold = True
    cw_lbl.font.name = 'Calibri'
    cw_lbl.font.size = Pt(9.5)
    cw_txt = cw_p.add_run("Machine Learning, Deep Learning, Data Structures & Algorithms, Design and Analysis of Algorithms, Database Management Systems, Operating Systems, Linear Algebra, Probability & Statistics, Object-Oriented Programming.")
    cw_txt.font.name = 'Calibri'
    cw_txt.font.size = Pt(9.5)
    cw_txt.font.color.rgb = RGBColor(0x37, 0x41, 0x51)

    # =========================================================================
    # PUBLICATIONS & RESEARCH
    # =========================================================================
    add_section_header(doc, "Publications & Research")

    pub_p = doc.add_paragraph()
    pub_p.paragraph_format.space_before = Pt(2)
    pub_p.paragraph_format.space_after = Pt(1)
    pub_p.paragraph_format.left_indent = Inches(0.25)
    pub_p.paragraph_format.line_spacing = 1.05

    pub_title = pub_p.add_run('"Towards Holistic Emotion Understanding: A Text-Anchored Cross-Attention Transformer for Multimodal Sentiment Analysis"\n')
    pub_title.bold = True
    pub_title.font.name = 'Calibri'
    pub_title.font.size = Pt(9.5)
    pub_title.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    pub_authors = pub_p.add_run("Authors: Aaron Alex Mathew, Dr. Anu Maria Sebastian (IIIT Kottayam) | ")
    pub_authors.font.name = 'Calibri'
    pub_authors.font.size = Pt(9)
    pub_authors.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    pub_status = pub_p.add_run("Status: Preprint (Available on request / ArXiv pending)\n")
    pub_status.italic = True
    pub_status.font.name = 'Calibri'
    pub_status.font.size = Pt(9)
    pub_status.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    pub_res = pub_p.add_run("Key Results: 80.03% binary classification accuracy, 0.8687 weighted F1 score on CMU-MOSEI benchmark (outperforming early and late fusion baselines).")
    pub_res.font.name = 'Calibri'
    pub_res.font.size = Pt(9)
    pub_res.font.color.rgb = RGBColor(0x37, 0x41, 0x51)

    # =========================================================================
    # LEADERSHIP, AWARDS & COMMUNITY CONTRIBUTIONS
    # =========================================================================
    add_section_header(doc, "Leadership & Community Contributions")

    add_bullet_point(
        doc,
        "Lead Web Developer & Core Contributor — μLearn PRC: ",
        "Architected, designed, and maintain the primary community web platform for μLearn PRC at Providence College of Engineering, supporting tech workshops, peer learning, and event registrations."
    )
    add_bullet_point(
        doc,
        "Research Selection — IIIT Kottayam: ",
        "Selected for competitive research internship at the Affective Computing Lab under Dr. Anu Maria Sebastian, conducting advanced transformer research on multimodal fusion."
    )
    add_bullet_point(
        doc,
        "Hackathon Competitor — Bharatiya Antariksh Hackathon (BAH 2026): ",
        "Architected Project Phoenix, a dual-engine space weather advisory system utilizing ISRO Aditya-L1 satellite telemetry (SoLEXS & HEL1OS) with Neupert Effect physics constraints for early solar flare forecasting.",
        space_after=2
    )

    # Save document
    output_path = os.path.join(os.path.dirname(__file__), "Aaron_CV_Honest.docx")
    doc.save(output_path)
    print(f"Production-ready honest CV successfully generated at: {output_path}")

if __name__ == "__main__":
    build_cv()
