#!/usr/bin/env python3
"""
generate_honest_cv.py
Automated generator for Aaron Alex Mathew's honest, ATS-friendly CV in .docx format.
"""

import os
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_element(name):
    return OxmlElement(name)

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

    # Set font family
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Calibri" w:hAnsi="Calibri"/>')
    rPr.append(rFonts)

    new_run.append(rPr)
    text_node = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def set_cell_margins(cell, top=50, bottom=50, left=50, right=50):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_section_header(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    
    run = p.add_run(title.upper())
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)  # Professional Navy

    # Add subtle bottom border line via XML
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="CBD5E1"/></w:pBdr>')
    pPr.append(pBdr)
    return p

def main():
    doc = Document()

    # Page Margins: 0.55 inches for optimal density & ATS readability
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Base Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    # Header Section
    header_p = doc.add_paragraph()
    header_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    header_p.paragraph_format.space_after = Pt(2)

    name_run = header_p.add_run("AARON ALEX MATHEW\n")
    name_run.bold = True
    name_run.font.name = 'Calibri'
    name_run.font.size = Pt(18)
    name_run.font.color.rgb = RGBColor(0x0F, 0x2C, 0x59)

    contact_p = doc.add_paragraph()
    contact_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    contact_p.paragraph_format.space_after = Pt(6)

    c_run = contact_p.add_run("Kerala, India  |  +91 80759 85482  |  ")
    c_run.font.size = Pt(9.5)

    add_hyperlink(contact_p, "mailto:aaronalexmathew48@gmail.com", "aaronalexmathew48@gmail.com")
    sep1 = contact_p.add_run("  |  ")
    sep1.font.size = Pt(9.5)

    add_hyperlink(contact_p, "https://github.com/RYuK006", "GitHub: RYuK006")
    sep2 = contact_p.add_run("  |  ")
    sep2.font.size = Pt(9.5)

    add_hyperlink(contact_p, "https://www.linkedin.com/in/aaron-alex-mathew/", "LinkedIn")

    # Output path
    output_path = os.path.join(os.path.dirname(__file__), "Aaron_CV_Honest.docx")
    doc.save(output_path)
    print(f"Base CV skeleton created at: {output_path}")

if __name__ == "__main__":
    main()
