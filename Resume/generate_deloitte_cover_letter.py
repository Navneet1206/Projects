"""
Cover Letter Generator - Navneet Vishwakarma
Tailored for: Deloitte - Full Stack Development - NextJS - Managed Services Engineer I (DFO&I - Customer)
Brand Colors: Deloitte Deep Forest Green (#004B49), Black (#000000), Charcoal (#2D2D2D)
Font Style: Arial (Clean Sans-Serif)
"""

import os
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Navneet_Vishwakarma_Deloitte_CoverLetter.docx")

# Deloitte Brand Color Palette
DARK = RGBColor(0x00, 0x00, 0x00)       # Black
ACCENT = RGBColor(0x00, 0x4B, 0x49)     # Deloitte Deep Forest Green (#004B49)
TEXT = RGBColor(0x2D, 0x2D, 0x2D)      # Charcoal
MUTED = RGBColor(0x55, 0x55, 0x55)     # Medium Gray
RULE_COLOR = "004B49"                  # Deloitte Deep Forest Green Hex
FONT_NAME = "Arial"


def remove_paragraph_spacing(paragraph):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(11.5)


def add_thin_rule(doc, color=RULE_COLOR, thickness="8000"):
    p = doc.add_paragraph()
    remove_paragraph_spacing(p)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="{thickness}" w:space="1" w:color="{color}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def styled_run(paragraph, text, font_name=FONT_NAME, size=Pt(10), color=TEXT,
               bold=False, italic=False):
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = size
    run.font.color.rgb = color
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    east_asia = rPr.find(qn('w:rFonts'))
    if east_asia is not None and east_asia.get(qn('w:eastAsiaTheme')) is not None:
        del east_asia.attrib[qn('w:eastAsiaTheme')]
    return run


def add_letter_paragraph(doc, text, bold_words=None, space_before=6, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(12)
    p.paragraph_format.left_indent = Inches(0)
    
    if bold_words:
        current_idx = 0
        for word, is_bold in bold_words:
            styled_run(p, word, size=Pt(9.5), bold=is_bold)
    else:
        styled_run(p, text, size=Pt(9.5))


def build_cover_letter():
    doc = Document()

    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = Pt(10)
    style.font.color.rgb = TEXT
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    # Header Contact Details
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_after = Pt(2)
    p_name.paragraph_format.line_spacing = Pt(18)
    styled_run(p_name, "NAVNEET VISHWAKARMA", font_name=FONT_NAME, size=Pt(16), color=DARK, bold=True)

    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_after = Pt(1)
    p_contact.paragraph_format.line_spacing = Pt(11)
    styled_run(p_contact, "Satna, Madhya Pradesh, India  |  +91 8435061006  |  rajvl132011@gmail.com", size=Pt(8.5), color=MUTED)

    p_links = doc.add_paragraph()
    p_links.paragraph_format.space_after = Pt(4)
    p_links.paragraph_format.line_spacing = Pt(11)
    styled_run(p_links, "LinkedIn: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "linkedin.com/in/navneet1206", size=Pt(8.5), color=ACCENT)
    styled_run(p_links, "  |  GitHub: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "github.com/Navneet1206", size=Pt(8.5), color=ACCENT)

    add_thin_rule(doc, thickness="12000")

    # Date and Recruiter Address
    p_date = doc.add_paragraph()
    p_date.paragraph_format.space_before = Pt(12)
    p_date.paragraph_format.space_after = Pt(6)
    styled_run(p_date, "June 24, 2026", size=Pt(9.5), color=TEXT)

    p_to = doc.add_paragraph()
    p_to.paragraph_format.space_before = Pt(4)
    p_to.paragraph_format.space_after = Pt(12)
    p_to.paragraph_format.line_spacing = Pt(11)
    styled_run(p_to, "To,\n", size=Pt(9.5), bold=True)
    styled_run(p_to, "The Recruiting Team\n", size=Pt(9.5))
    styled_run(p_to, "Deloitte Consulting India Private Limited\n", size=Pt(9.5))
    styled_run(p_to, "Bengaluru / Hyderabad / Pune / Chennai", size=Pt(9.5))

    # Subject Line
    p_subj = doc.add_paragraph()
    p_subj.paragraph_format.space_before = Pt(6)
    p_subj.paragraph_format.space_after = Pt(12)
    styled_run(p_subj, "Subject: Application for Full Stack Development - NextJS - Managed Services Engineer I (Requisition Code: 353063)", size=Pt(9.5), color=DARK, bold=True)

    # Salutation
    p_sal = doc.add_paragraph()
    p_sal.paragraph_format.space_after = Pt(8)
    styled_run(p_sal, "Dear Hiring Committee / Recruiting Team,", size=Pt(9.5), color=TEXT, bold=True)

    # Paragraph 1: Intro
    add_letter_paragraph(
        doc,
        "I am writing to express my strong interest in the Full Stack Development - NextJS - Managed Services Engineer I position at Deloitte Consulting, under the Digital Foundry Operate & Innovations (DFO&I) - Customer group. With over 1.3 years of experience as a Full Stack Developer at Savayas Life and Balance, coupled with a deep specialization in Next.js 15, React 19, and scalable API/operations management, I am confident in my ability to deliver high-quality digital products and manage ongoing services that elevate customer value at Deloitte."
    )

    # Paragraph 2: Technical depth
    p2_words = [
        ("During my tenure at Savayas Life and Balance, I successfully developed, optimized, and maintained two production-grade web applications using ", False),
        ("Next.js and React", True),
        (". I have extensive hands-on experience with Next.js rendering strategies (SSR, SSG, ISR), dynamic routing, and caching mechanisms. In one of my key performance optimization initiatives, I improved the company's ", False),
        ("Core Web Vitals scores by 40%", True),
        (" by leveraging runtime optimizations, bundle split techniques, and asset lazy-loading. Furthermore, my development workflows strictly align with accessibility standards, enabling me to achieve a ", False),
        ("92% accessibility compliance score", True),
        (" using semantic HTML and WCAG 2.1 / ARIA patterns.", False)
    ]
    add_letter_paragraph(doc, "", bold_words=p2_words)

    # Paragraph 3: Quality, API integration & AI assisted dev
    p3_words = [
        ("Quality assurance and robust system integrations form the core of my development practices. I am highly proficient in writing unit and integration tests using ", False),
        ("Jest and React Testing Library", True),
        (". I regularly integrate frontends with ", False),
        ("RESTful and GraphQL APIs", True),
        (", secure user authentication patterns using JWT and OAuth/OIDC, and state management solutions. In addition, I am well-versed in Git workflows and CI/CD pipelines, and I regularly use GenAI tools like ", False),
        ("GitHub Copilot", True),
        (" to accelerate feature delivery while maintaining strict secure coding practices and intellectual property rules.", False)
    ]
    add_letter_paragraph(doc, "", bold_words=p3_words)

    # Paragraph 4: Operations/Managed Services alignment
    p4_words = [
        ("The 'Managed Services' nature of this role perfectly aligns with my cross-functional skill set. Alongside my development capabilities, I possess practical experience in systems engineering, including ", False),
        ("Docker containerization, automated infrastructure management using Ansible, and Linux server operations", True),
        (". This dual capability as a Full Stack Developer and Systems Specialist allows me to not only build premium web applications but also troubleshoot production incidents, conduct Root Cause Analysis (RCA), and ensure long-term platform reliability and uptime.", False)
    ]
    add_letter_paragraph(doc, "", bold_words=p4_words)

    # Paragraph 5: Closing
    add_letter_paragraph(
        doc,
        "Deloitte's inclusive culture and focus on professional growth and making an impact that matters resonate deeply with me. I am excited about the opportunity to collaborate with your multidisciplinary teams of engineers, strategists, and designers to solve complex business problems. Thank you for your time and consideration. I look forward to the opportunity to discuss how my qualifications align with your team's objectives."
    )

    # Sign-off
    p_close = doc.add_paragraph()
    p_close.paragraph_format.space_before = Pt(12)
    p_close.paragraph_format.space_after = Pt(2)
    p_close.paragraph_format.line_spacing = Pt(11)
    styled_run(p_close, "Sincerely,\n", size=Pt(9.5))
    styled_run(p_close, "Navneet Vishwakarma", size=Pt(9.5), color=DARK, bold=True)

    doc.save(OUTPUT_FILE)
    print(f"Cover Letter saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_cover_letter()
