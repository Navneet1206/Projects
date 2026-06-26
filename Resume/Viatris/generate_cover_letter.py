"""
Viatris Application - Cover Letter Generator
Generates a professional cover letter .docx for IT Support / IT Administrator role at Viatris.
"""

import os
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Navneet_Vishwakarma_Cover_Letter.docx")

DARK = RGBColor(0x1A, 0x1A, 0x2E)
ACCENT = RGBColor(0x16, 0x53, 0x8D)
TEXT = RGBColor(0x2D, 0x2D, 0x2D)
MUTED = RGBColor(0x55, 0x55, 0x55)
RULE_COLOR = "165390"
FONT = "Calibri"


def styled_run(paragraph, text, size=Pt(10.5), color=TEXT,
               bold=False, italic=False):
    run = paragraph.add_run(text)
    run.font.name = FONT
    run.font.size = size
    run.font.color.rgb = color
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('w:rFonts'))
    if ea is not None and ea.get(qn('w:eastAsiaTheme')) is not None:
        del ea.attrib[qn('w:eastAsiaTheme')]
    return run


def add_thin_rule(doc):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(4)
    pf.line_spacing = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="8000" w:space="1" w:color="{RULE_COLOR}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def para(doc, text, size=Pt(10.5), color=TEXT, bold=False, italic=False,
         align=None, space_before=Pt(0), space_after=Pt(4), line_spacing=Pt(14)):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_before = space_before
    p.paragraph_format.space_after = space_after
    p.paragraph_format.line_spacing = line_spacing
    styled_run(p, text, size=size, color=color, bold=bold, italic=italic)
    return p


def bullet(doc, label, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = Pt(13)
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.first_line_indent = -Inches(0.15)
    styled_run(p, "▪  ", size=Pt(7), color=ACCENT)
    styled_run(p, f"{label}: ", size=Pt(10), bold=True)
    styled_run(p, text, size=Pt(10))


def build():
    doc = Document()

    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    style = doc.styles['Normal']
    style.font.name = FONT
    style.font.size = Pt(10.5)
    style.font.color.rgb = TEXT
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    # ── Header ────────────────────────────────────────────────
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_name.paragraph_format.space_after = Pt(2)
    styled_run(p_name, "NAVNEET VISHWAKARMA", size=Pt(16), color=ACCENT, bold=True)

    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_after = Pt(1)
    p_contact.paragraph_format.line_spacing = Pt(13)
    styled_run(p_contact, "+91-84350-61006  |  rajvl132011@gmail.com", size=Pt(9.5))

    p_loc = doc.add_paragraph()
    p_loc.paragraph_format.space_after = Pt(1)
    p_loc.paragraph_format.line_spacing = Pt(13)
    styled_run(p_loc, "Satna, Madhya Pradesh, India (Willing to Relocate)", size=Pt(9.5))

    p_links = doc.add_paragraph()
    p_links.paragraph_format.space_after = Pt(6)
    p_links.paragraph_format.line_spacing = Pt(13)
    styled_run(p_links, "LinkedIn: ", size=Pt(9.5))
    styled_run(p_links, "linkedin.com/in/navneet1206", size=Pt(9.5), color=ACCENT)
    styled_run(p_links, "  |  GitHub: ", size=Pt(9.5))
    styled_run(p_links, "github.com/navneetvishwakarma", size=Pt(9.5), color=ACCENT)

    add_thin_rule(doc)

    # ── Date ──────────────────────────────────────────────────
    para(doc, "June 26, 2026", space_before=Pt(6), space_after=Pt(10))

    # ── Recipient ─────────────────────────────────────────────
    para(doc, "Ms. Kajal Singh", bold=True, space_after=Pt(1))
    para(doc, "Viatris India", space_after=Pt(1))
    para(doc, "Email: kajal.singh@viatris.com", color=ACCENT, space_after=Pt(8))

    # ── Subject ───────────────────────────────────────────────
    p_subj = doc.add_paragraph()
    p_subj.paragraph_format.space_after = Pt(10)
    p_subj.paragraph_format.line_spacing = Pt(14)
    styled_run(p_subj, "Subject: ", bold=True)
    styled_run(p_subj, "Application for IT Support / IT Administrator Role – Referred by Mr. Chirag Sen")

    # ── Salutation ────────────────────────────────────────────
    para(doc, "Dear Ms. Singh,", space_after=Pt(8))

    # ── Body ──────────────────────────────────────────────────
    para(doc,
         "I am writing to express my strong interest in the IT Support / IT Administrator "
         "position at Viatris, as referred by Mr. Chirag Sen. As a Linux System Administrator "
         "with over 2 years of hands-on experience in infrastructure operations, technical "
         "support, and system troubleshooting, I am confident that my technical skills, "
         "dedication, and eagerness to learn make me a strong fit for this position.",
         space_after=Pt(8))

    # Why Viatris
    para(doc, "WHY VIATRIS?", size=Pt(10.5), color=ACCENT, bold=True, space_after=Pt(4))
    para(doc,
         "Viatris's commitment to improving healthcare access globally resonates deeply with "
         "my professional values. I am particularly drawn to the organization's emphasis on "
         "operational excellence and its India-based IT operations supporting worldwide "
         "healthcare delivery. The opportunity to contribute my technical expertise to a "
         "company where technology directly impacts patient outcomes is truly exciting to me.",
         space_after=Pt(8))

    # What I Bring
    para(doc, "WHAT I BRING TO THE ROLE:", size=Pt(10.5), color=ACCENT, bold=True, space_after=Pt(4))
    para(doc,
         "In my current role at Savayas Life and Balance Pvt. Ltd., I administer 15+ Linux "
         "servers (SLES 15, Ubuntu 22.04) maintaining 99.5% uptime across production and "
         "staging environments. My experience includes:",
         space_after=Pt(4))

    bullet(doc, "Incident Resolution",
           "Resolved 75+ P1/P2 operational incidents through systematic log analysis and "
           "root cause investigation, with an average resolution time under 2 hours.")
    bullet(doc, "Automation & Efficiency",
           "Built Ansible playbooks and Bash scripts that reduced manual administration "
           "effort by 40%, saving 15+ hours weekly and allowing the team to focus on "
           "strategic initiatives.")
    bullet(doc, "Security & Compliance",
           "Implemented CIS/NIST-aligned server hardening across the entire infrastructure, "
           "including SSH hardening, firewall enforcement, audit logging, and legacy service "
           "disablement.")
    bullet(doc, "End-User Support",
           "Provided L1/L2 technical support to 50+ end-users, achieving a 95% first-call "
           "resolution rate and ensuring minimal business disruption.")
    bullet(doc, "Documentation",
           "Maintained comprehensive operational documentation including runbooks, SOPs, "
           "incident reports, and knowledge base articles for team reference and continuity.")

    para(doc,
         "My technical toolkit includes SUSE Linux Enterprise Server, Ubuntu, Ansible, Docker, "
         "Bash scripting, TCP/IP networking, firewall management, and system monitoring – all "
         "directly applicable to maintaining Viatris's secure, reliable, and scalable IT "
         "infrastructure.",
         space_before=Pt(6), space_after=Pt(8))

    # Beyond Technical Skills
    para(doc, "BEYOND TECHNICAL SKILLS:", size=Pt(10.5), color=ACCENT, bold=True, space_after=Pt(4))
    para(doc,
         "I pride myself on clear communication, collaborative problem-solving, and a "
         "customer-first approach to IT support. I understand that in a healthcare-focused "
         "organization like Viatris, every minute of system downtime can impact critical "
         "operations. I am committed to ensuring maximum availability, rapid incident response, "
         "and proactive system monitoring.",
         space_after=Pt(8))

    para(doc,
         "If given the opportunity, I will give my best and contribute effectively to your "
         "team. I am eager to bring my skills in Linux administration, infrastructure "
         "automation, and technical support to support Viatris's mission of empowering people "
         "to live healthier at every stage of life.",
         space_after=Pt(6))

    para(doc,
         "I would be grateful for the opportunity to discuss my candidature further. Thank "
         "you for your time and consideration. I look forward to hearing from you.",
         space_after=Pt(12))

    # ── Sign-off ──────────────────────────────────────────────
    para(doc, "Warm regards,", space_after=Pt(10))

    para(doc, "Navneet Vishwakarma", bold=True, space_after=Pt(1))
    para(doc, "+91-84350-61006", size=Pt(9.5), space_after=Pt(1))
    para(doc, "rajvl132011@gmail.com", size=Pt(9.5), color=ACCENT, space_after=Pt(1))
    p_sl = doc.add_paragraph()
    p_sl.paragraph_format.space_after = Pt(1)
    p_sl.paragraph_format.line_spacing = Pt(13)
    styled_run(p_sl, "LinkedIn: ", size=Pt(9.5))
    styled_run(p_sl, "linkedin.com/in/navneet1206", size=Pt(9.5), color=ACCENT)
    p_sg = doc.add_paragraph()
    p_sg.paragraph_format.space_after = Pt(8)
    p_sg.paragraph_format.line_spacing = Pt(13)
    styled_run(p_sg, "GitHub: ", size=Pt(9.5))
    styled_run(p_sg, "github.com/navneetvishwakarma", size=Pt(9.5), color=ACCENT)

    para(doc, "Referred by: Mr. Chirag Sen", italic=True, color=MUTED, size=Pt(9.5))

    doc.save(OUTPUT_FILE)
    print(f"Cover Letter saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    build()
