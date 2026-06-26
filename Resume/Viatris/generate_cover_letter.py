"""
Viatris Application - Cover Letter (Single Page)
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
HIGHLIGHT_BG = "E8F0FE"  # light blue background for referral highlight
RULE_COLOR = "165390"
FONT = "Calibri"


def styled_run(paragraph, text, size=Pt(9.5), color=TEXT, bold=False, italic=False):
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
    pf.space_after = Pt(3)
    pf.line_spacing = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="6000" w:space="1" w:color="{RULE_COLOR}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def para(doc, text, size=Pt(9.5), color=TEXT, bold=False, italic=False,
         space_before=Pt(0), space_after=Pt(3), line_spacing=Pt(12.5)):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = space_before
    p.paragraph_format.space_after = space_after
    p.paragraph_format.line_spacing = line_spacing
    styled_run(p, text, size=size, color=color, bold=bold, italic=italic)
    return p


def bullet_item(doc, label, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0.5)
    p.paragraph_format.space_after = Pt(0.5)
    p.paragraph_format.line_spacing = Pt(11.5)
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = -Inches(0.15)
    styled_run(p, "▪  ", size=Pt(6), color=ACCENT)
    styled_run(p, f"{label}: ", size=Pt(9), bold=True)
    styled_run(p, text, size=Pt(9))


def build():
    doc = Document()

    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles['Normal']
    style.font.name = FONT
    style.font.size = Pt(9.5)
    style.font.color.rgb = TEXT
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    # ── Header ────────────────────────────────────────────────
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_after = Pt(1)
    styled_run(p_name, "NAVNEET VISHWAKARMA", size=Pt(14), color=ACCENT, bold=True)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_after = Pt(1)
    p_info.paragraph_format.line_spacing = Pt(12)
    styled_run(p_info, "+91-84350-61006  |  rajvl132011@gmail.com  |  Satna, MP, India (Willing to Relocate)", size=Pt(8.5))

    p_links = doc.add_paragraph()
    p_links.paragraph_format.space_after = Pt(4)
    p_links.paragraph_format.line_spacing = Pt(12)
    styled_run(p_links, "LinkedIn: ", size=Pt(8.5))
    styled_run(p_links, "linkedin.com/in/navneet1206", size=Pt(8.5), color=ACCENT)
    styled_run(p_links, "  |  GitHub: ", size=Pt(8.5))
    styled_run(p_links, "github.com/navneetvishwakarma", size=Pt(8.5), color=ACCENT)

    add_thin_rule(doc)

    # ── Date & Recipient (compact) ────────────────────────────
    para(doc, "June 26, 2026", space_after=Pt(5), size=Pt(9))

    p_to = doc.add_paragraph()
    p_to.paragraph_format.space_after = Pt(0)
    p_to.paragraph_format.line_spacing = Pt(12)
    styled_run(p_to, "Ms. Kajal Singh  |  ", size=Pt(9))
    styled_run(p_to, "Viatris India", size=Pt(9), bold=True)
    styled_run(p_to, "  |  kajal.singh@viatris.com", size=Pt(9), color=ACCENT)

    # ── Subject with referral highlight ───────────────────────
    p_subj = doc.add_paragraph()
    p_subj.paragraph_format.space_before = Pt(5)
    p_subj.paragraph_format.space_after = Pt(5)
    p_subj.paragraph_format.line_spacing = Pt(12)
    styled_run(p_subj, "Subject: ", size=Pt(9), bold=True)
    styled_run(p_subj, "Application for IT Support / IT Administrator Role – ", size=Pt(9))
    # Highlighted referral
    ref_run = p_subj.add_run("Referred by Mr. Chirag Sen")
    ref_run.font.name = FONT
    ref_run.font.size = Pt(9)
    ref_run.font.color.rgb = ACCENT
    ref_run.bold = True
    ref_run.italic = False
    # Add shading/highlight to the run
    rPr = ref_run._r.get_or_add_rPr()
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="{HIGHLIGHT_BG}"/>')
    rPr.append(shading)

    # ── Salutation ────────────────────────────────────────────
    para(doc, "Dear Ms. Singh,", space_after=Pt(4), size=Pt(9.5))

    # ── Body ──────────────────────────────────────────────────
    para(doc,
         "I am writing to express my strong interest in the IT Support / IT Administrator "
         "position at Viatris, as referred by Mr. Chirag Sen. As a Linux System Administrator "
         "with over 2 years of hands-on experience in infrastructure operations, technical "
         "support, and system troubleshooting, I am confident that my technical skills, "
         "dedication, and eagerness to learn make me a strong fit for this position.",
         size=Pt(9), space_after=Pt(4), line_spacing=Pt(12))

    # Why Viatris (compact)
    para(doc, "WHY VIATRIS?", size=Pt(9), color=ACCENT, bold=True, space_after=Pt(2))
    para(doc,
         "Viatris's commitment to improving healthcare access globally resonates deeply with "
         "my professional values. I am drawn to the organization's emphasis on operational "
         "excellence and its India-based IT operations supporting worldwide healthcare delivery. "
         "The opportunity to contribute where technology directly impacts patient outcomes is "
         "truly exciting.",
         size=Pt(9), space_after=Pt(4), line_spacing=Pt(12))

    # What I Bring
    para(doc, "WHAT I BRING TO THE ROLE:", size=Pt(9), color=ACCENT, bold=True, space_after=Pt(2))
    para(doc,
         "In my current role at Savayas Life and Balance Pvt. Ltd., I administer 15+ Linux "
         "servers (SLES 15, Ubuntu 22.04) maintaining 99.5% uptime. My experience includes:",
         size=Pt(9), space_after=Pt(2), line_spacing=Pt(12))

    bullet_item(doc, "Incident Resolution",
                "Resolved 75+ P1/P2 incidents via log analysis and RCA, avg resolution under 2 hours.")
    bullet_item(doc, "Automation",
                "Ansible playbooks and Bash scripts reduced manual effort by 40%, saving 15+ hrs/week.")
    bullet_item(doc, "Security",
                "CIS/NIST-aligned hardening: SSH, firewalls, audit logging, legacy service disablement.")
    bullet_item(doc, "End-User Support",
                "L1/L2 support to 50+ users with 95% first-call resolution rate.")
    bullet_item(doc, "Documentation",
                "Maintained runbooks, SOPs, incident reports, and KB articles.")

    # Beyond Technical
    para(doc, "BEYOND TECHNICAL SKILLS:", size=Pt(9), color=ACCENT, bold=True,
         space_before=Pt(3), space_after=Pt(2))
    para(doc,
         "I pride myself on clear communication, collaborative problem-solving, and a "
         "customer-first approach. In a healthcare organization like Viatris, every minute of "
         "uptime matters. I am committed to maximum availability, rapid incident response, "
         "and proactive monitoring.",
         size=Pt(9), space_after=Pt(3), line_spacing=Pt(12))

    para(doc,
         "I would be grateful for the opportunity to discuss my candidature further. Thank "
         "you for your time and consideration.",
         size=Pt(9), space_after=Pt(6), line_spacing=Pt(12))

    # ── Sign-off ──────────────────────────────────────────────
    para(doc, "Warm regards,", size=Pt(9.5), space_after=Pt(4))

    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_after = Pt(1)
    p_sign.paragraph_format.line_spacing = Pt(12)
    styled_run(p_sign, "Navneet Vishwakarma", size=Pt(9.5), bold=True)

    p_sc = doc.add_paragraph()
    p_sc.paragraph_format.space_after = Pt(1)
    p_sc.paragraph_format.line_spacing = Pt(11)
    styled_run(p_sc, "+91-84350-61006  |  rajvl132011@gmail.com", size=Pt(8.5))

    p_sl = doc.add_paragraph()
    p_sl.paragraph_format.space_after = Pt(4)
    p_sl.paragraph_format.line_spacing = Pt(11)
    styled_run(p_sl, "LinkedIn: ", size=Pt(8.5))
    styled_run(p_sl, "linkedin.com/in/navneet1206", size=Pt(8.5), color=ACCENT)
    styled_run(p_sl, "  |  GitHub: ", size=Pt(8.5))
    styled_run(p_sl, "github.com/navneetvishwakarma", size=Pt(8.5), color=ACCENT)

    # ── Referral highlight at bottom ──────────────────────────
    p_ref = doc.add_paragraph()
    p_ref.paragraph_format.space_before = Pt(2)
    p_ref.paragraph_format.space_after = Pt(0)
    p_ref.paragraph_format.line_spacing = Pt(12)
    ref_run2 = p_ref.add_run("  Referred by: Mr. Chirag Sen  ")
    ref_run2.font.name = FONT
    ref_run2.font.size = Pt(9.5)
    ref_run2.font.color.rgb = ACCENT
    ref_run2.bold = True
    rPr2 = ref_run2._r.get_or_add_rPr()
    shading2 = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="{HIGHLIGHT_BG}"/>')
    rPr2.append(shading2)

    doc.save(OUTPUT_FILE)
    print(f"Cover Letter saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    build()
