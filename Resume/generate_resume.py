"""
Resume Generator - Navneet Vishwakarma
Generates a single-page, ATS-optimized .docx resume.
"""

import os
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Navneet_Vishwakarma_Resume.docx")

# ── Color Palette ──────────────────────────────────────────────
DARK = RGBColor(0x1A, 0x1A, 0x2E)       # deep navy for headings
ACCENT = RGBColor(0x16, 0x53, 0x8D)     # strong blue for name/links
TEXT = RGBColor(0x2D, 0x2D, 0x2D)       # near-black body text
MUTED = RGBColor(0x55, 0x55, 0x55)      # grey for dates/locations
RULE_COLOR = "165390"                    # hex for decorative lines

# ── Font Names ─────────────────────────────────────────────────
FONT_HEADING = "Calibri"
FONT_BODY = "Calibri"


def set_cell_shading(cell, color_hex):
    """Apply background shading to a table cell."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def remove_paragraph_spacing(paragraph):
    """Zero out all spacing on a paragraph for tight layout."""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(11.5)


def add_thin_rule(doc, color=RULE_COLOR, thickness="8000"):
    """Insert a thin horizontal rule below the last paragraph."""
    p = doc.add_paragraph()
    remove_paragraph_spacing(p)
    p.paragraph_format.space_after = Pt(3)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="{thickness}" w:space="1" w:color="{color}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def styled_run(paragraph, text, font_name=FONT_BODY, size=Pt(10), color=TEXT,
               bold=False, italic=False):
    """Add a styled run to an existing paragraph."""
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = size
    run.font.color.rgb = color
    run.bold = bold
    run.italic = italic
    # Clear East-Asian fallback theme if present
    rPr = run._r.get_or_add_rPr()
    east_asia = rPr.find(qn('w:rFonts'))
    if east_asia is not None and east_asia.get(qn('w:eastAsiaTheme')) is not None:
        del east_asia.attrib[qn('w:eastAsiaTheme')]
    return run


def add_section_heading(doc, title):
    """Add a styled section heading with a thin rule underneath."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = Pt(13)
    styled_run(p, title.upper(), font_name=FONT_HEADING, size=Pt(10.5),
               color=ACCENT, bold=True)
    add_thin_rule(doc)


def add_bullet(doc, text, indent=Inches(0.25)):
    """Add an ATS-safe bullet point (plain dash prefix, no list style)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0.5)
    p.paragraph_format.space_after = Pt(0.5)
    p.paragraph_format.line_spacing = Pt(11)
    p.paragraph_format.left_indent = indent
    p.paragraph_format.first_line_indent = -Inches(0.15)
    styled_run(p, "▪  ", size=Pt(7), color=ACCENT)
    styled_run(p, text, size=Pt(9))


def build_resume():
    doc = Document()

    # ── Page Setup ─────────────────────────────────────────────
    section = doc.sections[0]
    section.page_width = Inches(8.27)   # A4
    section.page_height = Inches(11.69)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(0.8)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    # Remove default paragraph style spacing
    style = doc.styles['Normal']
    style.font.name = FONT_BODY
    style.font.size = Pt(10)
    style.font.color.rgb = TEXT
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    # ════════════════════════════════════════════════════════════
    # HEADER — Name & Contact
    # ════════════════════════════════════════════════════════════
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_after = Pt(2)
    p_name.paragraph_format.line_spacing = Pt(22)
    styled_run(p_name, "NAVNEET VISHWAKARMA", font_name=FONT_HEADING,
               size=Pt(18), color=ACCENT, bold=True)

    # Title line
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(3)
    p_title.paragraph_format.line_spacing = Pt(12)
    styled_run(p_title,
               "Linux Support Engineer  |  Linux Administration  |  SUSE Linux  |  Ansible  |  Technical Support",
               size=Pt(8.5), color=MUTED, bold=False)

    # Contact line
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_after = Pt(1)
    p_contact.paragraph_format.line_spacing = Pt(11)
    styled_run(p_contact, "+91 8435061006  |  Rajvl132011@gmail.com  |  Satna, Madhya Pradesh, India",
               size=Pt(8.5), color=TEXT)

    # Links line
    p_links = doc.add_paragraph()
    p_links.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_links.paragraph_format.space_after = Pt(2)
    p_links.paragraph_format.line_spacing = Pt(11)
    styled_run(p_links, "LinkedIn: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "linkedin.com/in/navneet1206", size=Pt(8.5), color=ACCENT)
    styled_run(p_links, "  |  GitHub: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "github.com/Navneet1206", size=Pt(8.5), color=ACCENT)

    add_thin_rule(doc, thickness="12000")

    # ════════════════════════════════════════════════════════════
    # PROFESSIONAL SUMMARY
    # ════════════════════════════════════════════════════════════
    add_section_heading(doc, "Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(1)
    p_sum.paragraph_format.line_spacing = Pt(11.5)
    styled_run(
        p_sum,
        "Computer Science Engineering student with hands-on experience in Linux Administration, "
        "Technical Support, Infrastructure Operations, and System Troubleshooting. Skilled in managing "
        "SUSE Linux Enterprise Server (SLES), Ubuntu, and openSUSE environments. Proficient in Bash "
        "Scripting, Ansible automation, Docker containerization, system monitoring, log analysis, and "
        "Root Cause Analysis (RCA). Demonstrated ability to resolve P1/P2 production incidents, harden "
        "server configurations to CIS/NIST baselines, and maintain high service availability. Seeking "
        "opportunities as a Linux Support Engineer, Linux Administrator, Technical Support Engineer, "
        "Infrastructure Support Engineer, NOC Engineer, or Junior DevOps Engineer.",
        size=Pt(9)
    )

    # ════════════════════════════════════════════════════════════
    # TECHNICAL SKILLS
    # ════════════════════════════════════════════════════════════
    add_section_heading(doc, "Technical Skills")

    skills_data = [
        ("Operating Systems", "SLES 12/15, Ubuntu Linux, openSUSE Leap, Windows Server"),
        ("Linux Administration", "User & Group Mgmt, File Systems, Package Mgmt, Process & Service Control, Security, Performance Monitoring"),
        ("Technical Support", "Troubleshooting, Incident Resolution, Root Cause Analysis, Log Analysis, Production & Infrastructure Support"),
        ("Automation & IaC", "Bash Shell Scripting, Ansible (Playbooks, Roles, Handlers)"),
        ("Containers & VCS", "Docker, Git, GitHub"),
        ("Networking", "TCP/IP, DNS, DHCP, SSH, HTTP/HTTPS, Firewall (UFW, firewalld)"),
    ]

    for label, value in skills_data:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(0.5)
        p.paragraph_format.line_spacing = Pt(11)
        p.paragraph_format.left_indent = Inches(0.1)
        styled_run(p, f"{label}:  ", size=Pt(9), color=DARK, bold=True)
        styled_run(p, value, size=Pt(9))

    # ════════════════════════════════════════════════════════════
    # PROFESSIONAL EXPERIENCE
    # ════════════════════════════════════════════════════════════
    add_section_heading(doc, "Professional Experience")

    # Job title line
    p_job = doc.add_paragraph()
    p_job.paragraph_format.space_before = Pt(2)
    p_job.paragraph_format.space_after = Pt(0)
    p_job.paragraph_format.line_spacing = Pt(12)
    styled_run(p_job, "Linux System Administrator", size=Pt(10), color=DARK, bold=True)

    # Company & date line
    p_co = doc.add_paragraph()
    p_co.paragraph_format.space_before = Pt(0)
    p_co.paragraph_format.space_after = Pt(2)
    p_co.paragraph_format.line_spacing = Pt(11)
    styled_run(p_co, "Savayas Life and Balance Pvt. Ltd.", size=Pt(9), color=TEXT, italic=True)
    styled_run(p_co, "  |  March 2024 – Present  |  Satna, MP", size=Pt(8.5), color=MUTED)

    bullets_exp = [
        "Administered 10+ Linux servers (SLES 15, Ubuntu 22.04) ensuring 99.5% uptime across production and staging environments.",
        "Resolved 50+ operational incidents through systematic log analysis (/var/log/messages, journalctl) and root cause investigation.",
        "Managed user accounts, sudo policies, file permissions, and SSH access controls across multi-server infrastructure.",
        "Automated routine system administration tasks using Bash scripts and Ansible playbooks, reducing manual effort by 40%.",
        "Deployed and maintained Docker-based application stacks, managing container lifecycle and resource allocation.",
        "Implemented server hardening measures aligned with CIS benchmarks: disabled root SSH, enforced key-based auth, configured firewall rules.",
        "Maintained operational documentation including runbooks, SOPs, and incident reports for knowledge transfer.",
    ]
    for b in bullets_exp:
        add_bullet(doc, b)

    # ════════════════════════════════════════════════════════════
    # PROJECTS
    # ════════════════════════════════════════════════════════════
    add_section_heading(doc, "Projects")

    # Project 1
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(2)
    p_p1.paragraph_format.space_after = Pt(1)
    p_p1.paragraph_format.line_spacing = Pt(12)
    styled_run(p_p1, "Linux Server Hardening with Ansible", size=Pt(9.5), color=DARK, bold=True)
    styled_run(p_p1, "  |  ", size=Pt(8.5), color=MUTED)
    styled_run(p_p1, "github.com/Navneet1206/Projects/.../Navneet-Server-Hardning",
               size=Pt(8), color=ACCENT)

    proj1_bullets = [
        "Built 6 Ansible roles to automate SSH hardening, firewall enforcement (UFW/firewalld), audit logging, and compliance scanning across Ubuntu and openSUSE targets.",
        "Enforced CIS/NIST-aligned security baselines: disabled legacy services (Telnet, FTP, rsh), configured PAM password policies, and enabled persistent journald logging.",
        "Integrated Lynis security auditing framework for automated post-hardening compliance verification.",
    ]
    for b in proj1_bullets:
        add_bullet(doc, b)

    # Project 2
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(4)
    p_p2.paragraph_format.space_after = Pt(1)
    p_p2.paragraph_format.line_spacing = Pt(12)
    styled_run(p_p2, "HA Cluster Incident Investigation & Root Cause Analysis", size=Pt(9.5), color=DARK, bold=True)
    styled_run(p_p2, "  |  ", size=Pt(8.5), color=MUTED)
    styled_run(p_p2, "github.com/Navneet1206/Projects/.../RCA_project",
               size=Pt(8), color=ACCENT)

    proj2_bullets = [
        "Conducted full Root Cause Analysis on a P1 outage in a 2-node SLES HA cluster (Pacemaker/Corosync/SBD), diagnosing a 5x watchdog reboot loop caused by network isolation and fencing failure.",
        "Analyzed supportconfig archives, correlated corosync token timeouts, ARP table anomalies, and SBD watchdog configuration mismatches to pinpoint cascading failure chain.",
        "Documented detailed incident timeline, technical findings, and remediation plan including Corosync QDevice and dual-ring network recommendations.",
    ]
    for b in proj2_bullets:
        add_bullet(doc, b)

    # ════════════════════════════════════════════════════════════
    # CERTIFICATIONS
    # ════════════════════════════════════════════════════════════
    add_section_heading(doc, "Certifications")

    certs = [
        "Linux Administration Training (CX-501) – Codenixia",
        "Linux Server Hardening & Security Training (CX-701) – Codenixia",
        "SUSE Linux Enterprise Server (SLES) Administration Training",
    ]
    for c in certs:
        add_bullet(doc, c)

    # ════════════════════════════════════════════════════════════
    # EDUCATION
    # ════════════════════════════════════════════════════════════
    add_section_heading(doc, "Education")

    p_deg = doc.add_paragraph()
    p_deg.paragraph_format.space_before = Pt(2)
    p_deg.paragraph_format.space_after = Pt(0)
    p_deg.paragraph_format.line_spacing = Pt(12)
    styled_run(p_deg, "Bachelor of Technology (B.Tech) – Computer Science & Engineering",
               size=Pt(9.5), color=DARK, bold=True)

    p_uni = doc.add_paragraph()
    p_uni.paragraph_format.space_before = Pt(0)
    p_uni.paragraph_format.space_after = Pt(1)
    p_uni.paragraph_format.line_spacing = Pt(11)
    styled_run(p_uni, "AKS University, Satna, MP", size=Pt(9), color=TEXT, italic=True)
    styled_run(p_uni, "  |  2022 – 2026  |  CGPA: 8.29 / 10", size=Pt(8.5), color=MUTED)

    # ── Save ───────────────────────────────────────────────────
    doc.save(OUTPUT_FILE)
    print(f"Resume saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_resume()
