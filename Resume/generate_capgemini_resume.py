"""
Resume Generator - Navneet Vishwakarma
Tailored for: Capgemini DevOps Engineer (Job ID: 461808-en_GB)
100% honest content — no fabricated skills or inflated numbers.
"""

import os
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Navneet_Vishwakarma_DevOps_Capgemini.docx")

DARK = RGBColor(0x1A, 0x1A, 0x2E)
ACCENT = RGBColor(0x16, 0x53, 0x8D)
TEXT = RGBColor(0x2D, 0x2D, 0x2D)
MUTED = RGBColor(0x55, 0x55, 0x55)
RULE_COLOR = "165390"
FONT_HEADING = "Calibri"
FONT_BODY = "Calibri"


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


def styled_run(paragraph, text, font_name=FONT_BODY, size=Pt(10), color=TEXT,
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


def add_section_heading(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = Pt(12)
    styled_run(p, title.upper(), font_name=FONT_HEADING, size=Pt(10),
               color=ACCENT, bold=True)
    add_thin_rule(doc)


def add_bullet(doc, text, indent=Inches(0.25)):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0.5)
    p.paragraph_format.space_after = Pt(0.5)
    p.paragraph_format.line_spacing = Pt(10.8)
    p.paragraph_format.left_indent = indent
    p.paragraph_format.first_line_indent = -Inches(0.15)
    styled_run(p, "▪  ", size=Pt(7), color=ACCENT)
    styled_run(p, text, size=Pt(8.5))


def build_resume():
    doc = Document()

    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(0.7)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    style = doc.styles['Normal']
    style.font.name = FONT_BODY
    style.font.size = Pt(10)
    style.font.color.rgb = TEXT
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    # ═══════════════════════════════════════════════════════════
    # HEADER
    # ═══════════════════════════════════════════════════════════
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_after = Pt(2)
    p_name.paragraph_format.line_spacing = Pt(20)
    styled_run(p_name, "NAVNEET VISHWAKARMA", font_name=FONT_HEADING,
               size=Pt(17), color=ACCENT, bold=True)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.line_spacing = Pt(11)
    styled_run(p_title,
               "Linux Administrator  |  DevOps Engineer  |  CI/CD  |  Ansible  |  Docker",
               size=Pt(8.5), color=MUTED)

    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_after = Pt(1)
    p_contact.paragraph_format.line_spacing = Pt(11)
    styled_run(p_contact, "+91 8435061006  |  Rajvl132011@gmail.com  |  Satna, Madhya Pradesh, India",
               size=Pt(8.5), color=TEXT)

    p_links = doc.add_paragraph()
    p_links.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_links.paragraph_format.space_after = Pt(2)
    p_links.paragraph_format.line_spacing = Pt(11)
    styled_run(p_links, "LinkedIn: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "linkedin.com/in/navneet1206", size=Pt(8.5), color=ACCENT)
    styled_run(p_links, "  |  GitHub: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "github.com/Navneet1206", size=Pt(8.5), color=ACCENT)

    add_thin_rule(doc, thickness="12000")

    # ═══════════════════════════════════════════════════════════
    # PROFESSIONAL SUMMARY
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(1)
    p_sum.paragraph_format.line_spacing = Pt(11)
    styled_run(
        p_sum,
        "Linux Administrator and aspiring DevOps Engineer with hands-on experience managing "
        "Linux servers (SLES, Ubuntu), deploying Docker containers, building CI/CD pipelines "
        "using GitHub Actions, and automating infrastructure with Ansible and Bash scripting. "
        "Experienced in server hardening, incident troubleshooting, log analysis, Root Cause "
        "Analysis (RCA), backup and recovery operations, and maintaining operational documentation. "
        "Familiar with Kubernetes fundamentals and KVM virtualization. Holds a B.Tech in Computer "
        "Science and is eager to contribute to platform reliability, continuous delivery, and "
        "infrastructure automation in a DevOps/SRE environment.",
        size=Pt(8.5)
    )

    # ═══════════════════════════════════════════════════════════
    # TECHNICAL SKILLS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Technical Skills")

    skills_data = [
        ("CI/CD & Version Control",
         "GitHub Actions (Pipeline Build & Maintenance), Git, GitHub"),
        ("Containers & Orchestration",
         "Docker (Build, Compose, Deployment), Kubernetes (Basic: Pods, Deployments, kubectl)"),
        ("Infrastructure as Code",
         "Ansible (Playbooks, Roles, Handlers), Bash Shell Scripting, Cron Scheduling"),
        ("Operating Systems",
         "SUSE Linux Enterprise Server (SLES), Ubuntu Server, openSUSE Leap"),
        ("Virtualization",
         "KVM"),
        ("Linux Administration",
         "User & Group Mgmt, File Systems, Package Mgmt (zypper, apt), Service Control (systemd), "
         "SSH Hardening, Firewall (firewalld, UFW), auditd, Log Analysis (journalctl, /var/log)"),
        ("Incident & Reliability",
         "Incident Troubleshooting, Root Cause Analysis (RCA), Backup & Recovery, "
         "Uptime Monitoring, Runbooks/SOPs"),
        ("Networking",
         "TCP/IP, DNS, DHCP, SSH, HTTP/HTTPS, Network Troubleshooting"),
    ]

    for label, value in skills_data:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0.5)
        p.paragraph_format.space_after = Pt(0.5)
        p.paragraph_format.line_spacing = Pt(10.5)
        p.paragraph_format.left_indent = Inches(0.1)
        styled_run(p, f"{label}:  ", size=Pt(8.5), color=DARK, bold=True)
        styled_run(p, value, size=Pt(8.5))

    # ═══════════════════════════════════════════════════════════
    # PROFESSIONAL EXPERIENCE
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Professional Experience")

    p_job = doc.add_paragraph()
    p_job.paragraph_format.space_before = Pt(2)
    p_job.paragraph_format.space_after = Pt(0)
    p_job.paragraph_format.line_spacing = Pt(12)
    styled_run(p_job, "Linux System Administrator",
               size=Pt(9.5), color=DARK, bold=True)

    p_co = doc.add_paragraph()
    p_co.paragraph_format.space_before = Pt(0)
    p_co.paragraph_format.space_after = Pt(2)
    p_co.paragraph_format.line_spacing = Pt(11)
    styled_run(p_co, "Savayas Life and Balance Pvt. Ltd.", size=Pt(9), color=TEXT, italic=True)
    styled_run(p_co, "  |  March 2024 – Present  |  Satna, MP", size=Pt(8.5), color=MUTED)

    bullets_exp = [
        "Manage 5 Linux servers (SLES, Ubuntu) across production and staging environments, "
        "handling server provisioning, updates, and day-to-day infrastructure operations.",

        "Built and maintain CI/CD pipelines using GitHub Actions for automated build and "
        "deployment workflows of web applications.",

        "Deploy and manage Docker-based containerized application stacks, handling container "
        "lifecycle, image builds, and multi-service compositions.",

        "Troubleshoot and resolve production incidents through systematic log analysis "
        "(journalctl, /var/log/messages) and root cause investigation; resolved 15+ operational "
        "issues since joining.",

        "Automate routine server administration tasks using Ansible playbooks and Bash scripts, "
        "covering user provisioning, package updates, and configuration management.",

        "Implement server hardening measures: SSH key-based authentication, firewall rules "
        "(firewalld/UFW), audit logging (auditd), and service lockdown following security best practices.",

        "Manage backup and recovery operations for application data and server configurations "
        "to ensure data protection and restoration readiness.",

        "Maintain operational documentation including runbooks, SOPs, and incident reports "
        "for team knowledge sharing and repeatable processes.",
    ]
    for b in bullets_exp:
        add_bullet(doc, b)

    # ═══════════════════════════════════════════════════════════
    # PROJECTS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Key Projects")

    # Project 1
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(2)
    p_p1.paragraph_format.space_after = Pt(1)
    p_p1.paragraph_format.line_spacing = Pt(11)
    styled_run(p_p1, "Infrastructure Hardening Automation with Ansible (IaC)",
               size=Pt(9), color=DARK, bold=True)
    styled_run(p_p1, "  |  ", size=Pt(8), color=MUTED)
    styled_run(p_p1, "github.com/Navneet1206/Projects/.../Navneet-Server-Hardning",
               size=Pt(7.5), color=ACCENT)

    proj1_bullets = [
        "Wrote 6 Ansible roles to automate SSH hardening, firewall enforcement, audit logging, "
        "legacy service removal, and compliance scanning across Ubuntu and openSUSE servers.",

        "Applied CIS/NIST-aligned security configurations: PAM password policies, persistent "
        "journald logging, and automated patch scheduling. Integrated Lynis for compliance verification.",
    ]
    for b in proj1_bullets:
        add_bullet(doc, b)

    # Project 2
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(3)
    p_p2.paragraph_format.space_after = Pt(1)
    p_p2.paragraph_format.line_spacing = Pt(11)
    styled_run(p_p2, "HA Cluster P1 Incident Investigation & Root Cause Analysis",
               size=Pt(9), color=DARK, bold=True)
    styled_run(p_p2, "  |  ", size=Pt(8), color=MUTED)
    styled_run(p_p2, "github.com/Navneet1206/Projects/.../RCA_project",
               size=Pt(7.5), color=ACCENT)

    proj2_bullets = [
        "Conducted Root Cause Analysis on a P1 outage in a 2-node SLES HA cluster "
        "(Pacemaker/Corosync/SBD), troubleshooting cascading failures across networking, "
        "fencing, and watchdog layers that triggered a 5x reboot loop.",

        "Analyzed supportconfig archives and system logs to reconstruct event timeline; "
        "delivered incident report with actionable remediation steps.",
    ]
    for b in proj2_bullets:
        add_bullet(doc, b)

    # ═══════════════════════════════════════════════════════════
    # CERTIFICATIONS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Certifications")

    certs = [
        "Linux Administration Training (CX-501) – Codenixia",
        "Linux Server Hardening & Security Training (CX-701) – Codenixia",
        "SUSE Linux Enterprise Server (SLES) Administration Training",
    ]
    for c in certs:
        add_bullet(doc, c)

    # ═══════════════════════════════════════════════════════════
    # EDUCATION
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Education")

    p_deg = doc.add_paragraph()
    p_deg.paragraph_format.space_before = Pt(2)
    p_deg.paragraph_format.space_after = Pt(0)
    p_deg.paragraph_format.line_spacing = Pt(11)
    styled_run(p_deg, "Bachelor of Technology (B.Tech) – Computer Science & Engineering",
               size=Pt(9), color=DARK, bold=True)

    p_uni = doc.add_paragraph()
    p_uni.paragraph_format.space_before = Pt(0)
    p_uni.paragraph_format.space_after = Pt(1)
    p_uni.paragraph_format.line_spacing = Pt(11)
    styled_run(p_uni, "AKS University, Satna, MP", size=Pt(8.5), color=TEXT, italic=True)
    styled_run(p_uni, "  |  2022 – 2026  |  CGPA: 8.29 / 10", size=Pt(8.5), color=MUTED)

    doc.save(OUTPUT_FILE)
    print(f"Resume saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_resume()
