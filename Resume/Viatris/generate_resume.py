"""
Viatris Application - Resume Generator
Generates a professional ATS-optimized .docx resume for IT Support / IT Administrator at Viatris.
"""

import os
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Navneet_Vishwakarma_Resume.docx")

DARK = RGBColor(0x1A, 0x1A, 0x2E)
ACCENT = RGBColor(0x16, 0x53, 0x8D)
TEXT = RGBColor(0x2D, 0x2D, 0x2D)
MUTED = RGBColor(0x55, 0x55, 0x55)
RULE_COLOR = "165390"
FONT = "Calibri"


def remove_paragraph_spacing(paragraph):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(11.5)


def add_thin_rule(doc, thickness="8000"):
    p = doc.add_paragraph()
    remove_paragraph_spacing(p)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="{thickness}" w:space="1" w:color="{RULE_COLOR}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def styled_run(paragraph, text, size=Pt(10), color=TEXT, bold=False, italic=False):
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


def add_section_heading(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = Pt(12)
    styled_run(p, title.upper(), size=Pt(10), color=ACCENT, bold=True)
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


def build():
    doc = Document()

    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Cm(0.9)
    section.bottom_margin = Cm(0.7)
    section.left_margin = Cm(1.4)
    section.right_margin = Cm(1.4)

    style = doc.styles['Normal']
    style.font.name = FONT
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
    styled_run(p_name, "NAVNEET VISHWAKARMA", size=Pt(17), color=ACCENT, bold=True)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.line_spacing = Pt(11)
    styled_run(p_title,
               "IT Support Engineer  |  Linux Administrator  |  Infrastructure Operations  |  Technical Support",
               size=Pt(8.5), color=MUTED)

    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_after = Pt(1)
    p_contact.paragraph_format.line_spacing = Pt(11)
    styled_run(p_contact,
               "+91-84350-61006  |  rajvl132011@gmail.com  |  Satna, MP, India (Willing to Relocate)",
               size=Pt(8.5))

    p_links = doc.add_paragraph()
    p_links.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_links.paragraph_format.space_after = Pt(2)
    p_links.paragraph_format.line_spacing = Pt(11)
    styled_run(p_links, "LinkedIn: ", size=Pt(8.5))
    styled_run(p_links, "linkedin.com/in/navneet1206", size=Pt(8.5), color=ACCENT)
    styled_run(p_links, "  |  GitHub: ", size=Pt(8.5))
    styled_run(p_links, "github.com/navneetvishwakarma", size=Pt(8.5), color=ACCENT)

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
        "IT Support & Linux System Administrator with 2+ years of hands-on experience in "
        "Linux Administration, Infrastructure Operations, and Technical Support. Skilled in "
        "managing SUSE Linux Enterprise Server (SLES), Ubuntu, and Windows Server environments. "
        "Proficient in Bash Scripting, Ansible automation, Docker containerization, system "
        "monitoring, log analysis, and Root Cause Analysis (RCA). Demonstrated ability to "
        "resolve P1/P2 production incidents, implement CIS/NIST security baselines, and "
        "maintain 99.5% service availability. Seeking to contribute to Viatris's IT operations "
        "in delivering secure, reliable, and scalable infrastructure support.",
        size=Pt(8.5)
    )

    # ═══════════════════════════════════════════════════════════
    # TECHNICAL SKILLS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Technical Skills")

    skills = [
        ("Operating Systems",
         "SLES 12/15, Ubuntu Linux, openSUSE Leap, Windows Server 2016/2019"),
        ("Linux Administration",
         "User & Group Mgmt, File Systems, Package Mgmt (RPM, APT), Service Control (systemd), "
         "Security Hardening, Performance Monitoring (top, htop, vmstat, iostat, sar)"),
        ("IT Support",
         "Incident Management, Root Cause Analysis, Log Analysis, Ticketing Systems, "
         "Hardware/Software Troubleshooting, Remote Support, End-User Support"),
        ("Automation & Scripting",
         "Bash Shell Scripting, Ansible (Playbooks, Roles, Handlers, Inventory), "
         "Cron Jobs, Systemd Timers"),
        ("Virtualization",
         "VMware vSphere, KVM"),
        ("Containers",
         "Docker, Docker Compose, Container Lifecycle Management"),
        ("Networking",
         "TCP/IP, DNS, DHCP, SSH, HTTP/HTTPS, Firewall (UFW, firewalld, iptables), "
         "VLANs, NAT, Port Forwarding"),
        ("Security & Compliance",
         "CIS Benchmarks, NIST Guidelines, Server Hardening, Audit Logging, "
         "Lynis Auditing, PAM Configuration, SSH Key-Based Auth"),
        ("Monitoring & Logging",
         "journalctl, rsyslog, /var/log analysis, Supportconfig"),
        ("Version Control",
         "Git, GitHub"),
    ]

    for label, value in skills:
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
    styled_run(p_job, "Linux System Administrator", size=Pt(9.5), color=DARK, bold=True)

    p_co = doc.add_paragraph()
    p_co.paragraph_format.space_before = Pt(0)
    p_co.paragraph_format.space_after = Pt(2)
    p_co.paragraph_format.line_spacing = Pt(11)
    styled_run(p_co, "Savayas Life and Balance Pvt. Ltd.", size=Pt(9), italic=True)
    styled_run(p_co, "  |  March 2024 – Present  |  Satna, Madhya Pradesh", size=Pt(8.5), color=MUTED)

    exp_bullets = [
        "Administered 15+ Linux servers (SLES 15, Ubuntu 22.04) ensuring 99.5% uptime across "
        "production and staging environments for business-critical applications.",

        "Resolved 75+ P1/P2 operational incidents through systematic log analysis "
        "(/var/log/messages, journalctl, supportconfig) and root cause investigation, "
        "achieving average resolution time under 2 hours.",

        "Provided L1/L2 technical support to 50+ end-users, resolving hardware, software, "
        "and network issues with 95% first-call resolution rate.",

        "Managed user accounts, sudo policies, file permissions, and SSH access controls "
        "across multi-server infrastructure; enforced key-based authentication and disabled "
        "root SSH access.",

        "Automated routine system administration tasks using Bash scripts and Ansible "
        "playbooks, reducing manual effort by 40% and saving 15+ hours per week.",

        "Deployed and maintained Docker-based application stacks, managing container "
        "lifecycle, resource allocation, and network configuration.",

        "Implemented server hardening measures aligned with CIS benchmarks: disabled legacy "
        "services (Telnet, FTP, rsh), configured PAM password policies, enforced firewall "
        "rules (UFW/firewalld), and enabled persistent journald logging.",

        "Maintained comprehensive operational documentation including runbooks, SOPs, "
        "incident reports, and knowledge base articles for team reference and training.",

        "Collaborated with cross-functional teams to deploy system patches, security "
        "updates, and configuration changes with zero downtime.",
    ]
    for b in exp_bullets:
        add_bullet(doc, b)

    # ═══════════════════════════════════════════════════════════
    # PROJECTS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Projects")

    # Project 1
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(2)
    p_p1.paragraph_format.space_after = Pt(1)
    p_p1.paragraph_format.line_spacing = Pt(11)
    styled_run(p_p1, "Automated Linux Server Hardening with Ansible",
               size=Pt(9), color=DARK, bold=True)

    p1_bullets = [
        "Built 6 Ansible roles to automate SSH hardening, firewall enforcement (UFW/firewalld), "
        "audit logging, and compliance scanning across Ubuntu and openSUSE target servers.",

        "Enforced CIS/NIST-aligned security baselines: disabled legacy services (Telnet, FTP, "
        "rsh), configured PAM password policies, and enabled persistent journald logging.",

        "Integrated Lynis security auditing framework for automated post-hardening compliance "
        "validation and reporting.",

        "Result: Reduced security audit preparation time by 60% and ensured 100% compliance "
        "across all managed servers.",
    ]
    for b in p1_bullets:
        add_bullet(doc, b)

    # Project 2
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(3)
    p_p2.paragraph_format.space_after = Pt(1)
    p_p2.paragraph_format.line_spacing = Pt(11)
    styled_run(p_p2, "High-Availability Cluster Incident Investigation & Root Cause Analysis",
               size=Pt(9), color=DARK, bold=True)

    p2_bullets = [
        "Conducted full Root Cause Analysis on a P1 outage in a 2-node SLES HA cluster "
        "(Pacemaker/Corosync/SBD), diagnosing a 5x watchdog reboot loop caused by network "
        "isolation and fencing failure.",

        "Analyzed supportconfig archives, correlated corosync token timeouts, ARP table "
        "anomalies, and SBD watchdog configuration mismatches to pinpoint the cascading "
        "failure chain.",

        "Documented detailed incident timeline, technical findings, and remediation plan "
        "including Corosync QDevice and dual-ring network architecture recommendations.",

        "Result: Prevented recurrence of similar outages and improved cluster stability by 80%.",
    ]
    for b in p2_bullets:
        add_bullet(doc, b)

    # ═══════════════════════════════════════════════════════════
    # CERTIFICATIONS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Certifications")
    certs = [
        "Linux Administration Training (CX-501) – Codenixia",
        "Linux Server Hardening & Security Training (CX-701) – Codenixia",
        "SUSE Linux Enterprise Server (SLES) Administration Training – Codenixia",
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
    styled_run(p_uni, "AKS University, Satna, Madhya Pradesh", size=Pt(8.5), italic=True)
    styled_run(p_uni, "  |  2022 – 2026  |  CGPA: 8.29 / 10", size=Pt(8.5), color=MUTED)

    p_cw = doc.add_paragraph()
    p_cw.paragraph_format.space_before = Pt(1)
    p_cw.paragraph_format.space_after = Pt(1)
    p_cw.paragraph_format.line_spacing = Pt(10.5)
    p_cw.paragraph_format.left_indent = Inches(0.1)
    styled_run(p_cw, "Relevant Coursework:  ", size=Pt(8.5), color=DARK, bold=True)
    styled_run(p_cw,
               "Operating Systems, Computer Networks, Database Management Systems, "
               "Information Security, System Administration, Data Structures & Algorithms",
               size=Pt(8.5))

    # ═══════════════════════════════════════════════════════════
    # ACHIEVEMENTS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Achievements")
    achievements = [
        "Maintained 99.5% server uptime across 15+ production Linux servers over 12+ months.",
        "Resolved 75+ P1/P2 incidents with average resolution time under 2 hours.",
        "Automated 40% of routine tasks, saving 15+ manual hours per week.",
        "Implemented CIS-compliant security hardening across entire server fleet.",
        "Achieved 95% first-call resolution rate for end-user technical support requests.",
    ]
    for a in achievements:
        add_bullet(doc, a)

    # ═══════════════════════════════════════════════════════════
    # LANGUAGES & REFERENCES
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Languages")
    p_lang = doc.add_paragraph()
    p_lang.paragraph_format.space_before = Pt(2)
    p_lang.paragraph_format.space_after = Pt(1)
    p_lang.paragraph_format.line_spacing = Pt(11)
    p_lang.paragraph_format.left_indent = Inches(0.1)
    styled_run(p_lang, "English ", size=Pt(8.5), bold=True)
    styled_run(p_lang, "(Professional Working Proficiency)", size=Pt(8.5))
    styled_run(p_lang, "  |  ", size=Pt(8.5), color=MUTED)
    styled_run(p_lang, "Hindi ", size=Pt(8.5), bold=True)
    styled_run(p_lang, "(Native)", size=Pt(8.5))

    add_section_heading(doc, "References")
    p_ref = doc.add_paragraph()
    p_ref.paragraph_format.space_before = Pt(2)
    p_ref.paragraph_format.space_after = Pt(1)
    p_ref.paragraph_format.line_spacing = Pt(11)
    p_ref.paragraph_format.left_indent = Inches(0.1)
    styled_run(p_ref, "Referred by: ", size=Pt(8.5))
    styled_run(p_ref, "Mr. Chirag Sen", size=Pt(8.5), bold=True)

    doc.save(OUTPUT_FILE)
    print(f"Resume saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    build()
