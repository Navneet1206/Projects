"""
Resume Generator - Navneet Vishwakarma
Tailored for: Deloitte - Full Stack Development - NextJS - Managed Services Engineer I (DFO&I - Customer)
Brand Colors: Deloitte Green (#86BC25), Black (#000000), Charcoal (#2D2D2D)
Font Style: Arial / Calibri (Clean Sans-Serif)
"""

import os
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Navneet_Vishwakarma_Deloitte_FullStack.docx")

# Deloitte Brand Color Palette
DARK = RGBColor(0x00, 0x00, 0x00)       # Black
ACCENT = RGBColor(0x00, 0x4B, 0x49)     # Deloitte Deep Forest Green (#004B49)
TEXT = RGBColor(0x2D, 0x2D, 0x2D)      # Charcoal
MUTED = RGBColor(0x55, 0x55, 0x55)     # Medium Gray
RULE_COLOR = "004B49"                  # Deloitte Deep Forest Green Hex for bottom rules
FONT_NAME = "Arial"                    # Clean Sans-Serif font matching brand guides


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


def add_section_heading(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = Pt(12)
    styled_run(p, title.upper(), font_name=FONT_NAME, size=Pt(10),
               color=DARK, bold=True)
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
    style.font.name = FONT_NAME
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
    styled_run(p_name, "NAVNEET VISHWAKARMA", font_name=FONT_NAME,
               size=Pt(17), color=DARK, bold=True)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.line_spacing = Pt(11)
    styled_run(p_title,
               "MANAGED SERVICES ENGINEER I  |  NEXT.JS / REACT DEVELOPER",
               size=Pt(8.5), color=MUTED, bold=True)

    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_after = Pt(1)
    p_contact.paragraph_format.line_spacing = Pt(11)
    styled_run(p_contact, "Satna, Madhya Pradesh, India  |  rajvl132011@gmail.com  |  +91 8435061006",
               size=Pt(8.5), color=TEXT)

    p_links = doc.add_paragraph()
    p_links.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_links.paragraph_format.space_after = Pt(2)
    p_links.paragraph_format.line_spacing = Pt(11)
    styled_run(p_links, "LinkedIn: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "linkedin.com/in/navneet1206", size=Pt(8.5), color=ACCENT)
    styled_run(p_links, "  |  GitHub: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "github.com/Navneet1206", size=Pt(8.5), color=ACCENT)
    styled_run(p_links, "  |  Portfolio: ", size=Pt(8.5), color=TEXT)
    styled_run(p_links, "devfolio.co/@Navneet1206", size=Pt(8.5), color=ACCENT)

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
        "Full Stack Developer with ",
        size=Pt(8.5)
    )
    styled_run(p_sum, "1.3+ years of experience", size=Pt(8.5), bold=True)
    styled_run(
        p_sum,
        " in building and optimizing web applications using ",
        size=Pt(8.5)
    )
    styled_run(p_sum, "Next.js and React", size=Pt(8.5), bold=True)
    styled_run(
        p_sum,
        ". Proficient in JavaScript (ES6+), HTML5, CSS3, Git workflows, CI/CD pipelines, Jest, and React Testing Library. "
        "Strong understanding of web performance fundamentals (Core Web Vitals) and accessibility best practices (WCAG/ARIA). "
        "Currently pursuing B.Tech in Computer Science & Engineering with CGPA 8.29/10.",
        size=Pt(8.5)
    )

    # ═══════════════════════════════════════════════════════════
    # TECHNICAL SKILLS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Technical Skills")

    skills_data = [
        ("Core Development", "Next.js (App Router, Pages Router, Server Components), React.js, JavaScript (ES6+), HTML5, CSS3"),
        ("Testing & QA", "Jest, React Testing Library, Unit Testing, Integration Testing"),
        ("Version Control & CI/CD", "Git, GitHub, Git Workflows, CI/CD Pipelines"),
        ("Performance & Accessibility", "Core Web Vitals (LCP, FID, CLS), Bundle Optimization, Runtime Optimization, WCAG 2.1, ARIA"),
        ("APIs & Authentication", "RESTful APIs, GraphQL, OAuth/OIDC, JWT, Session Management"),
        ("AI-Assisted Development", "GitHub Copilot, Secure Coding Practices, IP/Data-Handling Rules")
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
    styled_run(p_job, "Full Stack Developer", size=Pt(9.5), color=DARK, bold=True)

    p_co = doc.add_paragraph()
    p_co.paragraph_format.space_before = Pt(0)
    p_co.paragraph_format.space_after = Pt(2)
    p_co.paragraph_format.line_spacing = Pt(11)
    styled_run(p_co, "Savayas Life and Balance Pvt. Ltd.", size=Pt(9), color=TEXT, italic=True)
    styled_run(p_co, "  |  March 2024 – Present  |  Satna, MP", size=Pt(8.5), color=MUTED)

    bullets_exp = [
        "Developed and maintained 2 production web applications (savayasyoga.com, savayasheal.com) using Next.js and React.",
        "Built reusable components and patterns for consistent UI/UX across applications to streamline development and align with design guides.",
        "Implemented automated testing using Jest and React Testing Library to ensure code reliability and quality.",
        "Optimized application performance focusing on Core Web Vitals (LCP, FID, CLS) through bundle and runtime optimization.",
        "Applied accessibility best practices (WCAG 2.1, ARIA) for inclusive and standard-compliant user experiences.",
        "Collaborated using Git workflows and participated in CI/CD pipelines for seamless deployments and releases.",
        "Integrated frontends with RESTful APIs and implemented authentication patterns (JWT, session management).",
        "Participated in Agile delivery (sprint planning, standups, demos, retrospectives) with clear estimates and risk visibility.",
        "Utilized GitHub Copilot for AI-assisted development while applying human verification and secure coding practices."
    ]
    for b in bullets_exp:
        add_bullet(doc, b)

    # Key Achievements Subsection
    p_ach = doc.add_paragraph()
    p_ach.paragraph_format.space_before = Pt(3)
    p_ach.paragraph_format.space_after = Pt(1)
    p_ach.paragraph_format.line_spacing = Pt(11)
    p_ach.paragraph_format.left_indent = Inches(0.15)
    styled_run(p_ach, "Key Achievements:", size=Pt(9), color=DARK, bold=True)

    ach_bullets = [
        "Successfully deployed 2 live production applications with 99.9% uptime.",
        "Improved Core Web Vitals scores by 40% through performance optimization.",
        "Achieved 92% accessibility compliance score by implementing ARIA attributes and semantic HTML."
    ]
    for b in ach_bullets:
        add_bullet(doc, b, indent=Inches(0.35))

    # ═══════════════════════════════════════════════════════════
    # PROJECTS
    # ═══════════════════════════════════════════════════════════
    add_section_heading(doc, "Key Projects")

    # Project 1
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(2)
    p_p1.paragraph_format.space_after = Pt(1)
    p_p1.paragraph_format.line_spacing = Pt(11)
    styled_run(p_p1, "HelloMyGym – Gym Management Platform", size=Pt(9), color=DARK, bold=True)
    styled_run(p_p1, "  |  Live Demo: ", size=Pt(8), color=MUTED)
    styled_run(p_p1, "hellomygym.com", size=Pt(8.5), color=ACCENT)

    proj1_bullets = [
        "Developed a mobile-first gym management SaaS platform for Indian gym owners.",
        "Built dynamic dashboards for attendance tracking, membership renewals, and expense management.",
        "Implemented responsive UI with modern JavaScript and CSS3.",
        "Optimized for Core Web Vitals with scores of 90+ for LCP, FID, and CLS."
    ]
    for b in proj1_bullets:
        add_bullet(doc, b)

    # Project 2
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(3)
    p_p2.paragraph_format.space_after = Pt(1)
    p_p2.paragraph_format.line_spacing = Pt(11)
    styled_run(p_p2, "Savayas Yoga – Wellness Platform", size=Pt(9), color=DARK, bold=True)
    styled_run(p_p2, "  |  Live: ", size=Pt(8), color=MUTED)
    styled_run(p_p2, "savayasyoga.com", size=Pt(8.5), color=ACCENT)

    proj2_bullets = [
        "Designed and developed a responsive wellness platform for yoga class bookings.",
        "Created reusable React components and established a component library.",
        "Set up automated testing with Jest and React Testing Library (85%+ coverage).",
        "Implemented JWT-based authentication for user sessions."
    ]
    for b in proj2_bullets:
        add_bullet(doc, b)

    # Project 3
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(3)
    p_p3.paragraph_format.space_after = Pt(1)
    p_p3.paragraph_format.line_spacing = Pt(11)
    styled_run(p_p3, "Savayas Heal – Holistic Wellness Platform", size=Pt(9), color=DARK, bold=True)
    styled_run(p_p3, "  |  Live: ", size=Pt(8), color=MUTED)
    styled_run(p_p3, "savayasheal.com", size=Pt(8.5), color=ACCENT)

    proj3_bullets = [
        "Built a content-rich wellness application with blog management.",
        "Implemented server-side rendering (SSR) for SEO optimization and fast initial load.",
        "Developed accessible UI following WCAG 2.1 AA standards with proper ARIA labels.",
        "Optimized images and assets using Next.js Image component."
    ]
    for b in proj3_bullets:
        add_bullet(doc, b)

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
    styled_run(p_uni, "AKS University, Satna, Madhya Pradesh", size=Pt(8.5), color=TEXT, italic=True)
    styled_run(p_uni, "  |  2022 – 2026 (Expected)  |  CGPA: 8.29 / 10", size=Pt(8.5), color=MUTED)

    add_bullet(doc, "Relevant Coursework: Data Structures, Algorithms, Web Development, Database Systems")

    doc.save(OUTPUT_FILE)
    print(f"Resume saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_resume()
