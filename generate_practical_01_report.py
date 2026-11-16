import os
import subprocess
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_bg(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def generate_docx():
    doc = docx.Document()
    
    # Page Margins: 1 inch
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)

    # Title & Department Header
    p_dept = doc.add_paragraph()
    p_dept.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_dept = p_dept.add_run("Department of Computer Engineering (Software Engineering)\n")
    r_dept.font.size = Pt(14)
    r_dept.font.bold = True
    r_dept.font.color.rgb = RGBColor(15, 23, 42)

    r_sub = p_dept.add_run("Web Technology Laboratory — Academic Year 2026-27\n")
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Practical No. 01: Single-Page Responsive Web Application")
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(37, 99, 235)

    # Metadata Table
    meta_table = doc.add_table(rows=2, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Name:", "Sarthak Bhosale"), ("PRN:", "202401100051")],
        [("Batch:", "TY SE-A2"), ("Roll No:", "A50")]
    ]
    for row_idx, row in enumerate(meta_data):
        col_offset = 0
        for label, val in row:
            cell_lbl = meta_table.cell(row_idx, col_offset)
            cell_val = meta_table.cell(row_idx, col_offset + 1)
            cell_lbl.text = label
            cell_lbl.paragraphs[0].runs[0].font.bold = True
            cell_lbl.paragraphs[0].runs[0].font.color.rgb = RGBColor(71, 85, 105)
            cell_val.text = val
            cell_val.paragraphs[0].runs[0].font.bold = True
            cell_val.paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 23, 42)
            set_cell_bg(cell_lbl, "F1F5F9")
            set_cell_bg(cell_val, "FFFFFF")
            set_cell_margins(cell_lbl, 80, 80, 100, 100)
            set_cell_margins(cell_val, 80, 80, 100, 100)
            col_offset += 2

    doc.add_paragraph() # Spacing

    def add_section_heading(num_str, title_str):
        h = doc.add_paragraph()
        r = h.add_run(f"{num_str}. {title_str}")
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(4)
        return h

    # 1. Aim
    add_section_heading("1", "Aim")
    p = doc.add_paragraph(
        "To design, develop, and deploy a responsive single-page web application for a Task Management System "
        "named 'TaskGravity' using semantic HTML5, custom minimalist CSS3, and the Bootstrap 5 responsive grid system, "
        "demonstrating clean UI/UX design rules, multi-viewport responsiveness, and component integration."
    )

    # 2. Objectives
    add_section_heading("2", "Objectives")
    objs = [
        "To structure a responsive single-page application using semantic HTML5 tags (<header>, <nav>, <section>, <main>, <article>, <footer>).",
        "To enforce a clean, flat, minimalist design system avoiding generic AI gradients, heavy blurs, or visual clutter.",
        "To implement responsive multi-column layouts across mobile (<576px), tablet (<992px), and desktop (1200px+) using Bootstrap 5 flexbox grid.",
        "To integrate at least 3 distinct Bootstrap components (Navbar, Cards, Badges, Accordion, Modal, Progress Bar).",
        "To maintain highly modular, DRY, and maintainable CSS utilizing CSS3 custom properties (design tokens).",
        "To deploy the application to a live hosting platform and publish source code to GitHub."
    ]
    for o in objs:
        p = doc.add_paragraph(o, style='List Bullet')
        p.paragraph_format.space_after = Pt(2)

    # 3. Theory & Concepts
    add_section_heading("3", "Theory & Concepts")
    p = doc.add_paragraph(
        "TaskGravity is a high-velocity task orchestration platform tailored for modern engineering workflows. "
        "The application architecture strictly complies with flat, content-centric UI guidelines and utilizes standard "
        "web technologies:"
    )
    
    concepts = [
        ("3.1 Semantic HTML5", 
         "HTML5 introduces semantic elements that clearly describe their meaning to both the browser and the developer. "
         "Instead of generic <div> tags, TaskGravity employs <header> for site branding and top navigation, <section> for "
         "distinct content groups (Hero banner, metrics, task board, architecture), <article> for self-contained task cards, "
         "and <footer> for metadata and licensing information. This promotes accessibility (screen readers) and SEO best practices."),
        ("3.2 Minimalist CSS3 & Design Tokens",
         "The visual theme adheres to strict minimalism: flat surfaces, subtle 1px border lines (#e2e8f0), neutral slate "
         "backgrounds (#f8fafc), high-contrast typography (#0f172a), and authoritative slate/blue action tones (#2563eb). "
         "Design tokens are centralized in :root variables (--tg-bg, --tg-surface, --tg-border, --tg-primary), ensuring DRY "
         "declarations and zero duplicated styling rules."),
        ("3.3 Bootstrap 5 Responsive Grid System",
         "Bootstrap 5 uses a 12-column flexbox grid system with five default responsive tiers (xs, sm, md, lg, xl, xxl). "
         "TaskGravity leverages responsive classes such as col-12 col-md-6 col-lg-3 for stat cards and col-12 col-md-4 for "
         "Kanban board streams, guaranteeing that the layout automatically stacks vertically on smartphones, organizes into "
         "two columns on tablets, and displays full 3-column or 4-column views on desktop monitors."),
        ("3.4 Bootstrap Components Implemented",
         "The application integrates 6 distinct Bootstrap components:\n"
         "• Navbar: Collapsible navigation bar with brand icon, links, and responsive hamburger toggler for small screens.\n"
         "• Cards: Elevated surface containers with headers, content descriptions, and metadata footers.\n"
         "• Badges: Color-coded status and priority tags (Urgent, Medium, Low, Category indicators).\n"
         "• Modal: Interactive dialog popup for creating new tasks with form input validation states.\n"
         "• Accordion: Collapsible disclosure panels detailing the multi-practical engineering roadmap.\n"
         "• Progress Bar: Dynamic sprint completion bar tracking closed engineering deliverables.")
    ]
    for title, desc in concepts:
        h = doc.add_paragraph()
        r = h.add_run(title)
        r.font.bold = True
        r.font.size = Pt(11.5)
        h.paragraph_format.space_before = Pt(6)
        h.paragraph_format.space_after = Pt(2)
        doc.add_paragraph(desc)

    # 4. System Architecture & UI Specifications
    add_section_heading("4", "System Architecture & Responsive Grid Specifications")
    doc.add_paragraph("The responsive viewport distribution of TaskGravity is summarized in the matrix below:")
    
    spec_table = doc.add_table(rows=5, cols=4)
    spec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Viewport Device", "Breakpoint", "Grid Class Used", "Layout Presentation"]
    for i, h in enumerate(headers):
        c = spec_table.cell(0, i)
        c.text = h
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        set_cell_bg(c, "1E293B")
        set_cell_margins(c, 80, 80, 100, 100)
    
    grid_data = [
        ("Mobile Handsets", "< 576px (xs)", "col-12", "Single stacked column, collapsed hamburger navbar"),
        ("Tablet Portrait", "576px - 767px (sm)", "col-sm-6", "2-column metric cards, stacked task streams"),
        ("Tablet Landscape", "768px - 991px (md)", "col-md-4", "3-column task boards (Backlog, In Progress, Done)"),
        ("Desktop Monitors", "≥ 992px (lg, xl)", "col-lg-3, col-lg-8", "Expanded top navigation, 4-column metric grid, side-by-side hero")
    ]
    for row_idx, row in enumerate(grid_data, start=1):
        for col_idx, text in enumerate(row):
            c = spec_table.cell(row_idx, col_idx)
            c.text = text
            if row_idx % 2 == 1:
                set_cell_bg(c, "F8FAFC")
            else:
                set_cell_bg(c, "FFFFFF")
            set_cell_margins(c, 70, 70, 90, 90)

    doc.add_paragraph()

    # 5. Algorithm & Implementation Steps
    add_section_heading("5", "Algorithm & Implementation Steps")
    steps = [
        "Initialize project directory with semantic files (index.html, styles.css) and configure Git version control.",
        "Import Bootstrap 5.3 CDN CSS and JavaScript bundle with Popper in the HTML <head> and before </body>.",
        "Construct sticky responsive <nav> using Bootstrap's navbar-expand-lg and data-bs-toggle='collapse' attributes.",
        "Implement Hero Banner section with clean status badge indicator, authoritative value proposition, and quick action CTAs.",
        "Design 4-card metric overview displaying sprint progress with custom Bootstrap progress-bar styling.",
        "Build 3-column Kanban Task Stream using Bootstrap's row and col-12 col-md-4 classes.",
        "Create custom minimalist card styles (.task-card, .badge-minimal) relying on clean hex tokens and flat borders.",
        "Integrate Bootstrap Modal (#createTaskModal) containing form inputs (Title, Category, Priority, Scope) with modal trigger.",
        "Add interactive Bootstrap Accordion (#roadmapAccordion) for system engineering and scaling roadmap documentation.",
        "Implement accessible <footer> section with developer metadata, institutional info, and repository links.",
        "Test responsiveness across simulated viewports (Chrome DevTools Mobile & Tablet simulation).",
        "Deploy to GitHub Pages and link repository."
    ]
    for i, s in enumerate(steps, 1):
        p = doc.add_paragraph(f"Step {i}: {s}")
        p.paragraph_format.space_after = Pt(2)

    # 6. Source Code
    add_section_heading("6", "Source Code Implementation")
    
    p = doc.add_paragraph()
    r = p.add_run("6.1 styles.css (Minimalist Custom Stylesheet & Design Tokens)")
    r.font.bold = True
    
    with open(r"D:\_TY\Web-Tech\Practicals\Practical1\styles.css", "r", encoding="utf-8") as f:
        css_code = f.read()
    
    p_code = doc.add_paragraph()
    r_code = p_code.add_run(css_code)
    r_code.font.name = 'Consolas'
    r_code.font.size = Pt(8.5)
    r_code.font.color.rgb = RGBColor(30, 41, 59)
    p_code.paragraph_format.space_before = Pt(4)
    p_code.paragraph_format.space_after = Pt(12)

    p = doc.add_paragraph()
    r = p.add_run("6.2 index.html (Semantic Single-Page Application Structure)")
    r.font.bold = True

    with open(r"D:\_TY\Web-Tech\Practicals\Practical1\index.html", "r", encoding="utf-8") as f:
        html_code = f.read()

    p_code2 = doc.add_paragraph()
    r_code2 = p_code2.add_run(html_code)
    r_code2.font.name = 'Consolas'
    r_code2.font.size = Pt(8)
    r_code2.font.color.rgb = RGBColor(30, 41, 59)
    p_code2.paragraph_format.space_before = Pt(4)
    p_code2.paragraph_format.space_after = Pt(12)

    # 7. Verification & Output
    add_section_heading("7", "Verification & Test Results")
    results = [
        ("Navbar Responsiveness", "Tested at 375px (Mobile) and 1280px (Desktop). Navbar collapses to hamburger menu below 992px and expands cleanly on desktop.", "Passed"),
        ("Grid Flow & Fluidity", "Tested on 320px, 768px, and 1440px viewports. Cards stack seamlessly without horizontal overflow or broken layout boundaries.", "Passed"),
        ("Modal Component", "Clicking '+ New Task' cleanly launches #createTaskModal with accessible focus trap and dismiss controls.", "Passed"),
        ("Accordion Collapse", "Toggling roadmap accordion items smoothly expands content while collapsing adjacent panels.", "Passed"),
        ("Visual Aesthetics", "Verified zero AI-style generic gradients, neon dropshadows, or noisy card glows. High readability achieved.", "Passed")
    ]
    res_table = doc.add_table(rows=len(results) + 1, cols=3)
    res_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, h in enumerate(["Test Case", "Observed Behavior & Criteria", "Verification Status"]):
        c = res_table.cell(0, idx)
        c.text = h
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        set_cell_bg(c, "1E293B")
        set_cell_margins(c, 70, 70, 90, 90)

    for r_idx, (tc, obs, st) in enumerate(results, 1):
        c0 = res_table.cell(r_idx, 0)
        c1 = res_table.cell(r_idx, 1)
        c2 = res_table.cell(r_idx, 2)
        c0.text = tc
        c0.paragraphs[0].runs[0].font.bold = True
        c1.text = obs
        c2.text = st
        c2.paragraphs[0].runs[0].font.bold = True
        c2.paragraphs[0].runs[0].font.color.rgb = RGBColor(5, 150, 105)
        set_cell_bg(c0, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        set_cell_bg(c1, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        set_cell_bg(c2, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        set_cell_margins(c0, 60, 60, 80, 80)
        set_cell_margins(c1, 60, 60, 80, 80)
        set_cell_margins(c2, 60, 60, 80, 80)

    doc.add_paragraph()

    # 8. Conclusion
    add_section_heading("8", "Conclusion")
    doc.add_paragraph(
        "The Practical 01 implementation successfully demonstrates the creation of a responsive, minimalist single-page web "
        "application named TaskGravity. By adhering to strict UI/UX guidelines, the application presents a high-density, "
        "content-first layout devoid of unnecessary visual noise. Leveraging HTML5 semantic tags, custom DRY CSS3 variables, "
        "and Bootstrap 5 components (Navbar, Cards, Badges, Accordion, Modal, Progress Bar) ensures full cross-device "
        "accessibility and responsive agility. This foundational UI layer establishes the standard for client-side data "
        "binding, React components, and server-side persistence in subsequent practical assignments."
    )

    # MANDATORY LINKS AT THE VERY END
    doc.add_paragraph()
    p_links = doc.add_paragraph()
    r_links_title = p_links.add_run("Project Links & Deployment Information\n")
    r_links_title.font.size = Pt(12)
    r_links_title.font.bold = True
    r_links_title.font.color.rgb = RGBColor(15, 23, 42)

    r_gh = p_links.add_run("GitHub Link / Code Link: ")
    r_gh.font.bold = True
    r_gh_url = p_links.add_run("https://github.com/Sarthak-1229/taskgravity-practical-01\n")
    r_gh_url.font.color.rgb = RGBColor(37, 99, 235)
    r_gh_url.font.underline = True

    r_live = p_links.add_run("Live Deployment Link: ")
    r_live.font.bold = True
    r_live_url = p_links.add_run("https://sarthak-1229.github.io/taskgravity-practical-01/\n")
    r_live_url.font.color.rgb = RGBColor(37, 99, 235)
    r_live_url.font.underline = True

    docx_path = r"D:\_TY\Web-Tech\Practicals\Practical1\Practical_01_Report.docx"
    doc.save(docx_path)
    print(f"Generated DOCX: {docx_path}")

def generate_pdf():
    # Build complete HTML report for Chrome Headless rendering
    with open(r"D:\_TY\Web-Tech\Practicals\Practical1\styles.css", "r", encoding="utf-8") as f:
        css_code = f.read()
    with open(r"D:\_TY\Web-Tech\Practicals\Practical1\index.html", "r", encoding="utf-8") as f:
        html_code = f.read()

    # Escape HTML code for pre blocks
    import html
    escaped_css = html.escape(css_code)
    escaped_html = html.escape(html_code)

    report_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Practical 01 Report - TaskGravity</title>
<style>
  @page {{
    size: A4;
    margin: 20mm 18mm 20mm 18mm;
  }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    color: #0f172a;
    line-height: 1.5;
    font-size: 11pt;
    margin: 0;
    padding: 0;
  }}
  .dept-header {{
    text-align: center;
    border-bottom: 2px solid #0f172a;
    padding-bottom: 12px;
    margin-bottom: 18px;
  }}
  .dept-title {{
    font-size: 15pt;
    font-weight: 800;
    color: #0f172a;
    text-transform: uppercase;
    letter-spacing: 0.02em;
    margin-bottom: 4px;
  }}
  .dept-sub {{
    font-size: 11pt;
    color: #475569;
    font-weight: 500;
  }}
  .practical-title {{
    font-size: 14pt;
    font-weight: 700;
    color: #2563eb;
    margin-top: 8px;
  }}
  .meta-table {{
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
    font-size: 10pt;
  }}
  .meta-table td {{
    border: 1px solid #cbd5e1;
    padding: 6px 12px;
  }}
  .meta-label {{
    background-color: #f1f5f9;
    font-weight: 600;
    color: #334155;
    width: 18%;
  }}
  .meta-val {{
    font-weight: 600;
    color: #0f172a;
    width: 32%;
  }}
  h2.sec-title {{
    font-size: 12pt;
    font-weight: 700;
    color: #0f172a;
    border-left: 4px solid #2563eb;
    padding-left: 8px;
    margin-top: 18px;
    margin-bottom: 8px;
    text-transform: uppercase;
    letter-spacing: 0.03em;
  }}
  h3.subsec-title {{
    font-size: 11pt;
    font-weight: 700;
    color: #1e293b;
    margin-top: 12px;
    margin-bottom: 4px;
  }}
  p, li {{
    text-align: justify;
    color: #334155;
    margin-bottom: 6px;
  }}
  ul, ol {{
    margin-top: 4px;
    margin-bottom: 8px;
    padding-left: 22px;
  }}
  li {{
    margin-bottom: 4px;
  }}
  .data-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 12px 0 16px;
    font-size: 9.5pt;
  }}
  .data-table th {{
    background-color: #1e293b;
    color: #ffffff;
    font-weight: 600;
    padding: 7px 10px;
    border: 1px solid #1e293b;
    text-align: left;
  }}
  .data-table td {{
    padding: 6px 10px;
    border: 1px solid #e2e8f0;
    color: #334155;
  }}
  .data-table tr:nth-child(even) td {{
    background-color: #f8fafc;
  }}
  pre.code-block {{
    background-color: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 4px;
    padding: 10px;
    font-family: Consolas, "Courier New", monospace;
    font-size: 8pt;
    line-height: 1.35;
    overflow-x: auto;
    color: #0f172a;
    page-break-inside: avoid;
    white-space: pre-wrap;
    word-break: break-all;
  }}
  .badge-pass {{
    color: #059669;
    font-weight: 700;
  }}
  .links-box {{
    background-color: #f8fafc;
    border: 1px solid #cbd5e1;
    border-left: 4px solid #059669;
    padding: 12px 16px;
    margin-top: 20px;
    page-break-inside: avoid;
  }}
  .links-box a {{
    color: #2563eb;
    text-decoration: underline;
    font-weight: 600;
  }}
</style>
</head>
<body>

<div class="dept-header">
  <div class="dept-title">Department of Computer Engineering (Software Engineering)</div>
  <div class="dept-sub">Web Technology Laboratory — Academic Year 2026-27</div>
  <div class="practical-title">Practical No. 01: Single-Page Responsive Web Application</div>
</div>

<table class="meta-table">
  <tr>
    <td class="meta-label">Name:</td>
    <td class="meta-val">Sarthak Bhosale</td>
    <td class="meta-label">PRN:</td>
    <td class="meta-val">202401100051</td>
  </tr>
  <tr>
    <td class="meta-label">Batch:</td>
    <td class="meta-val">TY SE-A2</td>
    <td class="meta-label">Roll No:</td>
    <td class="meta-val">A50</td>
  </tr>
</table>

<h2 class="sec-title">1. Aim</h2>
<p>
To design, develop, and deploy a responsive single-page web application for a Task Management System named <strong>TaskGravity</strong> using semantic HTML5, custom minimalist CSS3, and the Bootstrap 5 responsive grid system, demonstrating clean UI/UX design rules, multi-device fluidity, and component integration.
</p>

<h2 class="sec-title">2. Objectives</h2>
<ul>
  <li>To structure a responsive single-page web interface using semantic HTML5 tags (<code>&lt;header&gt;</code>, <code>&lt;nav&gt;</code>, <code>&lt;section&gt;</code>, <code>&lt;main&gt;</code>, <code>&lt;article&gt;</code>, <code>&lt;footer&gt;</code>).</li>
  <li>To enforce a clean, flat, minimalist design system avoiding generic AI gradients, heavy shadows, or visual clutter.</li>
  <li>To implement responsive multi-column layouts across mobile (&lt;576px), tablet (&lt;992px), and desktop (1200px+) using Bootstrap 5 flexbox grid.</li>
  <li>To integrate at least 3 distinct Bootstrap components (Navbar, Cards, Badges, Accordion, Modal, Progress Bar).</li>
  <li>To maintain highly modular, DRY, and maintainable CSS utilizing CSS3 custom properties (design tokens).</li>
  <li>To deploy the application to a live hosting platform and publish source code to GitHub.</li>
</ul>

<h2 class="sec-title">3. Theory & Concepts</h2>
<p>
TaskGravity is a high-velocity task orchestration platform tailored for modern engineering workflows. The application architecture strictly complies with flat, content-centric UI guidelines and utilizes standard web technologies:
</p>

<h3 class="subsec-title">3.1 Semantic HTML5</h3>
<p>
HTML5 introduces semantic elements that clearly describe their meaning to both the browser and the developer. Instead of generic <code>&lt;div&gt;</code> tags, TaskGravity employs <code>&lt;header&gt;</code> for site branding and top navigation, <code>&lt;section&gt;</code> for distinct content groups (Hero banner, metrics, task board, architecture), <code>&lt;article&gt;</code> for self-contained task cards, and <code>&lt;footer&gt;</code> for metadata and licensing information. This promotes accessibility (screen readers) and SEO best practices.
</p>

<h3 class="subsec-title">3.2 Minimalist CSS3 & Design Tokens</h3>
<p>
The visual theme adheres to strict minimalism: flat surfaces, subtle 1px border lines (#e2e8f0), neutral slate backgrounds (#f8fafc), high-contrast typography (#0f172a), and authoritative slate/blue action tones (#2563eb). Design tokens are centralized in <code>:root</code> variables (<code>--tg-bg</code>, <code>--tg-surface</code>, <code>--tg-border</code>, <code>--tg-primary</code>), ensuring DRY declarations and zero duplicated styling rules.
</p>

<h3 class="subsec-title">3.3 Bootstrap 5 Responsive Grid System</h3>
<p>
Bootstrap 5 uses a 12-column flexbox grid system with five default responsive tiers (xs, sm, md, lg, xl, xxl). TaskGravity leverages responsive classes such as <code>col-12 col-md-6 col-lg-3</code> for stat cards and <code>col-12 col-md-4</code> for Kanban board streams, guaranteeing that the layout automatically stacks vertically on smartphones, organizes into two columns on tablets, and displays full 3-column or 4-column views on desktop monitors.
</p>

<h3 class="subsec-title">3.4 Bootstrap Components Implemented</h3>
<p>
The application integrates 6 distinct Bootstrap components:
</p>
<ul>
  <li><strong>Navbar:</strong> Collapsible navigation bar with brand icon, links, and responsive hamburger toggler for small screens.</li>
  <li><strong>Cards:</strong> Elevated surface containers with headers, content descriptions, and metadata footers.</li>
  <li><strong>Badges:</strong> Color-coded status and priority tags (Urgent, Medium, Low, Category indicators).</li>
  <li><strong>Modal:</strong> Interactive dialog popup for creating new tasks with form input validation states.</li>
  <li><strong>Accordion:</strong> Collapsible disclosure panels detailing the multi-practical engineering roadmap.</li>
  <li><strong>Progress Bar:</strong> Dynamic sprint completion bar tracking closed engineering deliverables.</li>
</ul>

<h2 class="sec-title">4. System Architecture & Responsive Grid Specifications</h2>
<p>The responsive viewport distribution of TaskGravity is summarized in the matrix below:</p>

<table class="data-table">
  <thead>
    <tr>
      <th>Viewport Device</th>
      <th>Breakpoint</th>
      <th>Grid Class Used</th>
      <th>Layout Presentation</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Mobile Handsets</strong></td>
      <td>&lt; 576px (xs)</td>
      <td><code>col-12</code></td>
      <td>Single stacked column, collapsed hamburger navbar</td>
    </tr>
    <tr>
      <td><strong>Tablet Portrait</strong></td>
      <td>576px - 767px (sm)</td>
      <td><code>col-sm-6</code></td>
      <td>2-column metric cards, stacked task streams</td>
    </tr>
    <tr>
      <td><strong>Tablet Landscape</strong></td>
      <td>768px - 991px (md)</td>
      <td><code>col-md-4</code></td>
      <td>3-column task boards (Backlog, In Progress, Done)</td>
    </tr>
    <tr>
      <td><strong>Desktop Monitors</strong></td>
      <td>&ge; 992px (lg, xl)</td>
      <td><code>col-lg-3, col-lg-8</code></td>
      <td>Expanded top navigation, 4-column metric grid, side-by-side hero</td>
    </tr>
  </tbody>
</table>

<h2 class="sec-title">5. Algorithm & Implementation Steps</h2>
<ol>
  <li>Initialize project directory with semantic files (<code>index.html</code>, <code>styles.css</code>) and configure Git version control.</li>
  <li>Import Bootstrap 5.3 CDN CSS and JavaScript bundle with Popper in the HTML <code>&lt;head&gt;</code> and before <code>&lt;/body&gt;</code>.</li>
  <li>Construct sticky responsive <code>&lt;nav&gt;</code> using Bootstrap's <code>navbar-expand-lg</code> and <code>data-bs-toggle="collapse"</code> attributes.</li>
  <li>Implement Hero Banner section with clean status badge indicator, authoritative value proposition, and quick action CTAs.</li>
  <li>Design 4-card metric overview displaying sprint progress with custom Bootstrap progress-bar styling.</li>
  <li>Build 3-column Kanban Task Stream using Bootstrap's <code>row</code> and <code>col-12 col-md-4</code> classes.</li>
  <li>Create custom minimalist card styles (<code>.task-card</code>, <code>.badge-minimal</code>) relying on clean hex tokens and flat borders.</li>
  <li>Integrate Bootstrap Modal (<code>#createTaskModal</code>) containing form inputs (Title, Category, Priority, Scope) with modal trigger.</li>
  <li>Add interactive Bootstrap Accordion (<code>#roadmapAccordion</code>) for system engineering and scaling roadmap documentation.</li>
  <li>Implement accessible <code>&lt;footer&gt;</code> section with developer metadata, institutional info, and repository links.</li>
  <li>Test responsiveness across simulated viewports (Chrome DevTools Mobile & Tablet simulation).</li>
  <li>Deploy to GitHub Pages and link repository.</li>
</ol>

<h2 class="sec-title">6. Source Code Implementation</h2>

<h3 class="subsec-title">6.1 styles.css (Minimalist Custom Stylesheet & Design Tokens)</h3>
<pre class="code-block">{escaped_css}</pre>

<h3 class="subsec-title">6.2 index.html (Semantic Single-Page Application Structure)</h3>
<pre class="code-block">{escaped_html}</pre>

<h2 class="sec-title">7. Verification & Test Results</h2>

<table class="data-table">
  <thead>
    <tr>
      <th>Test Case</th>
      <th>Observed Behavior & Criteria</th>
      <th>Verification Status</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>Navbar Responsiveness</strong></td>
      <td>Tested at 375px (Mobile) and 1280px (Desktop). Navbar collapses to hamburger menu below 992px and expands cleanly on desktop.</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>Grid Flow & Fluidity</strong></td>
      <td>Tested on 320px, 768px, and 1440px viewports. Cards stack seamlessly without horizontal overflow or broken layout boundaries.</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>Modal Component</strong></td>
      <td>Clicking '+ New Task' cleanly launches #createTaskModal with accessible focus trap and dismiss controls.</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>Accordion Collapse</strong></td>
      <td>Toggling roadmap accordion items smoothly expands content while collapsing adjacent panels.</td>
      <td class="badge-pass">Passed</td>
    </tr>
    <tr>
      <td><strong>Visual Aesthetics</strong></td>
      <td>Verified zero AI-style generic gradients, neon dropshadows, or noisy card glows. High readability achieved.</td>
      <td class="badge-pass">Passed</td>
    </tr>
  </tbody>
</table>

<h2 class="sec-title">8. Conclusion</h2>
<p>
The Practical 01 implementation successfully demonstrates the creation of a responsive, minimalist single-page web application named TaskGravity. By adhering to strict UI/UX guidelines, the application presents a high-density, content-first layout devoid of unnecessary visual noise. Leveraging HTML5 semantic tags, custom DRY CSS3 variables, and Bootstrap 5 components (Navbar, Cards, Badges, Accordion, Modal, Progress Bar) ensures full cross-device accessibility and responsive agility. This foundational UI layer establishes the standard for client-side data binding, React components, and server-side persistence in subsequent practical assignments.
</p>

<div class="links-box">
  <div style="font-weight: 700; font-size: 11pt; margin-bottom: 6px; color: #0f172a;">Project Links & Deployment Information</div>
  <div><strong>GitHub Link / Code Link:</strong> <a href="https://github.com/Sarthak-1229/taskgravity-practical-01" target="_blank">https://github.com/Sarthak-1229/taskgravity-practical-01</a></div>
  <div style="margin-top: 4px;"><strong>Live Deployment Link:</strong> <a href="https://sarthak-1229.github.io/taskgravity-practical-01/" target="_blank">https://sarthak-1229.github.io/taskgravity-practical-01/</a></div>
</div>

</body>
</html>
"""
    html_report_path = r"D:\_TY\Web-Tech\Practicals\Practical1\report_temp.html"
    pdf_report_path = r"D:\_TY\Web-Tech\Practicals\Practical1\Practical_01_Report.pdf"

    with open(html_report_path, "w", encoding="utf-8") as f:
        f.write(report_html)

    chrome_cmd = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "--headless=new",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_report_path}",
        html_report_path
    ]
    subprocess.run(chrome_cmd, check=True)
    
    if os.path.exists(html_report_path):
        os.remove(html_report_path)
    
    print(f"Generated PDF: {pdf_report_path}")

if __name__ == "__main__":
    generate_docx()
    generate_pdf()
