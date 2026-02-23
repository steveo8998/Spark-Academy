#!/usr/bin/env python3
"""
Spark Docs — DOCX to HTML Converter
Reads a .docx file and produces a styled, mobile-friendly HTML file.
Faithfully maps document colors and structure from the Spark design system.

Usage:
    python3 convert.py input.docx
    python3 convert.py input.docx output.html
"""

import sys
import os
import re
from docx import Document
from docx.oxml.ns import qn
from html import escape

# ── Color map: docx fill hex → CSS class ─────────────────────────
# Maps the exact fill colors used in Spark documents to semantic classes
COLOR_CLASS = {
    "1E4D8C": "fill-navy",
    "2E86C1": "fill-blue",
    "1A7A4A": "fill-green",
    "B7860B": "fill-gold",
    "D6E4F0": "fill-sky",
    "F4F6F8": "fill-lightgray",
    "FEF9EC": "fill-goldtint-a",
    "FFFDF4": "fill-goldtint-b",
    "FFFFFF": "fill-white",
    # Pool A audit colors
    "1B2A4A": "fill-navy",
    "C8E6C9": "fill-green-light",
    "DCEDC8": "fill-green-pale",
    "FFF9C4": "fill-yellow",
    "FFCDD2": "fill-red-light",
    "EDE7F6": "fill-purple-light",
    "D6E4F7": "fill-sky",
    "F2F7FF": "fill-sky-pale",
    "EEF4FF": "fill-sky-pale",
    "F9F9F9": "fill-lightgray",
}

def get_cell_fill(cell):
    """Extract fill color hex from a table cell."""
    try:
        tc = cell._tc
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is not None:
            shd = tcPr.find(qn('w:shd'))
            if shd is not None:
                fill = shd.get(qn('w:fill'))
                if fill and fill.upper() not in ('AUTO', 'NONE', ''):
                    return fill.upper()
    except Exception:
        pass
    return None

def get_run_color(run):
    """Extract font color hex from a run."""
    try:
        rPr = run._r.find(qn('w:rPr'))
        if rPr is not None:
            color = rPr.find(qn('w:color'))
            if color is not None:
                val = color.get(qn('w:val'))
                if val and val.upper() not in ('AUTO', ''):
                    return val.upper()
    except Exception:
        pass
    return None

def get_para_shading(para):
    """Extract background shading from a paragraph."""
    try:
        pPr = para._p.find(qn('w:pPr'))
        if pPr is not None:
            shd = pPr.find(qn('w:shd'))
            if shd is not None:
                fill = shd.get(qn('w:fill'))
                if fill and fill.upper() not in ('AUTO', 'NONE', ''):
                    return fill.upper()
    except Exception:
        pass
    return None

def runs_to_html(paragraph):
    """Convert paragraph runs to inline HTML, preserving bold/italic/color."""
    parts = []
    for run in paragraph.runs:
        text = escape(run.text)
        if not text:
            continue
        bold = run.bold
        italic = run.italic
        color = get_run_color(run)
        
        if bold:
            text = f"<strong>{text}</strong>"
        if italic:
            text = f"<em>{text}</em>"
        if color and color not in ("000000", "1A1A2E", "333333", "FFFFFF", "WHITE"):
            css_color = f"#{color.lower()}"
            # Map known colors to CSS vars
            color_map = {
                "2E86C1": "var(--blue)",
                "1E4D8C": "var(--navy)",
                "1A7A4A": "var(--green)",
                "B7860B": "var(--gold)",
                "C8960C": "var(--gold)",
                "888888": "var(--mid)",
                "555555": "var(--mid)",
                "6B7A8D": "var(--mid)",
            }
            css_color = color_map.get(color, css_color)
            text = f'<span style="color:{css_color}">{text}</span>'
        parts.append(text)
    return "".join(parts)

def classify_table(table):
    """
    Identify the semantic role of a table based on its structure and fill colors.
    Returns a string type identifier.
    """
    if not table.rows:
        return "generic"
    
    rows = table.rows
    cols = len(table.columns)
    
    # Get fill of first cell
    first_fill = get_cell_fill(rows[0].cells[0])
    
    if not first_fill:
        return "generic"
    
    f = first_fill.upper()
    
    # School name bar: 1 row, 2 cols, blue fill
    if len(rows) == 1 and cols == 2 and f in ("2E86C1",):
        return "school-name-bar"
    
    # Section label: 1 row, 1 col, navy/blue/green/gold fill
    if len(rows) == 1 and cols == 1 and f in ("1E4D8C", "2E86C1", "1A7A4A", "B7860B", "1B2A4A"):
        return "section-label"
    
    # Green header for "What They Got Right"
    if len(rows) == 1 and cols == 2 and f in ("1A7A4A",):
        return "got-right-header"
    
    # Two-column content table (got right items, alternating rows)
    if cols == 2 and len(rows) > 1 and f in ("FFFFFF", "F4F6F8"):
        return "two-col-content"
    
    # Takeaway box: 1 col, gold tint rows
    if cols == 1 and len(rows) >= 1 and f in ("FEF9EC", "FFFDF4"):
        return "takeaway-box"
    
    # Info grid: 2 cols, sky/white alternating (school index on intro page)
    if cols == 2 and f in ("D6E4F0", "D6E4F7"):
        return "info-grid"
    
    # Principle header (synthesis page): 1 col, navy fill
    if cols == 1 and f in ("1E4D8C", "1B2A4A"):
        return "principle-header"
    
    # Score table (audit docs): has dark navy header row
    if f in ("1B2A4A",) and cols >= 3:
        return "score-table"

    # Score legend (audit docs)  
    if len(rows) == 1 and cols >= 4 and f in ("C8E6C9",):
        return "score-legend"
    
    return "generic"

def table_to_html(table):
    """Convert a docx table to HTML based on its semantic type."""
    t = classify_table(table)
    
    if t == "school-name-bar":
        name_cell = table.rows[0].cells[0]
        loc_cell  = table.rows[0].cells[1]
        name = escape(name_cell.text.strip())
        loc  = escape(loc_cell.text.strip())
        return f'<div class="school-name-bar"><span class="school-name">{name}</span><span class="school-loc">{loc}</span></div>'
    
    if t == "section-label":
        cell  = table.rows[0].cells[0]
        text  = escape(cell.text.strip())
        fill  = get_cell_fill(cell) or "1E4D8C"
        css   = COLOR_CLASS.get(fill.upper(), "fill-navy")
        return f'<div class="section-label {css}">{text}</div>'
    
    if t == "got-right-header":
        return '<div class="got-right-header">What This Site Does Well</div>'
    
    if t == "two-col-content":
        rows_html = []
        for i, row in enumerate(table.rows):
            left  = row.cells[0].text.strip()
            right = row.cells[1].text.strip()
            cls   = "row-alt" if i % 2 == 1 else ""
            if left or right:
                rows_html.append(
                    f'<div class="two-col-row {cls}">'
                    f'<div class="two-col-cell">{escape(left)}</div>'
                    f'<div class="two-col-cell">{escape(right)}</div>'
                    f'</div>'
                )
        return f'<div class="two-col-table">{"".join(rows_html)}</div>'
    
    if t == "takeaway-box":
        items_html = []
        for i, row in enumerate(table.rows):
            text = row.cells[0].text.strip()
            if text:
                # Split "1.  text" pattern
                text = re.sub(r'^\d+\.\s+', '', text)
                items_html.append(
                    f'<div class="takeaway-item">'
                    f'<span class="takeaway-num">{i+1}</span>'
                    f'<span class="takeaway-text">{escape(text)}</span>'
                    f'</div>'
                )
        return f'<div class="takeaway-box">{"".join(items_html)}</div>'
    
    if t == "info-grid":
        rows_html = []
        for row in table.rows:
            cells_html = []
            for cell in row.cells:
                fill = get_cell_fill(cell)
                css  = COLOR_CLASS.get((fill or "FFFFFF").upper(), "fill-white")
                # Handle multi-paragraph cells
                lines = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
                content = "<br>".join(escape(l) for l in lines)
                cells_html.append(f'<div class="info-cell {css}">{content}</div>')
            rows_html.append(f'<div class="info-row">{"".join(cells_html)}</div>')
        return f'<div class="info-grid">{"".join(rows_html)}</div>'
    
    if t == "principle-header":
        cell = table.rows[0].cells[0]
        text = cell.text.strip()
        # Parse "Principle N:  Title" pattern
        m = re.match(r'(Principle \d+:)\s+(.*)', text)
        if m:
            label = escape(m.group(1))
            title = escape(m.group(2))
            return f'<div class="principle-header"><span class="principle-label">{label}</span> {title}</div>'
        return f'<div class="principle-header">{escape(text)}</div>'
    
    if t == "score-table":
        rows_html = []
        for i, row in enumerate(table.rows):
            cells_html = []
            for cell in row.cells:
                fill = get_cell_fill(cell)
                css  = COLOR_CLASS.get((fill or "FFFFFF").upper(), "")
                txt  = escape(cell.text.strip())
                tag  = "th" if i == 0 else "td"
                cells_html.append(f'<{tag} class="{css}">{txt}</{tag}>')
            rows_html.append(f'<tr>{"".join(cells_html)}</tr>')
        return f'<div class="table-wrap"><table class="score-table">{"".join(rows_html)}</table></div>'
    
    if t == "score-legend":
        row = table.rows[0]
        cells_html = []
        for cell in row.cells:
            fill = get_cell_fill(cell)
            css  = COLOR_CLASS.get((fill or "FFFFFF").upper(), "")
            txt  = escape(cell.text.strip())
            cells_html.append(f'<div class="legend-cell {css}">{txt}</div>')
        return f'<div class="score-legend">{"".join(cells_html)}</div>'
    
    # Generic fallback: render as a simple responsive table
    rows_html = []
    for i, row in enumerate(table.rows):
        cells_html = []
        for cell in row.cells:
            fill = get_cell_fill(cell)
            css  = COLOR_CLASS.get((fill or "FFFFFF").upper(), "")
            txt  = escape(cell.text.strip())
            tag  = "th" if i == 0 else "td"
            cells_html.append(f'<{tag} class="{css}">{txt}</{tag}>')
        rows_html.append(f'<tr>{"".join(cells_html)}</tr>')
    return f'<div class="table-wrap"><table>{"".join(rows_html)}</table></div>'


def para_to_html(para):
    """Convert a paragraph to an HTML element."""
    text = para.text.strip()
    if not text:
        return None
    
    style = ""
    try:
        style = para.style.name if para.style else ""
    except Exception:
        style = ""
    
    fill  = get_para_shading(para)
    inner = runs_to_html(para) or escape(text)
    
    # Heading styles
    if style == "Heading 1":
        return f'<h1>{inner}</h1>'
    if style == "Heading 2":
        return f'<h2>{inner}</h2>'
    if style == "Heading 3":
        return f'<h3>{inner}</h3>'
    
    # List paragraph
    if style == "List Paragraph":
        return f'<li>{inner}</li>'
    
    # Paragraph with background shading = styled block
    if fill and fill.upper() in ("1E4D8C", "2E86C1", "1A7A4A", "B7860B", "1B2A4A"):
        css = COLOR_CLASS.get(fill.upper(), "fill-navy")
        return f'<div class="shaded-para {css}">{inner}</div>'
    
    return f'<p>{inner}</p>'


def doc_to_html_body(docx_path):
    """
    Convert a docx file to an HTML body string.
    Walks the document body, handling tables inline in document order.
    """
    doc = Document(docx_path)
    body = doc.element.body
    
    html_parts = []
    in_list = False
    
    # Walk direct children of body to preserve table/paragraph order
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        
        if tag == 'p':
            # Find matching paragraph object
            from docx.text.paragraph import Paragraph as DocxPara
            para = DocxPara(child, doc)
            
            result = para_to_html(para)
            
            if result is None:
                if in_list:
                    html_parts.append('</ul>')
                    in_list = False
                continue
            
            if result.startswith('<li>'):
                if not in_list:
                    html_parts.append('<ul>')
                    in_list = True
                html_parts.append(result)
            else:
                if in_list:
                    html_parts.append('</ul>')
                    in_list = False
                html_parts.append(result)
        
        elif tag == 'tbl':
            if in_list:
                html_parts.append('</ul>')
                in_list = False
            from docx.table import Table as DocxTable
            table = DocxTable(child, doc)
            html_parts.append(table_to_html(table))
    
    if in_list:
        html_parts.append('</ul>')
    
    return '\n'.join(html_parts)


CSS = """
:root {
  --navy:   #1E3A5F;
  --blue:   #2E6DA4;
  --green:  #1A7A4A;
  --gold:   #C8960C;
  --sky:    #D6E4F0;
  --ink:    #1A1A2E;
  --mid:    #6B7A8D;
  --mist:   #F5F7FA;
  --border: #DDE3EC;
  --white:  #FFFFFF;
  --serif:  'DM Serif Display', Georgia, serif;
  --sans:   'DM Sans', system-ui, sans-serif;
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: var(--sans);
  background: var(--mist);
  color: var(--ink);
  font-size: 15px;
  line-height: 1.75;
}

/* ── Header ── */
.site-header {
  background: var(--navy);
  padding: 16px 24px;
  border-bottom: 3px solid var(--gold);
  display: flex;
  align-items: center;
  gap: 12px;
  position: sticky;
  top: 0;
  z-index: 100;
}
.logo-mark {
  width: 34px; height: 34px;
  background: var(--gold);
  border-radius: 7px;
  display: flex; align-items: center; justify-content: center;
  font-family: var(--serif);
  font-size: 18px;
  color: var(--navy);
  font-weight: 700;
  flex-shrink: 0;
}
.header-text h1 {
  font-family: var(--serif);
  font-size: 16px;
  color: #fff;
  line-height: 1.2;
}
.header-text p { font-size: 11px; color: rgba(255,255,255,0.5); margin-top: 1px; }

/* ── Layout ── */
.doc-wrapper {
  max-width: 780px;
  margin: 0 auto;
  padding: 32px 20px 80px;
}
.doc-card {
  background: var(--white);
  border-radius: 10px;
  border: 1px solid var(--border);
  overflow: hidden;
  box-shadow: 0 2px 12px rgba(0,0,0,0.06);
}
.doc-title-bar {
  background: var(--navy);
  padding: 28px 32px;
  border-bottom: 4px solid var(--gold);
}
.doc-title-bar h2 {
  font-family: var(--serif);
  font-size: clamp(20px, 4vw, 28px);
  color: #fff;
  line-height: 1.2;
}
.doc-title-bar p { font-size: 13px; color: rgba(255,255,255,0.5); margin-top: 6px; }
.doc-body { padding: 36px 32px 56px; }

/* ── Typography ── */
h1 {
  font-family: var(--serif);
  font-size: clamp(20px, 4vw, 26px);
  color: var(--navy);
  margin: 40px 0 10px;
  padding-bottom: 10px;
  border-bottom: 2px solid var(--sky);
  line-height: 1.25;
}
h1:first-child { margin-top: 0; }
h2 {
  font-family: var(--serif);
  font-size: clamp(17px, 3vw, 21px);
  color: var(--blue);
  margin: 30px 0 8px;
  line-height: 1.3;
}
h3 {
  font-size: 12px;
  font-weight: 700;
  color: var(--navy);
  text-transform: uppercase;
  letter-spacing: 0.07em;
  margin: 24px 0 6px;
}
p { margin: 0 0 14px; }
p:last-child { margin-bottom: 0; }
strong { font-weight: 600; color: var(--navy); }
em { font-style: italic; color: var(--mid); }
ul { padding-left: 22px; margin: 0 0 16px; }
li { margin-bottom: 6px; }

/* ── Section label (colored header bar) ── */
.section-label {
  padding: 10px 16px;
  font-size: 11px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.1em;
  margin: 28px 0 12px;
  border-radius: 4px;
  color: var(--white);
}

/* ── School name bar ── */
.school-name-bar {
  display: flex;
  align-items: stretch;
  border-radius: 6px;
  overflow: hidden;
  margin: 28px 0 20px;
}
.school-name {
  background: var(--blue);
  color: #fff;
  font-family: var(--serif);
  font-size: clamp(16px, 3vw, 20px);
  padding: 14px 20px;
  flex: 1;
  display: flex;
  align-items: center;
}
.school-loc {
  background: var(--navy);
  color: rgba(255,255,255,0.7);
  font-size: 12px;
  font-style: italic;
  padding: 14px 16px;
  display: flex;
  align-items: center;
  white-space: nowrap;
}

/* ── What They Got Right header ── */
.got-right-header {
  background: var(--green);
  color: #fff;
  font-weight: 700;
  font-size: 14px;
  padding: 11px 16px;
  border-radius: 4px 4px 0 0;
  margin-top: 24px;
  margin-bottom: 0;
}

/* ── Two-column content table ── */
.two-col-table {
  border: 1px solid var(--border);
  border-radius: 0 0 6px 6px;
  overflow: hidden;
  margin-bottom: 24px;
  font-size: 14px;
}
.two-col-row {
  display: flex;
  border-bottom: 1px solid var(--border);
}
.two-col-row:last-child { border-bottom: none; }
.two-col-row.row-alt { background: var(--mist); }
.two-col-cell {
  flex: 1;
  padding: 10px 14px;
  line-height: 1.55;
  border-right: 1px solid var(--border);
}
.two-col-cell:last-child { border-right: none; }

/* ── Takeaway box ── */
.takeaway-box {
  border: 1px solid #E8D98A;
  border-radius: 6px;
  overflow: hidden;
  margin: 12px 0 24px;
}
.takeaway-item {
  display: flex;
  gap: 14px;
  padding: 14px 16px;
  border-bottom: 1px solid #E8D98A;
  line-height: 1.6;
  font-size: 14px;
}
.takeaway-item:nth-child(even) { background: #FFFDF4; }
.takeaway-item:nth-child(odd)  { background: #FEF9EC; }
.takeaway-item:last-child { border-bottom: none; }
.takeaway-num {
  width: 26px; height: 26px;
  background: var(--gold);
  border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-size: 12px;
  font-weight: 700;
  color: var(--navy);
  flex-shrink: 0;
  margin-top: 1px;
}
.takeaway-text { flex: 1; color: var(--ink); }

/* ── Info grid (school index) ── */
.info-grid { margin: 16px 0 24px; display: flex; flex-direction: column; gap: 4px; }
.info-row  { display: flex; gap: 4px; }
.info-cell {
  flex: 1;
  padding: 10px 14px;
  font-size: 14px;
  border-radius: 4px;
  line-height: 1.5;
}

/* ── Principle headers (synthesis page) ── */
.principle-header {
  background: var(--navy);
  color: #fff;
  padding: 13px 20px;
  border-radius: 4px;
  font-size: 15px;
  font-weight: 600;
  border-left: 5px solid var(--gold);
  margin: 28px 0 8px;
}
.principle-label {
  color: var(--gold);
  font-weight: 700;
  margin-right: 4px;
}

/* ── Shaded paragraph ── */
.shaded-para {
  padding: 10px 16px;
  border-radius: 4px;
  color: #fff;
  font-size: 11px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.1em;
  margin: 24px 0 10px;
}

/* ── Score table (audit docs) ── */
.table-wrap { overflow-x: auto; margin: 16px 0 20px; }
table { width: 100%; border-collapse: collapse; font-size: 14px; min-width: 400px; }
th, td { border: 1px solid var(--border); padding: 9px 12px; vertical-align: top; line-height: 1.5; }
th { background: var(--navy); color: #fff; font-weight: 600; font-size: 13px; }
tr:nth-child(even) td:not([class]) { background: var(--mist); }

/* ── Score legend ── */
.score-legend {
  display: flex;
  gap: 6px;
  flex-wrap: wrap;
  margin: 12px 0 20px;
}
.legend-cell {
  padding: 5px 12px;
  border-radius: 20px;
  font-size: 12px;
  font-weight: 600;
  border: 1px solid var(--border);
}

/* ── Fill color classes ── */
.fill-navy       { background: #1E3A5F; color: #fff; }
.fill-blue       { background: #2E6DA4; color: #fff; }
.fill-green      { background: #1A7A4A; color: #fff; }
.fill-gold       { background: #C8960C; color: #fff; }
.fill-sky        { background: #D6E4F0; color: var(--ink); }
.fill-lightgray  { background: #F4F6F8; color: var(--ink); }
.fill-goldtint-a { background: #FEF9EC; color: var(--ink); }
.fill-goldtint-b { background: #FFFDF4; color: var(--ink); }
.fill-white      { background: #FFFFFF; color: var(--ink); }
.fill-green-light{ background: #C8E6C9; color: #1A4A1E; }
.fill-green-pale { background: #DCEDC8; color: #1A4A1E; }
.fill-yellow     { background: #FFF9C4; color: #5A4A00; }
.fill-red-light  { background: #FFCDD2; color: #7A1A1A; }
.fill-purple-light{ background: #EDE7F6; color: #3A1A6A; }
.fill-sky-pale   { background: #EEF4FF; color: var(--ink); }

/* ── Footer ── */
.site-footer {
  text-align: center;
  padding: 16px;
  font-size: 12px;
  color: var(--mid);
  border-top: 1px solid var(--border);
  background: var(--white);
}

/* ── Mobile ── */
@media (max-width: 600px) {
  .doc-body { padding: 22px 18px 40px; }
  .doc-title-bar { padding: 20px 20px; }
  .two-col-row { flex-direction: column; }
  .two-col-cell { border-right: none; border-bottom: 1px solid var(--border); }
  .two-col-cell:last-child { border-bottom: none; }
  .school-name-bar { flex-direction: column; }
  .school-loc { justify-content: flex-start; }
  .info-row { flex-direction: column; }
  .score-legend { gap: 4px; }
}
"""

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title} — Spark Academy</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,300;1,400&display=swap" rel="stylesheet">
<style>{css}</style>
</head>
<body>
<header class="site-header">
  <div class="logo-mark">S</div>
  <div class="header-text">
    <h1>Spark Docs</h1>
    <p>Spark Academy Website Project</p>
  </div>
</header>
<div class="doc-wrapper">
  <div class="doc-card">
    <div class="doc-title-bar">
      <h2>{title}</h2>
      <p>Spark Academy &middot; Internal Document</p>
    </div>
    <div class="doc-body">
{body}
    </div>
  </div>
</div>
<footer class="site-footer">
  Spark Academy Website Project &nbsp;&middot;&nbsp; For internal use only
</footer>
</body>
</html>"""


def convert(docx_path, output_path=None):
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)

    print(f"Converting: {docx_path}")
    body = doc_to_html_body(docx_path)

    title = os.path.basename(docx_path).replace('.docx', '').replace('_', ' ').replace('-', ' ')
    html  = HTML_TEMPLATE.format(title=escape(title), css=CSS, body=body)

    if output_path is None:
        slug        = os.path.basename(docx_path).replace('.docx', '').replace(' ', '-').lower()
        output_path = slug + '.html'

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"Done: {output_path}")
    return output_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 convert.py input.docx [output.html]")
        sys.exit(1)
    
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    convert(inp, out)
