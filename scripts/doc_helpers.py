#!/usr/bin/env python3
"""
doc_helpers.py

Standardized Word Document (DOCX) styling and XML helper functions
for Elsevier / Vehicular Communications submissions.
"""

import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set inner cell padding in twips (dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m_name, m_val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m_name}')
        node.set(qn('w:w'), str(m_val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_apa_table_borders(table):
    """
    Format table with formal 3-line rules:
    top rule (1 pt), header bottom rule (0.75 pt), table bottom rule (1 pt).
    No vertical rules.
    """
    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblBorders'):
            tblPr.remove(child)
            
    tblBorders = OxmlElement('w:tblBorders')
    
    # Top rule (1 pt = sz 8)
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '8')
    top.set(qn('w:space'), '0')
    top.set(qn('w:color'), '000000')
    tblBorders.append(top)
    
    # Bottom rule (1 pt = sz 8)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '8')
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), '000000')
    tblBorders.append(bottom)
    
    # No vertical, left, right, or inside borders
    for b_name in ['left', 'right', 'insideH', 'insideV']:
        el = OxmlElement(f'w:{b_name}')
        el.set(qn('w:val'), 'none')
        tblBorders.append(el)
        
    tblPr.append(tblBorders)
    
    # Header bottom rule (0.75 pt = sz 6) on row 0 cells
    if len(table.rows) > 0:
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            b = OxmlElement('w:bottom')
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'), '6')
            b.set(qn('w:space'), '0')
            b.set(qn('w:color'), '000000')
            tcBorders.append(b)
            tcPr.append(tcBorders)

def add_page_number_to_footer(section):
    """Add right-aligned page number field to section footer."""
    footer = section.footer
    f_p = footer.paragraphs[0]
    f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f_p.text = ""
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'PAGE')
    f_p._p.append(fldSimple)
    for r in f_p.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10)

def apply_text_styling(paragraph, font_name="Times New Roman", font_size=12, line_spacing=2.0, space_before=0, space_after=0, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    """Apply standard typography to a paragraph."""
    paragraph.alignment = align
    p_format = paragraph.paragraph_format
    p_format.line_spacing = line_spacing
    p_format.space_before = Pt(space_before)
    p_format.space_after = Pt(space_after)
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
