#!/usr/bin/env python3
"""
build_highlights.py

Creates standalone Highlights.docx for Elsevier / Vehicular Communications submission.
Contains strictly 3-5 bullet points, each strictly <= 85 characters including spaces.
"""

from pathlib import Path
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_highlights(output_path):
    doc = docx.Document()
    
    # Page setup: Standard Letter, 2.5 cm margins
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # Document Header / Title
    p_hdr = doc.add_paragraph()
    r_hdr = p_hdr.add_run("Highlights")
    r_hdr.bold = True
    r_hdr.font.name = "Calibri"
    r_hdr.font.size = Pt(14)
    r_hdr.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(12)

    # Subtitle / Manuscript Title Reference
    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit and Cross-Platform Domain Analysis")
    r_sub.italic = True
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(10.5)
    r_sub.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)

    highlights = [
        "Time-blocked split reveals genuine CAN-IDS generalization without lookahead bias.",
        "Single-feature ID tree captures 76.6% of multi-feature gain over random chance.",
        "Direct model transfer across vehicle architectures collapses to 91.5% FPR.",
        "Multi-frame whole-bus debounce suppresses burst false alarms by up to 99.8%.",
        "Supervised few-shot adaptation converges to 0.00% FPR within 1.0 s of driving."
    ]

    for idx, h in enumerate(highlights, 1):
        # Character count validation
        char_count = len(h)
        assert char_count <= 85, f"Highlight {idx} exceeds 85 characters: '{h}' ({char_count} chars)"
        
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.2
        p.paragraph_format.left_indent = Inches(0.25)
        
        r_bullet = p.add_run("• ")
        r_bullet.bold = True
        r_bullet.font.name = "Calibri"
        r_bullet.font.size = Pt(11)
        r_bullet.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        
        r_text = p.add_run(h)
        r_text.font.name = "Calibri"
        r_text.font.size = Pt(11)
        r_text.font.color.rgb = RGBColor(0x26, 0x26, 0x26)

    doc.save(str(output_path))
    alt_path = Path(output_path).parent / "1_manuscripts_and_drafts" / Path(output_path).name
    if alt_path.parent.is_dir():
        doc.save(str(alt_path))
    print(f"Successfully generated Highlights: {output_path} (and copy in {alt_path})")

if __name__ == '__main__':
    base_dir = Path(__file__).resolve().parent
    out_file = base_dir / "Highlights.docx"
    create_highlights(out_file)
