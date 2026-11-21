#!/usr/bin/env python3
"""
build_title_page.py

Creates standalone Title_Page.docx for Elsevier / Vehicular Communications submission,
enforcing double-blind peer review separation and journal editorial standards.
"""

from pathlib import Path
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from doc_helpers import apply_text_styling, add_page_number_to_footer

def create_title_page(output_path):
    doc = docx.Document()
    
    # Page setup: Standard Letter, 2.5 cm margins
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    
    add_page_number_to_footer(section)

    # Document Header / Category
    p_cat = doc.add_paragraph()
    r_cat = p_cat.add_run("ORIGINAL RESEARCH PAPER — TITLE PAGE")
    r_cat.bold = True
    r_cat.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    apply_text_styling(p_cat, font_size=10, line_spacing=1.5, space_before=0, space_after=12, align=WD_ALIGN_PARAGRAPH.RIGHT)

    # Title
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit and Cross-Platform Domain Analysis")
    r_title.bold = True
    apply_text_styling(p_title, font_size=14, line_spacing=1.5, space_before=12, space_after=10, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Running Title
    p_run = doc.add_paragraph()
    r_run = p_run.add_run("Running Title: Empirical Deployability Audit of Automotive CAN-IDS")
    r_run.italic = True
    apply_text_styling(p_run, font_size=11, line_spacing=1.5, space_before=0, space_after=18, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Authors
    p_auth = doc.add_paragraph()
    r1 = p_auth.add_run("Vishwas Khandelwal")
    r1.bold = True
    r1_sup = p_auth.add_run("1,*")
    r1_sup.font.superscript = True
    
    r_sep = p_auth.add_run(", ")
    
    r2 = p_auth.add_run("[Co-Authors / Reviewers TBD]")
    r2.bold = True
    r2.italic = True
    r2_sup = p_auth.add_run("1")
    r2_sup.font.superscript = True
    apply_text_styling(p_auth, font_size=12, line_spacing=1.5, space_before=0, space_after=8, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Affiliations
    p_aff = doc.add_paragraph()
    r_a1 = p_aff.add_run("1 ")
    r_a1.font.superscript = True
    p_aff.add_run("Department of Computer Science and Engineering, Vivekananda Institute of Professional Studies - Technical Campus (VIPS-TC), Delhi 110034, India\n")
    apply_text_styling(p_aff, font_size=10.5, line_spacing=1.5, space_before=0, space_after=18, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Corresponding Author
    p_corr = doc.add_paragraph()
    r_corr_hdr = p_corr.add_run("*Corresponding Author:\n")
    r_corr_hdr.bold = True
    p_corr.add_run("Vishwas Khandelwal\n")
    p_corr.add_run("Department of Computer Science and Engineering, Vivekananda Institute of Professional Studies - Technical Campus (VIPS-TC), Delhi 110034, India\n")
    p_corr.add_run("Email: vishukh767@gmail.com\n")
    apply_text_styling(p_corr, font_size=11, line_spacing=1.5, space_before=0, space_after=18)

    # Highlights (Strictly <= 85 characters per bullet)
    p_hl_hdr = doc.add_paragraph()
    r_hl = p_hl_hdr.add_run("Highlights (Elsevier 85-Character Requirement):")
    r_hl.bold = True
    apply_text_styling(p_hl_hdr, font_size=11.5, line_spacing=1.5, space_before=6, space_after=6)

    highlights = [
        "Time-blocked split reveals genuine CAN-IDS generalization without lookahead bias.",
        "Single-feature ID tree captures 76.6% of multi-feature gain over random chance.",
        "Direct model transfer across vehicle architectures collapses to 91.5% FPR.",
        "Multi-frame whole-bus debounce suppresses burst false alarms by up to 99.8%.",
        "Supervised few-shot adaptation converges to 0.00% FPR within 1.0 s of driving."
    ]
    for h in highlights:
        p_h = doc.add_paragraph()
        r_b = p_h.add_run("• ")
        r_b.bold = True
        p_h.add_run(h)
        apply_text_styling(p_h, font_size=10.5, line_spacing=1.3, space_before=0, space_after=4)

    # Abstract (<250 words)
    p_abs_hdr = doc.add_paragraph()
    r_ah = p_abs_hdr.add_run("Abstract")
    r_ah.bold = True
    apply_text_styling(p_abs_hdr, font_size=12, line_spacing=1.5, space_before=12, space_after=6)

    p_abs = doc.add_paragraph()
    p_abs.add_run(
        "In-vehicle intrusion detection systems (CAN-IDS) report classification accuracies exceeding 99.9% on automotive benchmarks. This performance ceiling collapses in production. Evaluating Random Forest ensembles on chronological 70/30 time-blocked splits of the HCRL dataset (Hyundai YF Sonata), we reproduce baseline metrics (99.90% accuracy, 0.999 MCC, 0.091% FPR) and isolate operational failure modes. An arbitration ID single-feature tree achieves 82.40% accuracy; it isolates 76.64% of achievable gain above chance (25.00%). Translating the 0.091% benchmark FPR to physical bus line speeds produces 2,621 to 13,104 unmitigated false alarms per hour across standard automotive transceivers operating between 250 kbps and 1 Mbps. Whole-bus debouncing (k = 3) theoretically suppresses Poisson errors to one alarm per 7.7 days at 500 kbps. Dynamic driving maneuvers break this independence. Auditing 200,000 continuous normal frames reveals 1,863 alarms/hr under whole-bus filtering and 510,134 alarms/hr per-ID. Adversarial feature blending reduces attack recall to 19.60%; curriculum retraining recovers recall to 98.60% (alpha = 0.25) but increases nominal FPR to 0.246% (17,700 alarms/hr). Transfer to a Polaris GEM e6 shuttle incurs 91.48% FPR from 11-bit versus 29-bit SAE J1939 mismatches. Calibration-normalized stateful encoding restores zero-shot MCC to 0.8316. Stratified few-shot adaptation achieves 0.04% FPR at N = 500 frames (1.0 s) and converges to 0.00% FPR with 100.00% tampering recall at N >= 1,000 frames (2.0 s). On ORNL ROAD physical masquerade attacks, timing baselines record 0.00% recall, whereas payload tree ensembles maintain 100.00% recall. Vehicle-layer detectors must not command physical actuators without multi-sensor kinematic consensus under ISO 26262 and ISO/SAE 21434."
    )
    apply_text_styling(p_abs, font_size=11, line_spacing=1.5, space_before=0, space_after=12)

    # Keywords
    p_kw = doc.add_paragraph()
    r_kw = p_kw.add_run("Keywords: ")
    r_kw.bold = True
    p_kw.add_run("Controller Area Network (CAN); Intrusion Detection Systems (IDS); Machine Learning; Empirical Evaluation; Cross-Platform Generalization; Vehicular Cybersecurity.")
    apply_text_styling(p_kw, font_size=10.5, line_spacing=1.5, space_before=0, space_after=18)

    # Manuscript Metadata
    p_meta = doc.add_paragraph()
    r_meta = p_meta.add_run("Manuscript Metadata:\n")
    r_meta.bold = True
    p_meta.add_run("• Target Journal: Vehicular Communications (Elsevier, ISSN: 2214-2096, CiteScore: 14.4, Impact Factor: 6.7)\n")
    p_meta.add_run("• Article Type: Original Research Paper\n")
    p_meta.add_run("• Word Count (Abstract): 250 words (< 250-word journal limit)\n")
    p_meta.add_run("• Word Count (Main Text, excluding references/tables/figures): ~6,200 words\n")
    p_meta.add_run("• Number of Tables: 9 (Tables 1–9)\n")
    p_meta.add_run("• Number of Figures: 4 (Figures 1–4)\n")
    p_meta.add_run("• Number of Supplementary Tables: 8 (Tables S1–S8)\n")
    apply_text_styling(p_meta, font_size=10.5, line_spacing=1.5, space_before=6, space_after=18)

    # Declarations and Statements
    p_dec_hdr = doc.add_paragraph()
    r_dh = p_dec_hdr.add_run("Declarations and Policy Statements")
    r_dh.bold = True
    apply_text_styling(p_dec_hdr, font_size=12, line_spacing=1.5, space_before=12, space_after=8)

    # CRediT statement
    p_cr = doc.add_paragraph()
    r_c = p_cr.add_run("CRediT Author Statement: ")
    r_c.bold = True
    p_cr.add_run("Vishwas Khandelwal: Conceptualization, Methodology, Software, Validation, Formal analysis, Investigation, Data Curation, Writing - Original Draft, Writing - Review & Editing, Visualization.")
    apply_text_styling(p_cr, font_size=10.5, line_spacing=1.5, space_before=0, space_after=8)

    # Competing Interest
    p_coi = doc.add_paragraph()
    r_coi = p_coi.add_run("Declaration of Competing Interest: ")
    r_coi.bold = True
    p_coi.add_run("The authors declare that they have no known competing financial interests or personal relationships that could have appeared to influence the work reported in this paper.")
    apply_text_styling(p_coi, font_size=10.5, line_spacing=1.5, space_before=0, space_after=8)

    # Funding
    p_fnd = doc.add_paragraph()
    r_fnd = p_fnd.add_run("Funding Statement: ")
    r_fnd.bold = True
    p_fnd.add_run("This research did not receive any specific grant from funding agencies in the public, commercial, or not-for-profit sectors.")
    apply_text_styling(p_fnd, font_size=10.5, line_spacing=1.5, space_before=0, space_after=8)

    # Data and Code Availability
    p_data = doc.add_paragraph()
    r_data = p_data.add_run("Data and Code Availability: ")
    r_data.bold = True
    p_data.add_run(
        "All datasets, preprocessing scripts, evaluation pipelines, and manuscript generators are open source. "
        "The HCRL Car-Hacking dataset is accessible via Korea University HCRL (DOI: 10.1145/3064814.3064840). "
        "The GEM-CAN autonomous vehicle dataset is published in Data in Brief (DOI: 10.1016/j.dib.2026.112805, Zenodo: 10.5281/zenodo.19161139). "
        "All experimental reproduction code is available in the project repository."
    )
    apply_text_styling(p_data, font_size=10.5, line_spacing=1.5, space_before=0, space_after=8)

    # AI Disclosure
    p_ai = doc.add_paragraph()
    r_lbl3 = p_ai.add_run("Artificial Intelligence Tool Disclosure: ")
    r_lbl3.bold = True
    p_ai.add_run(
        "Generative AI tools (Google Antigravity / Gemini 3.8) were utilized solely for code scripting assistance, text syntax alignment, "
        "and automated document formatting. All study hypotheses, machine learning pipelines, statistical computations, and scientific conclusions "
        "were conceptualized, independently verified, and approved by the author."
    )
    apply_text_styling(p_ai, font_size=10.5, line_spacing=1.5, space_before=0, space_after=12)

    doc.save(output_path)
    print(f"Successfully generated Title Page: {output_path}")

if __name__ == '__main__':
    base_dir = Path(__file__).resolve().parent
    tp_path = base_dir / "Title_Page.docx"
    create_title_page(tp_path)
