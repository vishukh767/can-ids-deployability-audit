#!/usr/bin/env python3
"""
build_cover_letter.py

Generates a formal academic cover letter in DOCX format for submission to
the Editor-in-Chief of Vehicular Communications (Elsevier).
"""

from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def create_cover_letter(output_path):
    doc = docx.Document()
    
    # 1.0 inch margins
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        
    def add_p(text, before=0, after=6, spacing=1.15, bold=False, italic=False, color=(0x26, 0x26, 0x26), size=11, align=WD_ALIGN_PARAGRAPH.LEFT):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = spacing
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(size)
        r.font.color.rgb = RGBColor(*color)
        r.bold = bold
        r.italic = italic
        return p

    # Sender Letterhead / Affiliation
    add_p("Department of Computer Science and Engineering", before=0, after=2, bold=True, size=11, color=(0x1F, 0x4E, 0x79))
    add_p("Vivekananda Institute of Professional Studies - Technical Campus (VIPS-TC)\nOuter Ring Road, AU Block, Pitampura, Delhi 110034, India\nEmail: vishukh767@gmail.com", before=0, after=14, size=10, color=(0x55, 0x55, 0x55))

    # Date
    add_p("October 19, 2026", before=0, after=12, size=11)

    # Recipient
    add_p("To:\nProfessor Nei Kato and Professor Celimuge Wu\nEditors-in-Chief, Vehicular Communications\nElsevier B.V.", before=0, after=14, bold=False, size=11)

    # Subject
    add_p("SUBJECT: Submission of Original Research Paper — 'The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit and Cross-Platform Domain Analysis'", before=0, after=12, bold=True, size=11, color=(0x1F, 0x4E, 0x79))

    # Salutation
    add_p("Dear Professors Kato and Wu, and Esteemed Members of the Editorial Board,", before=0, after=8, size=11)

    # Opening Paragraph
    add_p(
        "We are pleased to submit our original research manuscript entitled \"The Brittleness of Vehicle-Layer AI Intrusion Detection: "
        "An Empirical Deployability Audit and Cross-Platform Domain Analysis\" for consideration for publication in Vehicular Communications "
        "(Elsevier, ISSN: 2214-2096, CiteScore: 14.4, Impact Factor: 6.7).",
        before=0, after=8
    )

    # Context & Problem Statement
    add_p(
        "In-vehicle intrusion detection systems (CAN-IDS) based on machine learning routinely report near-perfect classification accuracies (>99.9%) "
        "on canonical laboratory benchmarks. However, deploying these classifiers onto production electronic control units (ECUs) governed by "
        "ISO 26262 (functional safety) and ISO/SAE 21434 (cybersecurity engineering) reveals profound operational fragilities. Offline evaluation "
        "often masks temporal lookahead leakage, creates unrealistic expectations regarding bus line-speed false-alarm frequencies, and fails to evaluate "
        "cross-platform protocol transferability. In this paper, we present an exhaustive empirical deployability audit evaluating tree-based classifiers, "
        "gradient-boosted models, and deep autoencoders across three distinct vehicular environments: the Hyundai Sonata (HCRL benchmark), "
        "the Polaris GEM e6 autonomous campus shuttle (GEM-CAN benchmark), and real physical masquerade attack traces from the Oak Ridge National "
        "Laboratory (ORNL ROAD benchmark).",
        before=0, after=8
    )

    # Bulleted Contributions
    add_p("Key Empirical Contributions and Findings of this Work:", before=4, after=4, bold=True, color=(0x1F, 0x4E, 0x79))

    points = [
        ("Elimination of Lookahead Temporal Leakage: ",
         "We prove that standard random train/test splits fragment periodic CAN streams, enabling models to memorize static periodic signatures. "
         "By implementing a rigorous 5-seed chronological 70/30 time-blocked split, we establish genuine generalization bounds and realistic baseline performance (99.90% accuracy, 0.999 MCC, 0.091% FPR)."),
        ("Single-Feature Allowlist Memorization Diagnostic: ",
         "We demonstrate that a single-feature decision tree thresholding strictly on arbitration CAN ID achieves 82.40% accuracy, capturing 76.6% of the multi-feature model's gain over random chance (25.0%) without inspecting payload data or timing intervals. This exposes that canonical benchmark success is heavily driven by identifier allowlisting rather than anomaly detection."),
        ("Operational Line-Speed Translation & Dynamic Burstiness Proof: ",
         "We translate the 0.091% evaluation FPR into operational line-speed alarm volumes across standard bitrates (250 kbps, 500 kbps, 1 Mbps), showing that unmitigated models produce 2,621 to 13,104 false alarms per driving hour. While a k=3 temporal debounce filter theoretically suppresses independent errors to one alarm every 7.7 operating days at 500 kbps (1.5 ms latency), empirical evaluation across 200,000 continuous normal frames reveals that transient vehicle dynamic maneuvers induce severe, correlated per-ID error bursts (14,514 false alerts), disproving independent Bernoulli assumptions."),
        ("Cross-Platform Transfer Collapse & Calibration-Normalized Solution: ",
         "Evaluating the Sonata-trained classifier directly on the Polaris GEM e6 shuttle results in a 91.48% FPR (35.57% accuracy, 0.1637 MCC) due to the architectural protocol mismatch between 11-bit standard and 29-bit extended SAE J1939 identifiers. We formulate a protocol-agnostic Calibration-Normalized Stateful Encoding that standardizes telemetry against nominal baseline percentiles, restoring cross-platform MCC to 0.8316. Stratified few-shot adaptation reaches 0.04% FPR at N = 500 frames (1.0 s) and converges to 0.00% FPR with 100.00% tampering recall at N >= 1,000 frames (2.0 s of driving data)."),
        ("Multi-Detector Audit & Physical Masquerade Validation: ",
         "We evaluate four distinct detector families (stateless timing rules, Random Forests, LightGBM, and Deep Reconstruction Autoencoders), verifying that cross-platform representation collapse is an inherent protocol boundary reality across all architectures. Against real-world timing-preserving masquerade attacks from the ORNL ROAD benchmark, stateless timing baselines collapse to 0.00% recall, whereas payload-aware tree ensembles achieve 100.00% detection (0.9993 MCC) through non-linear drivetrain signal correlation without requiring proprietary DBC schemas.")
    ]

    for p_title, p_desc in points:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Inches(0.25)
        p_item.paragraph_format.space_before = Pt(2)
        p_item.paragraph_format.space_after = Pt(4)
        p_item.paragraph_format.line_spacing = 1.15
        r_bullet = p_item.add_run("• ")
        r_bullet.bold = True
        r_bullet.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        r_t = p_item.add_run(p_title)
        r_t.bold = True
        r_t.font.name = "Calibri"
        r_t.font.size = Pt(10.5)
        r_d = p_item.add_run(p_desc)
        r_d.font.name = "Calibri"
        r_d.font.size = Pt(10.5)

    # Scope & Fit
    add_p(
        "This manuscript aligns directly with the core scope of Vehicular Communications, particularly in the areas of in-vehicle network "
        "security, connected and autonomous vehicle safety architectures, machine learning deployability, and safety-security co-engineering. "
        "Our work provides actionable design guidelines for automotive OEMs and Tier-1 suppliers reconciling ISO 26262 fail-safe and ISO/SAE 21434 "
        "fail-secure mandates.",
        before=6, after=8
    )

    # Compliance & Declarations
    add_p(
        "Mandatory Journal Declarations:\n"
        "1. Originality: This manuscript describes original research that has not been published previously and is not currently under consideration for publication elsewhere.\n"
        "2. Authorship: The manuscript has been approved for submission by all authors.\n"
        "3. Competing Interests: The authors declare no competing financial or personal interests.\n"
        "4. Open Science & Reproducibility: All parsed datasets, empirical logs, and programmatic reproduction scripts are fully documented in the manuscript and archived for open scientific verification.",
        before=4, after=12, size=10, color=(0x44, 0x44, 0x44)
    )

    # Sign-off
    add_p("Thank you very much for your time, consideration, and management of the peer-review process.", before=4, after=14)
    add_p("Sincerely,\n\nVishwas Khandelwal\nCorresponding Author\nDepartment of Computer Science and Engineering\nVivekananda Institute of Professional Studies - Technical Campus (VIPS-TC), Delhi, India\nEmail: vishukh767@gmail.com", before=0, after=0, bold=True, size=10.5, color=(0x1F, 0x4E, 0x79))

    doc.save(str(output_path))
    alt_path = Path(output_path).parent / "1_manuscripts_and_drafts" / Path(output_path).name
    if alt_path.parent.is_dir():
        doc.save(str(alt_path))
    print(f"Successfully generated Cover Letter: {output_path}")

if __name__ == '__main__':
    base_dir = Path(__file__).resolve().parent
    out_file = base_dir / "Cover_Letter.docx"
    create_cover_letter(out_file)
