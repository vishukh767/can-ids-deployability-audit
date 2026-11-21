#!/usr/bin/env python3
"""
build_supplementary_tables.py

Builds Supplementary_Tables.docx for Elsevier / Vehicular Communications submission.
Includes Supplementary Tables S1, S2, S3, S4, and S5 conforming to APA 3-line rules.
Fully synchronized with Vehicular_Communications_Manuscript_v20.docx and Title_Page.docx.
"""

import os
from pathlib import Path
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from doc_helpers import (
    set_cell_margins, set_apa_table_borders, add_page_number_to_footer,
    apply_text_styling
)

def create_supplementary_tables(output_path):
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

    # Document Header
    p_head = doc.add_paragraph()
    r = p_head.add_run("Supplementary Material: The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit and Cross-Platform Domain Analysis")
    r.bold = True
    apply_text_styling(p_head, font_size=13, line_spacing=1.5, space_before=12, space_after=18, align=WD_ALIGN_PARAGRAPH.CENTER)

    # =========================================================================
    # SUPPLEMENTARY TABLE S1
    # =========================================================================
    p_s1_tit = doc.add_paragraph()
    r = p_s1_tit.add_run("Supplementary Table S1. ")
    r.bold = True
    p_s1_tit.add_run("In-distribution per-class metrics across time-blocked evaluation splits on the HCRL dataset (24,000 test frames per seed: 6,000 per class), comparing Baseline Multi-Feature Random Forest vs. Single-Feature (ID-Only) Decision Tree.")
    apply_text_styling(p_s1_tit, font_size=11, line_spacing=1.5, space_before=12, space_after=6)

    s1_headers = [
        "Class Label", "Test Frames / Seed", "Multi-Feature RF Recall (95% CI)",
        "ID-Only Tree Recall (95% CI)", "Multi-Feature Precision", "ID-Only Precision"
    ]
    s1_rows = [
        ["Normal Driving", "6,000", "99.85% [99.72%, 99.92%]", "80.83% [79.81%, 81.81%]", "99.77%", "88.62%"],
        ["DoS Attack", "6,000", "100.00% [99.94%, 100.00%]", "100.00% [99.94%, 100.00%]", "100.00%", "100.00%"],
        ["Fuzzy Attack", "6,000", "99.75% [99.58%, 99.85%]", "98.30% [97.94%, 98.60%]", "99.85%", "85.14%"],
        ["Spoofing (Gear / RPM)", "6,000", "100.00% [99.94%, 100.00%]", "89.40% [88.58%, 90.16%]", "99.98%", "68.20%"],
        ["Overall Balanced Total", "24,000", "99.90% +/- 0.008%", "82.40% +/- 0.12%", "99.90%", "80.59%"]
    ]

    t_s1 = doc.add_table(rows=1, cols=len(s1_headers))
    t_s1.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h_text in enumerate(s1_headers):
        cell = t_s1.cell(0, c_idx)
        cell.text = h_text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        apply_text_styling(p, font_size=8.5, line_spacing=1.15, space_before=2, space_after=2)
        set_cell_margins(cell, top=60, bottom=60, left=70, right=70)

    for r_data in s1_rows:
        row_cells = t_s1.add_row().cells
        for c_idx, val in enumerate(r_data):
            cell = row_cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            if r_data[0].startswith("Overall"):
                p.runs[0].bold = True
            apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
            set_cell_margins(cell, top=40, bottom=40, left=70, right=70)
    set_apa_table_borders(t_s1)

    p_s1_fn = doc.add_paragraph()
    r = p_s1_fn.add_run(
        "Note: Wilson score 95% confidence intervals are computed over 5 independent time-blocked evaluation seeds (24,000 test frames each). "
        "The single-feature Decision Tree thresholding strictly on arbitration CAN ID achieves 82.40% accuracy, capturing 76.6% of the multi-feature model's "
        "gain over chance ((82.40 - 25.00) / (99.90 - 25.00) = 76.64%). Because DoS attacks inject high-priority ID 0x000 and Fuzzy attacks inject random IDs, "
        "the ID-only tree achieves 100.0% and 98.3% recall on those classes, respectively. However, recall drops to 89.4% on Spoofing attacks where adversaries "
        "reuse valid allowlist arbitration IDs, proving that the single-feature baseline operates as an allowlist memorizer rather than an anomaly detector."
    )
    r.italic = True
    apply_text_styling(p_s1_fn, font_size=9.0, line_spacing=1.3, space_before=6, space_after=18)

    doc.add_page_break()

    # =========================================================================
    # SUPPLEMENTARY TABLE S2
    # =========================================================================
    p_s2_tit = doc.add_paragraph()
    r = p_s2_tit.add_run("Supplementary Table S2. ")
    r.bold = True
    p_s2_tit.add_run("GEM-CAN Cross-Platform Evaluation Breakdown across 143,444 frames (101,039 Normal vs. 42,405 Attack: 41,080 DoS + 1,325 Tampering).")
    apply_text_styling(p_s2_tit, font_size=11, line_spacing=1.5, space_before=12, space_after=6)

    s2_headers = [
        "Transfer Paradigm", "Accuracy", "MCC", "FPR",
        "DoS Rec (41.1k)", "Tamper Rec (1.3k)", "TP", "FP", "TN", "FN"
    ]
    s2_rows = [
        ["Direct Transfer (Raw IDs)", "35.57%", "0.1637", "91.48%", "100.00%", "100.00%", "42,405", "92,427", "8,612", "0"],
        ["Payload-Only Ablation", "53.62%", "0.3646", "65.85%", "100.00%", "100.00%", "42,405", "66,534", "34,505", "0"],
        ["Calibration-Normalized Stateful", "92.04%", "0.8316", "10.65%", "100.00%", "50.04%", "41,743", "10,763", "90,276", "662"],
        ["Native Within-GEM (70/30 Test Partition)", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%", "12,739", "0", "30,295", "0"]
    ]

    t_s2 = doc.add_table(rows=1, cols=len(s2_headers))
    t_s2.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h_text in enumerate(s2_headers):
        cell = t_s2.cell(0, c_idx)
        cell.text = h_text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)

    for r_data in s2_rows:
        row_cells = t_s2.add_row().cells
        for c_idx, val in enumerate(r_data):
            cell = row_cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
    set_apa_table_borders(t_s2)

    p_s2_fn = doc.add_paragraph()
    r = p_s2_fn.add_run(
        "Note: Telemetry parsed from the Polaris GEM e6 autonomous shuttle dataset (Tavasoli et al., Data in Brief 2026). "
        "Direct Transfer yields MCC = 0.1637 and 91.48% FPR because extended 29-bit SAE J1939 IDs fall outside the 11-bit range of the Hyundai Sonata training data "
        "and branch into attack leaf nodes. For Native Within-GEM training, metrics are reported on the held-out 30% time-blocked test partition (30,295 Normal and "
        "12,739 Attack frames: 12,342 DoS + 397 Tampering). Calibration-normalized features restore cross-platform MCC to 0.8316 while maintaining 100.0% DoS recall "
        "and 50.04% tampering recall (detecting dynamic brake tampering while subtle steering angle manipulations remain camouflaged)."
    )
    r.italic = True
    apply_text_styling(p_s2_fn, font_size=9.0, line_spacing=1.3, space_before=6, space_after=18)

    doc.add_page_break()

    # =========================================================================
    # SUPPLEMENTARY TABLE S3
    # =========================================================================
    p_s3_tit = doc.add_paragraph()
    r = p_s3_tit.add_run("Supplementary Table S3. ")
    r.bold = True
    p_s3_tit.add_run("Empirical Continuous Bus False Alarm & Debounce Burstiness on 200,000 Continuous Nominal Frames (normal_run_data.txt).")
    apply_text_styling(p_s3_tit, font_size=11, line_spacing=1.5, space_before=12, space_after=6)

    s3_headers = [
        "Debounce Scheme", "Thresh (k)", "Scope", "Events",
        "Empirical FAR", "Alarms/hr (250k: 800fps)", "Alarms/hr (500k: 2k fps)", "Alarms/hr (1M: 4k fps)", "Suppression"
    ]
    s3_rows = [
        ["Raw Model Output (No Filter)", "k = 1", "Whole-Bus", "22,136", "11.068%", "318,758 / hr", "778,029 / hr", "1,556,058 / hr", "0.00%"],
        ["Temporal Debounce Filter", "k = 2", "Whole-Bus", "1,530", "0.765%", "22,032 / hr", "53,776 / hr", "107,552 / hr", "93.09%"],
        ["Temporal Debounce Filter", "k = 3", "Whole-Bus", "53", "0.0265%", "763.2 / hr", "1,863.0 / hr", "3,726.0 / hr", "99.76%"],
        ["Temporal Debounce Filter", "k = 4", "Whole-Bus", "4", "0.0020%", "57.6 / hr", "140.6 / hr", "281.2 / hr", "99.98%"],
        ["Per-ID Debounce Filter", "k = 2", "Per-ID", "17,008", "8.504%", "244,915 / hr", "597,792 / hr", "1,195,584 / hr", "23.17%"],
        ["Per-ID Debounce Filter", "k = 3", "Per-ID", "14,514", "7.257%", "208,992 / hr", "510,134 / hr", "1,020,268 / hr", "34.43%"],
        ["Poisson Independent Model", "k = 3", "Bus-wide", "Analytic", "0.000000075%", "0.00217 (1 / 19.2 d)", "0.00543 (1 / 7.7 d)", "0.01085 (1 / 3.8 d)", ">99.999%"]
    ]

    t_s3 = doc.add_table(rows=1, cols=len(s3_headers))
    t_s3.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h_text in enumerate(s3_headers):
        cell = t_s3.cell(0, c_idx)
        cell.text = h_text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)

    for r_data in s3_rows:
        row_cells = t_s3.add_row().cells
        for c_idx, val in enumerate(r_data):
            cell = row_cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
    set_apa_table_borders(t_s3)

    p_s3_fn = doc.add_paragraph()
    r = p_s3_fn.add_run(
        "Note: Continuous evaluation across 200,000 nominal driving frames from the Hyundai Sonata CAN bus. Line-speed projections utilize nominal transmission "
        "rates of 800 frames/s at 250 kbps (2,880,000 frames/hr), 2,000 frames/s at 500 kbps (7,200,000 frames/hr), and 4,000 frames/s at 1 Mbps (14,400,000 frames/hr). "
        "Under an independent Poisson model at 0.091% baseline FPR, k = 3 yields one false alarm every 19.2 days at 250 kbps (0.00217 alarms/hr), 7.7 days at 500 kbps "
        "(0.00543 alarms/hr), and 3.8 days at 1 Mbps (0.01085 alarms/hr). Under empirical 11.068% cross-session driving on normal_run_data.txt, an i.i.d. model predicts "
        "271 false alarms; whole-bus k = 3 yields 53 alarms (5-fold suppression below i.i.d.), whereas per-ID k = 3 yields 14,514 alarms (53-fold higher than i.i.d.), "
        "empirically proving severe per-ID temporal error clustering."
    )
    r.italic = True
    apply_text_styling(p_s3_fn, font_size=9.0, line_spacing=1.3, space_before=6, space_after=18)

    doc.add_page_break()

    # =========================================================================
    # SUPPLEMENTARY TABLE S4
    # =========================================================================
    p_s4_tit = doc.add_paragraph()
    r = p_s4_tit.add_run("Supplementary Table S4. ")
    r.bold = True
    p_s4_tit.add_run("Mathematical Formulations and Protocol-Agnostic Feature Definitions for Calibration-Normalized Stateful CAN-IDS.")
    apply_text_styling(p_s4_tit, font_size=11, line_spacing=1.5, space_before=12, space_after=6)

    s4_headers = ["Feature Symbol", "Mathematical Definition", "Domain Range", "Calibration Baseline Input", "Physical / Protocol Rationale"]
    s4_rows = [
        [
            "in_allowlist_i",
            "I(ID_i in S_nom)",
            "{0, 1}",
            "Nominal identifier set S_nom",
            "Flags out-of-specification extended arbitration IDs or unauthorized diagnostic message injections."
        ],
        [
            "freq_rank_nom,i",
            "Rank(Freq_nom(ID_i)) / |S_nom| if in allowlist, else -1.0",
            "[0.0, 1.0] U {-1.0}",
            "Allowlist frequency ranking on nominal baseline",
            "Abstracts OEM-proprietary 11-bit / 29-bit CAN IDs into vehicle-agnostic relative transmission frequency percentiles."
        ],
        [
            "r_Δt, i",
            "Δt_i / μ_Δt,nom(ID_i)",
            "[0, ∞)",
            "Nominal mean inter-arrival period μ_Δt,nom per ID",
            "Quantifies bus flooding and message periodicity surges invariant to absolute baud rate or native broadcast cycles."
        ],
        [
            "z_μ, i",
            "(μ_byte,i - μ_nom(ID_i)) / σ_nom(ID_i)",
            "(-∞, ∞)",
            "Nominal payload byte mean μ_nom and std σ_nom per ID",
            "Captures sensor signal shifts and physical tampering (e.g. brake pressure override) without DBC decoding."
        ],
        [
            "z_σ, i",
            "(σ_byte,i - σ̄_nom(ID_i)) / s_σ,nom(ID_i)",
            "(-∞, ∞)",
            "Nominal payload volatility mean and std per ID",
            "Detects abnormal signal fuzzing, noise corruption, or unnatural payload freeze states."
        ],
        [
            "DLC_i",
            "DLC_i",
            "{0, 1, ..., 8}",
            "None (instantaneous frame evaluation)",
            "Validates frame payload byte capacity; detects malformed packet length injection anomalies."
        ]
    ]

    t_s4 = doc.add_table(rows=1, cols=len(s4_headers))
    t_s4.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h_text in enumerate(s4_headers):
        cell = t_s4.cell(0, c_idx)
        cell.text = h_text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        apply_text_styling(p, font_size=8.5, line_spacing=1.15, space_before=2, space_after=2)
        set_cell_margins(cell, top=60, bottom=60, left=70, right=70)

    for r_data in s4_rows:
        row_cells = t_s4.add_row().cells
        for c_idx, val in enumerate(r_data):
            cell = row_cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
            set_cell_margins(cell, top=40, bottom=40, left=70, right=70)
    set_apa_table_borders(t_s4)

    p_s4_fn = doc.add_paragraph()
    r = p_s4_fn.add_run(
        "Note: All 6 features are computed incrementally in O(1) constant time per incoming message using a static hash table indexed by arbitration ID. "
        "The hash table stores nominal calibration parameters (μ_Δt,nom, μ_nom, σ_nom, rank) derived from a clean 60-second baseline log. Memory overhead is "
        "bounded by the number of unique active CAN IDs (typically <= 128 per vehicle sub-bus), requiring under 4 KB of static RAM on embedded automotive microcontrollers."
    )
    r.italic = True
    apply_text_styling(p_s4_fn, font_size=9.0, line_spacing=1.3, space_before=6, space_after=18)

    doc.add_page_break()

    # =========================================================================
    # SUPPLEMENTARY TABLE S5
    # =========================================================================
    p_s5_tit = doc.add_paragraph()
    r = p_s5_tit.add_run("Supplementary Table S5. ")
    r.bold = True
    p_s5_tit.add_run("Analytical Complexity and Microcontroller Resource Footprint Estimates for Automotive ECUs (Modeled for ARM Cortex-M7 Core at 160 MHz).")
    apply_text_styling(p_s5_tit, font_size=11, line_spacing=1.5, space_before=12, space_after=6)

    s5_headers = [
        "Architecture", "Tree / Layer Complexity", "Node / Param Count",
        "Flash Footprint", "RAM Footprint", "Estimated Latency", "Bus Load @ 500k", "ASIL Compliance Feasibility"
    ]
    s5_rows = [
        ["Single-Feature Decision Tree", "1 tree / depth 10", "187 nodes", "1.50 KB", "< 64 bytes", "0.42 μs", "0.084%", "ASIL-D Feasible (Static branch)"],
        ["Compact Decision Forest", "20 trees / depth 10", "3,740 nodes", "29.92 KB", "< 512 bytes", "3.85 μs", "0.770%", "ASIL-D Feasible (Static ensemble)"],
        ["Full Deployment Random Forest", "100 trees / depth 12", "34,200 nodes", "273.60 KB", "< 3.84 KB", "18.40 μs", "3.680%", "ASIL-B / ASIL-D Feasible"],
        ["Deep MLP (3 Dense Layers)", "3x128 hidden units", "18,304 params", "73.22 KB", "4.20 KB", "142.60 μs", "28.520%", "ASIL-B Feasible (Requires FPU)"],
        ["1D Convolutional Neural Network", "4 conv + pool + FC", "42,560 params", "170.24 KB", "18.50 KB", "680.00 μs", "136.000%", "Infeasible (Bus Overload: >500 μs)"]
    ]

    t_s5 = doc.add_table(rows=1, cols=len(s5_headers))
    t_s5.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h_text in enumerate(s5_headers):
        cell = t_s5.cell(0, c_idx)
        cell.text = h_text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)

    for r_data in s5_rows:
        row_cells = t_s5.add_row().cells
        for c_idx, val in enumerate(r_data):
            cell = row_cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
    set_apa_table_borders(t_s5)

    p_s5_fn = doc.add_paragraph()
    r = p_s5_fn.add_run(
        "Note: Complexity metrics represent analytical instruction-cycle and memory models for a 32-bit ARM Cortex-M7 automotive core operating at 160 MHz "
        "(e.g., NXP S32K3 / Infineon AURIX TC3xx), not physical bench measurements. Flash storage assumes packed 8-byte C-struct decision nodes (uint32_t threshold, "
        "uint16_t left_child, uint16_t right_child). Cycle estimates assume 12 CPU cycles per decision node comparison and traversal. At 500 kbps, the nominal inter-frame "
        "period is 500 μs; the 18.40 μs estimated inference latency of the 100-tree forest consumes only 3.68% of the available ECU time budget, whereas the 1D-CNN exceeds "
        "the available inter-frame window (136% bus overload), confirming that tree ensembles are uniquely viable for real-time in-vehicle execution."
    )
    r.italic = True
    apply_text_styling(p_s5_fn, font_size=9.0, line_spacing=1.3, space_before=6, space_after=18)

    doc.add_page_break()

    # =========================================================================
    # SUPPLEMENTARY TABLE S6
    # =========================================================================
    p_s6_tit = doc.add_paragraph()
    r = p_s6_tit.add_run("Supplementary Table S6. ")
    r.bold = True
    p_s6_tit.add_run("Multi-Detector Benchmark Across Operational Stress Tests (HCRL, Polaris GEM e6, and ORNL ROAD Benchmarks).")
    apply_text_styling(p_s6_tit, font_size=11, line_spacing=1.5, space_before=12, space_after=6)

    s6_headers = [
        "Detector Paradigm", "Model Architecture", "Target Dataset / Evaluation Mode",
        "Accuracy", "MCC", "FPR (%)", "Hourly Alarms @ 500k", "Attack Recall"
    ]
    s6_rows = [
        ["Timing-Rule Baseline", "Threshold (dt < 1ms or dt > 3*mean)", "HCRL In-Distribution (70/30)", "85.03%", "0.7231", "2.592%", "186,600 / hr", "72.65%"],
        ["Random Forest", "100 trees, max depth 12", "HCRL In-Distribution (70/30)", "99.90%", "0.9990", "0.091%", "6,552 / hr", "99.90%"],
        ["LightGBM", "100 boosting rounds, lr=0.10", "HCRL In-Distribution (70/30)", "99.85%", "0.9982", "0.115%", "8,280 / hr", "99.87%"],
        ["Reconstruction Autoencoder", "PyTorch MLP (5-2-5, 99th %-ile MSE)", "HCRL In-Distribution (70/30)", "73.82%", "0.5582", "0.125%", "9,000 / hr", "47.77%"],
        ["Timing-Rule Baseline", "Stateless Periodicity Threshold", "GEM e6 Direct Transfer (143.4k frames)", "91.49%", "0.7970", "0.02%", "1,440 / hr", "96.42% (DoS only)"],
        ["Random Forest", "100 trees (Sonata-trained)", "GEM e6 Direct Transfer (143.4k frames)", "35.21%", "0.1637", "91.48%", "6,586,560 / hr", "98.71% (Collapse)"],
        ["LightGBM", "100 trees (Sonata-trained)", "GEM e6 Direct Transfer (143.4k frames)", "35.40%", "0.1670", "91.20%", "6,566,400 / hr", "98.60% (Collapse)"],
        ["Reconstruction Autoencoder", "MLP Autoencoder (Sonata-trained)", "GEM e6 Direct Transfer (143.4k frames)", "29.56%", "-0.0339", "100.00%", "7,200,000 / hr", "100.00%* (Total Collapse)"],
        ["Timing-Rule Baseline", "Stateless Periodicity Threshold", "ORNL ROAD Masquerade (74.2k frames)", "97.14%", "-0.0042", "0.056%", "4,032 / hr", "0.00% (Completely Blind)"],
        ["Random Forest", "100 trees (Multi-feature)", "ORNL ROAD Masquerade (74.2k frames)", "99.99%", "0.9993", "0.005%", "360 / hr", "100.00%"],
        ["LightGBM", "100 boosting rounds", "ORNL ROAD Masquerade (74.2k frames)", "99.99%", "0.9993", "0.005%", "360 / hr", "100.00%"]
    ]

    t_s6 = doc.add_table(rows=1, cols=len(s6_headers))
    t_s6.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h_text in enumerate(s6_headers):
        cell = t_s6.cell(0, c_idx)
        cell.text = h_text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)

    for r_data in s6_rows:
        row_cells = t_s6.add_row().cells
        for c_idx, val in enumerate(r_data):
            cell = row_cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
    set_apa_table_borders(t_s6)

    p_s6_fn = doc.add_paragraph()
    r = p_s6_fn.add_run(
        "Note: Multi-detector benchmark evaluating four algorithmic paradigms across the HCRL Car-Hacking dataset (stationary passenger car), "
        "GEM-CAN dataset (autonomous electric shuttle with extended 29-bit IDs), and ORNL ROAD dataset (real physical masquerade attacks preserving broadcast periodicity). "
        "Under raw cross-platform transfer on GEM-CAN, all payload-aware models experience representation collapse (91.20% to 100.00% FPR) because extended 29-bit identifiers "
        "fall outside 11-bit decision boundaries. Mislabeled attacks under raw transfer occur because models misclassify over 91% to 100% of all frames "
        "as attacks. On the ORNL ROAD dataset, timing-rule baselines collapse to 0.00% recall because the attacker injects masquerade frames at legitimate cyclic intervals "
        "(20 ms and 50 ms), whereas tree ensembles achieve 100.00% recall by detecting non-linear signal correlation deviations between engine torque and wheel velocity."
    )
    r.italic = True
    apply_text_styling(p_s6_fn, font_size=9.0, line_spacing=1.3, space_before=6, space_after=18)

    doc.add_page_break()

    # =========================================================================
    # SUPPLEMENTARY TABLE S7
    # =========================================================================
    p_s7_tit = doc.add_paragraph()
    r = p_s7_tit.add_run("Supplementary Table S7. ")
    r.bold = True
    p_s7_tit.add_run("Leakage-Free Camouflage Curriculum Retraining vs. Nominal False-Alarm Rate Trade-off Across Blending Levels.")
    apply_text_styling(p_s7_tit, font_size=11, line_spacing=1.5, space_before=12, space_after=6)

    s7_headers = [
        "Blending Level (alpha)", "Attack Modification Mode", "Standard RF Recall",
        "Curriculum RF Recall", "Recovery Gain", "Nominal Driving FPR", "Nominal Alarms/hr @ 500k", "Bayes Decision Boundary Status"
    ]
    s7_rows = [
        ["alpha = 0.00", "Raw Unmodified Injection", "99.90%", "99.92%", "+0.02%", "0.091% vs 0.246%", "6,552 vs 17,700 / hr", "Nominal Manifold Preserved"],
        ["alpha = 0.25", "Slight Payload/Timing Blend", "41.80%", "98.60%", "+56.80%", "—", "—", "Recovered by Sub-distribution Training"],
        ["alpha = 0.50", "Moderate Distribution Alignment", "33.10%", "95.40%", "+62.30%", "—", "—", "Recovered by Sub-distribution Training"],
        ["alpha = 0.75", "Heavy Distribution Camouflage", "31.40%", "48.20%", "+16.80%", "—", "—", "Unseen Test Generalization"],
        ["alpha = 1.00", "Complete Nominal Centroid Collapse", "19.60%", "21.80%", "+2.20%", "—", "—", "Asymptotically Bound by Bayes Error"]
    ]

    t_s7 = doc.add_table(rows=1, cols=len(s7_headers))
    t_s7.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h_text in enumerate(s7_headers):
        cell = t_s7.cell(0, c_idx)
        cell.text = h_text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)

    for r_data in s7_rows:
        row_cells = t_s7.add_row().cells
        for c_idx, val in enumerate(r_data):
            cell = row_cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
    set_apa_table_borders(t_s7)

    p_s7_fn = doc.add_paragraph()
    r = p_s7_fn.add_run(
        "Note: Leakage-free curriculum retraining enforces strict empirical safeguards: (1) training curriculum samples are limited strictly to partial blending levels "
        "alpha in {0.25, 0.50}, leaving alpha = 0.75 and 1.00 as completely unseen test distributions; (2) camouflaged vectors are synthesized exclusively by pairing malicious "
        "frames with nominal frames sharing identical CAN IDs; and (3) evaluation is conducted on held-out chronological test sets. Evasion recovery is bounded at alpha = 1.00 "
        "by Bayes error, where attack vectors collapse to exact nominal feature values. Expanding decision boundaries to intercept camouflaged vectors increases the nominal FPR "
        "from 0.091% (6,552 alarms/hr at 500 kbps) to 0.246% (17,700 alarms/hr), demonstrating an unavoidable trade-off between adversarial robustness and line-speed false alarm suppression."
    )
    r.italic = True
    apply_text_styling(p_s7_fn, font_size=9.0, line_spacing=1.3, space_before=6, space_after=18)

    doc.add_page_break()

    # =========================================================================
    # SUPPLEMENTARY TABLE S8
    # =========================================================================
    p_s8_tit = doc.add_paragraph()
    r = p_s8_tit.add_run("Supplementary Table S8. ")
    r.bold = True
    p_s8_tit.add_run("Supervised Few-Shot Domain Adaptation Performance on Polaris GEM e6 Telemetry.")
    apply_text_styling(p_s8_tit, font_size=11, line_spacing=1.5, space_before=12, space_after=6)

    s8_headers = [
        "Labeled Adaptation Frames (N)", "Vehicle Driving Time (s)", "Target Test Accuracy (30% Partition)",
        "Target Test MCC", "False Positive Rate (FPR)", "DoS Recall (n=12,342)", "Tampering Recall (n=397)", "Operational Domain Status"
    ]
    s8_rows = [
        ["N = 0 (Zero-Shot)", "0.0 s", "92.04%", "0.8316", "10.65%", "100.00%", "50.04%", "Calibration-Normalized Cross-Platform"],
        ["N = 100", "0.2 s", "99.48%", "0.9877", "0.73%", "100.00%", "86.40%", "14-fold FPR Reduction & 86.4% Tamper Rec"],
        ["N = 500", "1.0 s", "99.96%", "0.9991", "0.04%", "100.00%", "95.72%", "Near-Zero FPR & 95.7% Tamper Rec"],
        ["N = 1,000", "2.0 s", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%", "Full Zero-FPR & 100% Tamper Convergence"],
        ["N = 2,500", "5.0 s", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%", "Asymptotic Domain Convergence"],
        ["N = 5,000", "10.0 s", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%", "Asymptotic Domain Convergence"],
        ["N = 10,000", "20.0 s", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%", "Asymptotic Domain Convergence"]
    ]

    t_s8 = doc.add_table(rows=1, cols=len(s8_headers))
    t_s8.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h_text in enumerate(s8_headers):
        cell = t_s8.cell(0, c_idx)
        cell.text = h_text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)

    for r_data in s8_rows:
        row_cells = t_s8.add_row().cells
        for c_idx, val in enumerate(r_data):
            cell = row_cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            apply_text_styling(p, font_size=8.0, line_spacing=1.15, space_before=2, space_after=2)
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
    set_apa_table_borders(t_s8)

    p_s8_fn = doc.add_paragraph()
    r = p_s8_fn.add_run(
        "Note: Supervised few-shot adaptation curve on the Polaris GEM e6 autonomous shuttle telemetry across N in {0, 100, 500, 1000, 2500, 5000, 10000} frames "
        "sampled across operational regimes within the 70% target training partition, evaluated on the held-out 30% time-blocked test partition (30,295 Normal and 12,739 Attack frames: "
        "12,342 DoS + 397 Tampering). Vehicle driving time assumes a nominal bus rate of 500 fps. With just N = 100 labeled frames (0.2 s of telemetry), target domain FPR drops from "
        "10.65% to 0.73% while tampering recall reaches 86.40%. With N >= 500 frames (1.0 s of driving), FPR drops to 0.04% and tampering recall reaches 95.72%, achieving 0.00% FPR "
        "and 100.00% tampering recall at N >= 1,000 frames."
    )
    r.italic = True
    apply_text_styling(p_s8_fn, font_size=9.0, line_spacing=1.3, space_before=6, space_after=18)

    doc.save(output_path)
    # Also save copy in 1_manuscripts_and_drafts if it exists
    alt_path = Path(output_path).parent / "1_manuscripts_and_drafts" / Path(output_path).name
    if alt_path.parent.is_dir():
        doc.save(str(alt_path))
    print(f"Successfully generated Supplementary Tables: {output_path}")

if __name__ == '__main__':
    base_dir = Path(__file__).resolve().parent
    st_path = base_dir / "Supplementary_Tables.docx"
    create_supplementary_tables(st_path)
