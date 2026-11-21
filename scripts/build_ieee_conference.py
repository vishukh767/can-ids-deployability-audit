#!/usr/bin/env python3
"""
compile_and_verify_ieee.py
Updates build_ieee_conference.py, compiles IEEE_Conference_Draft_v1.docx,
exports to PDF via Word COM, and reports exact physical page count.
"""

from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import win32com.client
import pypdf

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name in ["scripts", "2_code_and_scripts"] else SCRIPT_DIR
DRAFTS_DIR = PROJECT_ROOT / "drafts" if (PROJECT_ROOT / "drafts").is_dir() else PROJECT_ROOT / "1_manuscripts_and_drafts"
RESULTS_DIR = PROJECT_ROOT / "images" if (PROJECT_ROOT / "images").is_dir() else PROJECT_ROOT / "3_results_and_figures"
OUTPUT_DOCX = DRAFTS_DIR / "IEEE_Conference_Draft_v1.docx"
OUTPUT_PDF = DRAFTS_DIR / "IEEE_Conference_Draft_v1.pdf"

def set_run(r, sz=10, rgb=(0, 0, 0), bold=False, italic=False, font_name="Times New Roman", small_caps=False):
    r.font.name = font_name
    r.font.size = Pt(sz)
    r.font.color.rgb = RGBColor(*rgb)
    r.bold = bold
    r.italic = italic
    if small_caps:
        r.font.small_caps = True

def add_p(d, txt, before=0, after=2.5, line_spacing=1.0, indent=0.15, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = d.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line_spacing
    if indent > 0:
        pf.first_line_indent = Inches(indent)
    r = p.add_run(txt)
    set_run(r, 10.0, (0, 0, 0), bold=False, italic=False)
    return p

def add_sec_h(d, num_str, title_str):
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(9.5)
    pf.space_after = Pt(3)
    pf.keep_with_next = True
    txt = f"{num_str}.  {title_str.upper()}" if num_str else title_str.upper()
    r = p.add_run(txt)
    set_run(r, 10.0, (0, 0, 0), bold=True, small_caps=True)
    return p

def add_subsec_h(d, letter_str, title_str):
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(6.5)
    pf.space_after = Pt(2)
    pf.keep_with_next = True
    r = p.add_run(f"{letter_str}.  {title_str}")
    set_run(r, 10.0, (0, 0, 0), bold=False, italic=True)
    return p

def add_formula(d, formula_txt, num_str):
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(3.5)
    pf.space_after = Pt(3.5)
    pf.line_spacing = 1.0
    
    r1 = p.add_run(formula_txt)
    set_run(r1, 9.5, (0, 0, 0), italic=True)
    
    r_sp = p.add_run(f"    ({num_str})")
    set_run(r_sp, 9.5, (0, 0, 0), italic=False)
    return p

def set_cell_margins(cell, top=35, bottom=35, left=45, right=45):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_tbl(d, headers, rows, col_widths, cap_num, cap_title):
    cp = d.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cpf = cp.paragraph_format
    cpf.space_before = Pt(7.5)
    cpf.space_after = Pt(1.5)
    cpf.keep_with_next = True
    
    r_cap = cp.add_run(f"TABLE {cap_num}:  ")
    set_run(r_cap, 8.0, (0, 0, 0), bold=True, small_caps=True)
    r_tit = cp.add_run(cap_title.upper())
    set_run(r_tit, 7.5, (0, 0, 0), bold=False, small_caps=True)
    
    tbl = d.add_table(rows=len(rows) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    tblPr = tbl._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>\n'
        f'  <w:insideV w:val="none"/>\n'
        f'  <w:left w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)
    
    hdr_row = tbl.rows[0]
    hdr_trPr = hdr_row._tr.get_or_add_trPr()
    hdr_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    for c_idx, h_text in enumerate(headers):
        cell = hdr_row.cells[c_idx]
        cell.width = Inches(col_widths[c_idx])
        set_cell_margins(cell, top=35, bottom=35, left=35, right=35)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(h_text)
        set_run(r, 7.5, (0, 0, 0), bold=True)
        
    for r_idx, row_data in enumerate(rows):
        row = tbl.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Inches(col_widths[c_idx])
            set_cell_margins(cell, top=25, bottom=25, left=30, right=30)
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 or len(val) > 18 else WD_ALIGN_PARAGRAPH.CENTER
            p.alignment = align
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(val)
            set_run(r, 7.5, (0, 0, 0), bold=False)
            
    sp_p = d.add_paragraph()
    sp_p.paragraph_format.space_before = Pt(0)
    sp_p.paragraph_format.space_after = Pt(3.5)
    sp_p.paragraph_format.line_spacing = 1.0
    return tbl

def add_fig(d, fig_path, fig_num, cap_text, width_in=3.35):
    p_img = d.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(5)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    
    cand = Path(fig_path)
    if not cand.exists():
        cand = RESULTS_DIR / cand.name
    if cand.exists():
        r_img = p_img.add_run()
        r_img.add_picture(str(cand), width=Inches(width_in))
    else:
        r_img = p_img.add_run(f"[Figure {fig_num}: {cand.name}]")
        set_run(r_img, 8.0, (0x7F, 0, 0), bold=True)
        
    p_cap = d.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(1)
    p_cap.paragraph_format.space_after = Pt(5)
    
    r_fn = p_cap.add_run(f"Fig. {fig_num}.  ")
    set_run(r_fn, 8.0, (0, 0, 0), bold=True)
    r_ct = p_cap.add_run(cap_text)
    set_run(r_ct, 8.0, (0, 0, 0), bold=False)
    return p_img

def generate():
    doc = docx.Document()
    
    # Section 1: Title Block (Single Column)
    s1 = doc.sections[0]
    s1.page_width = Inches(8.5)
    s1.page_height = Inches(11.0)
    s1.top_margin = Inches(0.75)
    s1.bottom_margin = Inches(0.75)
    s1.left_margin = Inches(0.625)
    s1.right_margin = Inches(0.625)
    
    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(14)
    r_t = p_title.add_run("The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit and Cross-Platform Domain Analysis")
    set_run(r_t, 22.0, (0, 0, 0), bold=False)
    
    # Author block completely omitted for double-blind peer review


    # Abstract
    p_abs = doc.add_paragraph()
    p_abs.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_abs.paragraph_format.space_before = Pt(0)
    p_abs.paragraph_format.space_after = Pt(4)
    p_abs.paragraph_format.line_spacing = 1.0
    p_abs.paragraph_format.left_indent = Inches(0.20)
    p_abs.paragraph_format.right_indent = Inches(0.20)
    
    r_absh = p_abs.add_run("Abstract—")
    set_run(r_absh, 9.0, (0, 0, 0), bold=True, italic=True)
    r_abst = p_abs.add_run(
        "In-vehicle intrusion detection systems (CAN-IDS) report classification accuracies exceeding 99.9% on automotive benchmarks. "
        "This performance ceiling collapses in production. "
        "Evaluating Random Forest ensembles on chronological 70/30 time-blocked splits of the HCRL dataset (Hyundai YF Sonata), we reproduce baseline metrics (99.90% accuracy, 0.999 MCC, 0.091% FPR) and isolate operational failure modes. "
        "An arbitration ID single-feature tree achieves 82.40% accuracy; it isolates 76.64% of achievable gain above chance (25.00%). "
        "Translating the 0.091% benchmark FPR to physical bus line speeds produces 2,621 to 13,104 unmitigated false alarms per hour across standard automotive transceivers operating between 250 kbps and 1 Mbps. "
        "Whole-bus debouncing (k = 3) theoretically suppresses Poisson errors to one alarm per 7.7 days at 500 kbps. "
        "Dynamic driving maneuvers break this independence. "
        "Auditing 200,000 continuous normal frames reveals 1,863 alarms/hr under whole-bus filtering and 510,134 alarms/hr per-ID. "
        "Adversarial feature blending reduces attack recall to 19.60%; curriculum retraining recovers recall to 98.60% (alpha = 0.25) but increases nominal FPR to 0.246% (17,700 alarms/hr). "
        "Transfer to a Polaris GEM e6 shuttle incurs 91.48% FPR from 11-bit versus 29-bit SAE J1939 mismatches. "
        "Calibration-normalized stateful encoding restores zero-shot MCC to 0.8316. "
        "Stratified few-shot adaptation achieves 0.04% FPR at N = 500 frames (1.0 s) and converges to 0.00% FPR with 100.00% tampering recall at N >= 1,000 frames (2.0 s). "
        "On ORNL ROAD physical masquerade attacks, timing baselines record 0.00% recall, whereas payload tree ensembles maintain 100.00% recall. "
        "Vehicle-layer detectors must not command physical actuators without multi-sensor kinematic consensus under ISO 26262 and ISO/SAE 21434."
    )
    set_run(r_abst, 9.0, (0, 0, 0), bold=False)
    
    # Index Terms
    p_idx = doc.add_paragraph()
    p_idx.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_idx.paragraph_format.space_before = Pt(2)
    p_idx.paragraph_format.space_after = Pt(8)
    p_idx.paragraph_format.line_spacing = 1.0
    p_idx.paragraph_format.left_indent = Inches(0.20)
    p_idx.paragraph_format.right_indent = Inches(0.20)
    
    r_idxh = p_idx.add_run("Index Terms—")
    set_run(r_idxh, 9.0, (0, 0, 0), bold=True, italic=True)
    r_idxt = p_idx.add_run("Controller Area Network (CAN), Automotive Cybersecurity, Intrusion Detection Systems (CAN-IDS), False Alarm Rate Translation, Cross-Platform Generalization, Temporal Debouncing, ISO 26262, ISO/SAE 21434.")
    set_run(r_idxt, 9.0, (0, 0, 0), bold=False)
    
    # Section 2: Two-Column Section
    s2 = doc.add_section(WD_SECTION_START.CONTINUOUS)
    s2.top_margin = Inches(0.75)
    s2.bottom_margin = Inches(0.75)
    s2.left_margin = Inches(0.625)
    s2.right_margin = Inches(0.625)
    
    sect2_pr = s2._sectPr
    cols = sect2_pr.xpath('./w:cols')
    if cols:
        cols[0].set(qn('w:num'), '2')
        cols[0].set(qn('w:space'), '360')
        cols[0].set(qn('w:equalWidth'), '1')
    else:
        cols_el = OxmlElement('w:cols')
        cols_el.set(qn('w:num'), '2')
        cols_el.set(qn('w:space'), '360')
        cols_el.set(qn('w:equalWidth'), '1')
        sect2_pr.append(cols_el)
        
    # --- I. INTRODUCTION ---
    add_sec_h(doc, "I", "Introduction")
    add_p(doc,
        "Machine learning models for in-vehicle intrusion detection report classification accuracies exceeding 99.9% across public benchmarks [1]. "
        "These figures generate dangerous confidence. "
        "Supervised classifiers achieve near-perfect evaluation scores within static, identically distributed datasets. "
        "Yet when these architectures encounter live vehicle bus operating conditions, their practical utility deteriorates rapidly. "
        "Standard automotive evaluation protocols fail to reflect physical transceiver constraints, electronic control unit (ECU) scheduling dynamics, and cross-platform protocol variations."
    )
    add_p(doc,
        "Broadcast arbitration in Controller Area Network (CAN) architectures operates without sender authentication or cryptographic message validation. "
        "Priority arbitration relies entirely on message identifiers. "
        "Dominant bits override recessive bits during bus contention, allowing compromised ECUs to manipulate safety-critical functions. "
        "Modern intrusion detection research deploys deep neural networks and tree ensembles to monitor bus traffic. "
        "However, reported benchmark metrics obscure fatal deployment liabilities, including line-speed false alarm multiplication, identifier memorization shortcuts, and cross-vehicle transfer failure."
    )
    add_p(doc,
        "Physical transceiver line speeds transform small statistical error rates into continuous alarm floods. "
        "An empirical false-positive rate of 0.091% appears negligible in offline research tables. "
        "Line speeds amplify statistical errors. "
        "Yet on a 500 kbps automotive drivetrain bus transmitting 2,000 frames per second, that exact error rate generates over 6,550 unmitigated false alarms every operating hour. "
        "Alert fatigue rapidly blinds human operators and automated safety supervisors. "
        "Evaluating algorithmic metrics in isolation from physical bus clock rates creates an incomplete appraisal of deployability."
    )
    add_p(doc,
        "Automotive control systems demand bounded execution latency and deterministic worst-case timing. "
        "Deep neural networks incur prohibitive instruction overheads on resource-constrained embedded microcontrollers lacking dedicated neural accelerators. "
        "Resource limits constrain onboard deployment. "
        "When security monitoring introduces contention on internal buses or starves vehicle control loops, the defense mechanism degrades vehicle safety. "
        "A deployable detector must satisfy physical memory footprints and execute within microsecond message arrival deadlines."
    )
    add_p(doc,
        "This deployability audit evaluates tree ensembles and deep autoencoders across 200,000 real highway driving frames, dual physical vehicle platforms (Hyundai Sonata and Polaris GEM e6 autonomous shuttle), and physical masquerade injections from the ORNL ROAD benchmark. "
        "Physical bus dynamics break common assumptions. "
        "We quantify operational failure modes across three axes: translation of statistical false-positive rates into line-speed alarm frequencies, adversarial feature-space camouflage bounds, and architectural cross-platform representation breakdown. "
        "Our audit delivers actionable engineering constraints for vehicle-layer sentinels."
    )

    # --- II. THREAT LANDSCAPE ---
    add_sec_h(doc, "II", "Threat Landscape and Automotive Attack Surface")
    add_subsec_h(doc, "A", "Connected Vehicle Attack Interfaces")
    add_p(doc,
        "Modern connected passenger vehicles expose extensive remote and physical attack surfaces that bridge external wireless channels to internal vehicle buses. "
        "Vulnerabilities span cellular telematics units, Bluetooth transceivers, Wi-Fi infotainment systems, and mandatory physical On-Board Diagnostics (OBD-II) ports. "
        "Physical exploits compromise passenger safety. "
        "Remote takeover demonstrations by Miller and Valasek [4] proved that cellular gateway compromise enables unauthorized firmware re-flashing and arbitrary message injection into internal drivetrain buses. "
        "Similarly, Koscher et al. [3] demonstrated that internal bus access grants complete control over throttle actuators and electronic braking systems regardless of driver inputs."
    )
    add_subsec_h(doc, "B", "CAN Bus Vulnerabilities and Exploitation Mechanisms")
    add_p(doc,
        "The Controller Area Network operates as a shared multi-master differential serial bus governed by lossless bitwise arbitration. "
        "Arbitration identifiers serve two distinct purposes: message prioritization and semantic routing. "
        "Zero-value dominant bits overwrite logic-one recessive bits. "
        "An adversary injecting frames with dominant identifier 0x000 causes total bus denial-of-service (DoS) by starving lower-priority functional messages. "
        "In Fuzzy attacks, randomized identifiers and payloads disrupt transceiver state machines. "
        "In semantic masquerade attacks, adversaries inject target payload manipulations while preserving valid arbitration identifiers and broadcast periodicities [5]."
    )
    add_p(doc,
        "The fundamental vulnerability of CAN stems from its unauthenticated broadcast medium. "
        "Every connected ECU receives every transmitted frame. "
        "Message transceivers lack hardware provisions to identify which physical chip originated a given transmission. "
        "Because physical chips lack hardware source attribution, an adversary who gains read-write access to a low-criticality telematics node can spoof high-criticality messages from the transmission control module or steering column without physical bus rejection."
    )
    add_p(doc,
        "Bit stuffing and transceiver synchronization further complicate anomaly detection on physical buses. "
        "Automotive transceivers insert complementary stuffing bits after five consecutive identical bit states. "
        "Noise pulses alter stuffing sequences. "
        "These physical signaling constraints bound the speed at which microcontrollers can parse raw bitstreams without specialized peripheral coprocessors. "
        "Intrusion detection models must process messages at line speeds without dropping frames during bursty traffic intervals."
    )
    add_p(doc,
        "Adversaries exploit physical bus arbitration characteristics to execute covert injection attacks. "
        "By synchronizing transmissions with legitimate broadcast intervals, attackers bypass simple threshold counters. "
        "Modern attacks mimic legitimate transmission timing. "
        "Detecting these sophisticated intrusions requires models capable of inspecting subtle inter-signal correlations across consecutive payload bytes rather than relying on crude volume spikes."
    )

    # --- III. RELATED WORK ---
    add_sec_h(doc, "III", "Related Work and Benchmark Positioning")
    add_subsec_h(doc, "A", "In-Vehicle Intrusion Detection Paradigms")
    add_p(doc,
        "In-vehicle intrusion detection research splits into timing-interval monitoring, deep payload inspection, and distributed architectures. "
        "Timing systems detect volumetric injections. "
        "Song et al. [2] developed inter-arrival interval analysis to flag high-frequency message bursts. "
        "While timing rules reliably detect volumetric flooding, they remain entirely blind to masquerade attacks that respect nominal broadcast schedules. "
        "Supervised machine learning detectors analyze multi-byte payload structures to identify subtle signal manipulations. "
        "However, deep learning models demand substantial computational memory, exceeding the strict hardware limitations of automotive microcontrollers lacking floating-point units."
    )
    add_p(doc,
        "Distributed and federated architectures attempt to address scale across modern multi-bus topologies. "
        "Althunayyan et al. [10] proposed hierarchical federated learning across vehicle sub-networks. "
        "Fleet architectures face serious synchronization delays. "
        "While distributed models distribute computational burden, they introduce communication synchronization latency and require stable transmission channels. "
        "Vehicle-layer sentinels must execute strictly within real-time microcontroller memory boundaries."
    )
    add_subsec_h(doc, "B", "Critical Positioning Against Automotive Benchmarks")
    add_p(doc,
        "Prior automotive intrusion detection benchmarks suffer from methodological shortcuts that artificially inflate reported detection performance. "
        "Stationary logging introduces severe artifacts. "
        "The widely cited HCRL Car-Hacking dataset [1] was captured from a stationary vehicle with the engine idling, suppressing dynamic road telemetry. "
        "Random k-fold partitioning splits consecutive periodic frames across training and test folds, which induces extreme data leakage. "
        "Classifiers easily memorize sequence timestamps. "
        "Recent benchmarks eliminate these shortcuts by capturing physical vehicle operation: the ORNL ROAD dataset [5] incorporates verified physical masquerade attacks, while the GEM-CAN benchmark [6] captures 29-bit extended SAE J1939 telemetry on an autonomous campus shuttle."
    )
    add_p(doc,
        "Evaluation methodology dictates algorithmic validity in safety-critical cyber-physical systems. "
        "Artificially clean datasets mask operational vulnerabilities. "
        "When models train on stationary idling traffic, they learn static baselines that fail under real highway driving dynamics. "
        "This empirical gap motivates our multi-benchmark deployability audit across diverse vehicle architectures."
    )
    
    # TABLE I
    t1_h = ["Benchmark / Study", "Platform", "Split Strategy", "Key Vulnerability Isolated"]
    t1_r = [
        ["Song et al. [2]", "Stationary bench", "Random K-fold", "Fails on masquerade injections"],
        ["Althunayyan et al. [10]", "Simulated fleet", "Random K-fold", "Unvalidated cross-bus drift"],
        ["Verma et al. (ROAD) [5]", "Physical car", "Chronological", "Stateless timing blind to masquerade"],
        ["Tavasoli et al. (GEM) [6]", "Autonomous shuttle", "Chronological", "Extended 29-bit SAE J1939 shifts"],
        ["Ours (Deployability Audit)", "Sonata + GEM e6", "70/30 Time-Blocked", "Line-speed alarms & transfer collapse"]
    ]
    add_tbl(doc, t1_h, t1_r, [1.0, 0.75, 0.8, 0.95], "I", "Comparative Positioning Against In-Vehicle Benchmarks")

    # --- IV. METHODOLOGY ---
    add_sec_h(doc, "IV", "Mathematical Formulation and Methodology")
    add_subsec_h(doc, "A", "CAN Bus Message Representation")
    add_p(doc,
        "A continuous CAN communication stream forms an ordered sequence of frames S = {f_1, f_2, ..., f_N}. "
        "Message frames obey strict physical structures. "
        "Each individual frame is defined as:"
    )
    add_formula(doc, "f_i = (t_i, ID_i, DLC_i, B_i,0, ..., B_i,DLC-1, y_i)", "1")
    add_p(doc,
        "where t_i in R^+ denotes the microsecond arrival timestamp, ID_i in N_0 denotes the arbitration identifier, DLC_i in {0,...,8} indicates payload byte length, B_i,j represents payload bytes, and y_i in {0,1,2,3} represents the ground truth class. "
        "Embedded processors require lightweight statistical representations. "
        "For each frame f_i, we compute the inter-arrival timing interval Delta t_i relative to the previous frame sharing the identical arbitration identifier:"
    )
    add_formula(doc, "Delta t_i = t_i - max { t_k | k < i, ID_k = ID_i }", "2")
    add_p(doc,
        "Order of operations matters here. "
        "We calculate Delta t_i over the unbroken raw packet capture before subsampling. "
        "Payload central tendency and dispersion are captured via byte mean mu_byte,i and byte standard deviation sigma_byte,i across available bytes:"
    )
    add_formula(doc, "mu_byte,i = (1 / DLC_i) * sum_{j=0}^{DLC_i - 1} B_i,j", "3")
    add_formula(doc, "sigma_byte,i = sqrt( (1 / DLC_i) * sum_{j=0}^{DLC_i - 1} (B_i,j - mu_byte,i)^2 )", "4")
    
    add_subsec_h(doc, "B", "70/30 Time-Blocked Partitioning and Signal Isolation")
    add_p(doc,
        "Periodic broadcast scheduling governs normal CAN bus traffic, with functional messages repeating at fixed 10 ms to 100 ms intervals. "
        "Random train-test splits leak identifier phase. "
        "When packets from the same periodic stream are randomly distributed between training and test sets, classifiers memorize sequence timestamps. "
        "To eliminate data leakage, we enforce a strict 70/30 chronological time-blocked partition across all datasets. "
        "The initial 70% of continuous driving telemetry trains the models, while the final 30% evaluates generalization."
    )
    add_p(doc,
        "Balanced statistical metrics ensure robust evaluation. "
        "Model discrimination is assessed using classification accuracy, precision, recall, and the Matthews Correlation Coefficient (MCC):"
    )
    add_formula(doc, "MCC = (TP * TN - FP * FN) / sqrt((TP+FP)(TP+FN)(TN+FP)(TN+FN))", "5")
    add_p(doc,
        "MCC bounds performance between -1 and +1. "
        "It provides a strictly calibrated metric across severe class imbalances where raw accuracy produces misleading optimism. "
        "We quantify performance stability across five random initializations, reporting mean metrics alongside Wilson score 95% confidence intervals."
    )
    add_p(doc,
        "To quantify the degree to which classifiers rely on identifier memorization rather than payload inspection, we isolate a single-feature baseline decision tree trained strictly on raw arbitration identifiers. "
        "We calculate the normalized gain above chance G_chance = (Acc_ID - 0.25) / (Acc_full - 0.25). "
        "Single-feature baselines isolate memorization shortcuts. "
        "This metric measures the proportion of model accuracy attributable purely to message header memorization."
    )
    add_subsec_h(doc, "C", "Embedded Complexity Model")
    add_p(doc,
        "We target an ARM Cortex-M7 automotive microcontroller (480 MHz, 2 MB Flash, 1 MB SRAM) representative of modern central gateway controllers. "
        "Tree ensembles fit embedded hardware constraints. "
        "Our 100-tree Random Forest architecture averages 342 decision nodes per tree (~34,200 nodes total). "
        "At 8 bytes per decision node, the complete forest consumes ~274 KB of non-volatile flash memory and requires <4 KB of stack RAM. "
        "Analytical instruction-cycle modeling yields an inference latency of 68.4 microseconds per frame, completing well within the 250 microsecond arrival window of 1 Mbps buses (Table S5)."
    )

    # --- V. IN-DISTRIBUTION BENCHMARK ---
    add_sec_h(doc, "V", "Empirical Deployability Audit: In-Distribution Benchmark")
    add_subsec_h(doc, "A", "Baseline In-Distribution Performance")
    add_p(doc,
        "Evaluating the Random Forest ensemble on the chronological 70/30 split of the HCRL dataset reproduces nominal benchmark performance: 99.90% accuracy, 0.999 MCC, and a 0.091% false-positive rate (Table II, Table S1). "
        "Across 24,000 held-out test frames, the model attains 100.0% recall on DoS and Spoofing injections and 99.60% recall on Fuzzy manipulation. "
        "Benchmark metrics appear exceptionally strong. "
        "Wilson score 95% confidence intervals establish tight error bounds: [0.034%, 0.198%] on n=6,000 nominal test frames, and [0.062%, 0.134%] across all N=30,000 nominal frames. "
        "Multi-class discrimination is illustrated in Fig. 1."
    )
    
    # FIG 1
    fig1_path = RESULTS_DIR / "figure3_hcrl_confusion_matrix.png"
    add_fig(doc, fig1_path, "1", "Multi-class confusion matrix of Random Forest on chronological HCRL split (99.90% accuracy, 0.091% FPR).", 3.2)

    add_subsec_h(doc, "B", "Single-Feature CAN ID Memorization")
    add_p(doc,
        "To determine whether classifiers learn true behavioral anomalies or memorize static message identifiers, we isolate a single-feature decision tree trained strictly on raw arbitration identifiers. "
        "The identifier tree achieves 82.40% accuracy. "
        "Chance accuracy across the four balanced classes is 25.00%. "
        "The single-feature baseline captures 76.64% of the achievable accuracy gain above chance ((82.40 - 25.0) / (99.90 - 25.0) = 0.7664). "
        "Per-class recall analysis exposes the underlying shortcut: the ID-only tree attains 100.0% recall on DoS and 99.6% on Fuzzy, but drops to 31.2% recall on Spoofing. "
        "Because spoofing attacks inject legitimate identifiers with falsified payloads, ID-only trees fail where behavioral inspection is required."
    )
    add_subsec_h(doc, "C", "Timing Feature Redundancy and Jitter Brittleness")
    add_p(doc,
        "Ablating inter-arrival timing Delta t_i reduces in-distribution ensemble accuracy by only 0.05% (from 99.90% to 99.85%). "
        "Timing features prove largely redundant for separating high-volume injections from benign idling frames in stationary captures. "
        "Deep decision trees place tight numerical split thresholds on Delta t_i and payload variance to bound normal leaves. "
        "These narrow thresholds prove brittle. "
        "When normal driving exhibits natural scheduling jitter or variable payload telemetry, test frames exit these tight boundaries and trigger false alarms."
    )
    add_subsec_h(doc, "D", "Multi-Detector Architectural Benchmark")
    add_p(doc,
        "We benchmark four distinct detector paradigms: stateless timing rules, Random Forests, LightGBM, and deep reconstruction autoencoders (Table II). "
        "In-distribution performance remains uniformly high across all algorithms (97.42% to 99.90% accuracy). "
        "All models perform well within their native distribution. "
        "However, direct cross-platform transfer onto unseen Polaris GEM e6 shuttle telemetry triggers immediate false-positive breakdowns across all payload-sensitive models (91.20% to 100.00% FPR). "
        "Representation breakdown reflects architectural protocol shifts rather than decision tree bias."
    )
    
    # TABLE II
    t2_h = ["Detector Architecture", "Feature Representation", "Accuracy", "MCC", "In-Dist FPR", "Transfer FPR"]
    t2_r = [
        ["Zero-Rule Baseline", "Majority Label", "25.00%", "0.0000", "0.000%", "0.00%"],
        ["Timing-Rule Baseline", "Inter-Arrival Period Delta t", "97.42%", "0.9658", "0.098%", "74.60%"],
        ["CAN ID Decision Tree", "Arbitration ID Only", "82.40%", "0.7712", "0.125%", "96.42%"],
        ["LightGBM Booster", "Full 5-D Vector", "99.88%", "0.9984", "0.095%", "91.20%"],
        ["Reconstruction AE", "Latent Error ||x - x_hat||", "98.45%", "0.9792", "0.142%", "100.00%"],
        ["Random Forest (Ours)", "Full 5-D Vector (70/30)", "99.90%", "0.9987", "0.091%", "91.48%"]
    ]
    add_tbl(doc, t2_h, t2_r, [1.05, 0.95, 0.45, 0.45, 0.45, 0.45], "II", "In-Distribution Benchmark and Multi-Detector Architectural Stress Test")

    # --- VI. LINE-SPEED PROJECTIONS ---
    add_sec_h(doc, "VI", "Operational Line-Speed Alarm Projections and Debouncing")
    add_subsec_h(doc, "A", "Translating Benchmark Rates to Physical Bus Frequencies")
    add_p(doc,
        "Offline academic benchmarks present error rates as static percentages. "
        "On operational automotive networks, error rates compound continuously with transceiver frame frequencies. "
        "Standard automotive CAN transceivers operate between 250 kbps (low-speed body bus, ~800 frames/s) and 1.0 Mbps (high-speed CAN FD / powertrain bus, ~4,000 frames/s). "
        "The operational false alarm frequency R_alarm scales directly with bus frame rate R_bus:"
    )
    add_formula(doc, "R_alarm = FPR * R_bus * 3600", "6")
    add_p(doc,
        "Translating our baseline 0.091% FPR generates 2,621 false alarms per hour at 250 kbps, 6,552 false alarms per hour at 500 kbps, and 13,104 false alarms per hour at 1.0 Mbps (Table III). "
        "Line speeds amplify statistical errors. "
        "A system triggering thousands of false alarms per hour saturates logging buffers and forces driver alert fatigue. "
        "Physical buses process millions of frames per shift."
    )
    add_subsec_h(doc, "B", "Temporal Debouncing vs. Dynamic Error Burstiness")
    add_p(doc,
        "Automotive engineers implement k-consecutive temporal debounce filters to suppress transient frame-level noise. "
        "Under independent Poisson arrival assumptions, debouncing requires k consecutive malicious classifications before raising a system alert. "
        "Theoretical debouncing models promise near-total suppression. "
        "At k = 3 and 500 kbps, theoretical Poisson false alarms fall to one event every 7.7 days (0.0054 alarms/hr). "
        "However, this formulation assumes statistical independence between adjacent frames. "
        "Real vehicle telemetry violates independence assumptions."
    )
    add_p(doc,
        "To quantify debouncing under real vehicle dynamics, we evaluated a continuous 200,000-frame highway driving log (normal_run_data.txt) from the HCRL dataset. "
        "Real driving produced 22,136 raw false positives (11.068% empirical FPR). "
        "Under whole-bus k = 3 filtering, clustered error bursts produced 1,863 false alarms per hour (Table III). "
        "Evaluating independent per-identifier debounce queues generated 510,134 false alarms per hour. "
        "Correlated bus telemetry shatters theoretical suppression models."
    )
    add_p(doc,
        "Physical error clustering stems from transient vehicle maneuvers and powertrain dynamics. "
        "Rapid acceleration shifts gear positions and throttle percentages simultaneously across multiple periodic messages. "
        "Engine transitions generate coordinated payload shifts. "
        "When these dynamic transitions occur, the classifier encounters multiple consecutive out-of-distribution frames, easily satisfying the k = 3 threshold and defeating the temporal filter."
    )
    add_p(doc,
        "Transceiver clock drift and bus load fluctuations contribute to temporal error correlation. "
        "When physical bus utilization climbs during heavy driving maneuvers, message arbitration introduces variable queuing latency. "
        "Timing jitter shifts inter-arrival intervals outside trained nominal bounds. "
        "Filtering architectures that assume independent random noise collapse when confronted with systematic vehicular state transitions."
    )
    
    # TABLE III
    t3_h = ["CAN Bus Speed", "Frame Rate", "Raw Alarms/Hr", "Poisson k=3", "Highway Driving (k=3)"]
    t3_r = [
        ["250 kbps (Body)", "800 fps", "2,621 / hr", "1 per 19.2 days", "1,863 / hr"],
        ["500 kbps (Drivetrain)", "2,000 fps", "6,552 / hr", "1 per 7.7 days", "1,863 / hr"],
        ["1.0 Mbps (High-Speed)", "4,000 fps", "13,104 / hr", "1 per 3.8 days", "1,863 / hr"]
    ]
    add_tbl(doc, t3_h, t3_r, [1.0, 0.65, 0.75, 0.85, 0.85], "III", "Operational Line-Speed False Alarm Projections and Temporal Debouncing")

    # --- VII. CAMOUFLAGE ---
    add_sec_h(doc, "VII", "Adversarial Camouflage and Retraining Tradeoffs")
    add_subsec_h(doc, "A", "Feature-Space Stealth Blending")
    add_p(doc,
        "Adversaries aware of machine learning surveillance shape attack feature distributions toward benign driving centroids. "
        "Adversarial blending conceals malicious frame signatures. "
        "We model stealth blending via convex interpolation parameter alpha in [0, 1]:"
    )
    add_formula(doc, "x_camouflaged(alpha) = (1 - alpha) * x_attack + alpha * mu_normal", "7")
    add_p(doc,
        "As alpha increases from 0.00 to 1.00, attack frames match nominal feature expectations. "
        "At alpha = 0.25, attack detection recall drops to 41.80% (Table IV). "
        "At centroid match (alpha = 1.00), recall collapses to 19.60% (Fig. 2). "
        "Adversarial blending exploits intrinsic decision boundary geometry. "
        "This empirical floor aligns with the theoretical Bayes error bound P_e^* = 0.5 - 0.25 * ||p_0 - p_1||_1, proving that statistical classifiers cannot differentiate identical feature distributions."
    )
    
    # FIG 2
    fig2_path = RESULTS_DIR / "figure5_stealth_attack_degradation.png"
    add_fig(doc, fig2_path, "2", "Degradation of attack recall and MCC under increasing feature-space stealth camouflage alpha.", 3.2)

    add_subsec_h(doc, "B", "Curriculum Retraining and the False-Alarm Tradeoff")
    add_p(doc,
        "To mitigate stealth evasion, we implement leakage-free curriculum retraining across blended attack variants. "
        "Retraining recovers detection capability. "
        "Curriculum models recover 98.60% recall at alpha = 0.25 (a +56.80% improvement) and 95.40% at alpha = 0.50 (Table IV). "
        "However, expanding decision boundaries introduces severe collateral costs. "
        "Nominal false-positive rates increase 2.7-fold from 0.091% to 0.246%, generating 17,700 false alarms per hour at 500 kbps. "
        "Robustness against evasion directly multiplies nominal false alarms."
    )
    add_p(doc,
        "This tradeoff exposes a fundamental geometry of adversarial learning in safety-critical networks. "
        "Expanding decision boundaries to enclose perturbed attack points inevitably swallows previously unobserved normal driving regions. "
        "Decision boundaries become overly permissive. "
        "System designers cannot simply harden models against stealth attacks without accepting orders-of-magnitude increases in operational false alarms."
    )
    
    # TABLE IV
    t4_h = ["Camouflage alpha", "Baseline Recall", "Curriculum Recall", "Recall Gain", "Nominal FPR (500 kbps)"]
    t4_r = [
        ["alpha = 0.00 (Raw Attack)", "99.90%", "99.92%", "+0.02%", "0.091% (6,552 / hr)"],
        ["alpha = 0.25 (Slight Blend)", "41.80%", "98.60%", "+56.80%", "0.246% (17,700 / hr)"],
        ["alpha = 0.50 (Moderate Blend)", "33.10%", "95.40%", "+62.30%", "0.246% (17,700 / hr)"],
        ["alpha = 0.75 (Heavy Blend)", "31.40%", "48.20%", "+16.80%", "0.246% (17,700 / hr)"],
        ["alpha = 1.00 (Centroid Match)", "19.60%", "21.80%", "+2.20%", "0.246% (17,700 / hr)"]
    ]
    add_tbl(doc, t4_h, t4_r, [1.0, 0.7, 0.7, 0.65, 0.95], "IV", "Feature-Space Stealth Camouflage and Curriculum Retraining Trade-Off")

    # --- VIII. CROSS-PLATFORM TRANSFER ---
    add_sec_h(doc, "VIII", "Cross-Platform Generalization and Validation")
    add_subsec_h(doc, "A", "Direct Transfer Collapse to Polaris GEM e6")
    add_p(doc,
        "We evaluate cross-platform transferability using the GEM-CAN benchmark [6], collected from an instrumented Polaris GEM e6 autonomous shuttle. "
        "Vehicle platforms differ fundamentally in network architecture. "
        "The shuttle telemetry comprises 143,444 frames: 100,000 nominal frames using 29-bit extended SAE J1939 identifiers, and 43,444 adversarial frames (41,080 DoS, 663 Brake tampering, 662 Steering tampering, and 1,039 nominal frames). "
        "Unified ground truth establishes 101,039 nominal frames and 42,405 attack frames. "
        "Standard 11-bit identifiers are completely absent from nominal shuttle operation."
    )
    add_p(doc,
        "Directly evaluating the Sonata-trained Random Forest on raw GEM e6 telemetry causes total performance collapse: 35.57% accuracy, 0.1637 MCC, and a 91.48% false-positive rate (Table V, Fig. 3). "
        "The detector flags almost everything. "
        "Out of 101,039 benign frames, 92,427 are falsely flagged as attacks. "
        "Decision trees trained on 11-bit integers (0 to 2,047) route 29-bit identifiers (>400,000,000) into out-of-range leaves assigned to Fuzzy attack labels. "
        "Ablating identifiers (Payload-Only) yields 53.62% accuracy and 65.85% FPR, proving electric shuttle motor bytes diverge sharply from passenger car combustion telemetry."
    )
    
    # FIG 3
    fig3_path = RESULTS_DIR / "figure4_gem_can_cross_platform_confusion.png"
    add_fig(doc, fig3_path, "3", "Cross-platform confusion matrix evaluating Sonata-trained model on Polaris GEM e6 telemetry (35.57% accuracy, 91.48% FPR).", 3.2)

    add_subsec_h(doc, "B", "Calibration-Normalized Stateful Encoding")
    add_p(doc,
        "To achieve protocol independence without proprietary OEM database (.dbc) files, we construct a Calibration-Normalized Stateful Encoding. "
        "Five statistical transformations enforce protocol invariance. "
        "We compute an Allowlist Membership Indicator I_allow,i in {0, 1}, a Normalized Frequency Rank R_freq,i = Rank(ID_i) / |A_nom|, a Dimensionless Timing Ratio r_dt,i = Delta t_i / mu_dt,nom(ID_i), a Z-Scored Byte Mean z_mu,i, and a Z-Scored Byte Dispersion z_sig,i. "
        "Evaluating calibration-normalized features under zero-shot transfer lifts accuracy to 92.04%, increases MCC to 0.8316, and curtails FPR from 91.48% to 10.65% (Table V). "
        "The model isolates 100.00% of DoS flooding and 50.04% of physical actuator tampering."
    )
    
    # FIG 4
    fig4_path = RESULTS_DIR / "figure6_cross_platform_comparison.png"
    add_fig(doc, fig4_path, "4", "Performance comparison across Native In-Distribution, Direct Transfer, Payload-Only, Calibration-Normalized Encoding, and Native GEM Baseline.", 3.35)

    add_subsec_h(doc, "C", "Physical Masquerade Validation (ORNL ROAD Benchmark)")
    add_p(doc,
        "We evaluate supervised models using a strict chronological 70/30 time-blocked split on correlated_signal_attack_1_masquerade.csv (74,152 frames: 72,066 nominal driving, 2,086 physically injected masquerade frames) from the ORNL ROAD benchmark [5]. "
        "The adversary modifies safety-critical drivetrain signals while maintaining legitimate 20 ms and 50 ms broadcast periods and target ECU identifiers. "
        "The injected traffic introduces zero timing disruption."
    )
    add_p(doc,
        "Stateless timing baselines collapse to 0.00% recall (-0.0042 MCC, 0.056% FPR, Table VI). "
        "Timing baselines miss every single masquerade frame. "
        "Because inter-arrival intervals remain within nominal boundaries, timing rules are completely blind. "
        "In contrast, Random Forest and LightGBM models achieve 100.00% recall (0.9993 MCC, 0.005% FPR). "
        "Tree ensembles split on non-linear payload byte relationships between engine torque and vehicle velocity without requiring reverse-engineered database files."
    )
    
    add_subsec_h(doc, "D", "Supervised Few-Shot Domain Convergence")
    add_p(doc,
        "Production automotive workflows permit brief calibration windows during vehicle assembly or routine dealer servicing. "
        "Target-domain convergence requires just one second. "
        "We trace supervised adaptation dynamics across the held-out 30% test partition of GEM e6 telemetry by incrementing labeled adaptation frames N from N = 0 to N = 10,000 (representing 20.0 seconds of operation at 500 fps). "
        "Allocating N = 100 stratified frames (0.2 s) suppresses FPR to 0.73% and lifts tampering recall to 86.40% (Table V). "
        "With N = 500 frames (1.0 s), FPR drops to 0.04% (99.96% accuracy, 0.9991 MCC, 95.72% tampering recall). "
        "At N >= 1,000 frames (2.0 s), the model converges to 100.00% accuracy, 1.0000 MCC, 0.00% FPR, and 100.00% tampering recall. "
        "Domain alignment converges within brief observation windows."
    )
    add_p(doc,
        "Few-shot domain adaptation eliminates the need to author vehicle-specific rules by hand. "
        "Capturing two seconds of nominal telemetry calibrates stateful feature baselines for unseen vehicle models. "
        "Target calibration restores high detection sensitivity. "
        "This empirical result establishes an efficient operational deployment path for heterogeneous commercial fleets."
    )
    
    # TABLE V
    t5_h = ["Adaptation Regime (N)", "Window", "Accuracy", "MCC", "Nominal FPR", "Tamper Recall"]
    t5_r = [
        ["Direct Transfer (N=0)", "0.0 s", "35.57%", "0.1637", "91.48%", "100.00%*"],
        ["Payload-Only (N=0)", "0.0 s", "53.62%", "0.3646", "65.85%", "100.00%*"],
        ["Calib-Normalized (N=0)", "0.0 s", "92.04%", "0.8316", "10.65%", "50.04%"],
        ["Few-Shot (N=100)", "0.2 s", "99.48%", "0.9877", "0.73%", "86.40%"],
        ["Few-Shot (N=500)", "1.0 s", "99.96%", "0.9991", "0.04%", "95.72%"],
        ["Few-Shot (N=1,000)", "2.0 s", "100.00%", "1.0000", "0.00%", "100.00%"],
        ["Native GEM Baseline", "200.0 s", "100.00%", "1.0000", "0.00%", "100.00%"]
    ]
    add_tbl(doc, t5_h, t5_r, [1.1, 0.5, 0.55, 0.45, 0.65, 0.75], "V", "Cross-Platform Transfer and Few-Shot Adaptation Performance on Polaris GEM e6")
    p_t5_fn = doc.add_paragraph()
    p_t5_fn.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_t5_fn.paragraph_format.space_before = Pt(1.5)
    p_t5_fn.paragraph_format.space_after = Pt(4.0)
    p_t5_fn.paragraph_format.line_spacing = 1.0
    r_fn = p_t5_fn.add_run("*Direct Transfer and Payload-Only trigger alerts on 91.48% and 65.85% of all incoming 29-bit frames due to out-of-range ID/byte shifts, yielding 100.00% apparent recall via indiscriminate flagging rather than anomaly discrimination.")
    set_run(r_fn, 8.0, (0x44, 0x44, 0x44), italic=True)

    # TABLE VI
    t6_h = ["Detector Architecture", "Masquerade Recall", "Nominal FPR", "MCC", "Operational Detection Mechanism"]
    t6_r = [
        ["Timing-Rule Baseline", "0.00%", "0.056%", "-0.0042", "Blind: inter-arrival periodicity preserved"],
        ["LightGBM Booster", "100.00%", "0.005%", "0.9993", "Gradient splits isolate torque/velocity shifts"],
        ["Random Forest (Proposed)", "100.00%", "0.005%", "0.9993", "Non-linear payload byte correlations"]
    ]
    add_tbl(doc, t6_h, t6_r, [1.0, 0.65, 0.55, 0.5, 1.3], "VI", "Empirical Validation Against Physical Masquerade Attacks on ORNL ROAD Benchmark")

    # --- IX. TWO-TIER ARCHITECTURE ---
    add_sec_h(doc, "IX", "Two-Tier Authority Architecture and Regulatory Alignment")
    add_subsec_h(doc, "A", "Safety vs. Security Tension (ISO 26262 vs. ISO/SAE 21434)")
    add_p(doc,
        "Automotive systems operate under competing functional safety and cybersecurity mandates [7], [8]. "
        "Safety and security mandates collide on live buses. "
        "Functional safety under ISO 26262 requires deterministic fail-safe behavior: upon detecting a fault, control logic must transition actuators into a predictable safe state through limp-home modes or emergency deceleration [7]. "
        "In parallel, ISO/SAE 21434 mandates fail-secure defenses that detect intrusions, isolate compromised nodes, and reject forged command frames [8]."
    )
    add_p(doc,
        "Granting machine learning models unilateral actuation authority creates an acute operational hazard. "
        "When an in-vehicle IDS triggers hundreds of false alarms per hour, direct wiring into braking, steering, or bus transceivers becomes a severe liability. "
        "An adversary can inject trivial bus noise to induce deliberate false positives. "
        "Under such conditions, an automated system would execute emergency stops or disable steer-by-wire subsystems during highway transit [7]. "
        "False alarms directly compromise physical vehicle safety."
    )
    add_subsec_h(doc, "B", "Deployment Architecture: Strict Two-Tier Authority Separation")
    add_p(doc,
        "To decouple safety-critical actuation from statistical classifier volatility, vehicle electronic architectures must enforce a strict two-tier authority separation. "
        "Single-frame ML predictions cannot command physical actuators. "
        "Under Tier 1, machine learning classifiers execute strictly as passive telemetry sentinels hosted on central gateway microcontrollers. "
        "These models maintain zero unilateral authority over physical steering mechanisms, friction brakes, or powertrain controllers. "
        "Gateway sentinels log cryptographically signed security audit records to non-volatile memory and stream alert telemetry to off-board vehicle security operation centers over isolated telematics links [14]."
    )
    add_p(doc,
        "Under Tier 2, physical safety interventions require multi-sensor kinematic consensus across independent physical measurements. "
        "Actuator commands—such as emergency deceleration or contactor isolation—must never trigger on frame-level CAN anomalies alone. "
        "Anomaly notifications must correlate with physical state deviations captured by wheel-speed Hall-effect encoders, 6-axis inertial measurement units, and optical vehicle odometry [7]. "
        "Tier 2 supervisory controllers combine multi-modal physical consensus with temporal debounce filtering (k >= 3) before activating protective maneuvers. "
        "Cross-modal validation protects passenger safety."
    )
    add_p(doc,
        "This architectural separation satisfies both functional safety and cybersecurity compliance standards. "
        "Tier 1 sentinels fulfill UN Regulation 155 monitoring mandates without introducing single points of systemic failure. "
        "Physical vehicle actuators remain protected from spurious algorithmic interventions."
    )

    # --- X. CONCLUSION ---
    add_sec_h(doc, "X", "Conclusion and Replication Note")
    add_p(doc,
        "Offline benchmark accuracy does not guarantee road readiness. "
        "Our empirical deployability audit proves that reported accuracies exceeding 99.9% mask fundamental operational vulnerabilities in vehicle-layer AI detectors. "
        "A single-feature arbitration identifier baseline reaches 82.40% accuracy; that single scalar captures 76.64% of the gain above chance through static header memorization. "
        "In-distribution false-positive rates of 0.091% generate 2,621 to 13,104 false alarms per hour across standard automotive transceivers. "
        "While theoretical debouncing achieves suppression under Poisson assumptions, dynamic vehicle maneuvers generate 1,863 alarms per hour in real highway driving. "
        "Curriculum retraining recovers stealth attack recall to 98.60% at alpha = 0.25 but induces a 2.7-fold increase in nominal false alarms (0.246% FPR, 17,700 alarms/hr)."
    )
    add_p(doc,
        "Protocol differences cause severe transfer collapse. "
        "Direct evaluation on unseen 29-bit SAE J1939 shuttle telemetry drives false positives to 91.48%. "
        "Calibration-normalized stateful encoding restores zero-shot MCC to 0.8316, and stratified few-shot adaptation achieves 0.04% FPR at N = 500 frames (1.0 s) and converges to 0.00% FPR at N >= 1,000 frames (2.0 s). "
        "Vehicle-layer AI detectors must function strictly as passive telemetry sentinels within a two-tier architecture. "
        "These models must withhold autonomous physical actuation authority unless corroborated by multi-sensor kinematic consensus under ISO 26262 and ISO/SAE 21434."
    )
    add_p(doc,
        "Replication Note: Evaluation scripts and benchmark partitions are archived for open replication (HCRL Car-Hacking dataset via Korea University, GEM-CAN dataset via Zenodo doi: 10.5281/zenodo.19161139, and ORNL ROAD dataset via Oak Ridge National Laboratory). "
        "All code and reproduction materials are available on GitHub. "
        "Replication suites include complete data splitting scripts."
    )

    # --- REFERENCES ---
    add_sec_h(doc, "", "References")
    
    references_ieee = [
        "[1] HCRL, \"Car-Hacking Dataset for Intrusion Detection,\" Information Security Applications Lab, Korea University, 2018. [Online]. Available: https://ocslab.hksecurity.net",
        "[2] H. M. Song, H. R. Kim, and H. K. Kim, \"Intrusion detection system based on the analysis of time intervals of CAN messages for in-vehicle network,\" in Proc. IEEE Int. Conf. Inf. Netw. (ICOIN), 2016, pp. 63–68.",
        "[3] K. Koscher, A. Czeskis, F. Roesner, S. Patel, T. Kohno, S. Checkoway, D. McCoy, B. Kantor, D. Anderson, H. Shacham, and S. Savage, \"Experimental security analysis of a modern automobile,\" in Proc. IEEE Symp. Secur. Privacy (SP), 2010, pp. 447–462.",
        "[4] C. Miller and C. Valasek, \"Remote exploitation of an unaltered passenger vehicle,\" Black Hat USA Technical Report, 2015, pp. 1–91.",
        "[5] M. E. Verma, M. D. Iannacone, R. A. Bridges, S. C. Hollifield, P. Kay, and F. L. Combs, \"Addressing the lack of comparability & testing in CAN intrusion detection research: The ROAD dataset,\" Comput. Secur., vol. 125, art. 103034, 2023.",
        "[6] M. Tavasoli, M. Rostami, H. Kebriaei, and V. Puig, \"GEM-CAN: A real-world dataset of CAN-bus attack scenarios on an autonomous vehicle for intrusion-detection research,\" Data in Brief, vol. 66, art. 112805, 2026. doi: 10.1016/j.dib.2026.112805.",
        "[7] ISO 26262:2018, \"Road Vehicles — Functional Safety,\" International Organization for Standardization, Geneva, Switzerland, 2018.",
        "[8] ISO/SAE 21434:2021, \"Road Vehicles — Cybersecurity Engineering,\" International Organization for Standardization / SAE International, Geneva, Switzerland, 2021.",
        "[9] UNECE, \"UN Regulation No. 155: Cyber Security and Cyber Security Management System,\" United Nations Economic Commission for Europe, Geneva, Switzerland, 2021.",
        "[10] M. Althunayyan, A. Javed, and O. Rana, \"A robust multi-stage intrusion detection system for in-vehicle network security using hierarchical federated learning,\" Veh. Commun., vol. 49, art. 100837, 2024.",
        "[11] B. Lampe and W. Meng, \"can-train-and-test: A curated CAN dataset for automotive intrusion detection,\" arXiv:2308.07008, 2023.",
        "[12] F. Aloraini, A. Javed, and O. Rana, \"Adversarial attacks on intrusion detection systems in in-vehicle networks of connected and autonomous vehicles,\" Sensors, vol. 24, no. 12, art. 3848, 2024.",
        "[13] S. Rajapaksha, H. Kalutarage, M. O. Al-Kadri, A. Petrovski, G. Madzudzo, and M. Cheah, \"AI-based intrusion detection systems for in-vehicle networks: A survey,\" ACM Comput. Surv., vol. 55, no. 11, art. 237, 2023.",
        "[14] A. Oseni, N. Moustafa, G. Creech, N. Sohrabi, A. Strelzoff, Z. Tari, and I. Linkov, \"An explainable deep learning framework for resilient in-vehicle intrusion detection,\" IEEE Trans. Intell. Transp. Syst., vol. 24, no. 10, pp. 11200–11214, 2023.",
        "[15] J. Akshya, M. Sundarrajan, R. Vijayakumar, and A. Nayyar, \"Explainable AI-driven intrusion detection for securing IoT-enabled autonomous transportation systems,\" Cluster Comput., vol. 28, art. 884, 2025.",
        "[16] K. G. Arachchige, G. Alkaabi, M. Murtaza, Q. E. U. Haq, A. Z. Abualkishik, and C.-C. Lee, \"Threat landscape and integrated cybersecurity framework for V2V and autonomous electric vehicles,\" World Electr. Veh. J., vol. 16, no. 8, art. 469, 2025.",
        "[17] Z. Ying, R. Jiang, Y. Liu, and K. Chen, \"A literature review on V2X communications security: Foundation, challenges, and trends,\" IEEE Trans. Intell. Transp. Syst., 2024.",
        "[18] F. Pedregosa, G. Varoquaux, A. Gramfort, V. Michel, B. Thirion, O. Grisel, M. Blondel, P. Prettenhofer, R. Weiss, V. Dubourg, J. Vanderplas, A. Passos, D. Cournapeau, M. Brucher, M. Perrot, and E. Duchesnay, \"Scikit-learn: Machine learning in Python,\" J. Mach. Learn. Res., vol. 12, pp. 2825–2830, 2011."
    ]
    
    for ref_str in references_ieee:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p_ref.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(2.5)
        pf.line_spacing = 1.0
        pf.left_indent = Inches(0.25)
        pf.first_line_indent = Inches(-0.25)
        r = p_ref.add_run(ref_str)
        set_run(r, 8.0, (0, 0, 0), bold=False)

    doc.save(OUTPUT_DOCX)
    print(f"Generated DOCX: {OUTPUT_DOCX}")
    


def verify_pages():
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    wdoc = word.Documents.Open(str(OUTPUT_DOCX))
    pages = wdoc.ComputeStatistics(2)
    words = wdoc.ComputeStatistics(0)
    wdoc.SaveAs2(str(OUTPUT_PDF), FileFormat=17)
    wdoc.Close(False)
    word.Quit()
    
    reader = pypdf.PdfReader(str(OUTPUT_PDF))
    pdf_pages = len(reader.pages)
    print(f"\n==================================================")
    print(f"VERIFICATION METRICS:")
    print(f"  Word Statistics Words: {words}")
    print(f"  Word COM Page Count:   {pages}")
    print(f"  PDF Render Page Count: {pdf_pages}")
    print(f"==================================================")

if __name__ == "__main__":
    generate()
    verify_pages()
