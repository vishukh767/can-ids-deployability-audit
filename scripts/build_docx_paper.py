"""
build_docx_paper.py: Programmatic DOCX Builder for Draft v21 Journal Manuscript
Target Venue: Vehicular Communications (Elsevier, CiteScore 14.4, Impact Factor 6.7)

Title: The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit and Cross-Platform Domain Analysis
Author: Vishwas Khandelwal*
Institution: Vivekananda Institute of Professional Studies - Technical Campus (VIPS-TC), Delhi, India

Incorporates:
- Full Part A & Part B Empirical Findings:
  * Table 1: Comparative Literature Positioning
  * Table 2: In-Distribution Benchmark Performance & Ablation (76.6% single-feature gain)
  * Table 3: Operational CAN Bus Line-Speed False Alarm Projections and Debounce (7.7 days @ 500k, 19.2 days @ 250k)
  * Table 4: Feature-Space Stealth Camouflage Stress Test Across Blending Levels
  * Table 5: Reconciled Cross-Platform Transferability, Feature Isolation, and Native Baseline (101,039 Normal vs. 42,405 Attack)
  * Table 6: Multi-Detector Architectural Benchmark (Timing-Rule, RF, LightGBM, Autoencoder across HCRL and GEM e6)
  * Table 7: Leakage-Free Camouflage Curriculum Retraining & Nominal False Alarm Trade-off
  * Table 8: Empirical Validation Against Physical Masquerade Attacks (ORNL ROAD Benchmark)
  * Table 9: Supervised Few-Shot Domain Adaptation Performance on Polaris GEM e6 (N = 0 to 10,000 frames)
- ISO 11898-1 CRC-15 transceiver discard physics accurately described
- Empirical post-debounce burstiness measured across 200,000 continuous normal frames (normal_run_data.txt)
- Embedded analytical complexity estimates (100 trees, 34,200 nodes, 274 KB flash, <4 KB stack)
- Verified citations: Tavasoli et al. (2026), Verma et al. / ROAD (2023), Oseni et al. (2023), Tariq et al. (2021), Shit & Subudhi (2025)
- Standard journal declarations (Data/Code Availability, Funding, Competing Interests, Author Contributions, AI-use disclosure)
"""

import os
import shutil
from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR if (SCRIPT_DIR / "4_datasets").is_dir() else SCRIPT_DIR.parent
OUTPUT_DIR = PROJECT_ROOT / "1_manuscripts_and_drafts"
RESULTS_DIR = PROJECT_ROOT / "3_results_and_figures"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

doc = docx.Document()

# ==============================================================================
# PAGE SETUP & NUMBERING
# ==============================================================================
for s in doc.sections:
    s.top_margin = Inches(1.0)
    s.bottom_margin = Inches(1.0)
    s.left_margin = Inches(1.0)
    s.right_margin = Inches(1.0)
    s.different_first_page_header_footer = True

    hp = s.header.paragraphs[0]
    hp.text = "The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit"
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if hp.runs:
        hp.runs[0].font.name = "Calibri"
        hp.runs[0].font.size = Pt(8.5)
        hp.runs[0].font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    fp = s.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run("Page ")
    fr.font.name = "Calibri"
    fr.font.size = Pt(9)
    fr.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    fp._element.append(fld)

def set_run(r, sz=11, rgb=(0x26, 0x26, 0x26), bold=False, italic=False, font_name="Calibri"):
    r.font.name = font_name
    r.font.size = Pt(sz)
    r.font.color.rgb = RGBColor(*rgb)
    r.bold = bold
    r.italic = italic

def add_p(d, txt, before=0, after=6, spacing=1.15):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = spacing
    r = p.add_run(txt)
    set_run(r, 11, (0x26, 0x26, 0x26))
    return p

def add_h1(d, txt):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(txt)
    set_run(r, 13.5, (0x1F, 0x4E, 0x79), bold=True)
    return p

def add_h2(d, txt):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(txt)
    set_run(r, 12, (0x2F, 0x55, 0x97), bold=True)
    return p

def add_h3(d, txt):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(txt)
    set_run(r, 11, (0x33, 0x33, 0x33), bold=True, italic=True)
    return p

def add_formula(d, formula_text, formula_num):
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(5)
    r1 = p.add_run(formula_text)
    set_run(r1, 10.5, (0x11, 0x11, 0x11), italic=True, font_name="Cambria Math")
    r2 = p.add_run(f"    ({formula_num})")
    set_run(r2, 10, (0x55, 0x55, 0x55), bold=True)
    return p

def add_tbl(d, headers, rows, widths, cap_num, cap_tit):
    p_cap = d.add_paragraph()
    p_cap.paragraph_format.space_before = Pt(10)
    p_cap.paragraph_format.space_after = Pt(3)
    p_cap.paragraph_format.keep_with_next = True
    r_n = p_cap.add_run(f"{cap_num}: ")
    set_run(r_n, 10, (0x1F, 0x4E, 0x79), bold=True)
    r_t = p_cap.add_run(cap_tit)
    set_run(r_t, 10, (0x26, 0x26, 0x26), bold=True)

    t = d.add_table(rows=len(rows)+1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0]
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        tcPr = c._element.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1F4E79"/>'))
        p = c.paragraphs[0]
        r = p.add_run(h)
        set_run(r, 9.5, (0xFF, 0xFF, 0xFF), bold=True)
    for r_i, r_data in enumerate(rows):
        row = t.rows[r_i + 1]
        bg = 'F9FAFC' if r_i % 2 == 1 else 'FFFFFF'
        for c_i, val in enumerate(r_data):
            c = row.cells[c_i]
            if bg != 'FFFFFF':
                c._element.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
            p = c.paragraphs[0]
            r = p.add_run(str(val))
            set_run(r, 9.0, (0x26, 0x26, 0x26))
    for row in t.rows:
        for c_i, w in enumerate(widths):
            row.cells[c_i].width = Inches(w)

def add_fig(d, fig_path, cap_num, cap_tit, width_in=5.4):
    p_img = d.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    resolved_path = None
    candidates = [Path(fig_path), PROJECT_ROOT / Path(fig_path).name, RESULTS_DIR / Path(fig_path).name]
    for p in candidates:
        if p.is_file():
            resolved_path = p
            break
    if resolved_path:
        r = p_img.add_run()
        r.add_picture(str(resolved_path), width=Inches(width_in))
    else:
        r = p_img.add_run(f"[Graphic file not found: {fig_path}]")
        set_run(r, 10, (0xC0, 0x00, 0x00), bold=True)

    p_cap = d.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(8)
    r_n = p_cap.add_run(f"Figure {cap_num}: ")
    set_run(r_n, 9.5, (0x1F, 0x4E, 0x79), bold=True)
    r_t = p_cap.add_run(cap_tit)
    set_run(r_t, 9.5, (0x33, 0x33, 0x33), italic=True)

# ==============================================================================
# TITLE & AUTHOR METADATA
# ==============================================================================
p_t = doc.add_paragraph()
p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_t.paragraph_format.space_before = Pt(0)
p_t.paragraph_format.space_after = Pt(6)
r_t = p_t.add_run("The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit and Cross-Platform Domain Analysis")
set_run(r_t, 16.5, (0x1F, 0x4E, 0x79), bold=True)

p_a = doc.add_paragraph()
p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_a.paragraph_format.space_before = Pt(2)
p_a.paragraph_format.space_after = Pt(12)
r_a = p_a.add_run(
    "Vishwas Khandelwal*\n"
    "Department of Computer Science and Engineering\n"
    "Vivekananda Institute of Professional Studies - Technical Campus (VIPS-TC), Delhi, India\n"
    "[Corresponding Author Email: vishukh767@gmail.com]"
)
set_run(r_a, 10, (0x33, 0x33, 0x33))

# ==============================================================================
# ABSTRACT & KEYWORDS (<250 words)
# ==============================================================================
add_h1(doc, "Abstract")
abstract_text = (
    "In-vehicle intrusion detection systems (CAN-IDS) report classification accuracies exceeding 99.9% on automotive benchmarks. "
    "This performance ceiling collapses in production. "
    "Evaluating Random Forest ensembles on chronological 70/30 time-blocked splits of the HCRL dataset (Hyundai YF Sonata), "
    "we reproduce baseline metrics (99.90% accuracy, 0.999 MCC, 0.091% FPR) and isolate operational failure modes. "
    "An arbitration ID single-feature tree achieves 82.40% accuracy; it isolates 76.64% of achievable gain above chance (25.00%). "
    "Translating the 0.091% benchmark FPR to physical bus line speeds produces 2,621 to 13,104 unmitigated false alarms per hour across standard automotive transceivers operating between 250 kbps and 1 Mbps. "
    "Whole-bus debouncing (k = 3) theoretically suppresses Poisson errors to one alarm per 7.7 days at 500 kbps. "
    "Dynamic driving maneuvers break this independence. "
    "Auditing 200,000 continuous normal frames reveals 1,863 alarms/hr under whole-bus filtering and 510,134 alarms/hr per-ID. "
    "Adversarial feature blending reduces attack recall to 19.60%; curriculum retraining recovers recall to 98.60% (alpha = 0.25) "
    "but increases nominal FPR to 0.246% (17,700 alarms/hr). "
    "Transfer to a Polaris GEM e6 shuttle incurs 91.48% FPR from 11-bit versus 29-bit SAE J1939 mismatches. "
    "Calibration-normalized stateful encoding restores zero-shot MCC to 0.8316. "
    "Stratified few-shot adaptation achieves 0.04% FPR at N = 500 frames (1.0 s) and converges to 0.00% FPR with 100.00% tampering recall at N >= 1,000 frames (2.0 s). "
    "On ORNL ROAD physical masquerade attacks, timing baselines record 0.00% recall, whereas payload tree ensembles maintain 100.00% recall. "
    "Vehicle-layer detectors must not command physical actuators without multi-sensor kinematic consensus under ISO 26262 and ISO/SAE 21434."
)
add_p(doc, abstract_text)

keywords_text = (
    "Keywords—Controller Area Network (CAN), Automotive Cybersecurity, Intrusion Detection Systems (IDS), Deployability Audit, "
    "False Alarm Rate Translation, Cross-Platform Generalization, Temporal Debouncing, ISO 26262, ISO/SAE 21434."
)
p_kw = add_p(doc, keywords_text, before=2, after=12)
if p_kw.runs:
    set_run(p_kw.runs[0], 10, (0x1F, 0x4E, 0x79), bold=True, italic=True)

# ==============================================================================
# 1 INTRODUCTION
# ==============================================================================
add_h1(doc, "1 Introduction")
add_p(doc,
    "Machine learning models for in-vehicle intrusion detection (CAN-IDS) report classification accuracies exceeding 99.9% alongside false-positive rates below 0.10% on automotive benchmarks [2], [4], [14]. "
    "This precision creates an immediate availability hazard [7]. "
    "On a 500 kbps Controller Area Network (CAN) bus operating at 50% nominal bus load (~2,000 frames per second), an empirical false-positive rate of 0.091% translates to 6,552 false alarms per driving hour. "
    "That equals 1.82 false alerts every second. "
    "Offline evaluation benchmarks mask this operational burden by evaluating classifiers on randomized train/test splits from stationary vehicle captures; models memorize static message arbitration identifiers rather than learning valid vehicular dynamics [12]."
)
add_p(doc,
    "Broadcast arbitration in CAN networks operates without sender authentication, payload encryption, or message integrity checks [17]. "
    "Lower numerical identifier values seize physical bus priority during arbitration [13], [18]. "
    "When a compromised electronic control unit (ECU) transmits identifier 0x000 at line rate, it asserts continuous dominant zeros that monopolize physical differential bus lines and force transceiver timeout faults [1], [11]. "
    "Legitimate safety-critical transmission slots starve. "
    "Deploying machine learning detectors directly onto resource-constrained automotive microcontrollers requires empirical stress tests that evaluate baseline shortcut memorization, line-speed error translation, feature-space stealth evasion, and cross-platform protocol boundaries under operational ISO 26262 constraints [14]."
)
add_p(doc,
    "We evaluate an ensemble Random Forest on the HCRL Car-Hacking dataset (Hyundai YF Sonata) [1] using a strict chronological 70/30 time-blocked split to eliminate temporal data leakage. "
    "Signal isolation uncovers severe shortcut dependence. "
    "A single-feature decision tree thresholding exclusively on arbitration ID achieves 82.40% accuracy; this simple boundary isolates 76.64% of the multi-feature model's gain above chance (25.00%) without inspecting payload data bytes or transmission timing intervals. "
    "Completely removing inter-arrival intervals (Delta t_i) shifts ensemble accuracy by only 0.05% (99.90% to 99.85%). "
    "Classifiers separate attack classes via identifier allowlists. "
    "Translating the nominal 0.091% false-positive rate across operational bitrates (250 kbps, 500 kbps, 1 Mbps) yields 2,621 to 13,104 unmitigated false alarms per driving hour. "
    "This burden equals 0.73 to 3.64 false alerts per second. "
    "Multi-frame sliding window debouncing (k = 3) theoretically suppresses independent Poisson errors to one alarm every 7.7 operating days at 500 kbps (1.5 ms latency). "
    "Real driving destroys this mathematical assumption. "
    "Evaluating 200,000 continuous normal frames from real driving logs (normal_run_data.txt) reveals 22,136 raw false positives (11.068% FPR, 778,029 alarms/hr), where dynamic vehicle maneuvers produce 14,514 correlated false alerts under per-ID filtering (510,134 alarms/hr). "
    "Under feature-space camouflage, blending malicious feature vectors toward nominal centroids degrades attack recall to 19.60%; curriculum retraining recovers recall to 98.60% at alpha = 0.25, but forces nominal FPR to 0.246% (17,700 false alarms/hr at 500 kbps)."
)
add_p(doc,
    "Cross-platform evaluation exposes an even steeper architectural barrier. "
    "Direct transfer of the Sonata-trained classifier to 143,444 real-world CAN frames from an autonomous Polaris GEM e6 shuttle (101,039 Normal vs. 42,405 Attack frames) [6] induces a 91.48% FPR and 35.57% accuracy. "
    "The 29-bit headers break every split. "
    "Decision trees trained on standard 11-bit integer boundaries branch into out-of-range leaves when processing extended 29-bit SAE J1939 identifiers whose numerical values exceed 400,000,000. "
    "Four distinct detector paradigms break similarly. "
    "Evaluating stateless timing rules, Random Forest, LightGBM, and Deep Reconstruction Autoencoders demonstrates that all feature-aware architectures suffer representation breakdown (74.60% to 100.00% FPR) under direct transfer because none of them encode platform-invariant protocol semantics. "
    "Our Calibration-Normalized Stateful Encoding standardizes telemetry against nominal operating distributions without proprietary DBC files; this representation restores zero-shot transfer MCC to 0.8316 and reduces cross-platform FPR to 10.65%. "
    "Supervised target adaptation converges within two seconds. "
    "Stratified few-shot training achieves 0.04% FPR at N = 500 frames (1.0 s) and reaches 0.00% FPR with 100.00% tampering recall at N >= 1,000 frames (2.0 s). "
    "Against real physical masquerade attacks from the Oak Ridge National Laboratory (ORNL) ROAD dataset [16], stateless timing baselines record 0.00% recall (-0.0042 MCC) because the attacker preserves 20 ms and 50 ms broadcast periodicity, whereas payload tree ensembles maintain 100.00% recall (0.9993 MCC)."
)
add_p(doc,
    "This investigation provides three tangible artifacts:"
)
add_p(doc,
    "• Open-Source Replication Suite: Full Python evaluation scripts, chronological splits, and model checkpoints are archived for the HCRL [1], Polaris GEM-CAN [6], and ORNL ROAD [16] datasets. Every pipeline is fully reproducible."
)
add_p(doc,
    "• Protocol-Agnostic Feature Engine: A calibration-normalized stateful encoding algorithm standardizes CAN 2.0A and SAE J1939 telemetry without reverse-engineered OEM DBC databases. This eliminates proprietary decoding barriers."
)
add_p(doc,
    "• Two-Tier Regulatory Architecture: A deployability framework reconciles ISO 26262 fail-safe and ISO/SAE 21434 fail-secure mandates; machine learning detectors serve as passive logging sentinels while physical actuation requires multi-sensor kinematic consensus. Actuation requires cross-modal sensory agreement."
)

# ==============================================================================
# 2 THREAT LANDSCAPE & ATTACK TAXONOMY
# ==============================================================================
add_h1(doc, "2 Threat Landscape and Automotive Attack Surface")
add_h2(doc, "2.1 Connected Vehicle Attack Surfaces")
add_p(doc,
    "Modern connected vehicles expose multiple physical and wireless attack interfaces that bridge external networks directly to internal Controller Area Network lines [5], [17]. "
    "Physical entry points center on the OBD-II port. "
    "On-Board Diagnostics (OBD-II) connectors provide unauthenticated electrical access to vehicle buses, frequently bridged by commercial fleet management dongles [11]. "
    "Remote attack vectors expand this perimeter through cellular Telematics Control Units (TCUs) executing over-the-air firmware updates, In-Vehicle Infotainment (IVI) units running Wi-Fi and Bluetooth stacks, and Vehicle-to-Everything (V2X) radio transceivers [13], [22], [23]. "
    "Remote exploits bridge these distinct subsystems. "
    "An adversary compromising any wireless module can inject arbitrary differential voltage pulses directly into safety-critical powertrain and chassis control domains."
)

add_h2(doc, "2.2 CAN Bus Arbitration and Vulnerabilities")
add_p(doc,
    "Channel access across differential CAN High and CAN Low lines is governed by lossless bitwise arbitration [13]. "
    "Bits are classified as dominant (logical 0) or recessive (logical 1). "
    "When two nodes transmit simultaneously, an ECU sending a recessive bit that senses a dominant state yields the bus immediately to the competing node [18]. "
    "Lower numeric identifiers hold higher transmission priority. "
    "Identifier 0x000 claims absolute physical priority. "
    "A compromised transceiver asserting zeros at maximum line rate monopolizes the physical medium. "
    "Legitimate ECUs cannot transmit periodic status messages [1]. "
    "Transmit error counters rapidly exceed 255. "
    "This condition forces physical bus-off state transitions, removing critical safety microcontrollers from the vehicle network entirely [1], [11]. "
    "The physical broadcast architecture lacks cryptographic message authentication, sender verification, and payload encryption [17]. "
    "Any node can inject arbitrary frame payloads."
)

add_h2(doc, "2.3 Attack Taxonomy in In-Vehicle Networks")
add_p(doc,
    "In-vehicle network exploits split into volumetric flooding and semantic masquerade injections [1], [16], [17]. "
    "Volumetric attacks target physical channel capacity. "
    "In denial-of-service (DoS) injections, an adversary transmits continuous bursts of high-priority arbitration identifiers, most commonly identifier 0x000 in standard 11-bit CAN 2.0A configurations. "
    "Because identifier 0x000 wins every arbitration contest, legitimate periodic messages from steering, braking, and powertrain controllers cannot access the physical differential bus lines [1]. "
    "Receiving ECUs register immediate communication timeout faults. "
    "Fuzzing injections spray randomized arbitration identifiers and random payload byte sequences across the differential lines at maximum transceiver line speeds without regard to message schema. "
    "This flood saturates hardware receive buffers, desynchronizes ECU state machines, and illuminates instrument cluster warning indicators across the dashboard [1]. "
    "Both volumetric vectors distort inter-arrival timing intervals. "
    "Stateless threshold detectors easily identify these anomalies due to the extreme message frequency spikes."
)
add_p(doc,
    "Semantic attacks execute without altering nominal bus timing. "
    "In spoofing injections, an adversary broadcasts valid arbitration identifiers assigned to real vehicular control loops, including engine speed, wheel velocity, transmission gear, and steering angle [1], [16]. "
    "Because the injected identifiers match legitimate ECU allowlists, static identifier filters accept the malicious frames without raising security flags. "
    "Detection requires analyzing multi-byte payload semantic correlations. "
    "Stealth camouflage elevates this evasion further. "
    "Camouflaged adversaries align message injections with the true cyclic transmission period of the target ECU (for example, transmitting precisely at 20 ms intervals) while constraining payload values within empirical nominal distributions [3], [16]. "
    "These injections avoid timing thresholds completely. "
    "Machine learning models evaluating statistical distributions classify these camouflaged frames as nominal background traffic unless trained on adversarial distributions."
)

add_h2(doc, "2.4 Grounding in Real-World Automotive Exploits")
add_p(doc,
    "Physical and remote takeover of production vehicles is an established empirical reality [11], [13]. "
    "Koscher et al. [11] demonstrated full internal control of passenger vehicles through direct OBD-II access, successfully disabling physical braking subsystems and preventing engine startup during dynamometer testing. "
    "Remote exploits present an even broader threat. "
    "Miller and Valasek [13] executed an unconstrained remote compromise of an unmodified 2014 Jeep Cherokee across public wireless cellular networks. "
    "They breached the telematics gateway remotely. "
    "By rewriting the gateway microcontroller firmware, they bridged from the infotainment bus onto the high-speed vehicle control bus, physically commanding steering actuators and killing the internal combustion engine at highway speeds. "
    "Modern regulatory mandates respond directly to these vulnerabilities. "
    "UNECE UN Regulation No. 155 (UN R155), ISO/IEC 5962 [8], and ISO/SAE 21434 [9], [21] enforce certified Cyber Security Management Systems (CSMS) alongside software bill of materials (SBOM) traceability across vehicle engineering lifecycles. "
    "Deploying intrusion detection systems on internal CAN buses forms a central technical requirement under these international standards."
)
# ==============================================================================
# 3 RELATED WORK & BENCHMARK POSITIONING
# ==============================================================================
add_h1(doc, "3 Related Work and Benchmark Positioning")
add_h2(doc, "3.1 In-Vehicle Intrusion Detection Paradigms")
add_p(doc,
    "Early in-vehicle intrusion detection systems monitored inter-message arrival intervals and cyclic broadcast periodicities [18]. "
    "Song et al. [18] flagged intrusions by identifying disruptions in transmission frequencies across active control loops. "
    "Volume floods break periodicity immediately. "
    "However, pure timing baselines cannot detect stealthy masquerade injections where an attacker preserves nominal broadcast intervals [16]. "
    "Subsequent research shifted toward payload byte inspection. "
    "Oseni et al. [14] formulated a deep learning pipeline combining autoencoders with feedforward neural networks for connected vehicles. "
    "Tariq et al. [20] introduced CAN-ADF; their architecture models temporal transition matrices across message arbitration identifiers. "
    "Akshya et al. [2] proposed tree ensembles paired with SHAP and LIME explainability frameworks. "
    "Embedded execution limits these explainable models [10]. "
    "Microcontroller cycle budgets prevent real-time SHAP kernel evaluation at line speed, and Aloraini et al. [3] proved that adversarial feature perturbations easily bypass deep payload neural networks."
)
add_p(doc,
    "Fleet-level intrusion detection introduces severe network telemetry constraints. "
    "Althunayyan et al. [4] investigated hierarchical federated learning architectures to aggregate threat signatures across vehicle clusters without transmitting raw CAN streams. "
    "Shit and Subudhi [19] evaluated blockchain-assisted distributed ledgers for tamper-resistant cryptographic anomaly logging across connected vehicle fleets. "
    "Cellular bandwidth limits raw frame transmission. "
    "Fleet-wide telemetry gateways cannot upload millions of uncompressed CAN frames per driving hour; therefore, detection must execute deterministically on local electronic control units before transmitting compressed security digests across wireless channels."
)

add_h2(doc, "3.2 Critical Audit of Automotive CAN Benchmarks")
add_p(doc,
    "The validity of machine learning detectors depends on dataset capture conditions and partitioning methodology [12], [16]. "
    "The HCRL Car-Hacking dataset (2018) [1] remains the most widely cited automotive benchmark, containing denial-of-service, fuzzing, gear spoofing, and RPM spoofing captures recorded from a stationary Hyundai YF Sonata. "
    "Injections in HCRL saturate the bus at high frame frequencies. "
    "High-rate bursts create artificial statistical shortcuts [12]. "
    "Classifiers evaluated on HCRL memorize high-density burst signatures rather than physical vehicle state transitions [12], [16]. "
    "Lampe and Meng [12] curated the multi-vehicle can-train-and-test benchmark and demonstrated that randomized train/test splits induce temporal data leakage; identical cyclic frame sequences appear in both splits and inflate test accuracy artificially."
)
add_p(doc,
    "Recent benchmarks eliminate these experimental shortcuts by capturing physical vehicle operation under realistic attack constraints. "
    "Verma et al. [16] released the Real Agenda-Driven (ROAD) dataset using physical driving telemetry recorded from an unmodified passenger vehicle across live road tests. "
    "ROAD introduces timing-preserving masquerade injections where falsified drivetrain signals conform strictly to nominal 20 ms and 50 ms broadcast periods. "
    "Timing rules miss these subtle manipulations [16]. "
    "Complementing passenger car telemetry, Tavasoli et al. [6] published the GEM-CAN dataset in Data in Brief, recording 143,444 raw frames from an autonomous Polaris GEM e6 electric shuttle. "
    "GEM-CAN captures real-world physical steering and brake tampering alongside denial-of-service floods across an extended 29-bit SAE J1939 network architecture. "
    "Table 1 contrasts these benchmarks against our unified deployability audit."
)

# Table 1
t1_h = ["Study / Benchmark", "Vehicle Scope", "Split Strategy", "Line-Speed Translation", "Cross-Platform Audit", "Primary Limitation Addressed"]
t1_r = [
    ["Song et al. (2016) [18]", "Single passenger car", "Sequential blocks", "Theoretical only", "No (single vehicle)", "Timing-only; blind to masquerade payload tampering"],
    ["Kang / HCRL (2018) [1]", "Hyundai YF Sonata", "Randomized k-fold", "Not addressed", "No (single vehicle)", "High injection frequencies; random split leakage risk"],
    ["Oseni et al. (2023) [14]", "Simulated CAV network", "Stratified random", "Evaluated latency", "No cross-vehicle test", "High accuracy claimed without operational false-alarm audit"],
    ["Tariq et al. (2021) [20]", "Passenger car (CAN-ADF)", "Sequential evaluation", "Partial audit", "No (single vehicle)", "Anomaly framework without multi-baud line-speed translation"],
    ["Verma et al. / ROAD [16]", "Passenger car", "Time-blocked captures", "Partial audit", "No (single vehicle)", "Focuses on masquerade; does not quantify line-speed alarm rates"],
    ["Lampe & Meng [12]", "Multiple vehicles", "Chronological blocks", "Not quantified", "Multi-vehicle evaluated", "Documents cross-vehicle drops; no debounce filter evaluation"],
    ["This Work", "Sonata + Polaris GEM e6 + ROAD", "Strict 70/30 time-blocked", "Empirically quantified", "Cross-architecture audit", "Full deployability audit: 11-vs-29 bit, line-speed FPR, debounce, few-shot"]
]
add_tbl(doc, t1_h, t1_r, [1.5, 1.1, 1.2, 1.1, 1.1, 1.6], "Table 1", "Comparative Positioning of Automotive CAN Intrusion Detection Literature")

# ==============================================================================
# 4 MATHEMATICAL FORMULATION & EXPERIMENTAL METHODOLOGY
# ==============================================================================
add_h1(doc, "4 Mathematical Formulation and Experimental Methodology")
add_h2(doc, "4.1 CAN Bus Message Representation and Feature Space")
add_p(doc,
    "A continuous CAN communication stream forms an ordered sequence of frames S = {f_1, f_2, ..., f_N}. "
    "Physical transceivers serialize each discrete packet. "
    "Each frame f_i is defined as:"
)
add_formula(doc, "f_i = (t_i, ID_i, DLC_i, B_i,0, B_i,1, ..., B_i,DLC-1, y_i)", "1")
add_p(doc,
    "where t_i in R^+ denotes the microsecond timestamp; ID_i in N_0 denotes the arbitration identifier; DLC_i in {0, ..., 8} denotes the payload length in bytes; "
    "B_i,j in {0x00, ..., 0xFF} represents the j-th payload data byte; and y_i represents the ground-truth classification label [1]."
)
add_p(doc,
    "Automotive microcontrollers lack hardware floating-point units for multi-layer tensor operations. "
    "Lightweight statistical features bound embedded execution overhead. "
    "For each arrival frame f_i, we extract a five-dimensional feature vector x_i in R^5 containing the raw integer arbitration identifier ID_i "
    "(0 to 2,047 for standard 11-bit CAN 2.0A; 0 to 536,870,911 for extended 29-bit CAN 2.0B) and the payload data length code DLC_i. "
    "Inter-arrival timing Delta t_i measures the elapsed time since the previous transmission of the identical arbitration identifier:"
)
add_formula(doc, "Delta t_i = t_i - max { t_k | k < i, ID_k = ID_i }", "2")
add_p(doc,
    "Order of operations matters here. "
    "We calculate Delta t_i over the unbroken raw packet capture before any dataset subsampling; calculating inter-arrival times after subsampling destroys the physical timing intervals of the automotive bus. "
    "Payload bytes are summarized by their arithmetic mean mu_byte,i and standard deviation sigma_byte,i:"
)
add_formula(doc, "mu_byte,i = (1 / DLC_i) * sum_{j=0}^{DLC_i - 1} B_i,j", "3")
add_formula(doc, "sigma_byte,i = sqrt( (1 / DLC_i) * sum_{j=0}^{DLC_i - 1} (B_i,j - mu_byte,i)^2 )", "4")
add_p(doc,
    "where mu_byte,i = 0 and sigma_byte,i = 0 when DLC_i = 0."
)

add_h2(doc, "4.2 Eliminating Data Leakage: 70/30 Time-Blocked Partitioning")
add_p(doc,
    "Periodic broadcast scheduling governs normal CAN bus traffic, with functional messages repeating at fixed 10 to 100 ms intervals [18]. "
    "Randomized data partitioning causes severe temporal leakage [12]. "
    "In a randomized split, identical cyclic messages from the same vehicle state distribute across both training and test subsets; the classifier memorizes static payload constants instead of learning valid operational transitions [12]. "
    "We enforce a strict chronological 70/30 time-blocked partition across each attack capture. "
    "For each log spanning time interval [T_start, T_end], the training partition is drawn strictly from [T_start, T_start + 0.70 * (T_end - T_start)]; the test partition evaluates exclusively on the subsequent chronological window (T_start + 0.70 * (T_end - T_start), T_end]. "
    "We sample 20,000 frames per class from each chronological block. "
    "This sampling strategy forms a balanced corpus of 80,000 frames per evaluation run (56,000 training frames and 24,000 test frames), repeated across five independent random seeds (mean +/- SD) (Table S3)."
)

add_h2(doc, "4.3 Statistical Metrics and Evaluation")
add_p(doc,
    "Multi-class model discrimination is assessed using classification accuracy, precision, recall, macro-averaged F1-score, and the Matthews Correlation Coefficient (MCC):"
)
add_formula(doc, "MCC = (TP * TN - FP * FN) / sqrt( (TP + FP)(TP + FN)(TN + FP)(TN + FN) )", "5")
add_p(doc,
    "MCC provides a balanced performance metric across skewed automotive attack distributions. "
    "Accuracy alone masks critical false alarm rates. "
    "We calculate 95% confidence intervals for the false-positive rate (FPR) using the Wilson score interval with continuity correction across n = 6,000 normal evaluation frames per seed: CI_95% = [0.034%, 0.198%] (pooled N = 30,000: [0.062%, 0.134%]). "
    "Multi-class Area Under the Precision-Recall Curve (AUPRC_macro = 1.000) is computed as the unweighted macro-average across all C = 4 classes."
)

add_h2(doc, "4.4 Single-Feature Baselines and Embedded Complexity Model")
add_p(doc,
    "We benchmark our multi-feature architecture against a single-feature decision tree operating strictly on the raw arbitration identifier ID_i [15]. "
    "The baseline tree isolates identifier memorization directly. "
    "The primary classifier is an ensemble Random Forest comprising 100 orthogonal decision trees with a maximum depth of 12; each tree splits internal nodes by minimizing Gini impurity [15]. "
    "We quantify embedded deployment feasibility using an analytical instruction-cycle and memory model targeted at an ARM Cortex-M7 automotive microcontroller (Table S5). "
    "Standard fixed-point C-struct node representations allocate 8 bytes per decision node. "
    "Across an average of 342 nodes per tree (~34,200 total nodes across 100 trees), this layout requires an estimated flash memory footprint of 273.6 KB (~274 KB) and consumes less than 4 KB of stack RAM. "
    "Per-frame inference executes in 18.4 microseconds. "
    "This execution speed remains well within standard automotive CAN message inter-arrival budgets (0.5 to 100 ms)."
)
# ==============================================================================
# 5 EMPIRICAL AUDIT: IN-DISTRIBUTION BENCHMARK
# ==============================================================================
add_h1(doc, "5 Empirical Deployability Audit: In-Distribution Benchmark")
add_h2(doc, "5.1 Baseline Benchmark Performance")
add_p(doc,
    "The time-blocked HCRL test set yields a mean classification accuracy of 99.90% +/- 0.008%, an MCC of 0.999 +/- 0.001, and an FPR of 0.091% across five random seeds (Table 2). "
    "Multi-class performance distributes across 24,000 test frames with 6,000 frames evaluated per class (Figure 1). "
    "Classification errors concentrate exclusively in Fuzzy injections. "
    "Out of 6,000 normal driving frames, the model misclassifies 9 frames as Fuzzy attacks; this represents a 0.150% single-seed FPR and a 0.091% five-seed mean FPR. "
    "Both DoS and Spoofing injections achieve 100.0% recall across all test blocks. "
    "Fuzzy injections attain 99.75% recall, where the model assigns 14 attack frames to Normal and 1 to Spoofing."
)

# Table 2
t2_h = ["Model Configuration", "Feature Subset", "Split Strategy", "Accuracy", "MCC", "Attack Recall", "False Positive Rate (95% CI)"]
t2_r = [
    ["Zero-Rule Heuristic", "None (Predict Normal)", "Time-Blocked (70/30)", "25.00%", "0.000", "0.00%", "0.00% (Majority-class baseline)"],
    ["Single-Feature Tree", "CAN ID Only (1 feature)", "Time-Blocked (70/30)", "82.40% +/- 0.12%", "0.781", "89.20%", "3.42% [2.98%, 3.91%]"],
    ["Random Forest (Ablated)", "Payload Only (ID/Timing removed)", "Time-Blocked (70/30)", "91.20% +/- 0.15%", "0.884", "94.60%", "1.84% [1.52%, 2.21%]"],
    ["Random Forest (Ablated)", "Full Features minus Timing", "Time-Blocked (70/30)", "99.85% +/- 0.01%", "0.998", "99.90%", "0.110% [0.045%, 0.225%]"],
    ["Random Forest (Proposed)", "Full 5 Features (ID, DLC, dt, mu, sigma)", "Time-Blocked (70/30)", "99.90% +/- 0.008%", "0.999", "99.92%", "0.091% [0.034%, 0.198%]"]
]
add_tbl(doc, t2_h, t2_r, [1.5, 1.6, 1.3, 1.0, 0.7, 0.9, 1.2], "Table 2", "In-Distribution Benchmark Performance and Feature Ablation on Hyundai Sonata Data")

# Figure 1
fig1_path = RESULTS_DIR / "figure3_hcrl_confusion_matrix.png"
add_fig(doc, fig1_path, "1", "Multi-Class Confusion Matrix of Random Forest on Chronological Time-Blocked HCRL Test Set (Seed 42: 99.90% Accuracy, 0.150% FPR | 5-Seed Mean FPR: 0.091% [95% CI: 0.034%, 0.198%]).", 4.8)

add_h2(doc, "5.2 Signal Isolation: Single-Feature ID Baseline Dissection")
add_p(doc,
    "Arbitration identifiers establish static message routing on automotive broadcast buses. "
    "Static identifier allowlisting drives the benchmark score. "
    "A single-feature decision tree thresholding strictly on raw CAN ID achieves 82.40% +/- 0.12% accuracy and 0.781 MCC (Table 2, Table S1). "
    "Against a four-class uniform chance baseline of 25.00%, this single scalar partition accounts for 76.64% of the total achievable performance gain ((82.40 - 25.00) / (99.90 - 25.00) = 57.40 / 74.90 = 76.64%) attained by the full 5-feature ensemble (99.90%). "
    "Class-specific metrics expose the underlying partition mechanism. "
    "The single-feature tree achieves 100.0% recall on DoS injections by isolating identifier 0x000 and attains 98.3% recall on Fuzzy traffic. "
    "On Spoofing attacks, recall drops to 89.4% with a 10.6% false-positive rate on benign frames sharing identifiers 0x316 and 0x43F."
)

add_h2(doc, "5.3 Timing Feature Ablation and Shortcut Verification")
add_p(doc,
    "Inter-arrival intervals (Delta t_i) capture transmission periodicity across active ECUs. "
    "Completely ablating Delta t_i from the feature set degrades full ensemble accuracy by only 0.05% (from 99.90% to 99.85%, Table 2). "
    "Payload statistics carry orthogonal signal. "
    "Within-session attack recall remains at 100.00% for DoS and Spoofing, while Fuzzy recall shifts by 0.08%. "
    "In stationary captures where injection floods saturate the bus, Delta t_i provides redundant discriminatory information alongside static identifier splits and payload byte distributions. "
    "Decision trees trained on the full feature set still assign split thresholds to Delta t_i and payload variance (mu_byte, sigma_byte) to bound Normal class leaves; across 100 trees, 14.2% of internal splits threshold on Delta t_i in [0.0008, 0.0210] s to delineate nominal cyclic transmissions."
)

add_h2(doc, "5.4 Multi-Detector Architectural Benchmark Across Operational Stress Tests")
add_p(doc,
    "We evaluate four algorithmic families across identical splits to isolate detector paradigms from dataset artifacts: (1) a stateless Timing-Rule baseline thresholding inter-arrival intervals (Delta t < 0.001 s or Delta t > 3 * mu_dt,nom), (2) our Random Forest classifier (100 orthogonal trees, max depth 12), (3) a gradient-boosted decision tree (LightGBM, 100 estimators, learning rate 0.10), and (4) a Deep Reconstruction Autoencoder. "
    "The autoencoder utilizes a PyTorch multi-layer perceptron with an asymmetric 5-to-2 bottleneck trained exclusively on nominal traffic. "
    "Frames exceeding the 99th percentile reconstruction mean squared error threshold are labeled as anomalies. "
    "Four distinct paradigms isolate algorithmic biases."
)
add_p(doc,
    "Within-distribution discrimination on the HCRL dataset reaches near-perfect accuracy for tree ensembles (Table 6, Table S4). "
    "Random Forest achieves 99.96% accuracy (0.9993 MCC, 0.004% FPR, 300 false alarms/hr at 500 kbps), while LightGBM attains 99.31% accuracy (0.9863 MCC, 1.296% FPR, 93,300 false alarms/hr). "
    "The stateless Timing-Rule baseline achieves 85.03% accuracy (0.7231 MCC, 72.65% recall); normal vehicle state transitions drive a 2.592% nominal FPR (186,600 false alarms/hr). "
    "The Reconstruction Autoencoder achieves 73.82% accuracy (0.5582 MCC, 47.77% recall, 0.125% FPR, 9,000 false alarms/hr). "
    "Shallow compression bottlenecks struggle on subtle payload shifts."
)
add_p(doc,
    "Protocol boundaries break every feature-aware architecture. "
    "When transferred directly to the Polaris GEM e6 shuttle without retraining, payload- and identifier-sensitive models experience representation failure (Table 6). "
    "Random Forest records an 86.29% FPR (39.22% accuracy, 0.2118 MCC); LightGBM records a 74.60% FPR (47.22% accuracy, 0.2914 MCC); the Reconstruction Autoencoder yields a 100.00% FPR (29.51% accuracy, -0.0339 MCC). "
    "The stateless Timing-Rule baseline attains 91.49% accuracy and 0.7970 MCC with a 0.02% FPR because 94.6% of the GEM-CAN attack capture consists of high-frequency DoS floods (ID 0x00000000 at ~0.0002 s intervals), which violate timing thresholds regardless of header width. "
    "However, timing rules register 0.00% recall on physical masquerade attacks (Section 8.4). "
    "Cross-platform transfer failures reflect structural identifier bit-width boundaries rather than inductive tree artifacts."
)

# Table 6
t6_h = ["Model Family", "HCRL Accuracy", "HCRL MCC", "HCRL FPR", "Alarms/hr @ 500k", "Raw GEM Transfer FPR", "Calibration GEM MCC"]
t6_r = [
    ["Timing-Rule Baseline", "85.03%", "0.7231", "2.592%", "186,600 / hr", "0.02%", "0.7970"],
    ["Random Forest (Proposed)", "99.96%", "0.9993", "0.004%", "300 / hr", "86.29%", "0.2118"],
    ["LightGBM", "99.31%", "0.9863", "1.296%", "93,300 / hr", "74.60%", "0.2914"],
    ["Reconstruction Autoencoder", "73.82%", "0.5582", "0.125%", "9,000 / hr", "100.00%", "-0.0339"]
]
add_tbl(doc, t6_h, t6_r, [1.6, 1.0, 1.0, 1.0, 1.2, 1.1, 1.1], "Table 6", "Multi-Detector Architectural Benchmark Across Operational Stress Tests and Cross-Platform Transfer")

# ==============================================================================
# 6 OPERATIONAL LINE-SPEED PROJECTIONS & DEBOUNCING
# ==============================================================================
add_h1(doc, "6 Operational Line-Speed False Alarm Projections and Temporal Debouncing")
add_h2(doc, "6.1 Translating Statistical Rates to Operational Bus Frequencies")
add_p(doc,
    "In operational vehicles, CAN classifiers evaluate frames continuously at bus line speeds. "
    "Physical bitrates dictate frame arrival volume. "
    "The expected false alarm volume per driving hour depends directly on transceiver baud rate and instantaneous bus load:"
)
add_formula(doc, "False Alarms per Operational Hour = FPR * R_bus * 3600", "6")
add_p(doc,
    "Percentage metrics obscure real-time bus volume. "
    "On a 500 kbps bus operating at 50% nominal load (~2,000 fps), an evaluation FPR of 0.091% generates 6,552 false alarms per operating hour (Table 3). "
    "This rate produces approximately 1.82 false alerts every second. "
    "At 250 kbps under 40% load (~800 fps assuming 128-bit average frame length with stuff bits), unmitigated false alarms total 2,621 per hour. "
    "At 1 Mbps under heavy 50% bus loading (~4,000 fps), unmitigated false alarms reach 13,104 per hour."
)

# Table 3
t3_h = ["Bus Baud Rate", "Bus Load (%)", "Approx. Bus Rate (fps)", "Hourly Frame Volume", "Unmitigated False Alarms/Hour", "Debounced Alarms/Hour (k=2)", "Debounced Alarms/Hour (k=3)"]
t3_r = [
    ["Low-Speed (250 kbps)", "40% Nominal", "~800 fps", "2,880,000 frames", "2,621 alarms/hr", "2.38 alarms/hr", "0.00217 alarms/hr (~1 per 19.2 days)"],
    ["High-Speed (500 kbps)", "50% Standard", "~2,000 fps", "7,200,000 frames", "6,552 alarms/hr", "5.96 alarms/hr", "0.00543 alarms/hr (~1 per 7.7 days)"],
    ["Classic CAN (1 Mbps)", "50% Heavy", "~4,000 fps", "14,400,000 frames", "13,104 alarms/hr", "11.92 alarms/hr", "0.01085 alarms/hr (~1 per 3.8 days)"]
]
add_tbl(doc, t3_h, t3_r, [1.4, 0.9, 1.1, 1.3, 1.3, 1.2, 1.3], "Table 3", "Operational CAN Bus Line-Speed False Alarm Projections and Temporal Debounce Suppression")

add_h2(doc, "6.2 Temporal Debounce Filtering and Independence Assumptions")
add_p(doc,
    "Automotive gateway microcontrollers can filter isolated false positives using a k-out-of-N sliding window debounce filter. "
    "Multi-frame filtering suppresses isolated spikes. "
    "Under an assumption of independent Bernoulli classification errors, the probability of asserting a false alarm across k consecutive frames equals P_debounce = (FPR)^k. "
    "At 500 kbps, a two-frame filter (k = 2) reduces theoretical false alarms to 5.96 per hour. "
    "A three-frame filter (k = 3) suppresses theoretical false alarms to 0.00543 per hour, which equals one alarm every 7.7 operating days (Table 3, Table S4)."
)
add_p(doc,
    "Detection latency introduces an operational trade-off across window configurations. "
    "Over a whole-bus window at 2,000 fps, a k = 3 filter incurs 1.5 ms of detection latency (3 / 2,000 fps). "
    "Per-ID filtering multiplies the delay. "
    "When debounce windows track individual arbitration identifiers, latency scales with message broadcast intervals; an ECU broadcasting every 20 ms incurs 60 ms of detection latency, with delays spanning 30 to 300 ms across slower body control loops."
)

add_h2(doc, "6.3 Empirical Measurement of Error Burstiness in Real Driving")
add_p(doc,
    "Physical CAN controllers operating under ISO 11898-1 automatically execute hardware CRC-15 and bit-stuffing verification. "
    "Transceiver hardware filters line noise first. "
    "Frames corrupted by electromagnetic interference or electrical transients fail CRC checks and trigger physical Error Frames; transceivers discard these frames at the silicon layer before host software receives them. "
    "False positives in production software originate instead from valid vehicle dynamic maneuvers (such as abrupt braking, rapid steering corrections, and sharp acceleration transients) where legitimate sensor telemetry deviates temporarily from steady-state statistics."
)
add_p(doc,
    "We evaluate debounce performance against 200,000 continuous, un-subsampled normal driving frames from the HCRL benchmark (normal_run_data.txt, spanning 102.42 seconds at 1,952.6 fps). "
    "Dynamic maneuvers violate baseline boundaries. "
    "Raw unmitigated classification (k = 1) yields 22,136 false positives, representing an 11.068% single-frame error rate and 778,029 alarms/hr. "
    "This single-frame FPR exceeds the 0.091% test FPR observed in Section 5.1 because normal_run_data.txt captures an independent driving session where inter-arrival intervals (Delta t_i) and dynamic payload bytes drift beyond the tight normal leaf boundaries identified in Section 5.3 (Delta t_i in [0.0008, 0.0210] s, mu_byte, sigma_byte)."
)
add_p(doc,
    "Per-ID debouncing fails under dynamic driving. "
    "Under an i.i.d. binomial model at p = 11.068%, the expected number of 3-frame false alarms is 200,000 * (0.11068)^3 = 271 alerts. "
    "Whole-bus debouncing at k = 3 suppresses false alarms to 53 alerts (a 99.76% reduction, leaving 1,863 alarms/hr); this represents a 5-fold suppression below the independent expectation because classification errors disperse across multiplexed ECU identifiers. "
    "Under per-ID filtering, a k = 3 window yields 14,514 false alerts (510,134 alarms/hr). "
    "This volume exceeds the independent binomial prediction by 53.5-fold because transient sensor states induce persistent error runs on specific arbitration IDs."
)

# ==============================================================================
# 7 STEALTH ATTACK STRESS TEST
# ==============================================================================
add_h1(doc, "7 Robustness Stress Tests: Feature-Space Stealth Camouflage")
add_h2(doc, "7.1 Motivation and Blending Formulation")
add_p(doc,
    "Adversaries executing masquerade injections can blend malicious payload bytes toward nominal distributions to bypass statistical detectors [3], [16]. "
    "Stealth evasion shifts malicious distributions. "
    "We parameterize this evasion using a feature-space camouflage model that shifts attack feature vectors x_attack toward nominal centroids by camouflage ratio alpha in [0.0, 1.0]:"
)
add_formula(doc, "x_camouflaged(alpha) = (1 - alpha) * x_attack + alpha * mu_normal", "7")
add_p(doc,
    "Here, mu_normal denotes the nominal centroid. "
    "This vector represents the empirical mean of benign driving traffic in the held-out test partition."
)

# Table 4
t4_h = ["Camouflage Level (alpha)", "Attack Modification", "Attack Recall (%)", "Precision (%)", "Macro F1-Score", "Matthews Corr. Coeff. (MCC)"]
t4_r = [
    ["0% (Raw Benchmark)", "Unmodified HCRL injections", "99.90%", "99.91%", "0.999", "0.999"],
    ["25% Camouflage", "Slight payload/timing blending", "41.80%", "98.10%", "0.623", "0.501"],
    ["50% Camouflage", "Moderate statistical drift", "33.10%", "97.40%", "0.551", "0.428"],
    ["75% Camouflage", "Heavy distribution alignment", "31.40%", "96.80%", "0.534", "0.406"],
    ["100% Camouflage", "Centroid payload alignment", "19.60%", "95.50%", "0.412", "0.252"]
]
add_tbl(doc, t4_h, t4_r, [1.5, 1.8, 1.2, 1.0, 1.0, 1.3], "Table 4", "Feature-Space Stealth Camouflage Stress Test Performance Across Blending Levels")

add_p(doc,
    "Detection recall drops under feature blending across all attack classes (Table 4, Figure 2). "
    "Under a slight 25% blend (alpha = 0.25), attack recall drops from 99.90% to 41.80% (0.501 MCC). "
    "Moderate blending at alpha = 0.50 reduces recall to 33.10% (0.428 MCC), while alpha = 0.75 yields 31.40% (0.406 MCC). "
    "At full alignment (alpha = 1.00), recall declines to 19.60% with an MCC of 0.252. "
    "Axis-aligned tree splits fail under centroid blending. "
    "Decision tree boundaries established on dense injection clusters lose discriminative power as malicious vectors converge toward nominal manifolds."
)

# Figure 2
fig2_path = RESULTS_DIR / "figure5_stealth_attack_degradation.png"
add_fig(doc, fig2_path, "2", "Degradation of Attack Detection Recall and Matthews Correlation Coefficient Under Feature-Space Camouflage Blending.", 5.0)

add_h2(doc, "7.2 Camouflage Curriculum Retraining and the Evasion-False Alarm Trade-off")
add_p(doc,
    "We implement a leakage-free curriculum retraining pipeline to evaluate whether supervised tree ensembles can adapt to continuous adversarial feature drift. "
    "Three safeguards prevent data leakage across training and evaluation partitions. "
    "Curriculum training samples incorporate only partial blending levels alpha in {0.25, 0.50}; distributions at alpha = 0.75 and alpha = 1.00 remain strictly unseen during training. "
    "Camouflaged vectors are synthesized by pairing malicious frames exclusively with nominal frames sharing identical arbitration IDs. "
    "Evaluations execute strictly on held-out chronological test windows."
)
add_p(doc,
    "Curriculum retraining recovers substantial detection capability across moderate blending levels (Table 7). "
    "The retrained model adapts rapidly. "
    "At alpha = 0.25, attack recall increases from 41.80% to 98.60% (+56.80% recovery gain), while at alpha = 0.50, recall increases from 33.10% to 95.40% (+62.30% recovery gain). "
    "On the unseen alpha = 0.75 distribution, the retrained ensemble maintains 48.20% recall (+16.80% gain over the 31.40% baseline). "
    "At alpha = 1.00, detection recall reaches 21.80% for the curriculum model versus 19.60% for the baseline; this limit is bounded by Bayes error because malicious feature vectors collapse entirely onto the empirical nominal centroid."
)
add_p(doc,
    "Expanding those decision boundaries carries a steep penalty. "
    "Intercepting blended feature vectors forces the model's nominal false-positive rate from 0.091% (6,552 alarms/hr at 500 kbps) to 0.246% (17,700 alarms/hr at 500 kbps), representing a 2.7-fold increase in false alerts during benign driving (Table 7, Table S7). "
    "On an operational CAN bus, this alert volume equates to 4.92 false alarms every second."
)

# Table 7
t7_h = ["Camouflage Level (alpha)", "Standard Model Recall", "Curriculum Retrained Recall", "Evasion Recovery Gain", "Standard FPR (Alarms/hr)", "Curriculum FPR (Alarms/hr)"]
t7_r = [
    ["alpha = 0.00 (Raw Attack)", "99.90%", "99.92%", "+0.02%", "0.091% (6,552 / hr)", "0.246% (17,700 / hr)"],
    ["alpha = 0.25 (Slight Blend)", "41.80%", "98.60%", "+56.80%", "—", "—"],
    ["alpha = 0.50 (Moderate Blend)", "33.10%", "95.40%", "+62.30%", "—", "—"],
    ["alpha = 0.75 (Heavy Blend)", "31.40%", "48.20%", "+16.80%", "—", "—"],
    ["alpha = 1.00 (Centroid Match)", "19.60%", "21.80%", "+2.20% (Bayes Bound)", "—", "—"]
]
add_tbl(doc, t7_h, t7_r, [1.5, 1.2, 1.3, 1.3, 1.3, 1.4], "Table 7", "Leakage-Free Camouflage Curriculum Retraining Performance Across Evasion Blends and Nominal False Alarm Trade-off")
# ==============================================================================
# 8 CROSS-PLATFORM TRANSFERABILITY
# ==============================================================================
add_h1(doc, "8 Cross-Platform Transferability: Sonata to Polaris GEM e6 Shuttle")
add_h2(doc, "8.1 Experimental Transfer Protocol and Ground-Truth Reconciliation")
add_p(doc,
    "We evaluate cross-platform generalization using the GEM-CAN benchmark (Tavasoli et al., 2026 [6]), captured from an instrumented Polaris GEM e6 autonomous campus shuttle. "
    "Vehicle platforms differ fundamentally in network architecture. The capture comprises 143,444 frames partitioned across two operational logs. "
    "The nominal run log contains 100,000 frames governed exclusively by 29-bit extended SAE J1939 identifiers across 13 unique Parameter Group Numbers (PGNs, such as 0x18FF00F9 and 0x18FF70F9). "
    "The adversarial log contains 43,444 frames. Standard 11-bit CAN identifiers are completely absent from nominal shuttle operation. "
    "Within the attack trace, an adversary executes high-frequency bus floods alongside targeted physical actuator manipulation: 41,080 frames execute a zero-identifier (0x00000000) DoS flood, "
    "663 frames tamper with brake controller commands on PGN 0x18FF00F9, 662 frames alter steering angle telemetry on PGN 0x18FF70F9, and 1,039 nominal frames interleave throughout the sequence."
)
add_p(doc,
    "Unified ground-truth reconciliation establishes 101,039 nominal frames and 42,405 adversarial frames across the complete shuttle telemetry corpus (41,080 DoS frames alongside 1,325 physical tampering frames). "
    "The target corpus enables rigorous cross-architecture evaluation. This balance provides an uncorrupted benchmark for empirical transferability audits."
)

add_h2(doc, "8.2 Cross-Platform Transfer Performance and Feature Isolation")
add_p(doc,
    "Direct evaluation of the Sonata-trained Random Forest model on raw GEM e6 shuttle telemetry yields an immediate collapse: 35.57% test accuracy, 0.1637 MCC, and a 91.48% false-positive rate on nominal shuttle driving (Table 5). "
    "Out of 101,039 benign frames, the classifier flags 92,427 as malicious attacks. The detector flags almost everything. "
    "Perfect nominal recall on DoS (100.00%) and physical tampering (100.00%) is an artifact of universal anomaly triggering rather than learned feature representation."
)

# Table 5
t5_h = ["Experiment Configuration", "Source Vehicle", "Target Vehicle", "Feature Representation", "Accuracy", "MCC", "FPR", "DoS Recall (n=41,080)", "Tamper Recall (n=1,325)"]
t5_r = [
    ["Direct Transfer (Raw IDs)", "Sonata (11-bit)", "GEM e6 (29-bit)", "Raw IDs, DLC, dt, mu, sigma", "35.57%", "0.1637", "91.48%", "100.00%", "100.00%"],
    ["Payload-Only Transfer", "Sonata (11-bit)", "GEM e6 (29-bit)", "DLC, Byte Mean, Byte Std Dev", "53.62%", "0.3646", "65.85%", "100.00%", "100.00%"],
    ["Calibration-Normalized (Proposed)", "Sonata (11-bit)", "GEM e6 (29-bit)", "Allowlist, Rank, dt_ratio, z_mu, z_sig", "92.04%", "0.8316", "10.65%", "100.00%", "50.04%"],
    ["Native Within-GEM Baseline", "GEM e6 (29-bit)", "GEM e6 (29-bit)", "Raw IDs, DLC, dt, mu, sigma (70/30)", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%"]
]
add_tbl(doc, t5_h, t5_r, [1.5, 1.0, 1.0, 1.5, 0.7, 0.6, 0.6, 1.0, 1.0], "Table 5", "Reconciled Cross-Platform Transferability, Feature Isolation, and Native Baseline Performance")

# Figure 3
fig3_path = RESULTS_DIR / "figure4_gem_can_cross_platform_confusion.png"
add_fig(doc, fig3_path, "3", "Cross-Platform Transfer Confusion Matrix (Sonata-Trained Model Evaluated on Polaris GEM e6 Telemetry: 35.57% Accuracy, 91.48% FPR, 0.1637 MCC).", 4.8)

add_h2(doc, "8.3 Root Cause Analysis and Protocol Normalization")
add_p(doc,
    "Direct transfer collapse originates from an architectural representation mismatch between standard CAN 2.0A (11-bit identifiers, numeric range 0–2,047) and extended CAN 2.0B / SAE J1939 frames (29-bit identifiers, numerical values exceeding 400,000,000). "
    "Decision trees trained on 11-bit integer splits route every 29-bit identifier into high-index leaves assigned to Fuzzy attack clusters during training. Extended 29-bit headers break every tree split. "
    "Eliminating identifiers entirely via payload-only features yields 53.62% test accuracy and a 65.85% false-positive rate (Table 5). Payload distributions cannot generalize across disparate powertrains. "
    "Electric shuttle motor controllers produce radically different byte patterns than internal combustion passenger vehicles."
)
add_p(doc,
    "To establish protocol invariance without proprietary OEM database (.dbc) files, we construct a Calibration-Normalized Stateful Encoding that standardizes incoming telemetry against a short nominal calibration window. "
    "Five statistical transformations enforce protocol invariance. First, an Allowlist Membership Indicator I_allow,i in {0, 1} evaluates whether frame identifier ID_i occurred during the nominal baseline calibration set A_nom. "
    "Second, a Normalized Frequency Rank R_freq,i = Rank(ID_i) / |A_nom| projects message broadcast intensity onto the unit interval according to transmission frequency during nominal driving. "
    "Third, a Dimensionless Timing Ratio r_dt,i = Delta t_i / mu_dt,nom(ID_i) normalizes inter-arrival intervals against each identifier empirical mean period mu_dt,nom(ID_i). "
    "Fourth, a Z-Scored Byte Mean z_mu,i = (mu_byte,i - mu_nom(ID_i)) / (sigma_nom(ID_i) + epsilon) standardizes payload central tendency against nominal expectations (epsilon = 10^-6). "
    "Fifth, a Z-Scored Byte Dispersion z_sig,i = (sigma_byte,i - bar_sigma_nom(ID_i)) / (Var_nom(ID_i)^(1/2) + epsilon) captures payload variance shifts."
)
add_p(doc,
    "Statistical standardization transforms raw telemetry into cross-platform signals. Evaluating this calibration-normalized feature set under zero-shot transfer lifts classification accuracy to 92.04%, increases MCC to 0.8316, and curtails the cross-platform false-positive rate from 91.48% down to 10.65% (Table 5). "
    "The normalized model maintains 100.00% recall on DoS flooding and isolates 50.04% of physical actuator tampering. Gross deviations in brake payload bytes trigger detection while stealthy steering adjustments remain undetected. "
    "Native retraining on a 70/30 time-blocked partition of GEM e6 telemetry achieves 100.00% accuracy, 1.0000 MCC, and 0.00% false-positive rate across both DoS and physical tampering attacks (Table 5, Figure 4)."
)

# Figure 4
fig4_path = RESULTS_DIR / "figure6_cross_platform_comparison.png"
add_fig(doc, fig4_path, "4", "Performance Comparison Across Native In-Distribution, Direct Transfer (Raw IDs), Payload-Only Transfer, Calibration-Normalized Encoding, and Native GEM Baseline.", 5.2)

add_h2(doc, "8.4 Validation Against Physical Masquerade Attacks (ORNL ROAD Benchmark)")
add_p(doc,
    "We evaluate supervised Random Forest and LightGBM architectures using a strict chronological 70/30 time-blocked split on correlated_signal_attack_1_masquerade.csv (74,152 frames: 72,066 nominal driving frames, 2,086 physically injected masquerade attack frames) from the Oak Ridge National Laboratory (ORNL) ROAD dataset [16]. "
    "In this trace, an adversary overrides safety-critical drivetrain signals while maintaining the nominal 20 ms and 50 ms broadcast periods and target ECU arbitration identifiers. The injected traffic introduces zero timing disruption."
)
add_p(doc,
    "Stateless timing-rule baselines register 0.00% recall (-0.0042 MCC, 0.056% FPR), as reported in Table 8. Timing baselines miss every single masquerade frame. "
    "Because the attacker injects frames at legitimate periodic intervals, inter-arrival periods remain within expected thresholds. In contrast, multi-feature Random Forest and LightGBM models achieve 100.00% recall, 0.9993 MCC, and 0.005% FPR (Table 8). "
    "Tree ensembles split on non-linear payload byte relationships between engine torque and vehicle velocity without requiring reverse-engineered OEM database files."
)

# Table 8
t8_h = ["Detector Paradigm", "ROAD Masquerade Recall", "Nominal Driving FPR", "Matthews Corr. Coeff. (MCC)", "Operational Detection Mechanism"]
t8_r = [
    ["Timing-Rule Baseline", "0.00%", "0.056%", "-0.0042", "Blind: Inter-arrival periodicity strictly preserved by attacker"],
    ["Random Forest (Proposed)", "100.00%", "0.005%", "0.9993", "Payload correlation: Identifies non-linear drivetrain byte shifts"],
    ["LightGBM", "100.00%", "0.005%", "0.9993", "Gradient splits: Isolates correlated speed/torque signal deviations"]
]
add_tbl(doc, t8_h, t8_r, [1.5, 1.2, 1.1, 1.2, 2.8], "Table 8", "Empirical Validation Against Physical Masquerade Attacks on the ORNL ROAD Benchmark")

add_h2(doc, "8.5 Supervised Few-Shot Domain Adaptation on the Polaris GEM e6")
add_p(doc,
    "Production automotive pipelines allow brief telemetry calibration windows during end-of-line vehicle assembly or scheduled dealer service intervals. Target-domain convergence requires just one second. "
    "We trace supervised adaptation dynamics across the held-out 30% test partition of GEM e6 telemetry by incrementing labeled target-domain adaptation frames N from N = 0 (zero-shot transfer) to N = 10,000 (representing 20.0 seconds of vehicle operation at 500 fps)."
)
add_p(doc,
    "Zero-shot transfer with calibration-normalized features (N = 0) achieves 92.04% test accuracy, 0.8316 MCC, and 10.65% FPR, while detecting 100.00% of DoS frames and 50.04% of physical tampering frames (Table 9, Table S8). "
    "Allocating N = 100 stratified adaptation frames (0.20 seconds of driving) suppresses the false-positive rate to 0.73%, lifts tampering recall to 86.40%, and yields an MCC of 0.9877. "
    "With N = 500 frames (1.0 second of driving), the false-positive rate drops to 0.04%, overall test accuracy reaches 99.96%, tampering recall improves to 95.72%, and MCC reaches 0.9991. "
    "At N >= 1,000 labeled frames (2.0 seconds of driving), the classifier achieves 100.00% test accuracy, 1.0000 MCC, 0.00% FPR, 100.00% DoS recall, and 100.00% tampering recall. Domain alignment converges within brief observation windows."
)

# Table 9
t9_h = ["Labeled Adaptation Frames (N)", "Driving Time Equivalent", "Test Accuracy (Held-out 30%)", "Test MCC", "False Positive Rate", "DoS Recall (n=12,342)", "Tampering Recall (n=397)"]
t9_r = [
    ["N = 0 (Zero-Shot)", "0.0 s", "92.04%", "0.8316", "10.65%", "100.00%", "50.04%"],
    ["N = 100", "0.2 s", "99.48%", "0.9877", "0.73%", "100.00%", "86.40%"],
    ["N = 500", "1.0 s", "99.96%", "0.9991", "0.04%", "100.00%", "95.72%"],
    ["N = 1,000", "2.0 s", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%"],
    ["N = 2,500", "5.0 s", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%"],
    ["N = 5,000", "10.0 s", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%"],
    ["N = 10,000", "20.0 s", "100.00%", "1.0000", "0.00%", "100.00%", "100.00%"]
]
add_tbl(doc, t9_h, t9_r, [1.5, 1.1, 1.2, 0.9, 1.0, 1.1, 1.2], "Table 9", "Supervised Few-Shot Domain Adaptation Performance on Polaris GEM e6 Telemetry")
p_t9_fn = doc.add_paragraph()
r_fn = p_t9_fn.add_run("*Note: Labeled adaptation frames are sampled across scenario regimes within the 70% target training partition. While zero-shot calibration-normalized features detect 50.04% of physical tampering, few-shot adaptation achieves 86.40% tampering recall at N = 100 frames and full 100.00% recall at N >= 1,000 frames.")
r_fn.italic = True
set_run(r_fn, 9.0, (0x55, 0x55, 0x55))

# ==============================================================================
# 9 ENGINEERING & REGULATORY IMPLICATIONS
# ==============================================================================
add_h1(doc, "9 Engineering and Regulatory Implications: Safety vs. Security Tension")
add_h2(doc, "9.1 Interaction Between ISO 26262 and ISO/SAE 21434")
add_p(doc,
    "Automotive systems operate under competing functional safety and cybersecurity mandates [7], [9]. Safety and security mandates collide on live buses. "
    "Functional safety under ISO 26262 requires deterministic fail-safe behavior: when a safety-critical subsystem experiences a fault, control logic must transition vehicle actuators into a predictable safe state through limp-home modes or emergency deceleration [7]. "
    "In parallel, ISO/SAE 21434 governs cybersecurity engineering and mandates fail-secure defenses. Under this standard, electronic architectures must detect unauthorized intrusions, isolate compromised nodes, and reject forged command frames [9]."
)
add_p(doc,
    "Granting machine learning models unilateral actuation authority creates an acute operational hazard. When a detector generates hundreds or thousands of false alarms per hour, an IDS wired directly into vehicular control controllers becomes an operational hazard. "
    "An adversary can inject trivial bus noise to induce deliberate false positives. Under such conditions, an automated system would execute emergency stops, disable steer-by-wire subsystems, or disconnect critical ECUs during highway transit [7]. "
    "False alarms directly compromise physical vehicle safety."
)

add_h2(doc, "9.2 Deployment Architecture: Strict Two-Tier Authority Separation")
add_p(doc,
    "To decouple safety-critical actuation from statistical classifier volatility, vehicle electronic architectures must enforce a strict two-tier authority separation. Single-frame ML predictions cannot command physical actuators. "
    "Under Tier 1, machine learning classifiers execute strictly as passive telemetry sentinels hosted on central gateway microcontrollers (MCUs). These models maintain zero unilateral authority over physical steering mechanisms, friction brakes, or powertrain controllers. "
    "Gateway sentinels log cryptographically signed security audit entries to non-volatile memory and route diagnostic alert telemetry to off-board vehicle security operation centers over isolated telematics links [14]."
)
add_p(doc,
    "Under Tier 2, physical safety interventions require multi-sensor kinematic consensus across independent physical measurements. Actuator commands—such as emergency deceleration, torque compensation, or high-voltage contactor isolation—must never trigger on frame-level CAN anomalies alone. "
    "Anomaly notifications must correlate with physical state deviations captured by wheel-speed Hall-effect encoders, 6-axis inertial measurement units (measuring lateral acceleration and yaw rates), and optical vehicle odometry [7]. "
    "Tier 2 supervisory controllers combine multi-modal physical consensus with temporal debounce filtering (k >= 3) before activating protective vehicle maneuvers. Cross-modal validation protects passenger safety."
)

# ==============================================================================
# 10 THREATS TO VALIDITY & LIMITATIONS
# ==============================================================================
add_h1(doc, "10 Threats to Validity and Limitations")
add_p(doc,
    "Three methodological boundaries frame the empirical findings of this audit. Internal validity concerns the synthetic abstraction of feature-space stealth perturbations. Feature-space perturbations do not capture physical transceiver constraints. "
    "Physical adversaries executing bus-level attacks must satisfy strict hardware bit-timing margins, arbitration phase synchronizations, and cyclic redundancy check (CRC) generation on bare physical transceivers [16]. "
    "Evaluating stealth evasion purely within statistical feature dimensions provides an analytical lower bound on classifier brittleness, though physical bus contention may restrict attacker injection flexibility."
)
add_p(doc,
    "External validity pertains to vehicle platform and protocol diversity. Our evaluation spans three distinct operational platforms: a Hyundai Sonata passenger car, an autonomous Polaris GEM e6 electric shuttle, and real physical masquerade captures from the ORNL ROAD benchmark. "
    "Vehicle network architectures continue to evolve. Future research must evaluate detector robustness across CAN Flexible Data-Rate (CAN FD) networks and high-bandwidth automotive Ethernet backbones (100BASE-T1) that feature larger payloads and switched topologies [23]."
)
add_p(doc,
    "Construct validity addresses the architectural coverage of our evaluation suite. We benchmarked four distinct detector paradigms in Section 5.4: stateless timing rules, random forests, gradient boosted trees, and deep reconstruction autoencoders. "
    "Protocol mismatches transcend specific algorithm families. Cross-platform transfer failure stems from systemic protocol differences rather than inductive biases specific to decision trees. "
    "Emerging detection paradigms—specifically sequence transformers and neuromorphic spiking architectures—warrant extensive field validation under operational bus constraints [17]."
)

# ==============================================================================
# 11 CONCLUSION
# ==============================================================================
add_h1(doc, "11 Conclusion and Future Research Directions")
add_p(doc,
    "Offline benchmark accuracy does not guarantee road readiness. Our empirical deployability audit reveals that reported benchmark accuracies exceeding 99.9% mask fundamental operational vulnerabilities in vehicle-layer AI detectors. "
    "A single-feature identifier baseline reaches 82.40% accuracy and captures 76.64% of the gain above chance. This baseline establishes heavy identifier memorization. "
    "In-distribution false-positive rates of 0.091% generate 2,621 to 13,104 false alarms per hour across 250 kbps to 1 Mbps line speeds. "
    "Temporal debounce filtering (k = 3) achieves theoretical suppression (one alarm every 7.7 days at 500 kbps), yet bursty vehicle errors in real driving generate 1,863 false alarms per hour (and 510,134 alarms per hour under per-identifier filtering). "
    "Feature-space camouflage reduces attack recall to 19.60%; curriculum retraining recovers 98.60% recall at 25% camouflage, but expands decision boundaries to force a 2.7-fold increase in nominal false alarms (0.246% FPR, 17,700 alarms per hour at 500 kbps)."
)
add_p(doc,
    "Protocol differences cause severe transfer collapse. Cross-platform domain shifts induce representation collapse across multiple detector architectures, driving false-positive rates to 91.48% on unseen 29-bit shuttle telemetry. "
    "Calibration-normalized stateful encoding restores zero-shot MCC to 0.8316 and suppresses false positives to 10.65%. "
    "Stratified few-shot adaptation achieves 0.04% FPR at N = 500 frames (1.0 s) and converges to 0.00% FPR and 100.00% tampering recall at N >= 1,000 frames (2.0 s of driving). "
    "On real physical masquerade attacks from the ORNL ROAD benchmark, stateless timing baselines collapse to 0.00% recall, whereas payload-aware tree ensembles maintain 100.00% detection. "
    "Vehicle-layer AI detectors must function strictly as passive telemetry sentinels within a two-tier architecture. "
    "These models must withhold autonomous physical actuation authority unless corroborated by multi-sensor kinematic consensus under ISO 26262 and ISO/SAE 21434."
)

# ==============================================================================
# DECLARATIONS
# ==============================================================================
add_h1(doc, "Declarations")
add_p(doc,
    "Data and Code Availability: The HCRL Car-Hacking dataset is publicly available from Korea University (https://ocslab.hksecurity.net). "
    "The GEM-CAN dataset is available on Zenodo (doi: 10.5281/zenodo.19161139). The ROAD dataset is available from Oak Ridge National Laboratory. "
    "Evaluation scripts and replication materials are archived on GitHub."
)
add_p(doc,
    "Funding: This research received no specific grant from any funding agency in the public, commercial, or not-for-profit sectors."
)
add_p(doc,
    "Competing Interests: The authors declare no competing financial or personal interests that could have appeared to influence the work reported in this paper."
)
add_p(doc,
    "Author Contributions: V. Khandelwal: Conceptualization, Methodology, Software, Validation, Formal Analysis, Investigation, Writing – Original Draft, Writing – Review & Editing."
)
add_p(doc,
    "AI-Use Disclosure: Generative AI tools were utilized for formatting code assistance, figure generation styling, and language polishing. All analytical formulations, experimental designs, and empirical findings were conducted and verified by the authors."
)

# ==============================================================================
# 12 REFERENCES
# ==============================================================================
add_h1(doc, "References")
references_list = [
    "[1] HCRL, 'Car-Hacking Dataset for Intrusion Detection,' Information Security Applications Lab, Korea University, 2018. [Online]. Available: https://ocslab.hksecurity.net/Datasets/car-hacking-dataset",
    "[2] J. Akshya, M. Sundarrajan, R. Vijayakumar, and A. Nayyar, 'Explainable AI-driven intrusion detection for securing IoT-enabled autonomous transportation systems,' Cluster Computing, vol. 28, art. 884, 2025. doi: 10.1007/s10586-025-05617-1.",
    "[3] F. Aloraini, A. Javed, and O. Rana, 'Adversarial attacks on intrusion detection systems in in-vehicle networks of connected and autonomous vehicles,' Sensors, vol. 24, no. 12, art. 3848, 2024. doi: 10.3390/s24123848.",
    "[4] M. Althunayyan, A. Javed, and O. Rana, 'A robust multi-stage intrusion detection system for in-vehicle network security using hierarchical federated learning,' Vehicular Communications, vol. 49, art. 100837, 2024. doi: 10.1016/j.vehcom.2024.100837.",
    "[5] K. G. Arachchige, G. Alkaabi, M. Murtaza, Q. E. U. Haq, A. Z. Abualkishik, and C.-C. Lee, 'Threat landscape and integrated cybersecurity framework for V2V and autonomous electric vehicles,' World Electric Vehicle Journal, vol. 16, no. 8, art. 469, 2025. doi: 10.3390/wevj16080469.",
    "[6] M. Tavasoli, M. Rostami, H. Kebriaei, and V. Puig, 'GEM-CAN: A Real-World Dataset of CAN-Bus Attack Scenarios On An Autonomous Vehicle for Intrusion-Detection Research,' Data in Brief, vol. 66, art. 112805, 2026. doi: 10.1016/j.dib.2026.112805. Zenodo: 10.5281/zenodo.19161139.",
    "[7] ISO 26262:2018, 'Road Vehicles — Functional Safety,' International Organization for Standardization, Geneva, Switzerland, 2018.",
    "[8] ISO/IEC 5962:2021, 'Information Technology — Software Package Data Exchange (SPDX) Specification,' International Organization for Standardization, Geneva, Switzerland, 2021.",
    "[9] ISO/SAE 21434:2021, 'Road Vehicles — Cybersecurity Engineering,' International Organization for Standardization / SAE International, Geneva, Switzerland, 2021.",
    "[10] V. Khandelwal, 'Building Trustworthy AI Systems: Developing Explainable Models for Transparent Decision-Making in Autonomous Vehicles,' Journal of Sustainable Solutions, vol. 1, no. 4, pp. 27–37, 2024.",
    "[11] K. Koscher, A. Czeskis, F. Roesner, S. Patel, T. Kohno, S. Checkoway, D. McCoy, B. Kantor, D. Anderson, H. Shacham, and S. Savage, 'Experimental security analysis of a modern automobile,' in Proceedings of the 2010 IEEE Symposium on Security and Privacy (SP), 2010, pp. 447–462. doi: 10.1109/SP.2010.34.",
    "[12] B. Lampe and W. Meng, 'can-train-and-test: A Curated CAN Dataset for Automotive Intrusion Detection,' arXiv preprint arXiv:2308.04972, 2023.",
    "[13] C. Miller and C. Valasek, 'Remote Exploitation of an Unaltered Passenger Vehicle,' Black Hat USA Technical White Paper, IOActive, 2015.",
    "[14] A. Oseni, N. Moustafa, G. Creech, N. Sohrabi, A. Strelzoff, Z. Tari, and I. Linkov, 'An Explainable Deep Learning Framework for Resilient Intrusion Detection in IoT-Enabled Transportation Networks,' IEEE Transactions on Intelligent Transportation Systems, vol. 24, no. 1, pp. 1000–1014, Jan. 2023. doi: 10.1109/TITS.2022.3188671.",
    "[15] F. Pedregosa, G. Varoquaux, A. Gramfort, V. Michel, B. Thirion, O. Grisel, M. Blondel, P. Prettenhofer, R. Weiss, V. Dubourg, J. Vanderplas, A. Passos, D. Cournapeau, M. Brucher, M. Perrot, and E. Duchesnay, 'Scikit-learn: Machine learning in Python,' Journal of Machine Learning Research, vol. 12, pp. 2825–2830, 2011.",
    "[16] M. E. Verma, M. D. Iannacone, R. A. Bridges, S. C. Hollifield, P. Kay, and F. L. Combs, 'Addressing the lack of comparability & testing in CAN intrusion detection research: The ROAD dataset,' Computers & Security, vol. 125, art. 103027, 2023. doi: 10.1016/j.cose.2022.103027.",
    "[17] S. Rajapaksha, H. Kalutarage, M. O. Al-Kadri, A. Petrovski, G. Madzudzo, and M. Cheah, 'AI-based intrusion detection systems for in-vehicle networks: A survey,' ACM Computing Surveys, vol. 55, no. 11, art. 237, 2023. doi: 10.1145/3570954.",
    "[18] H. M. Song, H. R. Kim, and H. K. Kim, 'Intrusion detection system based on the analysis of time intervals of CAN messages for in-vehicle network,' in Proceedings of the 2016 International Conference on Information Networking (ICOIN), 2016, pp. 63–68. doi: 10.1109/ICOIN.2016.7427086.",
    "[19] R. C. Shit and S. Subudhi, 'Scalable Hierarchical AI-Blockchain Framework for Real-Time Anomaly Detection in Large-Scale Autonomous Vehicle Networks,' arXiv preprint arXiv:2511.12648, 2025.",
    "[20] S. Tariq, S. Lee, Y. Shin, M. S. Lee, O. Y. Ryu, and K. Chung, 'CAN-ADF: An Anomaly Detection Framework for Controller Area Networks,' IEEE Transactions on Vehicular Technology, vol. 70, no. 6, pp. 5240–5253, Jun. 2021. doi: 10.1109/TVT.2021.3073740.",
    "[21] UNECE, 'UN Regulation No. 155: Cyber Security and Cyber Security Management System,' United Nations Economic Commission for Europe, World Forum for Harmonization of Vehicle Regulations, Geneva, Switzerland, 2021.",
    "[22] R. W. van der Heijden, T. Lukaseder, and F. Kargl, 'VeReMi: A Dataset for Comparable Evaluation of Misbehavior Detection in VANETs,' arXiv preprint arXiv:1804.06701, 2018.",
    "[23] Z. Ying, R. Jiang, Y. Liu, and K. Chen, 'A literature review on V2X communications security: Foundation, solutions, status, and future,' IET Communications, vol. 18, no. 20, pp. 1421–1439, 2024. doi: 10.1049/cmu2.12778.",
    "[24] S. Zhao, L. Zhang, Y. Wang, X. Liu, and H. Chen, 'Securing the future: AI-driven cybersecurity in the age of autonomous IoT,' Frontiers in the Internet of Things, vol. 4, art. 1658273, 2025. doi: 10.3389/friot.2025.1658273."
]

for ref in references_list:
    p_r = doc.add_paragraph()
    p_r.paragraph_format.left_indent = Inches(0.3)
    p_r.paragraph_format.first_line_indent = Inches(-0.3)
    p_r.paragraph_format.space_before = Pt(0)
    p_r.paragraph_format.space_after = Pt(4)
    p_r.paragraph_format.line_spacing = 1.05
    r = p_r.add_run(ref)
    set_run(r, 9.5, (0x33, 0x33, 0x33))

# ==============================================================================
# SAVE DOCUMENT DELIVERABLES
# ==============================================================================
output_v21_full = OUTPUT_DIR / "draft_v21_full_journal.docx"
output_v21_short = OUTPUT_DIR / "draft_v21.docx"
output_v21_root = OUTPUT_DIR / "Vehicular_Communications_Manuscript_v21.docx"
output_v20_root = OUTPUT_DIR / "Vehicular_Communications_Manuscript_v20.docx"

doc.save(str(output_v21_full))
doc.save(str(output_v21_short))
doc.save(str(output_v21_root))
doc.save(str(output_v20_root))
print(f"Generated Draft v21 Journal Manuscript: {output_v21_full}")
print(f"Generated Root Manuscript: {output_v21_root}")