import re
import docx
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name in ["scripts", "2_code_and_scripts"] else SCRIPT_DIR
doc_path = PROJECT_ROOT / "drafts" / "Vehicular_Communications_Manuscript_v21.docx"
if not doc_path.exists():
    doc_path = PROJECT_ROOT / "1_manuscripts_and_drafts" / "Vehicular_Communications_Manuscript_v21.docx"
if not doc_path.exists():
    doc_path = PROJECT_ROOT / "Vehicular_Communications_Manuscript_v21.docx"

doc = docx.Document(doc_path)

banned_words = [
    'crucially', 'notably', 'importantly', 'significantly', 'in stark contrast',
    'in sharp contrast', 'conversely', 'consequently', 'furthermore', 'moreover',
    'these findings conclusively prove', 'this confirms that', 'this demonstrates that',
    'this establishes a vital', 'catastrophic representation collapse',
    'collapses catastrophically', 'plummet', 'surge', 'unavoidable operational trade-off',
    'indispensable', 'near-universal characteristic', 'near-universal', 'focal point',
    'orchestrated', 'harmonized', 'delves', 'catastrophic'
]

participial_tail_re = re.compile(
    r',\s*(?:[\w-]+\s+)*(?:demonstrating|highlighting|confirming|showing|indicating|rendering|isolating|withholding|verifying|revealing|mandating|detecting|causing|leading|yielding|suggesting|implying)\b[^.?!]*[.?!]$',
    re.IGNORECASE
)

body_section_headers = [
    "Abstract",
    "1 Introduction",
    "2 Threat Landscape and Automotive Attack Surface",
    "3 Related Work and Benchmark Positioning",
    "4 Mathematical Formulation and Experimental Methodology",
    "5 Empirical Deployability Audit: In-Distribution Benchmark",
    "6 Operational Line-Speed False Alarm Projections and Temporal Debouncing",
    "7 Robustness Stress Tests: Feature-Space Stealth Camouflage",
    "8 Cross-Platform Transferability: Sonata to Polaris GEM e6 Shuttle",
    "9 Engineering and Regulatory Implications: Safety vs. Security Tension",
    "10 Threats to Validity and Limitations",
    "11 Conclusion and Future Research Directions"
]

subheaders = [
    "2.1 Connected Vehicle Attack Surfaces",
    "2.2 CAN Bus Arbitration and Vulnerabilities",
    "2.3 Attack Taxonomy in In-Vehicle Networks",
    "2.4 Grounding in Real-World Automotive Exploits",
    "3.1 In-Vehicle Intrusion Detection Paradigms",
    "3.2 Critical Audit of Automotive CAN Benchmarks",
    "4.1 CAN Bus Message Representation and Feature Space",
    "4.2 Eliminating Data Leakage: 70/30 Time-Blocked Partitioning",
    "4.3 Statistical Metrics and Evaluation",
    "4.4 Single-Feature Baselines and Embedded Complexity Model",
    "5.1 Baseline Benchmark Performance",
    "5.2 Signal Isolation: Single-Feature ID Baseline Dissection",
    "5.3 Timing Feature Ablation and Shortcut Verification",
    "5.4 Multi-Detector Architectural Benchmark Across Operational Stress Tests",
    "6.1 Translating Statistical Rates to Operational Bus Frequencies",
    "6.2 Temporal Debounce Filtering and Independence Assumptions",
    "6.3 Empirical Measurement of Error Burstiness in Real Driving",
    "7.1 Motivation and Blending Formulation",
    "7.2 Camouflage Curriculum Retraining and the Evasion-False Alarm Trade-off",
    "8.1 Experimental Transfer Protocol and Ground-Truth Reconciliation",
    "8.2 Cross-Platform Transfer Performance and Feature Isolation",
    "8.3 Root Cause Analysis and Protocol Normalization",
    "8.4 Validation Against Physical Masquerade Attacks (ORNL ROAD Benchmark)",
    "8.5 Supervised Few-Shot Domain Adaptation on the Polaris GEM e6",
    "9.1 Interaction Between ISO 26262 and ISO/SAE 21434",
    "9.2 Deployment Architecture: Strict Two-Tier Authority Separation",
    "Declarations",
    "References"
]

sections = {h: [] for h in body_section_headers}
curr_sec = None

for p in doc.paragraphs:
    txt = p.text.strip()
    if not txt:
        continue
    if txt in body_section_headers:
        curr_sec = txt
        continue
    if txt in subheaders or any(txt.startswith(sh) for sh in subheaders):
        if txt == "Declarations" or txt == "References":
            curr_sec = None
        continue
    if curr_sec is None:
        continue
    # Filter out metadata, captions, footnotes, formulas, and term definition lines
    if txt.startswith("Keywords") or txt.startswith("Keywords—") or txt.startswith("The Brittleness"):
        continue
    if txt.startswith("Table ") or txt.startswith("Figure ") or txt.startswith("*Note:"):
        continue
    # Filter formulas
    if re.search(r'\(\d\)\s*$', txt) or txt.startswith("Eq. ("):
        continue
    if txt.startswith("Δt") or txt.startswith("Delta t_i") or txt.startswith("f_i =") or txt.startswith("μ_byte") or txt.startswith("mu_byte") or txt.startswith("σ_byte") or txt.startswith("sigma_byte") or txt.startswith("MCC =") or txt.startswith("x_camouflaged"):
        continue
    if txt.startswith("where ") or txt.startswith("where:") or txt.endswith("Coefficient (MCC):"):
        continue
    sections[curr_sec].append(txt)

print("=" * 80)
print("GLOBAL FULL-MANUSCRIPT AUDIT REPORT")
print("=" * 80)

total_paragraphs = 0
total_sentences = 0
total_words = 0
banned_found = []
tail_ing_found = []
no_short_sents = []

for sec in body_section_headers:
    paras = sections.get(sec, [])
    print(f"\n--- Section: {sec} ({len(paras)} paragraphs) ---")
    for p_idx, p_txt in enumerate(paras):
        total_paragraphs += 1
        words_in_p = re.findall(r'\b[\w-]+\b', p_txt)
        total_words += len(words_in_p)
        
        # Banned words check
        p_lower = p_txt.lower()
        for bw in banned_words:
            pattern = r'\b' + re.escape(bw) + r'\b'
            if re.search(pattern, p_lower):
                banned_found.append((sec, p_idx+1, bw))
                print(f"  [!] BANNED WORD in P{p_idx+1}: '{bw}'")
                    
        # Sentences split
        sents = re.split(r'(?<=[.!?])\s+', p_txt.strip())
        total_sentences += len(sents)
        has_short = False
        sent_lens = []
        for s in sents:
            w_list = re.findall(r'\b[\w-]+\b', s)
            w_cnt = len(w_list)
            sent_lens.append(w_cnt)
            if 5 <= w_cnt <= 9:
                has_short = True
            if participial_tail_re.search(s):
                m = participial_tail_re.search(s)
                tail_ing_found.append((sec, p_idx+1, m.group(0), s))
                print(f"  [!] PARTICIPIAL TAIL in P{p_idx+1}: {m.group(0)}")
                
        if not has_short:
            no_short_sents.append((sec, p_idx+1, sent_lens))
            print(f"  [!] NO 5-9w SENTENCE in P{p_idx+1}: lengths = {sent_lens}")

print("\n" + "=" * 80)
print("AUDIT SUMMARY METRICS:")
print(f"Total Narrative Paragraphs Analyzed: {total_paragraphs}")
print(f"Total Sentences Analyzed:           {total_sentences}")
print(f"Total Narrative Words:               {total_words}")
print(f"Banned Word Violations:              {len(banned_found)}")
print(f"Participial Tail Violations:         {len(tail_ing_found)}")
print(f"Paragraphs Missing 5-9w Sentences:   {len(no_short_sents)}")
print("=" * 80)
