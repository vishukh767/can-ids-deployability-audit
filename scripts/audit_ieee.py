import re
import docx
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name in ["scripts", "2_code_and_scripts"] else SCRIPT_DIR
doc_path = PROJECT_ROOT / "drafts" / "IEEE_Conference_Draft_v1.docx"
if not doc_path.exists():
    doc_path = PROJECT_ROOT / "1_manuscripts_and_drafts" / "IEEE_Conference_Draft_v1.docx"
if not doc_path.exists():
    doc_path = PROJECT_ROOT / "IEEE_Conference_Draft_v1.docx"

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

print("=" * 80)
print("COMPREHENSIVE AUDIT REPORT: IEEE_Conference_Draft_v1.docx")
print("=" * 80)

# Filter narrative paragraphs
narrative_paras = []
for p in doc.paragraphs:
    txt = p.text.strip()
    if not txt:
        continue
    # skip title, author, abstract, index terms, headings, table/fig captions, formulas, refs
    if txt.startswith("The Brittleness") or txt.startswith("Vishwas") or txt.startswith("Department") or "vips-tc" in txt.lower():
        continue
    if txt.startswith("Abstract—") or txt.startswith("Index Terms—"):
        continue
    if re.match(r'^[I|V|X]+\.\s+', txt) or re.match(r'^[A-Z]\.\s+', txt):
        continue
    if txt.startswith("TABLE ") or txt.startswith("Fig. ") or txt.startswith("*"):
        continue
    if txt == "REFERENCES":
        continue
    if re.search(r'\(\d\)\s*$', txt) or txt.startswith("f_i =") or txt.startswith("Delta t_i") or txt.startswith("mu_byte") or txt.startswith("sigma_byte") or txt.startswith("MCC =") or txt.startswith("R_alarm =") or txt.startswith("x_camouflaged"):
        continue
    if re.match(r'^\[\d+\]', txt):
        continue
    narrative_paras.append(txt)

print(f"Total Narrative Paragraphs: {len(narrative_paras)}")

total_words = 0
total_sentences = 0
banned_violations = []
tail_violations = []
missing_short = []

for idx, p_txt in enumerate(narrative_paras):
    words = re.findall(r'\b[\w-]+\b', p_txt)
    total_words += len(words)
    
    # check banned
    p_lower = p_txt.lower()
    for bw in banned_words:
        pattern = r'\b' + re.escape(bw) + r'\b'
        if re.search(pattern, p_lower):
            banned_violations.append((idx+1, bw))
            print(f"  [!] P{idx+1} Banned Word: '{bw}'")
            
    # check sentences
    sents = re.split(r'(?<=[.!?])\s+', p_txt.strip())
    total_sentences += len(sents)
    has_short = False
    sent_lens = []
    for s in sents:
        w_list = re.findall(r'\b[\w-]+\b', s)
        cnt = len(w_list)
        sent_lens.append(cnt)
        if 5 <= cnt <= 9:
            has_short = True
        if participial_tail_re.search(s):
            m = participial_tail_re.search(s)
            tail_violations.append((idx+1, m.group(0), s))
            print(f"  [!] P{idx+1} Participial Tail: {m.group(0)}")
            
    if not has_short:
        missing_short.append((idx+1, sent_lens))
        print(f"  [!] P{idx+1} Missing 5-9w sentence: lengths = {sent_lens}")

print("\n" + "=" * 80)
print("AUDIT SUMMARY:")
print(f"Narrative Paragraphs:       {len(narrative_paras)}")
print(f"Narrative Sentences:        {total_sentences}")
print(f"Narrative Word Count:       {total_words}")
print(f"Banned Word Violations:     {len(banned_violations)}")
print(f"Participial Tail Violations:{len(tail_violations)}")
print(f"Paragraphs Missing 5-9w:    {len(missing_short)}")
print("=" * 80)

# Check all tables, figures, equations, metrics
print("\n--- INVARIANT VERIFICATION ---")
all_text = "\n".join([p.text for p in doc.paragraphs])
table_text = "\n".join(["\t".join([c.text for c in row.cells]) for t in doc.tables for row in t.rows])
corpus = all_text + "\n" + table_text

invariants = {
    "TABLE I present": "TABLE I" in corpus,
    "TABLE II present": "TABLE II" in corpus,
    "TABLE III present": "TABLE III" in corpus,
    "TABLE IV present": "TABLE IV" in corpus,
    "TABLE V present": "TABLE V" in corpus,
    "TABLE VI present": "TABLE VI" in corpus,
    "Fig. 1 present": "Fig. 1." in corpus or "Fig. 1" in corpus,
    "Fig. 2 present": "Fig. 2." in corpus or "Fig. 2" in corpus,
    "Fig. 3 present": "Fig. 3." in corpus or "Fig. 3" in corpus,
    "Fig. 4 present": "Fig. 4." in corpus or "Fig. 4" in corpus,
    "Equation (1)": "(1)" in corpus,
    "Equation (2)": "(2)" in corpus,
    "Equation (3)": "(3)" in corpus,
    "Equation (4)": "(4)" in corpus,
    "Equation (5)": "(5)" in corpus,
    "Equation (6)": "(6)" in corpus,
    "Equation (7)": "(7)" in corpus,
    "99.90% accuracy": "99.90%" in corpus,
    "0.091% FPR": "0.091%" in corpus,
    "82.40% ID-only": "82.40%" in corpus,
    "76.64% gain above chance": "76.64%" in corpus,
    "143,444 GEM frames": "143,444" in corpus,
    "101,039 GEM normal frames": "101,039" in corpus,
    "42,405 GEM attack frames": "42,405" in corpus,
    "35.57% direct transfer accuracy": "35.57%" in corpus,
    "91.48% direct transfer FPR": "91.48%" in corpus,
    "0.1637 direct transfer MCC": "0.1637" in corpus,
    "92.04% calibration-normalized accuracy": "92.04%" in corpus,
    "10.65% calibration-normalized FPR": "10.65%" in corpus,
    "0.8316 calibration-normalized MCC": "0.8316" in corpus,
    "74,152 ROAD frames": "74,152" in corpus,
    "ROAD 0.00% timing recall": "0.00%" in corpus,
    "ROAD 100.00% tree recall": "100.00%" in corpus,
    "ROAD 0.005% FPR": "0.005%" in corpus,
    "ROAD 0.9993 MCC": "0.9993" in corpus,
    "Few-shot N=500 0.04% FPR": "0.04%" in corpus,
    "Few-shot N>=1,000 0.00% FPR": "0.00%" in corpus,
    "1,863 alarms/hr in real driving": "1,863" in corpus,
    "510,134 alarms/hr per-ID": "510,134" in corpus,
    "7.7 days debouncing at 500 kbps": "7.7 days" in corpus,
    "Wilson CI [0.034%, 0.198%]": "[0.034%, 0.198%]" in corpus,
    "Wilson CI [0.062%, 0.134%]": "[0.062%, 0.134%]" in corpus,
    "ARM Cortex-M7 specs": "Cortex-M7" in corpus,
    "~34,200 nodes": "34,200" in corpus,
    "~274 KB flash": "274 KB" in corpus or "274" in corpus,
    "<4 KB stack RAM": "4 KB" in corpus
}

all_invariants_ok = True
for k, v in invariants.items():
    st = "PASS" if v else "FAIL"
    if not v:
        all_invariants_ok = False
    print(f"  [{st}] {k}")

# Citations [1] to [18]
refs_ok = True
for i in range(1, 19):
    ref_tag = f"[{i}]"
    if ref_tag not in corpus:
        print(f"  [FAIL] Missing citation {ref_tag}")
        refs_ok = False
        all_invariants_ok = False

if refs_ok:
    print("  [PASS] All Citations [1] through [18] verified present.")

print(f"\nFinal Invariant Status: {'ALL PASSED' if all_invariants_ok else 'SOME FAILED'}")
