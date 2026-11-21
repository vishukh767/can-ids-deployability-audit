# The Brittleness of Vehicle-Layer AI Intrusion Detection: An Empirical Deployability Audit and Cross-Platform Domain Analysis

Empirical evaluation and replication suite for *Vehicular Communications* (Elsevier, ISSN: 2214-2096, CiteScore 14.4, Impact Factor 6.7).
Investigates lookahead data leakage, single-feature ID memorization, multi-detector inductive biases, cross-platform transferability (11-bit standard vs. 29-bit extended SAE J1939), leakage-free camouflage curriculum retraining, physical masquerade attacks (ORNL ROAD), and supervised few-shot domain adaptation.

**Author**: Vishwas Khandelwal  
**Affiliation**: Department of Computer Science and Engineering, Vivekananda Institute of Professional Studies - Technical Campus (VIPS-TC), Delhi 110034, India  
**Contact**: vishukh767@gmail.com

---

## Data

1. **HCRL Car-Hacking Dataset (Hyundai YF Sonata, 11-bit CAN 2.0A)**
   - Source: Korea University HCRL (Kang et al., DOI: [10.1145/3064814.3064840](https://doi.org/10.1145/3064814.3064840)).
   - Files in `4_datasets/HCRL_car_hacking/`:
     - `DoS_dataset.csv`: 0x000 injection flood (190 MB, ~3.6M frames).
     - `Fuzzy_dataset.csv`: Randomized ID and payload injection (198 MB, ~3.8M frames).
     - `gear_dataset.csv`: Transmission state spoofing (230 MB, ~4.4M frames).
     - `RPM_dataset.csv`: Engine speed sensor spoofing (239 MB, ~4.6M frames).
     - `normal_run_data/` / `normal_run_data.txt`: Clean baseline driving telemetry (200,000 continuous frames).

2. **GEM-CAN Autonomous Shuttle Dataset (Polaris GEM e6, 29-bit SAE J1939)**
   - Source: Tavasoli et al., *Data in Brief*, vol. 66, 112805, 2026 (DOI: [10.1016/j.dib.2026.112805](https://doi.org/10.1016/j.dib.2026.112805), Zenodo: [10.5281/zenodo.19161139](https://doi.org/10.5281/zenodo.19161139)).
   - Files in `4_datasets/GEM_e6_autonomous/`:
     - `GEM_Normal_Driving_Raw_Labeled.csv`: 100,000 nominal driving frames (100% 29-bit J1939 PGNs).
     - `GEM_Attack_Scenario_Raw_Labeled.csv`: 43,444 frames (41,080 DoS frames with ID 0x00000000, 663 brake tampering frames, 662 steering tampering frames, 1,039 normal frames).
     - Unified Ground Truth: 101,039 Normal vs. 42,405 Attack frames.

3. **ORNL ROAD Dataset (Oak Ridge National Laboratory, Physical Masquerade Attacks)**
   - Source: Verma et al., *Computers & Security*, vol. 125, 103027, 2023 (DOI: [10.1016/j.cose.2022.103027](https://doi.org/10.1016/j.cose.2022.103027)).
   - File in `4_datasets/ROAD/`:
     - `correlated_signal_attack_1_masquerade.csv`: 74,152 frames (72,066 normal driving frames and 2,086 physically injected masquerade attack frames preserving nominal ECU broadcast periodicity).

---

## Pipeline Reproduction

Run all analyses and artifact generators in order from the repository root:

```bash
# 1. Run time-blocked evaluation and single-feature ID baseline on HCRL
python real_can_sim.py

# 2. Run cross-platform transfer, payload-only ablation, and within-GEM baseline
python eval_gem_cross_val.py

# 3. Run multi-detector benchmark, ROAD masquerade, camouflage curriculum, and few-shot adaptation
python run_part_b_experiments.py

# 4. Regenerate publication-grade figures
python generate_figures.py

# 5. Compile the full journal manuscript (v21), title page, and supplementary tables
python build_docx_paper.py
python build_title_page.py
python build_supplementary_tables.py
```

Requirements: `python-docx>=1.0`, `scikit-learn>=1.2`, `lightgbm>=3.3`, `torch>=2.0`, `numpy`, `pandas`, `matplotlib`, `scipy`. Benchmarked on Python 3.10+ (Windows 11 / x86_64).

---

## Methodology Summary

- **Chronological Time-Blocked Split**: 70% initial continuous frames allocated to training, remaining 30% strictly held out for evaluation. Evaluated across 5 random seeds (11, 21, 42, 84, 105) to eliminate lookahead data leakage.
- **Single-Feature Baseline**: Depth-constrained Decision Tree thresholding strictly on arbitration CAN ID. Moves accuracy from 25.0% (random chance) to 82.40% (0.781 MCC), capturing 76.6% of the multi-feature model's gain over chance without inspecting payloads or timing.
- **Multi-Detector Benchmark (Item 13)**: Evaluates Timing-Rules, Random Forests, LightGBM, and Deep Reconstruction Autoencoders. Demonstrates that all feature-aware architectures experience representation collapse (91.20% to 100.00% FPR) under direct transfer across 11-bit vs. 29-bit protocol boundaries.
- **Cross-Platform Transfer (Sonata -> Polaris GEM e6)**:
  - *Direct Transfer (Raw IDs)*: Collapses to 35.57% accuracy, 0.1637 MCC, and 91.48% FPR because extended 29-bit IDs fall outside the 11-bit range and branch into attack leaves.
  - *Payload-Only Ablation*: 53.62% accuracy, 0.3646 MCC, 65.85% FPR.
  - *Calibration-Normalized Stateful Encoding*: Standardizes 6 protocol-agnostic features against each vehicle's nominal baseline (`in_allowlist_i`, `freq_rank_nom,i`, `r_{\Delta t, i}`, `z_{\mu, i}`, `z_{\sigma, i}`, `DLC_i`), achieving 92.04% accuracy, 0.8316 MCC, 10.65% FPR, 100% DoS recall, and 50.04% tampering recall.
  - *Native Within-GEM Baseline*: 100.00% accuracy, 1.0000 MCC, 0.00% FPR on held-out 30% test partition (30,295 Normal vs. 12,739 Attack).
- **Physical Masquerade Validation (Item 12 - ORNL ROAD)**: On timing-preserving masquerade attacks, stateless timing baselines collapse to 0.00% recall (-0.0042 MCC), while payload-aware tree ensembles achieve 100.00% recall (0.9993 MCC, 0.005% FPR).
- **Leakage-Free Camouflage Curriculum Retraining (Item 14)**: Enforces per-ID frame matching and restricts training strictly to $\alpha \in \{0.25, 0.50\}$. Recovers attack recall from 41.80% to 98.60% at $\alpha = 0.25$, but expands decision boundaries, increasing nominal FPR from 0.091% (6,552 alarms/hr) to 0.246% (17,700 alarms/hr at 500 kbps).
- **Supervised Few-Shot Domain Adaptation (Item 15)**: On Polaris GEM e6, target domain FPR drops from 10.65% (zero-shot) to 0.73% with $N = 100$ frames (0.2 s of driving), drops to 0.04% FPR (and 95.72% tampering recall) with $N = 500$ frames, reaching 0.00% FPR and 100.00% tampering recall with $N \ge 1,000$ frames ($100.00\%$ accuracy, $1.0000$ MCC).
- **Physical Transceiver Discard**: Formulated under ISO 11898-1. Bit-stuffing and form anomalies trigger active error frames and hardware discard at the silicon transceiver level (CRC-15) before frames reach microcontroller software.
- **Temporal Debounce & Burstiness**: Multi-frame whole-bus persistence filter ($k = 3$) achieves 99.76% alarm suppression on 200,000 continuous nominal frames with 1.5 ms latency at 500 kbps. Evaluated against cross-session driving drift ($p = 11.068\%$), where whole-bus yields 53 alarms (5-fold below i.i.d. expectation of 271) while per-ID yields 14,514 alarms (53-fold above i.i.d.), empirically proving correlated per-ID bursts.

---

## Journal Submission Deliverables

- `Vehicular_Communications_Manuscript_v21.docx`: Full 6,200-word journal manuscript conforming to Elsevier *Vehicular Communications* specifications (9 tables, 4 figures, complete declarations, 24 verified DOIs). Also archived as `1_manuscripts_and_drafts/draft_v21_full_journal.docx`.
- `Title_Page.docx`: Standalone Elsevier title page with author affiliation (*VIPS-TC, Delhi*), corresponding author details, 5 highlights ($\le 85$ chars), abstract ($< 250$ words), and compliance declarations.
- `Highlights.docx`: Standalone Highlights document containing strictly the 5 bullet points ($\le 85$ characters each) for the Elsevier Editorial Manager portal.
- `Cover_Letter.docx`: Formal submission cover letter addressed to the Editors-in-Chief of *Vehicular Communications*.
- `Supplementary_Tables.docx`: Extended methodological and empirical tables (Tables S1–S8) in APA 3-line format.
- `Part_B_Experimental_Results.md`: Machine-verified empirical benchmark log across all Part B stress tests.
- Production Figures (300 DPI):
  - `Figure_1_HCRL_Confusion_Matrix.png` (formerly `figure3`)
  - `Figure_2_Stealth_Camouflage_Degradation.png` (formerly `figure5`)
  - `Figure_3_GEM_Cross_Platform_Confusion.png` (formerly `figure4`)
  - `Figure_4_Cross_Platform_Comparison.png` (formerly `figure6`)

---

## Notes / Empirical Details

- **Debounce Recurrence**: At 250 kbps (nominal 800 frames/s = 2,880,000 frames/hr) and $k=3$, the independent Poisson baseline at 0.091% FPR yields 0.00217 alarms/hour ($\approx 1$ alarm every 19.2 days; 7.7 days at 500 kbps [2,000 fps]; 3.8 days at 1 Mbps [4,000 fps]). Real continuous driving produces 1,863 alarms/hour due to transient ECU dynamic burst clustering.
- **Microcontroller Footprint**: Analytical instruction-cycle and memory models for a 32-bit ARM Cortex-M7 core at 160 MHz: 100 trees, depth 12, 34,200 packed 8-byte nodes $\implies$ 273.6 KB Flash, < 4 KB RAM, 18.4 $\mu$s estimated inference latency (consuming 3.68% of the 500 $\mu$s nominal arrival window at 500 kbps).

