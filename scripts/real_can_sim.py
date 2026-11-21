"""
real_can_sim.py: Hardened In-Distribution Evaluation Pipeline (HCRL Sonata)

Implements:
1. Dynamic dataset path resolution (runs from any directory).
2. Chronological time-blocked train/test split (no random split leakage).
3. Exact train-test duplicate feature-vector audit.
4. Single-feature baseline (CAN ID alone) vs. Random Forest.
5. Timing feature ablation (Full vs. Ablated without dt_since_last_same_id).
6. 5-seed evaluation reporting mean +/- SD.
"""

import csv
import sys
import time
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import (
    accuracy_score, precision_recall_fscore_support,
    matthews_corrcoef, confusion_matrix
)

# --- 1. Locate Datasets Dynamically ---
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR if (SCRIPT_DIR / "4_datasets").is_dir() else SCRIPT_DIR.parent

CANDIDATE_DIRS = [
    PROJECT_ROOT / "4_datasets" / "HCRL_car_hacking",
    PROJECT_ROOT / "4_datasets",
    SCRIPT_DIR.parent,
    SCRIPT_DIR,
    Path.cwd(),
]

def find_file(filename: str) -> Path:
    for folder in CANDIDATE_DIRS:
        target = folder / filename
        if target.is_file():
            return target
    matches = list(PROJECT_ROOT.glob(f"**/{filename}"))
    if matches:
        return matches[0]
    raise FileNotFoundError(
        f"Could not locate '{filename}'. Searched in:\n" +
        "\n".join(str(p) for p in CANDIDATE_DIRS)
    )

FILES = {
    "DoS_dataset.csv": "DoS",
    "Fuzzy_dataset.csv": "Fuzzy",
    "gear_dataset.csv": "Spoofing",
    "RPM_dataset.csv": "Spoofing",
}
PER_CLASS_TARGET = 20000

def parse_hcrl_file(path: Path, attack_name: str) -> pd.DataFrame:
    """Parses raw HCRL logs and computes inter-arrival dt on the FULL chronological stream."""
    timestamps, can_ids, dlcs, byte_means, byte_stds, labels = [], [], [], [], [], []
    last_seen = {}

    with open(path, newline="", encoding="utf-8", errors="ignore") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) < 4:
                continue
            try:
                ts = float(row[0])
                can_id = int(row[1], 16)
                dlc = int(row[2])
                label_flag = row[-1].strip()
            except ValueError:
                continue

            data_fields = row[3:3 + dlc]
            vals = [int(x, 16) for x in data_fields if x.strip() != ""] if data_fields else []
            bmean = sum(vals) / len(vals) if vals else 0.0
            bstd = (sum((v - bmean) ** 2 for v in vals) / len(vals)) ** 0.5 if vals else 0.0

            dt = ts - last_seen.get(can_id, ts - 0.05)
            last_seen[can_id] = ts

            timestamps.append(ts)
            can_ids.append(can_id)
            dlcs.append(dlc)
            byte_means.append(bmean)
            byte_stds.append(bstd)
            labels.append(attack_name if label_flag == "T" else "Normal")

    df = pd.DataFrame({
        "t": timestamps,
        "can_id": can_ids,
        "dlc": dlcs,
        "byte_mean": byte_means,
        "byte_std": byte_stds,
        "label": labels
    })
    # Compute true per-ID delta on full unbroken capture
    df["dt_since_last_same_id"] = df.groupby("can_id")["t"].diff().fillna(df["t"].diff().median())
    return df

print("=== STEP 1: Ingesting & Extracting Features on Full Streams ===")
frames = []
for fname, attack_name in FILES.items():
    file_path = find_file(fname)
    print(f"Reading {fname} from: {file_path}")
    fdf = parse_hcrl_file(file_path, attack_name)
    
    # Chronological sample per class to avoid temporal shuffling
    parts = []
    for lbl, group in fdf.groupby("label"):
        n = min(PER_CLASS_TARGET, len(group))
        parts.append(group.iloc[:n])
    frames.append(pd.concat(parts, ignore_index=True))

full_df = pd.concat(frames, ignore_index=True)
parts = []
for lbl, group in full_df.groupby("label"):
    n = min(PER_CLASS_TARGET, len(group))
    parts.append(group.iloc[:n])
dataset = pd.concat(parts, ignore_index=True)
print(f"Dataset assembled. Total records: {len(dataset):,}")

feature_cols = ["can_id", "dlc", "dt_since_last_same_id", "byte_mean", "byte_std"]

# --- Time-Blocked Chronological Split ---
print("\n=== STEP 2: Time-Blocked Train/Test Split (70/30) ===")
train_parts, test_parts = [], []
for lbl, grp in dataset.groupby("label"):
    grp_sorted = grp.sort_values("t").reset_index(drop=True)
    split_pt = int(0.70 * len(grp_sorted))
    train_parts.append(grp_sorted.iloc[:split_pt])
    test_parts.append(grp_sorted.iloc[split_pt:])

train_df = pd.concat(train_parts, ignore_index=True)
test_df = pd.concat(test_parts, ignore_index=True)

# Leakage check: exact feature-vector duplicates between train and test
train_tuples = set(map(tuple, train_df[feature_cols].round(4).values))
test_tuples = [tuple(x) for x in test_df[feature_cols].round(4).values]
duplicates = sum(1 for x in test_tuples if x in train_tuples)
dup_rate = (duplicates / len(test_df)) * 100
print(f"Train size: {len(train_df):,} | Test size: {len(test_df):,}")
print(f"Exact feature vector overlap: {duplicates}/{len(test_df)} ({dup_rate:.2f}%)")

X_train, y_train = train_df[feature_cols], train_df["label"]
X_test, y_test = test_df[feature_cols], test_df["label"]

# --- Baseline 1: Single-Feature CAN ID Baseline ---
print("\n=== STEP 3: Evaluating Single-Feature Baseline (CAN ID Only) ===")
dt_id = DecisionTreeClassifier(max_depth=3, random_state=42)
dt_id.fit(X_train[["can_id"]], y_train)
y_pred_id = dt_id.predict(X_test[["can_id"]])
acc_id = accuracy_score(y_test, y_pred_id)
mcc_id = matthews_corrcoef(y_test, y_pred_id)
print(f"CAN ID-only Decision Tree Baseline -> Accuracy: {acc_id*100:.2f}%, MCC: {mcc_id:.4f}")

# --- Primary Random Forest (Multi-Seed Evaluation) ---
print("\n=== STEP 4: Evaluating Random Forest Across 5 Random Seeds ===")
seeds = [11, 21, 42, 84, 105]
accs, mccs, fprs = [], [], []

for s in seeds:
    rf = RandomForestClassifier(n_estimators=150, max_depth=12, random_state=s, n_jobs=-1)
    rf.fit(X_train, y_train)
    preds = rf.predict(X_test)
    
    cm = confusion_matrix(y_test, preds, labels=["Normal", "DoS", "Fuzzy", "Spoofing"])
    tn = cm[0, 0]
    fp = cm[0, 1:].sum()
    fpr = fp / (fp + tn) if (fp + tn) > 0 else 0.0
    
    accs.append(accuracy_score(y_test, preds))
    mccs.append(matthews_corrcoef(y_test, preds))
    fprs.append(fpr)

print(f"Random Forest (Full Features, Chronological Split):")
print(f"  Accuracy: {np.mean(accs)*100:.3f}% +/- {np.std(accs)*100:.3f}%")
print(f"  MCC:      {np.mean(mccs):.4f} +/- {np.std(mccs):.4f}")
print(f"  FPR:      {np.mean(fprs)*100:.4f}% +/- {np.std(fprs)*100:.4f}%")

# --- Timing Ablation ---
print("\n=== STEP 5: Timing Feature Ablation ===")
ablated_cols = ["can_id", "dlc", "byte_mean", "byte_std"]
rf_ablated = RandomForestClassifier(n_estimators=150, max_depth=12, random_state=42, n_jobs=-1)
rf_ablated.fit(X_train[ablated_cols], y_train)
preds_abl = rf_ablated.predict(X_test[ablated_cols])
print(f"Ablated RF (No Timing Feature) -> Accuracy: {accuracy_score(y_test, preds_abl)*100:.2f}%, MCC: {matthews_corrcoef(y_test, preds_abl):.4f}")