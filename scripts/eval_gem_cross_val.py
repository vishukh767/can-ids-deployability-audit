"""
eval_gem_cross_val.py: Reconciled Cross-Platform Domain Transfer & Feature Isolation Pipeline

Evaluates:
1. Reconciled GEM-CAN Ground-Truth (101,039 Normal vs. 42,405 Attack frames)
2. Raw 11-bit vs. 29-bit CAN ID cross-platform evaluation (Sonata -> GEM e6)
3. Payload-only feature evaluation (ID and timing ablated)
4. Calibration-Normalized Stateful Encoding (Item 11 fixed: in_allowlist, freq_rank, dt_ratio, z_mu, z_sig)
5. Native Within-GEM Baseline (70/30 time-blocked split within each file)
6. Separate evaluation of DoS (n=41,080) vs. Physical Tampering (Brake + Steering, n=1,325) recall
"""

import os
import sys
import csv
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, precision_recall_fscore_support, matthews_corrcoef, confusion_matrix

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR if (SCRIPT_DIR / "4_datasets").is_dir() else SCRIPT_DIR.parent

CANDIDATE_DIRS = [
    PROJECT_ROOT / "4_datasets" / "HCRL_car_hacking",
    PROJECT_ROOT / "4_datasets" / "GEM_e6_autonomous",
    PROJECT_ROOT / "4_datasets",
    PROJECT_ROOT,
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
    raise FileNotFoundError(f"Could not locate '{filename}'.")

print("========================================================================")
print("  STEP 1: Training Baseline Random Forest on HCRL Sonata Telemetry      ")
print("========================================================================")

FILES = {
    "DoS_dataset.csv": "DoS",
    "Fuzzy_dataset.csv": "Fuzzy",
    "gear_dataset.csv": "Spoofing",
    "RPM_dataset.csv": "Spoofing",
}

hcrl_train_list = []
hcrl_nom_clean = []

for fname, atk in FILES.items():
    fpath = find_file(fname)
    rows = []
    with open(fpath, newline="", encoding="utf-8", errors="ignore") as f:
        reader = csv.reader(f)
        for r in reader:
            if len(r) < 4: continue
            try:
                ts = float(r[0])
                cid = int(r[1], 16)
                dlc = int(r[2])
                flag = r[-1].strip()
                vals = [int(x, 16) for x in r[3:3+dlc] if x.strip()!=''] if dlc>0 else []
                bm = sum(vals)/len(vals) if vals else 0.0
                bs = (sum((v-bm)**2 for v in vals)/len(vals))**0.5 if vals else 0.0
                rows.append((ts, cid, dlc, bm, bs, atk if flag=='T' else 'Normal'))
            except: continue
    df_f = pd.DataFrame(rows, columns=['t', 'can_id', 'dlc', 'byte_mean', 'byte_std', 'label'])
    df_f['dt'] = df_f.groupby('can_id')['t'].diff().fillna(df_f['t'].diff().median())
    
    # 70% time-block split
    split_idx = int(len(df_f) * 0.70)
    train_part = df_f.iloc[:split_idx]
    for lbl, grp in train_part.groupby('label'):
        hcrl_train_list.append(grp.iloc[:min(15000, len(grp))])
        if lbl == 'Normal':
            hcrl_nom_clean.append(grp)

hcrl_train = pd.concat(hcrl_train_list, ignore_index=True)
hcrl_nom = pd.concat(hcrl_nom_clean, ignore_index=True)

# Balance HCRL train to 15,000 per class (60,000 total)
parts = []
for lbl, grp in hcrl_train.groupby('label'):
    parts.append(grp.iloc[:min(15000, len(grp))])
train_hcrl = pd.concat(parts, ignore_index=True)
print(f"HCRL training pool: {len(train_hcrl):,} balanced frames.")

print("\n========================================================================")
print("  STEP 2: Parsing Reconciled GEM-CAN Telemetry                          ")
print("========================================================================")

gem_norm_file = find_file("GEM_Normal_Driving_Raw_Labeled.csv")
gem_att_file = find_file("GEM_Attack_Scenario_Raw_Labeled.csv")

df_gnorm = pd.read_csv(gem_norm_file)
df_gatt = pd.read_csv(gem_att_file)

def parse_gem_raw(df, is_atk):
    cids = df['arbitration_id'].apply(lambda x: int(str(x).strip(), 16) if str(x).strip().startswith('0x') or any(c in str(x).lower() for c in 'abcdef') else int(str(x).strip()))
    data_cols = [f'data{i}' for i in range(8)]
    mat = df[data_cols].copy()
    for c in mat.columns:
        mat[c] = mat[c].apply(lambda x: int(str(x), 16) if pd.notnull(x) and str(x).strip() != '' else 0)
    bm = mat.mean(axis=1)
    bs = mat.std(axis=1).fillna(0.0)
    t = df['timestamp'].astype(float)
    if is_atk:
        sub_type = df['Attack_Type'].astype(str).str.strip()
        label = sub_type.apply(lambda x: 'Normal' if x.lower() == 'normal' else 'Attack')
    else:
        sub_type = pd.Series('Normal', index=df.index)
        label = pd.Series('Normal', index=df.index)
    res = pd.DataFrame({'t': t, 'can_id': cids, 'dlc': df['dlc'].astype(int), 'byte_mean': bm, 'byte_std': bs, 'label': label, 'sub_type': sub_type})
    res['dt'] = res.groupby('can_id')['t'].diff().fillna(res['t'].diff().median())
    return res

p_gnorm = parse_gem_raw(df_gnorm, is_atk=False)
p_gatt = parse_gem_raw(df_gatt, is_atk=True)
gem_test = pd.concat([p_gnorm, p_gatt], ignore_index=True)

print(f"Total GEM frames parsed: {len(gem_test):,}")
print(f"Normal: {(gem_test['label']=='Normal').sum():,} | Attack: {(gem_test['label']=='Attack').sum():,}")
print("Attack sub-types:", gem_test['sub_type'].value_counts().to_dict())

print("\n========================================================================")
print("  STEP 3: Computing Calibration-Normalized Profiles                    ")
print("========================================================================")

# Sonata Nominal Calibration Profile
sonata_stats = {}
nom_ids_sonata = hcrl_nom['can_id'].value_counts()
for rank, (cid, cnt) in enumerate(nom_ids_sonata.items()):
    c_df = hcrl_nom[hcrl_nom['can_id'] == cid]
    sonata_stats[cid] = {
        'rank': rank / len(nom_ids_sonata),
        'mean_dt': c_df['dt'].mean(),
        'mean_mu': c_df['byte_mean'].mean(),
        'std_mu': max(c_df['byte_mean'].std(), 1e-4),
        'mean_sigma': c_df['byte_std'].mean(),
        'std_sigma': max(c_df['byte_std'].std(), 1e-4),
    }

# GEM Nominal Calibration Profile (from first 70% of clean normal log)
gem_nom_cal = p_gnorm.iloc[:int(len(p_gnorm)*0.70)]
gem_stats = {}
nom_ids_gem = gem_nom_cal['can_id'].value_counts()
for rank, (cid, cnt) in enumerate(nom_ids_gem.items()):
    c_df = gem_nom_cal[gem_nom_cal['can_id'] == cid]
    gem_stats[cid] = {
        'rank': rank / len(nom_ids_gem),
        'mean_dt': c_df['dt'].mean(),
        'mean_mu': c_df['byte_mean'].mean(),
        'std_mu': max(c_df['byte_mean'].std(), 1e-4),
        'mean_sigma': c_df['byte_std'].mean(),
        'std_sigma': max(c_df['byte_std'].std(), 1e-4),
    }

def apply_calib(df, stats):
    in_allow, freq_rank, dt_ratio, z_mu, z_sig = [], [], [], [], []
    for _, row in df.iterrows():
        cid = row['can_id']
        if cid in stats:
            s = stats[cid]
            in_allow.append(1.0)
            freq_rank.append(s['rank'])
            dt_ratio.append(min(row['dt'] / s['mean_dt'], 10.0))
            z_mu.append(np.clip((row['byte_mean'] - s['mean_mu']) / s['std_mu'], -10.0, 10.0))
            z_sig.append(np.clip((row['byte_std'] - s['mean_sigma']) / s['std_sigma'], -10.0, 10.0))
        else:
            in_allow.append(0.0)
            freq_rank.append(-1.0)
            dt_ratio.append(0.0)
            z_mu.append(0.0)
            z_sig.append(0.0)
    df_out = df.copy()
    df_out['in_allowlist'] = in_allow
    df_out['freq_rank'] = freq_rank
    df_out['dt_ratio'] = dt_ratio
    df_out['z_mu'] = z_mu
    df_out['z_sig'] = z_sig
    return df_out

train_hcrl_cal = apply_calib(train_hcrl, sonata_stats)
gem_test_cal = apply_calib(gem_test, gem_stats)

print("Calibration profiles applied successfully.")

print("\n========================================================================")
print("  STEP 4: Executing Reconciled Comparative Experiments                  ")
print("========================================================================")

def eval_preds(y_true, y_pred, sub_types):
    cm = confusion_matrix(y_true, y_pred, labels=['Normal', 'Attack'])
    tn, fp, fn, tp = cm.ravel()
    acc = accuracy_score(y_true, y_pred)
    mcc = matthews_corrcoef(y_true, y_pred)
    fpr = fp / (fp + tn)
    
    dos_mask = (sub_types == 'DoS')
    tamp_mask = sub_types.isin(['Brake Tampering', 'Steering Tampering'])
    
    dos_rec = (y_pred[dos_mask] == 'Attack').mean() if dos_mask.sum() > 0 else 0.0
    tamp_rec = (y_pred[tamp_mask] == 'Attack').mean() if tamp_mask.sum() > 0 else 0.0
    return acc, mcc, fpr, dos_rec, tamp_rec, cm

# EXP 1: Raw IDs
rf_raw = RandomForestClassifier(n_estimators=100, max_depth=12, random_state=42, n_jobs=-1)
raw_feats = ['can_id', 'dlc', 'dt', 'byte_mean', 'byte_std']
rf_raw.fit(train_hcrl[raw_feats], train_hcrl['label'])
pred_raw = pd.Series(rf_raw.predict(gem_test[raw_feats])).apply(lambda x: 'Normal' if x == 'Normal' else 'Attack')
acc1, mcc1, fpr1, drec1, trec1, cm1 = eval_preds(gem_test['label'], pred_raw, gem_test['sub_type'])

# EXP 2: Payload-Only
rf_pay = RandomForestClassifier(n_estimators=100, max_depth=12, random_state=42, n_jobs=-1)
pay_feats = ['dlc', 'byte_mean', 'byte_std']
rf_pay.fit(train_hcrl[pay_feats], train_hcrl['label'])
pred_pay = pd.Series(rf_pay.predict(gem_test[pay_feats])).apply(lambda x: 'Normal' if x == 'Normal' else 'Attack')
acc2, mcc2, fpr2, drec2, trec2, cm2 = eval_preds(gem_test['label'], pred_pay, gem_test['sub_type'])

# EXP 3: Calibration-Normalized Stateful (Item 11 Fixed)
rf_cal = RandomForestClassifier(n_estimators=100, max_depth=12, random_state=42, n_jobs=-1)
cal_feats = ['in_allowlist', 'freq_rank', 'dt_ratio', 'z_mu', 'z_sig', 'dlc']
rf_cal.fit(train_hcrl_cal[cal_feats], train_hcrl_cal['label'])
pred_cal = pd.Series(rf_cal.predict(gem_test_cal[cal_feats])).apply(lambda x: 'Normal' if x == 'Normal' else 'Attack')
acc3, mcc3, fpr3, drec3, trec3, cm3 = eval_preds(gem_test_cal['label'], pred_cal, gem_test_cal['sub_type'])

# EXP 4: Native Within-GEM Baseline (70/30 Time-Blocked within each file)
norm_s = int(len(p_gnorm) * 0.70)
att_s = int(len(p_gatt) * 0.70)
train_g = pd.concat([p_gnorm.iloc[:norm_s], p_gatt.iloc[:att_s]], ignore_index=True)
test_g = pd.concat([p_gnorm.iloc[norm_s:], p_gatt.iloc[att_s:]], ignore_index=True)
rf_nat = RandomForestClassifier(n_estimators=100, max_depth=12, random_state=42, n_jobs=-1)
rf_nat.fit(train_g[raw_feats], train_g['label'])
pred_nat = pd.Series(rf_nat.predict(test_g[raw_feats])).apply(lambda x: 'Normal' if x == 'Normal' else 'Attack')
acc4, mcc4, fpr4, drec4, trec4, cm4 = eval_preds(test_g['label'], pred_nat, test_g['sub_type'])

print("\n" + "="*80)
print("  FINAL RECONCILED EXPERIMENTAL MATRIX (GROUND TRUTH: 101,039 N / 42,405 A)")
print("="*80)
print(f"1. Direct Cross-Platform (Raw IDs):")
print(f"   Accuracy: {acc1*100:.2f}% | MCC: {mcc1:.4f} | FPR: {fpr1*100:.2f}% | DoS Rec: {drec1*100:.2f}% | Tamp Rec: {trec1*100:.2f}%")
print(f"   Confusion Matrix (Normal, Attack):\n{cm1}\n")

print(f"2. Payload-Only (ID & Timing Ablated):")
print(f"   Accuracy: {acc2*100:.2f}% | MCC: {mcc2:.4f} | FPR: {fpr2*100:.2f}% | DoS Rec: {drec2*100:.2f}% | Tamp Rec: {trec2*100:.2f}%")
print(f"   Confusion Matrix (Normal, Attack):\n{cm2}\n")

print(f"3. Calibration-Normalized Stateful (Item 11 Fixed):")
print(f"   Accuracy: {acc3*100:.2f}% | MCC: {mcc3:.4f} | FPR: {fpr3*100:.2f}% | DoS Rec: {drec3*100:.2f}% | Tamp Rec: {trec3*100:.2f}%")
print(f"   Confusion Matrix (Normal, Attack):\n{cm3}\n")

print(f"4. Native Within-GEM Baseline (70/30 Time-Blocked):")
print(f"   Accuracy: {acc4*100:.2f}% | MCC: {mcc4:.4f} | FPR: {fpr4*100:.2f}% | DoS Rec: {drec4*100:.2f}% | Tamp Rec: {trec4*100:.2f}%")
print(f"   Confusion Matrix (Normal, Attack):\n{cm4}\n")
print(f"================================================================================")
print(f"  EVALUATION COMPLETE - ALL ARTIFACTS PROCESSED IN DOWNLOADS FOLDER")
print(f"================================================================================")