"""
Synthetic CAN-bus intrusion-detection proof-of-concept.

IMPORTANT: This is a SYNTHETIC dataset generated to mimic the statistical
structure of the well-known CAN-intrusion-dataset (HCRL) and can-train-and-test
benchmarks, because this environment has no internet access to download the
real captured vehicle data. Attack signatures (DoS = abnormally high-frequency
identical CAN ID; Fuzzy = random ID + random payload; Spoofing = legitimate ID
reused with implausible payload) are modeled after the documented behavior of
those real datasets, but the numbers below are NOT from a real vehicle.
"""

import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

rng = np.random.default_rng(42)

# --- 1. Define a small set of "normal" CAN IDs, each with a typical period (ms) ---
normal_ids = {
    0x0C0: 10,   # e.g., wheel speed
    0x120: 20,   # e.g., engine RPM
    0x140: 50,   # e.g., steering angle
    0x316: 100,  # e.g., dashboard/odometer
    0x329: 100,  # e.g., gear position
}

def gen_normal_traffic(duration_ms):
    rows = []
    for can_id, period in normal_ids.items():
        t = rng.uniform(0, period)
        base_val = rng.integers(20, 200)
        while t < duration_ms:
            jitter = rng.normal(0, period * 0.03)
            data_bytes = np.clip(
                base_val + rng.normal(0, 3, size=8), 0, 255
            ).astype(int)
            rows.append((t, can_id, 8, *data_bytes, "Normal"))
            t += period + jitter
    return rows

def gen_dos_attack(start_ms, duration_ms):
    rows = []
    t = start_ms
    while t < start_ms + duration_ms:
        data_bytes = np.clip(rng.normal(2, 2, size=8), 0, 255).astype(int)
        rows.append((t, 0x000, 8, *data_bytes, "DoS"))
        t += rng.uniform(0.3, 1.2)  # flood, but with some jitter/variability
    return rows

def gen_fuzzy_attack(start_ms, duration_ms, n_msgs):
    rows = []
    times = np.sort(rng.uniform(start_ms, start_ms + duration_ms, n_msgs))
    for t in times:
        can_id = int(rng.integers(0x000, 0x7FF))
        data_bytes = rng.integers(0, 256, size=8)
        rows.append((t, can_id, 8, *data_bytes, "Fuzzy"))
    return rows

def gen_spoofing_attack(start_ms, duration_ms, target_id=0x120):
    # Reuses a legitimate ID (engine RPM) at a pinned value, but only moderately
    # faster than normal — a subtler, more realistic spoofing signature than an
    # obvious flood, so it is not trivially separable from normal traffic.
    rows = []
    t = start_ms
    spoof_val = 180  # plausible-looking but pinned value
    while t < start_ms + duration_ms:
        data_bytes = np.clip(spoof_val + rng.normal(0, 5, size=8), 0, 255).astype(int)
        rows.append((t, target_id, 8, *data_bytes, "Spoofing"))
        t += rng.uniform(8, 15)  # only moderately faster than the real 20ms period
    return rows

# --- 2. Assemble a timeline: normal traffic with attack windows injected ---
TOTAL_MS = 60_000  # 60 seconds of simulated bus traffic
all_rows = gen_normal_traffic(TOTAL_MS)

attack_windows = [
    (10_000, 2_000, "dos"),
    (20_000, 2_000, "fuzzy"),
    (35_000, 2_000, "spoofing"),
    (45_000, 1_500, "dos"),
    (52_000, 1_500, "fuzzy"),
]
for start, dur, kind in attack_windows:
    if kind == "dos":
        all_rows += gen_dos_attack(start, dur)
    elif kind == "fuzzy":
        all_rows += gen_fuzzy_attack(start, dur, n_msgs=int(dur * 2))
    elif kind == "spoofing":
        all_rows += gen_spoofing_attack(start, dur)

cols = ["t_ms", "can_id", "dlc"] + [f"b{i}" for i in range(8)] + ["label"]
df = pd.DataFrame(all_rows, columns=cols).sort_values("t_ms").reset_index(drop=True)

# --- 3. Feature engineering (mirrors common CAN-IDS feature choices in the literature) ---
df["dt_since_last_same_id"] = df.groupby("can_id")["t_ms"].diff().fillna(1000)
df["byte_mean"] = df[[f"b{i}" for i in range(8)]].mean(axis=1)
df["byte_std"] = df[[f"b{i}" for i in range(8)]].std(axis=1).fillna(0)

# Add measurement/sensor noise so classes are not perfectly separable
df["dt_since_last_same_id"] += rng.normal(0, 1.5, size=len(df))
df["byte_mean"] += rng.normal(0, 6, size=len(df))

# Simulate realistic label noise (~2%): a few messages are ambiguous/mislabeled,
# mirroring the imperfect ground truth common in real captured CAN datasets.
noise_idx = rng.choice(df.index, size=int(0.02 * len(df)), replace=False)
df.loc[noise_idx, "label"] = rng.choice(
    ["Normal", "DoS", "Fuzzy", "Spoofing"], size=len(noise_idx)
)

feature_cols = ["can_id", "dlc", "dt_since_last_same_id", "byte_mean", "byte_std"]
X = df[feature_cols]
y = df["label"]

print("Class distribution:\n", y.value_counts())

# --- 4. Train / test split + Random Forest classifier ---
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.3, random_state=42, stratify=y
)

clf = RandomForestClassifier(n_estimators=150, max_depth=12, random_state=42, n_jobs=-1)
clf.fit(X_train, y_train)
y_pred = clf.predict(X_test)

acc = accuracy_score(y_test, y_pred)
report = classification_report(y_test, y_pred, digits=3)
print(f"\nOverall accuracy: {acc:.4f}\n")
print(report)

labels_order = ["Normal", "DoS", "Fuzzy", "Spoofing"]
cm = confusion_matrix(y_test, y_pred, labels=labels_order)

# Save results to files for the write-up
with open("results.txt", "w") as f:
    f.write(f"Overall accuracy: {acc:.4f}\n\n")
    f.write(report)
    f.write("\nConfusion matrix (rows=true, cols=predicted), order = " + str(labels_order) + "\n")
    f.write(str(cm))

# --- 5. Plot confusion matrix ---
fig, ax = plt.subplots(figsize=(5.2, 4.6))
im = ax.imshow(cm, cmap="Blues")
ax.set_xticks(range(len(labels_order))); ax.set_xticklabels(labels_order)
ax.set_yticks(range(len(labels_order))); ax.set_yticklabels(labels_order)
ax.set_xlabel("Predicted label"); ax.set_ylabel("True label")
ax.set_title("Confusion Matrix — Vehicle-Layer IDS\n(synthetic CAN-bus proof-of-concept)")
for i in range(len(labels_order)):
    for j in range(len(labels_order)):
        val = cm[i, j]
        ax.text(j, i, str(val), ha="center", va="center",
                color="white" if val > cm.max() / 2 else "black", fontsize=11)
fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
plt.tight_layout()
plt.savefig("confusion_matrix.png", dpi=150)
print("\nSaved confusion matrix to confusion_matrix.png")

# Per-class accuracy for a simple bar chart too
from sklearn.metrics import precision_recall_fscore_support
prec, rec, f1, support = precision_recall_fscore_support(y_test, y_pred, labels=labels_order)
print("\nPer-class precision/recall/f1:")
for lbl, p, r, f in zip(labels_order, prec, rec, f1):
    print(f"  {lbl:10s} precision={p:.3f} recall={r:.3f} f1={f:.3f}")

with open("per_class.csv", "w") as f:
    f.write("label,precision,recall,f1,support\n")
    for lbl, p, r, fsc, s in zip(labels_order, prec, rec, f1, support):
        f.write(f"{lbl},{p:.4f},{r:.4f},{fsc:.4f},{s}\n")
