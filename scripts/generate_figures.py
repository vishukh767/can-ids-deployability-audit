"""
generate_figures.py: High-Resolution Publication Graphics Generator

Outputs figures directly to project root and 3_results_and_figures/.
"""

from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR if (SCRIPT_DIR / "4_datasets").is_dir() else SCRIPT_DIR.parent
RESULTS_DIR = PROJECT_ROOT / "images" if (PROJECT_ROOT / "images").is_dir() else PROJECT_ROOT / "3_results_and_figures"

plt.rcParams['font.size'] = 11
plt.rcParams['figure.dpi'] = 300
plt.rcParams['savefig.dpi'] = 300
plt.rcParams['font.family'] = 'serif'

def save_fig(fig, filename):
    p = RESULTS_DIR / filename
    fig.savefig(p, bbox_inches='tight')
    print(f"Saved: {p}")
    plt.close(fig)

# FIGURE 3: HCRL Time-Blocked Confusion Matrix
def gen_fig3():
    labels = ['Normal', 'DoS', 'Fuzzy', 'Spoofing']
    cm = np.array([
        [5991, 0, 9, 0],
        [0, 6000, 0, 0],
        [14, 0, 5985, 1],
        [0, 0, 0, 6000]
    ])
    fig, ax = plt.subplots(figsize=(6, 5))
    im = ax.imshow(cm, cmap='Blues')
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    ax.set(xticks=range(4), yticks=range(4), xticklabels=labels, yticklabels=labels,
           xlabel='Predicted label', ylabel='True label')
    ax.set_title('HCRL Time-Blocked Confusion Matrix\n(Seed 42: 99.90% Acc, 0.150% FPR | 5-Seed Mean FPR: 0.091%)', fontsize=11, pad=14)
    for i in range(4):
        for j in range(4):
            ax.text(j, i, f"{cm[i, j]:,}", ha='center', va='center',
                    color='white' if cm[i, j] > cm.max()/2 else 'black')
    plt.tight_layout()
    save_fig(fig, 'figure3_hcrl_confusion_matrix.png')
    import shutil
    shutil.copyfile(RESULTS_DIR / 'figure3_hcrl_confusion_matrix.png', RESULTS_DIR / 'Figure_1_HCRL_Confusion_Matrix.png')

# FIGURE 4: Cross-Platform Confusion Matrix (Direct Transfer: Sonata -> GEM e6)
def gen_fig4():
    labels = ['Normal', 'Attack']
    # Reconciled Ground Truth: 101,039 Normal vs 42,405 Attack
    # TN = 8,612, FP = 92,427, FN = 0, TP = 42,405
    cm = np.array([
        [8612, 92427],
        [0,    42405]
    ])
    fig, ax = plt.subplots(figsize=(6, 5))
    im = ax.imshow(cm, cmap='Reds')
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    ax.set(xticks=range(2), yticks=range(2), xticklabels=labels, yticklabels=labels,
           xlabel='Predicted label', ylabel='True label',
           title='Cross-Platform Confusion Matrix (Sonata -> GEM e6)\nRaw IDs: Acc 35.57%, FPR 91.48%, MCC 0.1637')
    for i in range(2):
        for j in range(2):
            ax.text(j, i, f"{cm[i, j]:,}", ha='center', va='center',
                    color='white' if cm[i, j] > cm.max()/2 else 'black', fontweight='bold')
    plt.tight_layout()
    save_fig(fig, 'figure4_gem_can_cross_platform_confusion.png')
    import shutil
    shutil.copyfile(RESULTS_DIR / 'figure4_gem_can_cross_platform_confusion.png', RESULTS_DIR / 'Figure_3_GEM_Cross_Platform_Confusion.png')

# FIGURE 5: Feature-Space Camouflage Stress Test Curve
def gen_fig5():
    blend = [0, 25, 50, 75, 100]
    recall = [99.9, 41.8, 33.1, 31.4, 19.6]
    mcc = [0.999, 0.501, 0.428, 0.406, 0.252]
    fig, ax = plt.subplots(figsize=(7, 4.5))
    ax.plot(blend, recall, 's-', color='#B22222', linewidth=2, label='Attack Recall (%)')
    ax.plot(blend, [m*100 for m in mcc], 'o--', color='#1F4E79', linewidth=2, label='MCC (x100)')
    ax.set_xlabel('Camouflage Level (% blend toward Normal feature statistics)')
    ax.set_ylabel('Performance Metric (%)')
    ax.set_title('Degradation Under Feature-Space Camouflage\n(Real HCRL test set, unchanged trained classifier)')
    ax.grid(True, linestyle=':', alpha=0.6)
    ax.legend(loc='upper right')
    ax.set_ylim(0, 105)
    plt.tight_layout()
    save_fig(fig, 'figure5_stealth_attack_degradation.png')
    import shutil
    shutil.copyfile(RESULTS_DIR / 'figure5_stealth_attack_degradation.png', RESULTS_DIR / 'Figure_2_Stealth_Camouflage_Degradation.png')

# FIGURE 6: Cross-Platform Dissection Bar Chart (5 Configurations)
def gen_fig6():
    experiments = [
        'HCRL Native\n(11-bit)',
        'Direct Transfer\n(Raw ID)',
        'Payload-Only\n(No ID/dt)',
        'Calibration-\nNormalized',
        'GEM Native\n(29-bit)'
    ]
    acc = [99.90, 35.57, 53.62, 92.04, 100.00]
    fpr = [0.091, 91.48, 65.85, 10.65, 0.00]
    
    x = np.arange(len(experiments))
    w = 0.35
    fig, ax = plt.subplots(figsize=(9, 4.8))
    rects1 = ax.bar(x - w/2, acc, w, label='Accuracy (%)', color='#1F4E79', edgecolor='black')
    rects2 = ax.bar(x + w/2, fpr, w, label='False Positive Rate (%)', color='#D9534F', edgecolor='black')
    
    ax.set_xticks(x)
    ax.set_xticklabels(experiments, fontsize=9.5)
    ax.set_ylabel('Percentage (%)')
    ax.set_title('Cross-Platform Dissection & Representation Normalization', pad=36, fontsize=11, fontweight='bold')
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.09), ncol=2, framealpha=0.95, fontsize=9.5)
    ax.set_ylim(0, 118)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    
    # Add values on top of bars
    for rect in rects1:
        height = rect.get_height()
        ax.annotate(f'{height:.1f}%',
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=8.5, fontweight='bold')
    for rect in rects2:
        height = rect.get_height()
        ax.annotate(f'{height:.1f}%',
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=8.5, fontweight='bold', color='#B22222')
                    
    plt.tight_layout()
    save_fig(fig, 'figure6_cross_platform_comparison.png')
    
    # Also save as Figure_4_Cross_Platform_Comparison.png
    import shutil
    shutil.copyfile(RESULTS_DIR / 'figure6_cross_platform_comparison.png', RESULTS_DIR / 'Figure_4_Cross_Platform_Comparison.png')

if __name__ == '__main__':
    print("Generating all figures with synchronized empirical numbers...")
    gen_fig3()
    gen_fig4()
    gen_fig5()
    gen_fig6()
    print("All figures successfully created and synchronized.")