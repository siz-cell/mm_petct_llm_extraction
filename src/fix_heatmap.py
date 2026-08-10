# -*- coding: utf-8 -*-
"""Fig3 heatmap — 7 PYCM fields matching Table 2/3/4 columns (no Exact Match)"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
DATA = PROJECT / "data"
OUT = PROJECT / "output" / "final"
OUT.mkdir(parents=True, exist_ok=True)

plt.rcParams.update({
    "font.family":"sans-serif","font.sans-serif":["Liberation Sans","Arial"],
    "font.size":7,"axes.linewidth":0.5,"axes.labelsize":7.5,"axes.titlesize":7.5,
    "xtick.labelsize":6.5,"ytick.labelsize":6.5,"xtick.direction":"out","ytick.direction":"out",
    "legend.fontsize":6,"legend.frameon":False,"pdf.fonttype":42,"savefig.dpi":300,
})

# 7 fields matching Table 2/3/4
FIELDS = [
    ("PET阳性/阴性","PET Positivity"),
    ("FLs-部位","FL Location"),
    ("FLs-数量","FL Count"),
    ("EM-部位","EM Location"),
    ("EM-数量","EM Count"),
    ("PM-数量","PM Count"),
    ("骨折-有/无","Fracture"),
]

# 12 models in order (same as T2)
MODELS = [
    "qwen3-235b-a22b-thinking-2507",   # Qwen3-235B-Thinking
    "qwen3-235b-a22b-Instruct-2507",   # Qwen3-235B-Instruct
    "gemini-2.5-flash-thinking",        # Gemini 2.5 Flash Thinking
    "o4-mini-2025-04-16-medium",        # o4-mini
    "deepseekr1",                       # DeepSeek-R1
    "deepseek-v3-250324",               # DeepSeek-V3
    "claude-haiku-4-5-20251001-thinking",# Claude Haiku 4.5 Thinking
    "gemini-2.5-flash-nothinking",      # Gemini 2.5 Flash
    "claude-haiku-4-5-20251001",        # Claude Haiku 4.5
    "gpt-4.1-2025-04-14",              # GPT-4.1
    "Qwen3-4B-Thinking-2507",           # Qwen3-4B-Thinking
    "Qwen3-4B-Instruct-2507",           # Qwen3-4B-Instruct
]
MODEL_EN = {
    "qwen3-235b-a22b-thinking-2507":"Qwen3-235B-Thinking",
    "qwen3-235b-a22b-Instruct-2507":"Qwen3-235B-Instruct",
    "gemini-2.5-flash-thinking":"Gemini 2.5 Flash Thinking",
    "o4-mini-2025-04-16-medium":"o4-mini",
    "deepseekr1":"DeepSeek-R1",
    "deepseek-v3-250324":"DeepSeek-V3",
    "claude-haiku-4-5-20251001-thinking":"Claude Haiku 4.5 Thinking",
    "gemini-2.5-flash-nothinking":"Gemini 2.5 Flash",
    "claude-haiku-4-5-20251001":"Claude Haiku 4.5",
    "gpt-4.1-2025-04-14":"GPT-4.1",
    "Qwen3-4B-Thinking-2507":"Qwen3-4B-Thinking",
    "Qwen3-4B-Instruct-2507":"Qwen3-4B-Instruct",
}

df = pd.read_excel(DATA/"汇总规范化数据.xlsx", sheet_name="PYCM字段级指标")

# Build matrix
matrix = []
row_labels = []
for model in MODELS:
    row = []
    for cn, en in FIELDS:
        sub = df[(df["模型"]==model)&(df["字段"]==cn)]
        if len(sub): row.append(sub.iloc[0]["F1"])
        else: row.append(np.nan)
    if not all(np.isnan(v) for v in row):
        matrix.append(row)
        row_labels.append(MODEL_EN.get(model, model))
matrix = np.array(matrix)
col_labels = [en for _, en in FIELDS]

fig, ax = plt.subplots(figsize=(11,5.5))
sns.heatmap(matrix, annot=True, fmt=".2f", xticklabels=col_labels, yticklabels=row_labels,
            cmap="coolwarm", vmin=0.35, vmax=1.0, linewidths=0.5, linecolor="white",
            cbar_kws={"label":"F1 Score","shrink":0.7}, annot_kws={"size":7}, ax=ax)
plt.setp(ax.get_xticklabels(), rotation=30, ha="right", fontsize=6.5)
plt.setp(ax.get_yticklabels(), rotation=0, fontsize=6.5)
ax.collections[0].colorbar.ax.tick_params(labelsize=6)
fig.tight_layout()
for ext in ["pdf","png"]: fig.savefig(OUT/f"fig3_heatmap.{ext}",bbox_inches="tight",dpi=300)
plt.close(fig); print("Fig3 heatmap done")

