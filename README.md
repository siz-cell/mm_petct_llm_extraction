# mm_petct_llm_extraction

**Transforming Free-Text PET/CT Reports into Structured, Traceable, and Clinically Actionable Data Using Large Language Models in Multiple Myeloma**


## Repository Structure

```
├── README.md                    # This file
├── requirements.txt             # Python dependencies
├── AGENTS.md                    # Agent instructions for reproducibility
│
├── src/                         # Core computation scripts
│   ├── 评测llm*.py             # LLM API evaluation (Qwen, DeepSeek, Claude, Gemini, GPT, etc.)
│   ├── czjxbidui.py            # Baseline assessment comparison
│   ├── jiegouhuabidui.py       # Structured table comparison
│   ├── mrdbidui.py             # MRD comparison
│   ├── hunxiao.py              # Confusion matrix computation
│   ├── maptest.py              # Follow-up mapping
│   ├── parse_all.py            # Text block → structured data normalization
│   ├── make_table1.py          # Generate Table 1 (Patient Baseline)
│   ├── make_table6.py          # Generate Table 6 (Interpretability)
│   ├── rebuild_tables.py       # Generate Tables T2-T4, T7-T9, T11
│   ├── make_heatmap.py         # Generate Fig 2 heatmap
│   ├── make_roc_pr.py          # Generate Fig 6C ROC/PR curves
│   ├── make_fig5_cm.py         # Generate Fig 6B confusion matrices
│   ├── make_fig5_6_9.py        # Generate Fig 6/Fig 9 radar charts
│   ├── make_composite_figure.py # Generate Fig 4-5 composite
│   ├── make_concept_figs.py    # Generate Fig 1/3/8 concept diagrams
│   ├── make_remaining_figs.py  # Generate Fig 5v2/6A/7
│   ├── make_remaining_supp.py  # Generate S8,S11,S12,S13 + 7 raw data CSVs
│   ├── make_data_dict.py       # Generate S1 data dictionary
│   ├── make_charts.py          # Lesion detail charts (Fig 1-7)
│   └── fix_all.py              # Run-all fix script (tables + charts)
│
├── notebooks/                   # Computation notebooks
│   ├── 初诊基线表_czjxmrd.ipynb       # Baseline evaluation
│   ├── 结构化表格_jiegouhua.ipynb     # Structured table extraction
│   ├── 结构化表格_8shot5biao.ipynb   # 8-shot evaluation
│   ├── mrd表格_czjxmrd.ipynb          # MRD assessment
│   ├── 随访评估表_map.ipynb          # Follow-up evaluation
│   └── 结构化表格_外部验证集_jiegouhua8shot4blora.ipynb  # External validation
│
├── prompts/                     # LLM prompt templates
│   ├── qwen3235b初诊基线表输入prompt.jsonl  # Baseline assessment prompt
│   └── 结构化表格输入prompt.jsonl           # Structured table prompt
│
├── config/                      # Configuration (add model_config.yaml)
├── data/                        # Data dictionary
│   └── 数据字典_S1.xlsx         # 55-field annotation schema
│
└── output/                      # Final tables and charts
    ├── tables/                  # 19 Word tables + CSV files
    │   ├── T1-T4, T6-T9, T11   # Main tables
    │   └── S2-S5, S8-S13       # Supplementary tables
    └── charts/                  # 28 PDF/PNG figures
        ├── Fig1-8              # Main figures
        └── S2-S9               # Supplementary figures
```

## Pipeline Overview

```
Raw PET/CT Reports (Excel)
    │
    ├─→ [评测llm*.py] ──→ LLM API calls ──→ Model predictions (JSONL)
    │
    ├─→ [czjxbidui.py / jiegouhuabidui.py] ──→ Comparison files (JSONL)
    │
    ├─→ [parse_all.py] ──→ Normalized data (汇总规范化数据.xlsx)
    │
    ├─→ [rebuild_tables.py / make_table*.py] ──→ Word tables (T1-T11, S2-S13)
    │
    └─→ [make_heatmap.py / make_roc_pr.py / ...] ──→ PDF/PNG charts (Fig1-8, S2-S9)
```

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### 1. LLM Evaluation
```bash
python src/评测llmqwen.py    # Qwen models
python src/评测llmdeepseek.py # DeepSeek models
python src/评测llm.py         # Other models
```

### 2. Compare Predictions to Gold Standard
```bash
python src/czjxbidui.py       # Baseline assessment
python src/jiegouhuabidui.py  # Structured table
python src/mrdbidui.py        # MRD
python src/hunxiao.py         # Confusion matrix
```

### 3. Normalize Data
```bash
python src/parse_all.py
```

### 4. Generate All Tables and Charts
```bash
python src/fix_all.py          # Tables T6-T8 + Charts F2, F6B
python src/rebuild_tables.py   # Tables T2-T4, T7-T9, T11
python src/make_table1.py      # Table T1
python src/make_concept_figs.py # Figures F1, F3, F8
python src/make_remaining_supp.py # Supplementary S8,S11-S13 + raw data CSVs
```

## Key Data Files

- **Gold Standard**: `初诊基线表/病灶明细表与结构化表格初诊基线MRD金标准_cleaned_fixed copy.xlsx` (364 reports × 14 columns)
- **Normalized Data**: `项目内部/data/汇总规范化数据.xlsx` (20 sheets, all metrics with 95% CI)
- **Model Outputs**: JSONL files in `初诊基线表/`, `结构化表格/`, `随访评估表/`, `mrd表格/`

