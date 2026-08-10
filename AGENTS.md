# AGENTS.md — npj Digital Medicine 论文数据项目

## 这是什么

多发性骨髓瘤 PET/CT 报告 LLM 结构化提取论文的数据处理与图表生成项目，目标期刊 **npj Digital Medicine**。不是软件工程项目——没有 build/test/lint。

## 目录结构

```
10112/bidui/jiegouh/多模型评测/
├── 初诊基线表/           # 模型输出 JSONL + 比对信息 (14模型)
├── 结构化表格/           # JSONL + 外部验证集评测
├── 随访评估表/           # 随访 JSONL
├── mrd表格/              # MRD JSONL
├── seaborn图表/
│   ├── 项目内部/          # ★ 主工作区
│   │   ├── data/          # 输入数据
│   │   ├── scripts/       # 生成脚本
│   │   ├── output/        # 最终产出
│   │   │   ├── charts/    # PDF/PNG/SVG 图表
│   │   │   ├── tables_final/  # Word 表格 (T1-T11)
│   │   │   └── data/      # 数据字典、映射文档、清单
│   │   └── NPJ_STYLE_GUIDE.md  # 配色/字号规范
│   ├── 病灶明细表图表/    # 病灶明细独立项目
│   ├── tupian/            # 早期图表 + notebook
│   └── 图表/表格word/     # 早期 Word 表
```

## 核心数据文件

| 文件 | 路径 | 用途 |
|------|------|------|
| 汇总规范化数据.xlsx | 项目内部/data/ | 20 sheets，所有归一化数据 |
| 临时对应图.xlsx | 项目内部/data/ | PYCM 文本块 + 结构化表 |
| 临时对应图表.xlsx | 项目内部/data/ | Word 表格源 + 外部验证 |
| metrics_overall.csv | 项目内部/data/ | 病灶明细 16 配置整体指标 |
| metrics_by_field.csv | 项目内部/data/ | 病灶明细 80 行字段级指标 |
| 金标准 Excel | 初诊基线表/ | 364 报告 × 14 列，患者+报告级数据 |

## 关键脚本

| 脚本 | 功能 |
|------|------|
| `scripts/rebuild_tables.py` | ★ 生成 T2/T3/T4/T7/T8/T9/T11 (10 列 PYCM F1 + CI) |
| `scripts/make_table1.py` | T1 患者基线 |
| `scripts/make_table6.py` | T6 可解释性 |
| `scripts/make_data_dict.py` | S1 数据字典 |
| `scripts/make_heatmap.py` | Fig3 热力图 |
| `scripts/make_roc_pr.py` | Fig4 ROC/PR |
| `scripts/make_fig5_cm.py` | Fig6B 4 类混淆矩阵 |
| `scripts/make_fig5_6_9.py` | Fig6+Fig9 雷达图 |
| `scripts/make_composite_figure.py` | Fig_composite 综合大图 |
| `scripts/parse_all.py` | 文本块 → 结构化数据 |

## 输入输出约定

- **输出路径**: 始终用相对 `PROJECT / "output" / ...`，不硬编码绝对路径
- **SVG**: 必须加 `"svg.fonttype": "none"` + 不传 `dpi`
- **Word 表格**: 用 `docx` 库 + npj 三线表边框
- **数据格式**: 全部 0-1 小数，不出现百分号

## 模型命名规范

标准英文名（表中使用）：
- `Qwen3-235B-Instruct` / `Qwen3-235B-Thinking`
- `Qwen3-4B-Instruct` / `Qwen3-4B-Thinking`
- `DeepSeek-V3` / `DeepSeek-R1`
- `Claude Haiku 4.5` / `Claude Haiku 4.5 Thinking`
- `Gemini 2.5 Flash` / `Gemini 2.5 Flash Thinking`
- `GPT-4.1` / `o4-mini`
- 变体加后缀: `0-shot`, `8-shot`, `LoRA 8-shot`

## 常见陷阱

1. **Python-docx 文件锁**: `doc.save()` 后文件可能被防病毒锁定。重用脚本前手动关掉 Word，或用新文件名/目录。
2. **模型名不一致**: PYCM 用 `deepseekr1`（无连字符），结构化表用 `deepseek-r1`（有连字符）。`get_f1()` 里有交叉映射表，新增模型需同步更新。
3. **Excel 读取慢**: 金标准 1.6MB，pandas read_excel 可能超时。用 `openpyxl(read_only=True, data_only=True)` 只读需要的列。
4. **Inline Python 中文问题**: `python -c` 含中文时常语法错误，写 `.py` 脚本文件执行。
5. **百分比转换**: 仅 `Exact Match Accuracy`、`Sensitivity`、`Specificity` 等明确为 % 的列需 /100。`SUVmax RMSE`（可 >1）和 `Kappa`（0-1）是例外。
6. **output/tables_final/ vs output/tables/**: 旧表在 `tables/` 可能被锁，新脚本输出到 `tables_final/`。

## 当前进度

17/23 完成。产出清单: `output/data/清单.csv`。剩余: F1/F3/F5v2/F6A/F7/F8 图表。
