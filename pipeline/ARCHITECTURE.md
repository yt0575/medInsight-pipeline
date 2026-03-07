# Pipeline Architecture

## 核心定位

本仓库已从“单病种报告”调整为“医学主题市场分析报告”流水线。

- 输入主键：`医学主题`
- 第4章数据：`data/<医学主题>.xlsx`
- 输出目录：`autofile/<医学主题>/`

## 主要模块

- `pipeline/core.py`
  - 统一管理 CLI、运行时配置、Excel 解析、Codex 提示资产、图表生成、DOCX 装配与 QA
  - 提供 `run`、`run_batch`、`run_stage1` ~ `run_stage5` 等 API
- `scripts/run_pipeline.py`
  - 轻量 CLI 入口，直接委托 `pipeline.core.main`
- `pipeline/disease_profiles.json`
  - 保留既有主题画像和图表规则；仍可用于 profile 匹配与 `fig_2_3` 语义控制

## 执行流程

1. 通过 `--topic` 或 `--from-readme` 解析医学主题名。
2. 若未显式传 `--xlsx`，默认定位到 `data/<医学主题>.xlsx`。
3. `run()` 先执行 `stage1`，生成证据池和参考文献。
4. 若缺失第4章结构化 JSON，自动调用标准 sheet 解析器生成 `ch04_codex_extract.json`。
5. 若工作簿仅提供医院端 sheet，自动把缺失渠道按 `0` 补齐到 `quarterly_channel`，并将缺失渠道的 Top10/CR5 保持为空。
6. 自动生成 Codex 前置资产，包括正文蓝图、`fig23_codex_spec_template.json`、`fig23_codex_prompt.txt`、`figure_specs_codex_template.json` 与 `figure_specs_codex_prompt.txt`。
7. 由当前 Codex 会话主导写入 `ch01~ch07.txt`、`summary.txt`、`fig23_codex_spec.json`，并按需回写 `figure_specs.json`。
8. 进入 assist 链路：`stage3(图表+清单) -> stage4(docx) -> stage5(QA)`。

## 批处理模式

`run_batch()` 会遍历 `data/*.xlsx`：

- 每个文件 stem 作为医学主题名
- 每个主题独立生成到 `autofile/<医学主题>/`
- 每个主题独立执行 QA
- 汇总结果写入 `autofile/batch_report_summary.csv`

## 质量控制

- TXT 闸门：总字数、分章字数、事实锚点、引用覆盖、重复句、医学密度
- 图表闸门：标题、来源、结构规则、关键图视觉可读性
- DOCX 闸门：页脚页码、图题、来源行、参考文献、结构一致性

## 扩展建议

- 新增主题规则时，优先扩展 `pipeline/disease_profiles.json`
- 新增图表语义时，优先复用 `figure_specs.json` 与 `fig23_codex_spec.json`
- 新增批处理策略时，优先扩展 `run_batch()`，避免分散到多个脚本

