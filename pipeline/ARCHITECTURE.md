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

## 当前真实流程

```mermaid
flowchart LR
  A["CLI: --topic / --all-topics"] --> B["configure_runtime"]
  B --> C["Stage1: 证据池 + refs + Codex 前置资产"]
  C --> D["Codex Prep: ch04 结构化抽取 + 叙事摘要 + gap/precheck"]
  D --> E["Codex 会话写正文 / summary / fig23 / 可选 figure_specs"]
  E --> F["Stage3: 图表 + manifest + txt gate"]
  F --> G["Stage4: 组装 final.docx"]
  G --> H["Stage5: 最终 QA"]
  H --> I["autofile/<医学主题>/"]
```

## 执行流程

1. 通过 `--topic` 或 `--from-readme` 解析医学主题名。
2. 若未显式传 `--xlsx`，默认定位到 `data/<医学主题>.xlsx`。
3. `run()` 固定执行 3 段：
   - `run_stage1_evidence()`：生成 `00_evidence.txt`、`refs.txt` 的通用 scaffold 和 Codex 前置提示资产；若当前主题已有 Codex 修订版 evidence/refs，则优先保留。
   - `ensure_codex_prep_assets_ready()`：准备第 4 章结构化数据和写作辅助文件。
   - `run_assist_pipeline()`：串行执行 `stage3 -> stage4 -> stage5`。
4. 第 4 章结构化数据来自 `ch04_codex_extract.json`：
   - 若不存在，脚本会从 `data/<医学主题>.xlsx` 自动抽取；
   - 若工作簿只提供医院端 sheet，脚本会把缺失渠道按 `0` 补齐到 `quarterly_channel`，缺失渠道的 Top10/CR5 保持为空。
5. `ensure_codex_prep_assets_ready()` 会生成或刷新以下辅助文件：
   - `ch04_excel_profile.txt`
   - `ch04_data_dictionary.txt`
   - `ch04_narrative_brief.txt`
   - `codex_gap_panel.txt`
   - `chapter_precheck.txt`
6. `write_codex_preflight_assets()` 会生成或刷新以下 Codex 写作/改写资产：
   - `codex_content_blueprint.txt`
   - `codex_rewrite_prompt.txt`
   - `fig23_codex_spec_template.json`
   - `fig23_codex_prompt.txt`
   - `figure_specs_codex_template.json`
   - `figure_specs_codex_prompt.txt`
   - `semantic_review_prompt.txt`
7. 当前流程是 **Codex-first**：
   - 脚本不再自动代写 `ch01~ch07.txt` 或 `summary.txt`
   - 脚本也不再自动代写 `fig23_codex_spec.json`
   - 这些文件必须由当前 Codex 会话主导写入
8. `stage3` 在出图前会先做 TXT 闸门：
   - 正文必须存在
   - `summary.txt` 必须存在
   - `fig23_codex_spec.json` 必须存在且可解析
   - 之后才会生成图表、`manifest_text.csv`、`manifest_fig.csv` 和 `ch04_agg_tables.xlsx`
9. `stage4` 负责把正文、summary、refs、图表和模板装配到最终 `final.docx`。
10. `stage5` 负责最终 QA，并输出 `qa_check.txt`。

## 写作辅助文件

- `codex_gap_panel.txt`
  - 写作前差值面板
  - 告诉当前总字数、各章缺口、各 block 缺口、优先补位顺序
- `chapter_precheck.txt`
  - 章节级轻量预检
  - 提前暴露 `PASS / WARN / FAIL`、引用、锚点覆盖、医学密度等问题
- `ch04_narrative_brief.txt`
  - 第 4 章叙事摘要
  - 把最新季度、YoY、长期趋势、Top 品种、CR5 等数字整理成可直接写作的 bullet brief

## 字数闸门

- 全文总字数仍要求在 `30000-34000`
- 分章字数增加了容差：
  - 默认允许 **低于章节下限 100 字内通过**
  - 只有超过这个容差，才会被视为必须补写
- 这条规则会同时体现在：
  - `codex_gap_panel.txt`
  - `chapter_precheck.txt`
  - `txt_stage_qa.txt`
  - `qa_check.txt`

## 推荐工作流

1. 先运行 `stage1` / `run()`，让脚本准备证据池和辅助文件。
2. 先看：
   - `codex_gap_panel.txt`
   - `chapter_precheck.txt`
   - `ch04_narrative_brief.txt`
3. 由当前 Codex 会话写：
   - `ch01.txt` ~ `ch07.txt`
   - `summary.txt`
   - `fig23_codex_spec.json`
   - 如有必要，再写 `figure_specs.json`
4. 再执行 `stage3 -> stage4 -> stage5`。

## 批处理注意事项

- `run_batch()` 会遍历 `data/*.xlsx`
- 但由于当前流程已改成 Codex-first，**批处理不再等于脚本自动代写正文**
- 如果某个主题缺少：
  - `ch01~ch07.txt`
  - `summary.txt`
  - `fig23_codex_spec.json`
  那么该主题会在 `stage3` 前置闸门处失败
- 因此，批处理更适合：
  - 正文已经由 Codex 写好之后统一装配；或
  - 在明确接受“先批量准备辅助资产，再逐主题写稿”的前提下使用

## 批处理模式

`run_batch()` 会遍历 `data/*.xlsx`：

- 每个文件 stem 作为医学主题名
- 每个主题独立生成到 `autofile/<医学主题>/`
- 每个主题独立执行 QA
- 汇总结果写入 `autofile/batch_report_summary.csv`

## 质量控制

- TXT 闸门：总字数、分章字数（带 100 字容差）、事实锚点、引用覆盖、重复句、医学密度
- 图表闸门：标题、来源、结构规则、关键图视觉可读性
- DOCX 闸门：页脚页码、图题、来源行、参考文献、结构一致性

## 图表渲染策略

- 图表 PNG 统一按 Word 插入宽度导出，减少插入 `.docx` 后因原始宽度差异造成的缩放差异。
- `图表x-x：xxx` 主标题与图片内嵌 `数据来源：xxxx` 使用统一样式；图内数据层文字可更小，但必须清晰可读。
- 流程图、时间线图、关系/机制图与定量图分开布局，不能把同一模板硬套到所有图表。
- 当流程图、时间线图或关系图的节点/事件标签过长时，优先自动换行、多排和避让，而不是压缩到单排。

## 扩展建议

- 新增主题规则时，优先扩展 `pipeline/disease_profiles.json`
- 新增图表语义时，优先复用 `figure_specs.json` 与 `fig23_codex_spec.json`
- 新增批处理策略时，优先扩展 `run_batch()`，避免分散到多个脚本
