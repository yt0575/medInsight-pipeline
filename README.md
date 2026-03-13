# medInsight Pipeline 快速入口

【医学主题占位符｜你只需改这一行】
医学主题：<<<在此填写医学主题>>>

## 1) 推荐工作流

- 把第四章数据放到 `data/` 目录，文件名固定为 `data/<医学主题>.xlsx`
- 运行单主题：`python scripts/run_pipeline.py --topic "<医学主题>"`
- 运行全量批处理：`python scripts/run_pipeline.py --all-topics`

脚本会自动完成以下动作：

- 生成证据池 scaffold、参考文献 scaffold 与 Codex 前置提示资产；若当前主题已有 Codex 会话修订版 `00_evidence.txt` / `refs.txt`，流程会优先保留
- 若缺失，自动抽取第4章结构化数据到 `autofile/<医学主题>/ch04_codex_extract.json`
- 自动写出 `fig23_codex_spec_template.json`、`fig23_codex_prompt.txt`、`figure_specs_codex_template.json`、`figure_specs_codex_prompt.txt`、`codex_gap_panel.txt`、`chapter_precheck.txt`、`ch04_narrative_brief.txt` 等辅助文件
- 自动写出 `codex_next_actions.txt`、`codex_block_cards.txt` 与 `codex_block_cards/`，便于按优先级和单 block 快速写稿或局部返工
- 由当前 Codex 会话主导写入 `ch01.txt` ~ `ch07.txt`、`summary.txt`、`fig23_codex_spec.json`，并按需回写 `figure_specs.json`
- 装配最终 Word，并执行严格 QA

## 2) 输入要求

- `template.docx`
- `data/<医学主题>.xlsx`
- 若某些主题仅提供 `医院品类/医院top`，流程会自动把缺失渠道按 0 补齐到第4章结构化数据，并在 `ch04_sheet_map.txt` 与 `ch04_codex_review.txt` 中标记

## 3) 输出位置

所有产物统一写入：

- `autofile/<医学主题>/`

核心交付物：

- `autofile/<医学主题>/《<医学主题>市场分析报告》_final.docx`
- `autofile/<医学主题>/qa_check.txt`
- `autofile/<医学主题>/ch04_codex_extract.json`
- `autofile/<医学主题>/figures/`

批量运行后还会生成：

- `autofile/batch_report_summary.csv`

## 4) 常用命令

```powershell
# 单主题，默认读取 data/<医学主题>.xlsx
python scripts/run_pipeline.py --topic "肠黏膜修复"

# 从 README 中的 `医学主题：` 配置行读取医学主题
python scripts/run_pipeline.py --from-readme

# 只刷新当前正文的写稿进度资产，不跑出图和 docx
python scripts/run_pipeline.py --topic "肠黏膜修复" --refresh-progress

# 指定 Excel
python scripts/run_pipeline.py --topic "肠黏膜修复" --xlsx "data/肠黏膜修复.xlsx"

# 批量处理 data 目录下全部主题
python scripts/run_pipeline.py --all-topics
```

## 5) 质量闸门

- 正式交付字数、分章门槛、三级标题与本章小结规则见 `AGENTS.md`
- 图表总量、第4章数据来源、页脚页码、图题、数据来源、参考文献链路需全部通过 QA
- `fig_2_3` 默认使用内容相关的分层路径图，避免关系网混乱与连线压框
- 图表图片统一按 Word 插入宽度导出；`图表x-x：xxx` 与 `数据来源：xxxx` 使用统一样式；流程图/时间线/关系图优先多排而非强行单排

## 6) 说明

- `--topic` 是主参数；`--disease` 仍保留为兼容别名
- `--from-readme` 同时兼容 `医学主题：` 与旧的 `疾病名：`
- 若你把很多 Excel 放进 `data/`，直接执行 `--all-topics` 即可逐份生成并逐份检查
- 详细参数见 `scripts/USAGE.md`
- 架构说明见 `pipeline/ARCHITECTURE.md`

