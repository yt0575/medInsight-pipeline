# medInsight Pipeline 快速入口

【医学主题占位符｜你只需改这一行】
医学主题：<<<在此填写医学主题>>>

## 1) 推荐工作流

- 把第四章数据放到 `data/` 目录，文件名固定为 `data/<医学主题>.xlsx`
- 运行单主题：`python scripts/run_pipeline.py --topic "<医学主题>"`
- 运行全量批处理：`python scripts/run_pipeline.py --all-topics`

脚本会自动完成以下动作：

- 生成证据池、参考文献与 Codex 前置提示资产
- 若缺失，自动抽取第4章结构化数据到 `autofile/<医学主题>/ch04_codex_extract.json`
- 自动写出 `fig23_codex_spec_template.json`、`fig23_codex_prompt.txt`、`figure_specs_codex_template.json`、`figure_specs_codex_prompt.txt` 等辅助文件
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

# 从 README 第一行配置读取医学主题
python scripts/run_pipeline.py --from-readme

# 指定 Excel
python scripts/run_pipeline.py --topic "肠黏膜修复" --xlsx "data/肠黏膜修复.xlsx"

# 批量处理 data 目录下全部主题
python scripts/run_pipeline.py --all-topics
```

## 5) 质量闸门

- 总字数必须在 `30000-34000`
- 各章必须达到脚本最低字数门槛
- 图表总量必须在既定范围内
- 第4章图表必须基于 `data/<医学主题>.xlsx`
- 页脚页码、图题、数据来源、参考文献链路必须全部通过 QA
- `fig_2_3` 默认使用内容相关的分层路径图，避免关系网混乱与连线压框

## 6) 说明

- `--topic` 是主参数；`--disease` 仍保留为兼容别名
- `--from-readme` 同时兼容 `医学主题：` 与旧的 `疾病名：`
- 若你把很多 Excel 放进 `data/`，直接执行 `--all-topics` 即可逐份生成并逐份检查
- 详细参数见 `scripts/USAGE.md`
- 架构说明见 `pipeline/ARCHITECTURE.md`

