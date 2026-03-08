# Usage

## 单主题运行

```powershell
python scripts/run_pipeline.py --topic "肠黏膜修复"
```

默认会读取：

- `data/肠黏膜修复.xlsx`

默认会输出到：

- `autofile/肠黏膜修复/`

## 批量运行

```powershell
python scripts/run_pipeline.py --all-topics
```

脚本会遍历 `data/*.xlsx`，并按文件名 stem 作为医学主题名逐份生成。

## 参数说明

- `--topic "<医学主题>"`：主参数，指定医学主题
- `--disease "<医学主题>"`：兼容旧参数，等同于 `--topic`
- `--all-topics`：遍历 `data` 目录下全部 `*.xlsx`
- `--data-dir "data"`：批量模式的数据目录
- `--from-readme`：从 `README.md` 读取 `医学主题：`；同时兼容旧的 `疾病名：`
- `--readme "README.md"`：指定读取配置的 README
- `--xlsx "data/<医学主题>.xlsx"`：覆盖默认 Excel 路径
- `--template "template.docx"`：覆盖模板路径
- `--out-base "autofile"`：覆盖输出根目录
- `--lite-output`：清理中间产物，仅保留最终 docx、QA 与日志

若某个 Excel 只包含 `医院品类/医院top`，脚本会自动把缺失的药店端/线上端季度值按 `0` 补齐，并将缺失渠道的 Top10/CR5 保持为空，同时在 `ch04_sheet_map.txt` 中标记。

## 自动准备行为

脚本会自动准备以下确定性或提示性资产：

- `00_evidence.txt`
- `refs.txt`
- `ch04_codex_extract.json`
- `fig23_codex_spec_template.json`
- `fig23_codex_prompt.txt`
- `figure_specs_codex_template.json`
- `figure_specs_codex_prompt.txt`
- `semantic_review_prompt.txt`
- `codex_gap_panel.txt`
- `chapter_precheck.txt`
- `ch04_narrative_brief.txt`

以下文件改为由当前 Codex 会话主导写入，不再由脚本自动代写：

- `fig23_codex_spec.json`
- `figure_specs.json`（按需）
- `ch01.txt` ~ `ch07.txt`
- `summary.txt`

## 验收输出

单主题运行后重点检查：

- `autofile/<医学主题>/《<医学主题>市场分析报告》_final.docx`
- `autofile/<医学主题>/qa_check.txt`

批量运行后重点检查：

- `autofile/batch_report_summary.csv`

若某个主题失败，汇总表中会给出失败原因，CLI 也会在最后抛出汇总错误。

