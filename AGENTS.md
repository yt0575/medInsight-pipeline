# Repository Guidelines

## Project Structure & Module Organization
This repository is template-and-data driven (not a traditional app codebase).

- `README.md`: master workflow and report constraints.
- `template.docx`: required Word template for final delivery.
- `*第四章数据.xlsx`: source dataset for Chapter 4 market analysis.
- `autofile/<疾病名>/`: generated artifacts directory (create if missing).
- `autofile/<疾病名>/figures/`: all exported figures (`fig_1_1.png`, `fig_4_1.png`, etc.).
- `autofile/<疾病名>/backup/`: timestamped backups before overwrites.

Keep all generated outputs inside `autofile/`; do not write artifacts to repository root.

## Build, Test, and Development Commands
No build system is configured. Use lightweight validation commands in PowerShell:

```powershell
# create output root
New-Item -ItemType Directory -Force autofile

# inspect expected outputs for one disease
Get-ChildItem "autofile/<疾病名>" -Recurse

# quick existence check for key files
@("00_evidence.txt","manifest_text.csv","manifest_fig.csv","ch04_agg_tables.xlsx") |
  ForEach-Object { Test-Path "autofile/<疾病名>/$_" }
```

## Coding Style & Naming Conventions
- Use UTF-8 text files and preserve Chinese domain terminology from the source materials.
- Prefer concise, deterministic filenames:
  - Chapter text: `ch01.txt` ... `ch07.txt`
  - Figures: `fig_<chapter>_<index>.png`
  - Final deliverable: `《<疾病名>疾病市场分析报告》_final.docx`
- Keep manifests machine-readable (`.csv`, stable column names, no ad-hoc renames).

## Testing & Validation Guidelines
There is no automated test framework yet. Treat validation as a release gate:

- Confirm required files exist in `autofile/<疾病名>/`.
- Verify Chapter 4 charts are derived from the Excel source and labeled consistently.
- Ensure overwrite operations create backups under `backup/`.
- Open final `.docx` and confirm TOC, heading levels, figure captions, and data-source lines render correctly.

## Commit & Pull Request Guidelines
No Git history is available in this folder, so use the following standard going forward:

- Commit format: `type(scope): summary` (for example, `docs(readme): clarify chapter-4 outputs`).
- Keep commits focused (template changes, data updates, and output examples should be separate when possible).
- PRs should include: purpose, changed files, validation steps run, and before/after screenshots for document layout changes.
- Link related issue/task IDs when available.
