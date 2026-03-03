# Pipeline Usage

## One-command full run

```powershell
python scripts/run_pipeline.py --disease "<疾病名>" --stage all
```

Or run with lightweight single-process role mode:

```powershell
python scripts/run_pipeline.py --disease "<疾病名>" --role all
```

Or run directly from README disease config:

```powershell
python scripts/run_pipeline.py --from-readme --role all
```

Optional arguments:

- `--xlsx "<疾病名>第四章数据.xlsx"`: override default chapter-4 excel path
- `--template "template.docx"`: override Word template
- `--out-base "autofile"`: override output root
- `--reuse-text` / `--no-reuse-text`: stage3 reuse existing `ch01~ch07` + `summary` by default; disable reuse to regenerate
- `--role "all/evidence/content/docx/qa"`: role mode (higher priority than `--stage`)
- `--from-readme`: read disease name from `README.md` line `疾病名：...` when `--disease` is omitted
- `--readme "README.md"`: override README path used by `--from-readme`

## Run by stages

```powershell
# Stage 1: evidence + refs
python scripts/run_pipeline.py --disease "<疾病名>" --stage 1

# Stage 2: chapter text + txt QA
python scripts/run_pipeline.py --disease "<疾病名>" --stage 2

# Stage 3: chapter-4 data line + figures + manifests
python scripts/run_pipeline.py --disease "<疾病名>" --stage 3

# Stage 4: assemble final docx from existing artifacts
python scripts/run_pipeline.py --disease "<疾病名>" --stage 4

# Stage 5: final QA
python scripts/run_pipeline.py --disease "<疾病名>" --stage 5
```

## Run by roles (single process, sequential)

```powershell
# Evidence Agent (stage1)
python scripts/run_pipeline.py --disease "<疾病名>" --role evidence

# Content Agent (stage2+stage3)
python scripts/run_pipeline.py --disease "<疾病名>" --role content

# Docx Agent (stage4)
python scripts/run_pipeline.py --disease "<疾病名>" --role docx

# QA Agent (stage5)
python scripts/run_pipeline.py --disease "<疾病名>" --role qa
```

## Backward-compatible entry

`generate_report.py` is now configurable:

```powershell
python generate_report.py --disease "<疾病名>" --stage all
```

Implementation modules:

- `pipeline/core.py`: core generation + QA logic
- `scripts/run_pipeline.py`: unified script entry for stage mode and role mode

