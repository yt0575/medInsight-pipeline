# Pipeline Architecture

## Consolidated layout

- `pipeline/core.py`
  - Single runtime module for full pipeline behavior.
  - Includes runtime configuration, stage dispatch, role dispatch, and stage implementations.
  - Owns parsing, text generation, figure generation, DOCX assembly, and QA checks.
  - Provides stable APIs: `run`, `run_stage1` ... `run_stage5`, role runner `run_role_pipeline`, and CLI `main`.

- `generate_report.py`
  - Backward-compatible wrapper entrypoint.
  - Delegates execution to `pipeline.core.main`.

## Execution flow

1. CLI arguments are parsed by `pipeline.core.parse_args`.
2. Disease name is resolved from `--disease` or `--from-readme --readme <path>`.
3. `pipeline.core.run` configures runtime paths (disease/xlsx/template/out-base).
4. If `--role` is provided, role dispatch executes sequential agents in one process:
   - `evidence -> content -> docx -> qa`
5. Otherwise, stage dispatch runs one stage or all stages directly inside `pipeline.core`.
6. Artifacts are written to `autofile/<疾病名>/`.
