#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pipeline import core


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Generate baseline figure set for visual regression tests")
    p.add_argument("--disease", default="颈椎病", help="疾病名，默认：颈椎病")
    p.add_argument("--xlsx", default=None, help="第四章Excel路径，默认：<疾病名>第四章数据.xlsx")
    p.add_argument(
        "--baseline-dir",
        default="tests/baselines/figures",
        help="baseline 输出目录，默认：tests/baselines/figures",
    )
    p.add_argument(
        "--work-out",
        default="autofile",
        help="临时工作输出根目录，默认：autofile",
    )
    return p.parse_args()


def main() -> None:
    args = parse_args()
    repo_root = ROOT
    xlsx = Path(args.xlsx) if args.xlsx else repo_root / f"{args.disease}第四章数据.xlsx"
    if not xlsx.exists():
        raise FileNotFoundError(f"Missing input excel: {xlsx}")

    core.configure_runtime(
        disease_name=args.disease,
        excel_path=xlsx,
        template_path=repo_root / "template.docx",
        out_base=repo_root / args.work_out,
    )
    core.ensure_runtime_dirs()
    ch4 = core.build_ch4_data(xlsx)
    core.generate_figures(ch4)

    baseline_dir = repo_root / args.baseline_dir
    baseline_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    for fig in sorted(core.FIG_DIR.glob("fig_*.png")):
        shutil.copy2(fig, baseline_dir / fig.name)
        count += 1
    print(f"baseline updated: {baseline_dir} ({count} files)")


if __name__ == "__main__":
    main()
