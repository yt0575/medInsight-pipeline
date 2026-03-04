from __future__ import annotations

import shutil
import uuid
from pathlib import Path

import matplotlib
import pytest

matplotlib.use("Agg", force=True)

from pipeline import core


REPO_ROOT = Path(__file__).resolve().parents[1]
XLSX_PATH = REPO_ROOT / "颈椎病第四章数据.xlsx"
TEMPLATE_PATH = REPO_ROOT / "template.docx"


@pytest.fixture()
def runtime_tmp() -> Path:
    out_base = REPO_ROOT / "autofile" / "_pytest_runtime" / uuid.uuid4().hex
    out_base.mkdir(parents=True, exist_ok=True)
    core.configure_runtime(
        disease_name="颈椎病",
        excel_path=XLSX_PATH,
        template_path=TEMPLATE_PATH,
        out_base=out_base,
    )
    core.ensure_runtime_dirs()
    try:
        yield out_base
    finally:
        shutil.rmtree(out_base, ignore_errors=True)


@pytest.fixture()
def ch4_data(runtime_tmp: Path):
    _ = runtime_tmp
    return core.build_ch4_data(XLSX_PATH)
