from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image

from pipeline import core


def test_generate_figures_outputs_all_png_files(ch4_data):
    fig_rows = core.generate_figures(ch4_data)
    assert len(fig_rows) == 26

    for row in fig_rows:
        fig_path = core.FIG_DIR / row["输出文件名"]
        assert fig_path.exists(), f"missing figure file: {fig_path.name}"


def test_generated_images_are_not_blank(ch4_data):
    fig_rows = core.generate_figures(ch4_data)

    for row in fig_rows:
        fig_path = core.FIG_DIR / row["输出文件名"]
        with Image.open(fig_path) as img:
            arr = np.asarray(img.convert("L"), dtype=np.float32)
            assert img.width >= 400 and img.height >= 240, f"unexpected small image: {fig_path.name}"
            assert float(arr.std()) > 2.0, f"image appears blank: {fig_path.name}"
            non_white_ratio = float((arr < 250).mean())
            assert non_white_ratio > 0.01, f"image appears empty/over-white: {fig_path.name}"


def test_chapter4_source_lines_are_miwang(ch4_data):
    fig_rows = core.generate_figures(ch4_data)
    ch4_rows = [r for r in fig_rows if r["fig_id"].startswith("fig_4_")]
    assert len(ch4_rows) == 8

    for row in ch4_rows:
        assert row["source_line"] == "数据来源：米内网"
        assert row["excel_sheet_or_table"] in {
            "quarterly_channel",
            "latest_share",
            "annual_channel",
            "latest_yoy",
            "top10_hospital",
            "top10_drugstore",
            "top10_online",
            "cr5_latest",
        }
