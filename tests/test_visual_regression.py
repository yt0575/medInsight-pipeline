from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest
from PIL import Image

from pipeline import core


BASELINE_DIR = Path(__file__).resolve().parent / "baselines" / "figures"


def _normalized_mae(curr: Path, base: Path) -> float:
    with Image.open(curr) as a_img, Image.open(base) as b_img:
        a = np.asarray(a_img.convert("L"), dtype=np.float32)
        b = np.asarray(b_img.convert("L"), dtype=np.float32)
    if a.shape != b.shape:
        # Keep comparison deterministic when dimensions differ slightly.
        h = min(a.shape[0], b.shape[0])
        w = min(a.shape[1], b.shape[1])
        a = a[:h, :w]
        b = b[:h, :w]
    return float(np.mean(np.abs(a - b)) / 255.0)


def test_visual_regression_if_baseline_exists(ch4_data):
    if not BASELINE_DIR.exists():
        pytest.skip("baseline figures not found; skip visual regression")

    core.generate_figures(ch4_data)

    baseline_files = sorted(BASELINE_DIR.glob("fig_*.png"))
    if not baseline_files:
        pytest.skip("baseline directory exists but no baseline figures")

    for base in baseline_files:
        curr = core.FIG_DIR / base.name
        assert curr.exists(), f"current figure missing: {base.name}"
        mae = _normalized_mae(curr, base)
        assert mae <= 0.20, f"visual regression exceeds threshold for {base.name}: mae={mae:.4f}"
