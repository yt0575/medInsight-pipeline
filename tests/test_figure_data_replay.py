from __future__ import annotations

import math
from pathlib import Path

import matplotlib.pyplot as plt
import pytest

from pipeline import core


def _assert_series_close(actual: list[float], expected: list[float], name: str) -> None:
    assert len(actual) == len(expected), f"{name} length mismatch"
    for idx, (a, e) in enumerate(zip(actual, expected)):
        if math.isnan(e):
            assert math.isnan(a), f"{name}[{idx}] expected NaN, got {a}"
        else:
            assert a == pytest.approx(e, rel=1e-6, abs=1e-6), f"{name}[{idx}] mismatch"


def test_ch4_figures_replay_data_points(monkeypatch, ch4_data):
    captured: dict[str, dict[str, list[float]] | list[float]] = {}

    def fake_save(path: Path, fig) -> None:
        name = Path(path).name
        ax = fig.axes[0]
        if name == "fig_4_1.png":
            lines = ax.get_lines()
            captured["fig_4_1"] = {
                line.get_label(): [float(v) for v in line.get_ydata()]
                for line in lines
                if line.get_label() in {"医院端", "药店端", "线上端"}
            }
        elif name == "fig_4_4.png":
            captured["fig_4_4"] = [float(p.get_height()) for p in ax.patches]
        elif name == "fig_4_8.png":
            captured["fig_4_8"] = [float(p.get_height()) for p in ax.patches]
        plt.close(fig)

    monkeypatch.setattr(core, "save_figure", fake_save)

    fig_rows = core.generate_figures(ch4_data)
    fig_map = {r["fig_id"]: r for r in fig_rows}

    assert fig_map["fig_4_1"]["excel_sheet_or_table"] == "quarterly_channel"
    assert fig_map["fig_4_4"]["excel_sheet_or_table"] == "latest_yoy"
    assert fig_map["fig_4_8"]["excel_sheet_or_table"] == "cr5_latest"

    q = ch4_data.quarterly
    replay_41 = captured["fig_4_1"]
    assert isinstance(replay_41, dict)
    _assert_series_close(replay_41["医院端"], q["hospital"].astype(float).tolist(), "fig_4_1.hospital")
    _assert_series_close(replay_41["药店端"], q["drugstore"].astype(float).tolist(), "fig_4_1.drugstore")
    _assert_series_close(replay_41["线上端"], q["online"].astype(float).tolist(), "fig_4_1.online")

    yoy_expected = ch4_data.yoy_latest["yoy_pct"].astype(float).tolist()
    replay_44 = captured["fig_4_4"]
    assert isinstance(replay_44, list)
    _assert_series_close(replay_44, yoy_expected, "fig_4_4.yoy")

    cr5_expected = ch4_data.cr5_latest["cr5_pct"].astype(float).tolist()
    replay_48 = captured["fig_4_8"]
    assert isinstance(replay_48, list)
    _assert_series_close(replay_48, cr5_expected, "fig_4_8.cr5")
