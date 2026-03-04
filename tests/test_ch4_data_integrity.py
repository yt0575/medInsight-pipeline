from __future__ import annotations

import numpy as np
import pytest

from pipeline import core


def test_build_ch4_data_is_sorted_and_non_empty(ch4_data):
    q = ch4_data.quarterly
    assert not q.empty
    assert list(q.columns)[:5] == ["quarter", "hospital", "drugstore", "online", "total"]

    qkeys = q["quarter"].map(core.qkey).tolist()
    assert qkeys == sorted(qkeys), "quarterly data must be sorted by quarter"


def test_latest_share_and_cr5_are_reasonable(ch4_data):
    share_sum = float(ch4_data.latest_share["share_pct"].sum())
    assert share_sum == pytest.approx(100.0, abs=1e-3)

    cr5 = ch4_data.cr5_latest["cr5_pct"].astype(float).to_numpy()
    assert np.nanmin(cr5) >= 0.0
    assert np.nanmax(cr5) <= 100.0


def test_yoy_rows_cover_three_channels(ch4_data):
    yoy = ch4_data.yoy_latest
    assert yoy.shape[0] == 3
    assert set(yoy["channel"].tolist()) == {"医院端", "药店端", "线上端"}
