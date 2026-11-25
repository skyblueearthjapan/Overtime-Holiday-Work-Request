"""休憩の時刻窓方式（worklog.gs:111-129。実装仕様 8-2・8-3）。"""
from __future__ import annotations

import pytest

from app.domain.breaks import calc_break_by_clock_time
from app.domain.worklog import compute_minutes
from test_domain_helpers import jst


def brk(h1, m1, h2, m2, s2=0, *, day2=26):
    return calc_break_by_clock_time(jst(2026, 11, 26, h1, m1), jst(2026, 11, day2, h2, m2, s2))


@pytest.mark.parametrize("end,expected", [
    ((19, 29, 59), 0),   # 19:29 として比べる → 19:20-19:30 を含まない
    ((19, 30, 0), 10),
    ((19, 30, 40), 10),  # 秒は捨てる
    ((19, 20, 0), 0),
    ((21, 40, 0), 20),
    ((21, 39, 0), 10),
])
def test_overtime_window_boundaries(end, expected):
    assert brk(17, 20, *end) == expected


def test_1710_window_not_counted_from_1720_start():
    assert brk(17, 20, 23, 0) == 20
    assert brk(17, 10, 17, 20) == 10


def test_full_day_windows():
    assert brk(8, 0, 17, 0) == 10 + 45 + 10
    assert brk(10, 0, 10, 10) == 10
    assert brk(10, 1, 13, 0) == 45
    assert brk(0, 0, 23, 59) == 10 + 45 + 10 + 10 + 10 + 10


def test_crossing_midnight_gives_zero():
    assert brk(17, 20, 1, 0, day2=27) == 0


def test_compute_minutes_rounding_and_floor_at_zero():
    s = jst(2026, 11, 26, 17, 20)
    assert compute_minutes(s, jst(2026, 11, 26, 19, 29, 59)) == (130, 0, 130)  # 129.98 分 → 130、休憩 0
    assert compute_minutes(s, jst(2026, 11, 26, 19, 30, 0)) == (130, 10, 120)
    assert compute_minutes(s, jst(2026, 11, 26, 19, 30, 40)) == (131, 10, 121)  # 130.67 → 131
    assert compute_minutes(s, jst(2026, 11, 26, 17, 20, 30)) == (1, 0, 1)  # 0.5 分 → 1（Math.round）
    assert compute_minutes(s, jst(2026, 11, 26, 17, 20, 29, 999)) == (0, 0, 0)
    assert compute_minutes(s, jst(2026, 11, 26, 17, 0)) == (0, 0, 0)  # 前でも 0
