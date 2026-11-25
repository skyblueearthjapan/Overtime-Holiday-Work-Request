"""時刻（決定 C・実装仕様 8-1）: JST の今日、年度の境目（3/15→3/16、3/31→4/1）、締め月、GAS の書き方。
app/ の中で date.today・datetime.now・datetime.utcnow・Python の round( を使っていないことも確かめる。"""
from __future__ import annotations

import re
from datetime import date, datetime, timezone
from pathlib import Path

import pytest

from app import timeutil
from app.timeutil import JST, closing_ym, fy_end_0401, fy_start_0316, fy_start_0401, to_jst_text
from conftest import freeze_now

APP_DIR = Path(__file__).resolve().parents[1] / "app"


def test_0630_jst_is_today_in_jst(monkeypatch):
    # 2026-11-26 06:30 JST は UTC では 2026-11-25 21:30
    freeze_now(monkeypatch, datetime(2026, 11, 25, 21, 30, tzinfo=timezone.utc))
    assert timeutil.today_jst() == date(2026, 11, 26)
    assert timeutil.now_jst().hour == 6 and timeutil.now_jst().minute == 30
    assert timeutil.now_utc().date() == date(2026, 11, 25)


def test_just_before_and_after_jst_midnight(monkeypatch):
    freeze_now(monkeypatch, datetime(2026, 11, 25, 14, 59, 59, tzinfo=timezone.utc))
    assert timeutil.today_jst() == date(2026, 11, 25)
    freeze_now(monkeypatch, datetime(2026, 11, 25, 15, 0, 0, tzinfo=timezone.utc))
    assert timeutil.today_jst() == date(2026, 11, 26)


@pytest.mark.parametrize("d,expected", [
    (date(2026, 3, 15), date(2025, 3, 16)),
    (date(2026, 3, 16), date(2026, 3, 16)),
    (date(2026, 3, 31), date(2026, 3, 16)),
    (date(2026, 4, 1), date(2026, 3, 16)),
    (date(2026, 1, 1), date(2025, 3, 16)),
    (date(2026, 12, 31), date(2026, 3, 16)),
    (date(2026, 2, 28), date(2025, 3, 16)),
])
def test_fiscal_year_0316(d, expected):
    assert fy_start_0316(d) == expected


@pytest.mark.parametrize("d,expected", [
    (date(2026, 3, 31), date(2025, 4, 1)),
    (date(2026, 4, 1), date(2026, 4, 1)),
    (date(2026, 3, 15), date(2025, 4, 1)),
    (date(2026, 3, 16), date(2025, 4, 1)),
    (date(2027, 1, 1), date(2026, 4, 1)),
])
def test_fiscal_year_0401(d, expected):
    assert fy_start_0401(d) == expected
    assert fy_end_0401(d) == date(expected.year + 1, 3, 31)


def test_fiscal_year_uses_jst_date_of_datetime():
    # 2026-03-15 15:30 UTC は JST では 3/16 00:30 → 3/16 始まりの新しい年度
    moment = datetime(2026, 3, 15, 15, 30, tzinfo=timezone.utc)
    assert fy_start_0316(moment) == date(2026, 3, 16)
    # 2026-03-31 15:30 UTC は JST では 4/1 00:30 → 4/1 始まりの新しい年度
    assert fy_start_0401(datetime(2026, 3, 31, 15, 30, tzinfo=timezone.utc)) == date(2026, 4, 1)
    assert fy_start_0401(datetime(2026, 3, 31, 14, 59, tzinfo=timezone.utc)) == date(2025, 4, 1)


@pytest.mark.parametrize("moment,expected", [
    (datetime(2026, 11, 15, 23, 30, tzinfo=JST), "2026-10"),
    (datetime(2026, 11, 16, 0, 30, tzinfo=JST), "2026-11"),
    (datetime(2026, 1, 10, 12, 0, tzinfo=JST), "2025-12"),
    (date(2026, 3, 16), "2026-03"),
])
def test_closing_month(moment, expected):
    assert closing_ym(moment) == expected


@pytest.mark.parametrize("pattern,expected", [
    ("yyyy-MM-dd", "2026-11-26"),
    ("yyyy-MM-dd HH:mm", "2026-11-26 06:05"),
    ("yyyy-MM-dd HH:mm:ss", "2026-11-26 06:05:09"),
    ("M/d", "11/26"),
    ("yyyy/MM/dd H:mm:ss", "2026/11/26 6:05:09"),
    ("HH:mm", "06:05"),
    ("yyyyMMdd'T'HHmm", "20261126T0605"),
])
def test_to_jst_text_gas_patterns(pattern, expected):
    moment = datetime(2026, 11, 25, 21, 5, 9, tzinfo=timezone.utc)  # JST 2026-11-26 06:05:09
    assert to_jst_text(moment, pattern) == expected


def test_to_jst_text_date_and_none():
    assert to_jst_text(date(2026, 3, 1), "M/d") == "3/1"
    assert to_jst_text(None) == ""
    # 時差のない値は UTC とみなす
    assert to_jst_text(datetime(2026, 11, 25, 15, 0), "yyyy-MM-dd HH:mm") == "2026-11-26 00:00"


# ===== 使ってはいけない書き方（実装仕様 8-1・8-3） =====

FORBIDDEN = [
    (re.compile(r"\bdate\.today\("), "date.today()"),
    (re.compile(r"\bdatetime\.now\("), "datetime.now()"),
    (re.compile(r"\bdatetime\.utcnow\("), "datetime.utcnow()"),
    (re.compile(r"(?<![\w.])round\("), "Python の round()"),
]


def test_app_does_not_use_wall_clock_or_python_round():
    offenders = []
    for path in APP_DIR.rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        for pattern, label in FORBIDDEN:
            if path.name == "timeutil.py" and path.parent == APP_DIR and label == "datetime.now()":
                continue  # now_utc() の中の 1 か所だけ
            if pattern.search(text):
                offenders.append(f"{path.relative_to(APP_DIR)}: {label}")
    assert offenders == []


def test_timeutil_has_single_wall_clock_call():
    text = (APP_DIR / "timeutil.py").read_text(encoding="utf-8")
    assert len(re.findall(r"\bdatetime\.now\(", text)) == 1
