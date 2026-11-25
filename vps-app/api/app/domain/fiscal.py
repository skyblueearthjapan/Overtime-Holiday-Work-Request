"""年度と締め月（実装仕様 8-1。決定 C）。すべて日本時間の日付で決める。

- 3/16 始まり（fy_start_0316）: 申請の日付の検査・スタッフ集計の年度・台帳の年度タブ（GAS が決める）。
- 4/1 始まり（fy_start_0401）: 総務の年度推移・特別条項。
- 締め月（closing_ym）: 日が 15 以下なら前月分、16 以上なら当月分。期間は M/16〜(M+1)/16 未満。
"""
from __future__ import annotations

from datetime import date, timedelta

from app.timeutil import closing_ym, fy_end_0401, fy_start_0316, fy_start_0401

__all__ = ["fy_start_0316", "fy_start_0401", "fy_end_0401", "closing_ym", "add_months", "closing_period",
           "month_range"]


def add_months(d: date, months: int) -> date:
    """new Date(y, m-1+months, day)（日は 1〜28 で使う）。"""
    index = d.year * 12 + (d.month - 1) + months
    return date(index // 12, index % 12 + 1, d.day)


def month_range(year: int, month: int) -> tuple[date, date]:
    """暦月の [1 日, 翌月 1 日)。"""
    start = date(year, month, 1)
    return start, add_months(start, 1)


def closing_period(year: int, month: int) -> tuple[date, date]:
    """締め月の [M/16, (M+1)/16)。"""
    start = date(year, month, 16)
    return start, add_months(start, 1)


def last_day(exclusive_end: date) -> date:
    return exclusive_end - timedelta(days=1)
