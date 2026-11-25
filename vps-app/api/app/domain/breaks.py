"""休憩（時刻窓方式。worklog.gs:111-129 の calcBreakByClockTime_ を写す。実装仕様 8-2）。

開始・終了を日本時間の 'HH:mm' にし（秒は捨てる）、分で比べる。窓をまるごと含む
（開始 <= 窓開始 かつ 終了 >= 窓終了）ときだけ、その窓の長さを足す。日付は見ない（日をまたぐと終了の分が小さくなり 0）。
9-1 #23 で Sebastian の残業完了もこれを使う（休憩マスタは使わない）。
"""
from __future__ import annotations

from datetime import datetime

from app import timeutil
from app.constants import BREAK_WINDOWS


def _minute_of_day(moment: datetime) -> int:
    text = timeutil.to_jst_text(moment, "yyyy-MM-dd HH:mm")
    return int(text[11:13]) * 60 + int(text[14:16])


def calc_break_by_clock_time(start: datetime, end: datetime) -> int:
    work_start = _minute_of_day(start)
    work_end = _minute_of_day(end)
    total = 0
    for sh, sm, eh, em in BREAK_WINDOWS:
        b_start = sh * 60 + sm
        b_end = eh * 60 + em
        if work_start <= b_start and work_end >= b_end:
            total += b_end - b_start
    return total
