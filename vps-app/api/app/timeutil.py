"""時刻（決定 C・実装仕様 8-1）。サーバーの「今」はここだけで決める。

- 「今」は now_utc() だけ。「今日」は today_jst()。date.today・datetime.now・datetime.utcnow を
  app/ のほかの場所で使わない（tests/test_timeutil.py が grep で拒む）。
- 各モジュールは `from app import timeutil` として `timeutil.now_utc()` を呼ぶ
  （試験で `monkeypatch.setattr(timeutil, "now_utc", ...)` が効くように）。
- DB に入れる日時は UTC。画面へ返す JSON は JST の文字（to_jst_text）。
- 年度は 2 通り（3/16 始まりと 4/1 始まり。実装仕様 8-1）。
"""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

JST = ZoneInfo("Asia/Tokyo")
UTC = timezone.utc


def now_utc() -> datetime:
    """今（UTC、時差付き）。試験ではこの関数を差し替える。"""
    return datetime.now(UTC)


def now_jst() -> datetime:
    """今（日本時間）。now_utc() から求める（差し替えが効くように）。"""
    return now_utc().astimezone(JST)


def today_jst() -> date:
    """今日（日本時間の日付）。06:30 JST は UTC では前日の 21:30 だが、今日は JST の日付。"""
    return now_jst().date()


def to_jst(value: datetime) -> datetime:
    """時差付きの日時を日本時間に。時差のない値は UTC とみなす。"""
    if value.tzinfo is None:
        value = value.replace(tzinfo=UTC)
    return value.astimezone(JST)


def jst_datetime(year: int, month: int, day: int, hour: int = 0, minute: int = 0, second: int = 0) -> datetime:
    """日本時間の壁時計の日時（時差付き）。"""
    return datetime(year, month, day, hour, minute, second, tzinfo=JST)


def jst_midnight(d: date) -> datetime:
    """その日の 0 時（日本時間、時差付き）。"""
    return datetime(d.year, d.month, d.day, tzinfo=JST)


# GAS の Utilities.formatDate の書き方（yyyy・MM・dd・HH・mm・ss・M・d・H・SSS）
_PATTERN_TOKEN = re.compile(r"yyyy|yy|MM|M|dd|d|HH|H|mm|ss|SSS|'[^']*'")


def to_jst_text(value: datetime | date | None, pattern: str = "yyyy-MM-dd HH:mm:ss") -> str:
    """日時・日付を GAS の書き方（'yyyy-MM-dd'・'yyyy-MM-dd HH:mm'・'yyyy-MM-dd HH:mm:ss'・'M/d' など）で文字にする。

    datetime は日本時間に直してから書く（時差のない値は UTC とみなす）。date はそのまま。None は ''。
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        moment = to_jst(value)
        h, mi, s, ms = moment.hour, moment.minute, moment.second, moment.microsecond // 1000
        y, mo, d = moment.year, moment.month, moment.day
    else:
        y, mo, d = value.year, value.month, value.day
        h = mi = s = ms = 0

    def repl(m: re.Match[str]) -> str:
        tok = m.group(0)
        if tok.startswith("'"):
            return tok[1:-1]
        return {
            "yyyy": f"{y:04d}",
            "yy": f"{y % 100:02d}",
            "MM": f"{mo:02d}",
            "M": str(mo),
            "dd": f"{d:02d}",
            "d": str(d),
            "HH": f"{h:02d}",
            "H": str(h),
            "mm": f"{mi:02d}",
            "ss": f"{s:02d}",
            "SSS": f"{ms:03d}",
        }[tok]

    return _PATTERN_TOKEN.sub(repl, pattern)


def jst_date_of(value: datetime) -> date:
    """日時の日本時間の日付。"""
    return to_jst(value).date()


# ===== 年度（実装仕様 8-1。どちらも日本時間の日付で決める） =====


def _as_date(d: date | datetime) -> date:
    return jst_date_of(d) if isinstance(d, datetime) else d


def fy_start_0316(d: date | datetime) -> date:
    """3/16 始まりの年度の初日。m > 3 か (m == 3 かつ 日 >= 16) ならその年の 3/16、でなければ前年の 3/16。

    使う所: 申請の日付の検査、スタッフ集計の年度、台帳の年度タブ（GAS が決める）。
    """
    day = _as_date(d)
    year = day.year if (day.month > 3 or (day.month == 3 and day.day >= 16)) else day.year - 1
    return date(year, 3, 16)


def fy_start_0401(d: date | datetime) -> date:
    """4/1 始まりの年度の初日（総務の年度推移・特別条項）。"""
    day = _as_date(d)
    year = day.year if day.month >= 4 else day.year - 1
    return date(year, 4, 1)


def fy_end_0401(d: date | datetime) -> date:
    """4/1 始まりの年度の最終日（3/31）。"""
    return date(fy_start_0401(d).year + 1, 3, 31)


def closing_ym(d: date | datetime) -> str:
    """締め月（'yyyy-MM'）。日が 15 以下なら前月分、16 以上なら当月分。"""
    day = _as_date(d)
    if day.day >= 16:
        return f"{day.year:04d}-{day.month:02d}"
    prev = day.replace(day=1) - timedelta(days=1)
    return f"{prev.year:04d}-{prev.month:02d}"
