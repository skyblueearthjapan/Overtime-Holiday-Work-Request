"""GAS（JavaScript の V8）と同じ結果を出すための小道具（日付の読み方・真偽・数）。

- js_truthy(v): JS の真偽（'', 0, None, False, NaN が偽。空の配列・辞書は真）。
- js_number(v): JS の Number(v)（'' → 0、'0x10' → 16、読めなければ NaN）。
- js_new_date(v): JS の new Date(v)。読めなければ None（Invalid Date）。V8 の実測（node）に合わせた:
    'yyyy-MM-dd' は UTC の 0 時、'yyyy-MM-ddTHH:mm(:ss)' に時差が無ければ日本時間、'yyyy/M/d'・'yyyy-MM-dd HH:mm:ss'
    （空白入り）は日本時間。月は 1〜12、日は 1〜31（2/30 は 3/2 に繰り越す）、時は 0〜23（24:00 は翌日 0 時）。
    null は 1970-01-01（epoch）、引数なし（MISSING）は Invalid。
- valid_ymd_text(s): new Date(s + 'T00:00:00+09:00') が有効か（申請・編集の「作業実施日が無効です」）。
- jst_at(ymd, 'HH:mm'): new Date(ymd + 'T' + hm + ':00+09:00')（時刻の形が違えば None）。
- label_md_dow(ymd): 'M/d(曜)'（TOP の targetDateLabel）、label_ja(ymd, today): fmtDateLabelJa_。
"""
from __future__ import annotations

import math
import re
from datetime import date, datetime, timedelta
from typing import Any

from app import timeutil
from app.jsutil import js_string, js_trim

JST = timeutil.JST
UTC = timeutil.UTC
EPOCH = datetime(1970, 1, 1, tzinfo=UTC)
DAY_NAMES = ("日", "月", "火", "水", "木", "金", "土")


class _Missing:
    """JS の undefined（引数を省いた）の印。"""

    def __repr__(self) -> str:  # pragma: no cover
        return "MISSING"

    def __bool__(self) -> bool:
        return False


MISSING: Any = _Missing()

YMD_RE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")
HM_RE = re.compile(r"^[0-9]{2}:[0-9]{2}$")


def js_truthy(value: Any) -> bool:
    if value is None or value is MISSING or value is False:
        return False
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return not (value == 0 or (isinstance(value, float) and math.isnan(value)))
    if isinstance(value, str):
        return value != ""
    return True


_NUM_RE = re.compile(r"^[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$")


def js_number(value: Any) -> float:
    """Number(value)。"""
    if value is None:
        return 0.0
    if value is MISSING:
        return math.nan
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        return math.nan
    text = js_trim(value)
    if text == "":
        return 0.0
    low = text.lower()
    for prefix, base in (("0x", 16), ("0o", 8), ("0b", 2)):
        if low.startswith(prefix):
            try:
                return float(int(text[2:], base))
            except ValueError:
                return math.nan
    if text in ("Infinity", "+Infinity"):
        return math.inf
    if text == "-Infinity":
        return -math.inf
    if _NUM_RE.match(text):
        return float(text)
    return math.nan


def js_int_or_float(value: float) -> int | float:
    """JSON にしたとき JS と同じ見た目になるように、整数の値は int にする。"""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


# ===== new Date(...) =====

_ISO_DATE = re.compile(r"^([0-9]{4})(?:-([0-9]{2})(?:-([0-9]{2}))?)?$")
_ISO_DATETIME = re.compile(
    r"^([0-9]{4})-([0-9]{2})-([0-9]{2})T([0-9]{2}):([0-9]{2})(?::([0-9]{2})(?:\.([0-9]{1,3}))?)?"
    r"(Z|[+-][0-9]{2}:[0-9]{2})?$"
)
_LEGACY = re.compile(
    r"^([0-9]{4})[/-]([0-9]{1,2})[/-]([0-9]{1,2})(?:[ T]+([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}))?)?$"
)


def _build(y: int, mo: int, d: int, h: int, mi: int, s: int, ms: int, tz) -> datetime | None:
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return None
    if h == 24:
        if mi or s or ms:
            return None
    elif not (0 <= h <= 23):
        return None
    if not (0 <= mi <= 59 and 0 <= s <= 59):
        return None
    base = datetime(y, mo, 1, tzinfo=tz) + timedelta(days=d - 1)
    return base + timedelta(hours=h, minutes=mi, seconds=s, milliseconds=ms)


def js_new_date(value: Any) -> datetime | None:
    """new Date(value)。Invalid Date は None。返す datetime は時差付き。"""
    if value is MISSING:
        return None
    if value is None:
        return EPOCH
    if isinstance(value, bool):
        return EPOCH + timedelta(milliseconds=1 if value else 0)
    if isinstance(value, datetime):
        return value if value.tzinfo else value.replace(tzinfo=UTC)
    if isinstance(value, date):
        return timeutil.jst_midnight(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return EPOCH + timedelta(milliseconds=int(value))
    if not isinstance(value, str):
        return None
    text = value.strip()
    m = _ISO_DATE.match(text)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2) or 1), int(m.group(3) or 1)
        return _build(y, mo, d, 0, 0, 0, 0, UTC)
    m = _ISO_DATETIME.match(text)
    if m:
        y, mo, d, h, mi = (int(m.group(i)) for i in range(1, 6))
        s = int(m.group(6) or 0)
        ms = int((m.group(7) or "0").ljust(3, "0"))
        off = m.group(8)
        if off is None:
            tz = JST
        elif off == "Z":
            tz = UTC
        else:
            sign = 1 if off[0] == "+" else -1
            hh, mm = int(off[1:3]), int(off[4:6])
            from datetime import timezone

            tz = timezone(sign * timedelta(hours=hh, minutes=mm))
        return _build(y, mo, d, h, mi, s, ms, tz)
    m = _LEGACY.match(text)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        h, mi, s = int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0)
        return _build(y, mo, d, h, mi, s, 0, JST)
    return None


def jst_ymd(moment: datetime) -> str:
    return timeutil.to_jst_text(moment, "yyyy-MM-dd")


def valid_ymd_text(ymd: str) -> bool:
    """new Date(ymd + 'T00:00:00+09:00') が有効か（形は呼び出し側で ^\\d{4}-\\d{2}-\\d{2}$ を確かめる）。"""
    return js_new_date(f"{ymd}T00:00:00+09:00") is not None


def jst_at(ymd: str, hm: str) -> datetime | None:
    """new Date(ymd + 'T' + hm + ':00+09:00')。"""
    return js_new_date(f"{ymd}T{hm}:00+09:00")


def ymd_to_date(ymd: str) -> date:
    """'yyyy-MM-dd'（有効な暦日）→ date。2/30 のような値は V8 と同じく繰り越す。"""
    dt = js_new_date(f"{ymd}T00:00:00+09:00")
    if dt is None:
        raise ValueError(ymd)
    return dt.astimezone(JST).date()


def strict_date(ymd: str) -> date | None:
    """暦の上で正しい 'yyyy-MM-dd' だけ date にする（DB の target_date 列へ入れるため）。"""
    try:
        return date.fromisoformat(ymd) if YMD_RE.match(ymd or "") else None
    except ValueError:
        return None


def label_md_dow(ymd: str) -> str:
    """new Date(ymd + 'T00:00:00') の (getMonth()+1) + '/' + getDate() + '(' + 曜 + ')'。"""
    dt = js_new_date(f"{ymd}T00:00:00")
    if dt is None:
        return "NaN/NaN(undefined)"
    d = dt.astimezone(JST)
    return f"{d.month}/{d.day}({DAY_NAMES[(d.weekday() + 1) % 7]})"


def label_ja(ymd: Any, today_ymd: Any) -> str:
    """fmtDateLabelJa_（今年なら '11/7(土)'、年が違えば '2025/11/7(土)'。読めなければそのまま）。"""
    text = js_string(ymd) if ymd is not None else ""
    m = re.match(r"^([0-9]{4})-([0-9]{2})-([0-9]{2})$", text)
    if not m:
        return text
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        dow_date = date(y, mo, 1) + timedelta(days=d - 1) if 1 <= mo <= 12 else None
    except ValueError:
        dow_date = None
    dow = DAY_NAMES[(dow_date.weekday() + 1) % 7] if dow_date else "undefined"
    same_year = (js_string(today_ymd) if today_ymd is not None else "")[:4] == m.group(1)
    return ("" if same_year else f"{y}/") + f"{mo}/{d}({dow})"


def is_retroactive(submitted_at: datetime | None, target_ymd: str) -> bool:
    """computeIsRetroactive_: 提出の JST の日付 > 対象日。"""
    if submitted_at is None or not target_ymd:
        return False
    return jst_ymd(submitted_at) > str(target_ymd)


def days_ago(target_ymd: str, today_ymd: str) -> int:
    """computeDaysAgo_: floor((今日 0 時 - 対象日 0 時) / 1 日)。未来・当日は 0。"""
    if not target_ymd or not today_ymd:
        return 0
    t = js_new_date(f"{target_ymd}T00:00:00+09:00")
    today = js_new_date(f"{today_ymd}T00:00:00+09:00")
    if t is None or today is None:
        return 0
    diff_ms = (today - t) // timedelta(milliseconds=1)
    if diff_ms <= 0:
        return 0
    return diff_ms // (24 * 60 * 60 * 1000)


def truncate_ms(moment: datetime) -> datetime:
    """JS の Date はミリ秒まで。"""
    return moment.replace(microsecond=(moment.microsecond // 1000) * 1000)


def iso_utc_ms(moment: datetime) -> str:
    """Date.prototype.toISOString（'2026-11-24T15:00:00.000Z'）。"""
    m = moment.astimezone(UTC)
    return m.strftime("%Y-%m-%dT%H:%M:%S.") + f"{m.microsecond // 1000:03d}Z"
