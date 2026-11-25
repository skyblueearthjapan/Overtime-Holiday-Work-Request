"""DB の値 → 画面へ返す JSON の値（実装仕様 4-1）。GAS の直列化の癖を写す小道具。

- 日時・日付は JST の文字（'T' も時差も付けない）。空は ''（null を返すのは仕様の表に null と書いた所だけ）。
- 数は int。
- GAS で undefined だったキーは出さない（omit_undefined に UNDEFINED を入れたキーを落とす）。
"""
from __future__ import annotations

from datetime import date, datetime
from typing import Any

from app.timeutil import to_jst_text

DATETIME_PATTERN = "yyyy-MM-dd HH:mm:ss"
DATETIME_MIN_PATTERN = "yyyy-MM-dd HH:mm"
DATE_PATTERN = "yyyy-MM-dd"


class _Undefined:
    """GAS の undefined（JSON に出さないキー）の印。"""

    _instance: "_Undefined | None" = None

    def __new__(cls) -> "_Undefined":
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __repr__(self) -> str:  # pragma: no cover
        return "UNDEFINED"

    def __bool__(self) -> bool:
        return False


UNDEFINED = _Undefined()


def fmt_datetime(value: datetime | None) -> str:
    """日時 → JST の 'yyyy-MM-dd HH:mm:ss'。空は ''。"""
    return to_jst_text(value, DATETIME_PATTERN) if value is not None else ""


def fmt_datetime_min(value: datetime | None) -> str:
    """日時 → JST の 'yyyy-MM-dd HH:mm'。空は ''。"""
    return to_jst_text(value, DATETIME_MIN_PATTERN) if value is not None else ""


def fmt_date(value: date | datetime | None) -> str:
    """日付 → 'yyyy-MM-dd'（datetime は JST の日付）。空は ''。"""
    return to_jst_text(value, DATE_PATTERN) if value is not None else ""


def omit_undefined(obj: dict[str, Any]) -> dict[str, Any]:
    """値が UNDEFINED のキーを落とす（GAS の JSON.stringify と同じ）。"""
    return {k: v for k, v in obj.items() if v is not UNDEFINED}


def json_default(value: Any) -> Any:
    """json.dumps の default。日時・日付は JST の文字。それ以外は TypeError。"""
    if isinstance(value, datetime):
        return fmt_datetime(value)
    if isinstance(value, date):
        return fmt_date(value)
    raise TypeError(f"JSON にできない値です: {type(value).__name__}")
