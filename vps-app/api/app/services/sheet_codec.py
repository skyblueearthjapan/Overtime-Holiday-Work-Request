"""DB の A 区分の表 ⇔ シートのセル（実装仕様 6-3・6-4。担当 D）。

行の反映（row_sync）・毎晩の照合（reconcile）・移行の道具（tools/migrate_from_sheets.py）が同じ決まりを使う。

- 列はシートの見出しの文字で対応させる（COLUMN_RULES。見出しは GAS の normalize_ と同じに直して比べる）。
- 書くときは Sheets API の updateCells のセル（CellData）を作る（encode_cell）。値は numberValue（シリアル値・分）か
  stringValue（入力として解釈されない＝RAW と同じ。'を付けない）。空は userEnteredValue を持たないセル（値を消す）。
- 読むときは UNFORMATTED_VALUE（日時はシリアル値）と FORMATTED_VALUE を一緒に読み、(値, 型) にする（decode_cell）。
- 型（kind）: 'serial'＝日付・日時のセル（数値）、'string'＝文字のセル。型が混ざる列は行ごとの型を
  行の cell_kinds[見出し] に持つ（無い・'auto' は「その列の最後のデータ行と同じ型」）。
- シリアル値は DB_SHEET_TIMEZONE の壁時計で 1899-12-30 からの日数（時刻は小数、ミリ秒まで）。

列の種類（ColumnRule.kind）:
  text           文字（自由入力を含む。'1/2' も文字のまま）
  int            分・回数（整数。NULL は空）
  datetime       日時のセル（submittedAt・approvedAt・approvedAt2・pdfGeneratedAt・BatchLogs 実行日時）。DB は timestamptz
  mixed_datetime 型が混ざる日時（Requests modifiedAt）。DB は timestamptz。文字のときは JST の 'yyyy-MM-dd HH:mm:ss'
  date           型が混ざる日付（Requests targetDate・BatchLogs 対象日）。DB は date ＋ 日付にならない元の文字
  text_datetime  型が混ざる日時を文字で持つ（WorkLogs updatedAt）。DB は JST の 'yyyy-MM-dd HH:mm:ss'
  jst_text       JST の文字（WorkLogs actualStartAt・actualEndAt）。書くときはいつも文字＋表示形式 '@'
                 （GAS の updateWorkLog_ と同じ）。日時のセルだった古い行は、GAS の toJstString_ と同じく
                 シートの時間帯の壁時計の文字で読む
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from typing import Any
from zoneinfo import ZoneInfo

from app.jsutil import js_string, js_trim, normalize_gas

UTC = timezone.utc
JST = ZoneInfo("Asia/Tokyo")
EPOCH = datetime(1899, 12, 30)
MS_PER_DAY = 86_400_000

SERIAL = "serial"
STRING = "string"
AUTO = "auto"
KINDS = (SERIAL, STRING, AUTO)

# 新しい行で、同じ列に参考にできる日付のセルが無いときの表示形式（実装仕様 6-4）
DEFAULT_DATE_FORMAT = {"type": "DATE", "pattern": "yyyy/mm/dd"}
DEFAULT_DATETIME_FORMAT = {"type": "DATE_TIME", "pattern": "yyyy/mm/dd h:mm:ss"}
TEXT_FORMAT = {"type": "TEXT", "pattern": "@"}
DATE_FORMAT_TYPES = ("DATE", "DATE_TIME", "TIME")

# 型が混ざる・日時の列（cell_kinds を持つ列）
DATE_LIKE = ("datetime", "mixed_datetime", "date", "text_datetime", "jst_text")


class CodecError(ValueError):
    """セルの値が決まりの型として読めない（移行は止める）。"""


@dataclass(frozen=True)
class ColumnRule:
    header: str
    attr: str
    kind: str

    @property
    def date_like(self) -> bool:
        return self.kind in DATE_LIKE

    @property
    def is_date_only(self) -> bool:
        return self.kind == "date"

    @property
    def default_kind(self) -> str:
        """cell_kinds に値が無いときの型。datetime はいつも日時のセル、混ざる列は auto。"""
        if self.kind == "datetime":
            return SERIAL
        if self.kind == "jst_text":
            return STRING
        if self.kind in ("mixed_datetime", "date", "text_datetime"):
            return AUTO
        return STRING if self.kind == "text" else SERIAL


@dataclass(frozen=True)
class TableSpec:
    table: str
    tab: str
    header_row: int  # 見出しの行（1 始まり）。データはその次の行から
    columns: tuple[ColumnRule, ...]
    id_headers: tuple[str, ...]  # 行の確かめに使う列

    @property
    def data_start_row(self) -> int:
        return self.header_row + 1

    def rule(self, header: str) -> ColumnRule:
        for c in self.columns:
            if c.header == header:
                return c
        raise KeyError(header)

    @property
    def headers(self) -> tuple[str, ...]:
        return tuple(c.header for c in self.columns)


def _rules(*items: tuple[str, str, str]) -> tuple[ColumnRule, ...]:
    return tuple(ColumnRule(h, a, k) for h, a, k in items)


REQUESTS = TableSpec(
    table="requests", tab="Requests", header_row=1, id_headers=("requestId",),
    columns=_rules(
        ("requestId", "request_id", "text"),
        ("requestType(overtime/holiday)", "request_type", "text"),
        ("status(submitted/approved/canceled)", "status", "text"),
        ("dept", "dept", "text"),
        ("workerCode", "worker_code", "text"),
        ("workerName", "worker_name", "text"),
        ("workerEmail", "worker_email", "text"),
        ("targetDate", "target_date", "date"),
        ("submittedAt", "submitted_at", "datetime"),
        ("approvedAt", "approved_at", "datetime"),
        ("approvedBy", "approved_by", "text"),
        ("approvedMinutes", "approved_minutes", "int"),
        ("reason", "reason", "text"),
        ("workContent", "work_content", "text"),
        ("jobId1", "job_id1", "text"),
        ("jobId2", "job_id2", "text"),
        ("jobId3", "job_id3", "text"),
        ("workNo1", "work_no1", "text"),
        ("workNo2", "work_no2", "text"),
        ("workNo3", "work_no3", "text"),
        ("orderNo1", "order_no1", "text"),
        ("orderNo2", "order_no2", "text"),
        ("orderNo3", "order_no3", "text"),
        ("customer1", "customer1", "text"),
        ("customer2", "customer2", "text"),
        ("customer3", "customer3", "text"),
        ("product1", "product1", "text"),
        ("product2", "product2", "text"),
        ("product3", "product3", "text"),
        ("hrMailSentAt", "hr_mail_sent_at", "text"),
        ("pdfGeneratedAt", "pdf_generated_at", "datetime"),
        ("pdfFileId", "pdf_file_id", "text"),
        ("pdfFolderId", "pdf_folder_id", "text"),
        ("exportError", "export_error", "text"),
        ("reasonDetail", "reason_detail", "text"),
        ("approvedBy2", "approved_by2", "text"),
        ("approvedAt2", "approved_at2", "datetime"),
        ("modifiedAt", "modified_at", "mixed_datetime"),
        ("modifiedBy", "modified_by", "text"),
    ),
)

WORK_LOGS = TableSpec(
    table="work_logs", tab="WorkLogs", header_row=2, id_headers=("requestId",),
    columns=_rules(
        ("requestId", "request_id", "text"),
        ("actualStartAt", "actual_start_at", "jst_text"),
        ("actualEndAt", "actual_end_at", "jst_text"),
        ("actualMinutes", "actual_minutes", "int"),
        ("breakMinutes", "break_minutes", "int"),
        ("netMinutes", "net_minutes", "int"),
        ("updatedAt", "updated_at", "text_datetime"),
        ("updatedBy", "updated_by", "text"),
    ),
)

BATCH_LOGS = TableSpec(
    table="batch_logs", tab="BatchLogs", header_row=1, id_headers=("実行日時", "バッチ名"),
    columns=_rules(
        ("実行日時", "executed_at", "datetime"),
        ("バッチ名", "batch_name", "text"),
        ("対象日", "target_date", "date"),
        ("成功", "success", "int"),
        ("スキップ", "skip", "int"),
        ("失敗", "fail", "int"),
        ("エラー詳細", "error_detail", "text"),
    ),
)

TABLES: dict[str, TableSpec] = {s.table: s for s in (REQUESTS, WORK_LOGS, BATCH_LOGS)}
TABLE_ORDER = ("requests", "work_logs", "batch_logs")
# 表 → 列の決まり（実装仕様 10-2 の名前）
COLUMN_RULES: dict[str, tuple[ColumnRule, ...]] = {t: s.columns for t, s in TABLES.items()}

KINDS_KEY = "__kinds__"  # synced_values の中で「シートにある型」を持つキー


# ===== シリアル値 =====


def wall_to_serial(wall: datetime) -> float:
    """時差のない壁時計 → シリアル値（ミリ秒まで）。"""
    delta = wall.replace(tzinfo=None) - EPOCH
    ms = delta.days * MS_PER_DAY + delta.seconds * 1000 + delta.microseconds // 1000
    return ms / MS_PER_DAY


def serial_to_wall(serial: float) -> datetime:
    """シリアル値 → 時差のない壁時計（ミリ秒に丸める）。"""
    ms = int(math.floor(float(serial) * MS_PER_DAY + 0.5))
    return EPOCH + timedelta(milliseconds=ms)


def datetime_to_serial(moment: datetime, tz: ZoneInfo) -> float:
    """時差付きの日時 → tz の壁時計のシリアル値。"""
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=UTC)
    return wall_to_serial(moment.astimezone(tz).replace(tzinfo=None))


def serial_to_datetime(serial: float, tz: ZoneInfo) -> datetime:
    """シリアル値（tz の壁時計）→ UTC の日時。"""
    return serial_to_wall(serial).replace(tzinfo=tz).astimezone(UTC)


def date_to_serial(d: date) -> float:
    return float((d - EPOCH.date()).days)


def serial_to_date(serial: float) -> date:
    return serial_to_wall(serial).date()


def _as_tz(tz: Any) -> ZoneInfo:
    return tz if isinstance(tz, ZoneInfo) else ZoneInfo(str(tz))


# ===== 文字の日時 =====

_DT_FORMATS = (
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M",
    "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%Y-%m-%dT%H:%M:%S",
)
_D_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y")
_YMD_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def parse_jst_datetime(text: str) -> datetime | None:
    """GAS が書いた JST の日時の文字（'yyyy-MM-dd HH:mm:ss' など）→ UTC の日時。読めなければ None。"""
    s = js_trim(text or "")
    if not s:
        return None
    m = re.match(r"^(.*?)(Z|[+-]\d{2}:?\d{2})$", s)
    if m and "T" in s:
        try:
            return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(UTC)
        except ValueError:
            return None
    for fmt in _DT_FORMATS + _D_FORMATS:
        try:
            return datetime.strptime(s, fmt).replace(tzinfo=JST).astimezone(UTC)
        except ValueError:
            continue
    return None


def parse_date_text(text: str) -> date | None:
    """'yyyy-MM-dd'・'yyyy/M/d'・'M/d/yyyy' → date。読めなければ None。"""
    s = js_trim(text or "")
    for fmt in _D_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def jst_text(moment: datetime | None) -> str:
    """UTC の日時 → JST の 'yyyy-MM-dd HH:mm:ss'（GAS の fmtDate_ と同じ）。"""
    if moment is None:
        return ""
    return moment.astimezone(JST).strftime("%Y-%m-%d %H:%M:%S")


def wall_text(wall: datetime) -> str:
    return wall.strftime("%Y-%m-%d %H:%M:%S")


_DATE_DISPLAY_RE = re.compile(r"^\s*(\d{1,4}[/\-.年]\d{1,2}[/\-.月]\d{1,4}|\d{1,2}:\d{2})")


def looks_like_date_display(formatted: Any) -> bool:
    """FORMATTED_VALUE の文字が日付・時刻の表示か（表示形式の無いシリアル値は '46352.29' のように出る）。"""
    return isinstance(formatted, str) and bool(_DATE_DISPLAY_RE.match(formatted))


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def is_blank(value: Any) -> bool:
    return value is None or value == ""


# ===== DB の行 ⇔ 論理の値 =====


def db_value(rule: ColumnRule, row: Any) -> Any:
    """ORM の行（または dict）から、その列の論理の値（date の列は date か文字）。"""
    get = row.get if isinstance(row, dict) else (lambda k, d=None: getattr(row, k, d))
    if rule.kind == "date":
        d = get("target_date")
        if d is not None:
            return d
        return get("target_date_text") or ""
    value = get(rule.attr)
    if rule.kind in ("datetime", "mixed_datetime", "int"):
        return value
    return "" if value is None else value


def assign_db_value(rule: ColumnRule, record: dict, value: Any) -> None:
    """論理の値を DB の列（dict）に入れる（移行）。"""
    if rule.kind == "date":
        if isinstance(value, date):
            record["target_date"] = value
            record["target_date_text"] = ""
        else:
            record["target_date"] = None
            record["target_date_text"] = value or ""
        return
    record[rule.attr] = value


def canonical(rule: ColumnRule, value: Any) -> Any:
    """synced_values に入れる・比べるための形（JSON にできる値）。

    日時は UTC の ISO（ミリ秒まで）、日付は 'yyyy-MM-dd'、整数は int（NULL は ''）、文字はそのまま。
    """
    if value is None or value == "":
        return ""
    if rule.kind in ("datetime", "mixed_datetime"):
        if isinstance(value, datetime):
            v = value if value.tzinfo else value.replace(tzinfo=UTC)
            return v.astimezone(UTC).isoformat(timespec="milliseconds")
        return js_string(value)
    if rule.kind == "date":
        return value.isoformat() if isinstance(value, date) else js_string(value)
    if rule.kind == "int":
        return int(value) if is_number(value) else js_string(value)
    return js_string(value)


def canonical_stored(rule: ColumnRule, stored: Any) -> Any:
    """synced_values に入っている値を canonical の形にそろえる。

    担当 M の gas_jobs.record_pdf_result は日時を datetime.isoformat()（マイクロ秒あり・なし）で入れるので、
    日時の列は ISO を読み直してミリ秒までの形にする（同じ時刻なら同じ値＝PDF の列を書き直さない）。
    """
    if rule.kind in ("datetime", "mixed_datetime") and isinstance(stored, str) and "T" in stored:
        try:
            parsed = datetime.fromisoformat(stored.replace("Z", "+00:00"))
        except ValueError:
            return stored
        return canonical(rule, parsed if parsed.tzinfo else parsed.replace(tzinfo=UTC))
    return stored


def snapshot(spec: TableSpec, row: Any) -> dict:
    """行の全列の canonical（見出し → 値）。synced_values の形。"""
    return {c.header: canonical(c, db_value(c, row)) for c in spec.columns}


def row_kinds(row: Any) -> dict:
    kinds = row.get("cell_kinds") if isinstance(row, dict) else getattr(row, "cell_kinds", None)
    return dict(kinds or {})


def expected_kind(rule: ColumnRule, kinds: dict | None, synced_kinds: dict | None = None) -> str:
    """書く・比べる型。行の cell_kinds → シートにあると分かっている型 → 列の既定（auto もあり）。"""
    k = (kinds or {}).get(rule.header)
    if k in (SERIAL, STRING):
        return k
    k = (synced_kinds or {}).get(rule.header)
    if k in (SERIAL, STRING):
        return k
    return rule.default_kind


def same_value(rule: ColumnRule, a: Any, b: Any) -> bool:
    """論理の値が同じか。日時は秒の単位（1 秒未満の差は同じ）、日付は日、ほかは文字として。"""
    if is_blank(a) and is_blank(b):
        return True
    if is_blank(a) != is_blank(b):
        return False
    if rule.kind in ("datetime", "mixed_datetime"):
        if isinstance(a, datetime) and isinstance(b, datetime):
            return abs((a - b).total_seconds()) < 1.0
        return canonical(rule, a) == canonical(rule, b)
    if rule.kind == "text_datetime":
        pa, pb = parse_jst_datetime(js_string(a)), parse_jst_datetime(js_string(b))
        if pa is not None and pb is not None:
            return abs((pa - pb).total_seconds()) < 1.0
        return js_string(a) == js_string(b)
    if rule.kind == "int":
        try:
            return int(a) == int(b)
        except (TypeError, ValueError):
            return js_string(a) == js_string(b)
    return canonical(rule, a) == canonical(rule, b)


# ===== 書く（CellData） =====


def _string_cell(text: str) -> dict:
    return {"userEnteredValue": {"stringValue": text}} if text != "" else {}


def _number_cell(number: float) -> dict:
    return {"userEnteredValue": {"numberValue": number}}


def default_format(rule: ColumnRule) -> dict:
    return DEFAULT_DATE_FORMAT if rule.is_date_only else DEFAULT_DATETIME_FORMAT


def encode_cell(rule: ColumnRule, value: Any, kind: str, tz: Any, number_format: dict | None = None) -> dict:
    """論理の値 → updateCells のセル。

    戻り値は {"cell": CellData, "fields": "userEnteredValue" か "userEnteredValue,userEnteredFormat.numberFormat",
    "kind": 実際に書いた型}。kind が auto のときは呼ぶ側が先に解く（解けなければ serial と同じ扱い）。
    日付・日時のセルは number_format（無ければ既定）を当てる。jst_text と、文字の text_datetime は '@'。
    """
    zone = _as_tz(tz)
    value_only = "userEnteredValue"
    with_format = "userEnteredValue,userEnteredFormat.numberFormat"
    if kind == AUTO:
        kind = SERIAL
    if rule.kind == "text":
        return {"cell": _string_cell(js_string(value)), "fields": value_only, "kind": STRING}
    if rule.kind == "int":
        if is_blank(value):
            return {"cell": {}, "fields": value_only, "kind": SERIAL}
        return {"cell": _number_cell(int(value)), "fields": value_only, "kind": SERIAL}
    if rule.kind == "jst_text":
        cell = _string_cell(js_string(value))
        cell["userEnteredFormat"] = {"numberFormat": dict(TEXT_FORMAT)}
        return {"cell": cell, "fields": with_format, "kind": STRING}
    if is_blank(value):
        return {"cell": {}, "fields": value_only, "kind": kind}

    fmt = dict(number_format or default_format(rule))
    if rule.kind in ("datetime", "mixed_datetime"):
        moment = value if isinstance(value, datetime) else parse_jst_datetime(js_string(value))
        if moment is None:
            return {"cell": _string_cell(js_string(value)), "fields": value_only, "kind": STRING}
        if kind == STRING:
            return {"cell": _string_cell(jst_text(moment)), "fields": value_only, "kind": STRING}
        cell = _number_cell(datetime_to_serial(moment, zone))
        cell["userEnteredFormat"] = {"numberFormat": fmt}
        return {"cell": cell, "fields": with_format, "kind": SERIAL}
    if rule.kind == "date":
        if not isinstance(value, date):
            return {"cell": _string_cell(js_string(value)), "fields": value_only, "kind": STRING}
        if kind == STRING:
            return {"cell": _string_cell(value.isoformat()), "fields": value_only, "kind": STRING}
        cell = _number_cell(date_to_serial(value))
        cell["userEnteredFormat"] = {"numberFormat": fmt}
        return {"cell": cell, "fields": with_format, "kind": SERIAL}
    if rule.kind == "text_datetime":
        text = js_string(value)
        moment = parse_jst_datetime(text)
        if kind == STRING or moment is None:
            cell = _string_cell(text)
            cell["userEnteredFormat"] = {"numberFormat": dict(TEXT_FORMAT)}
            return {"cell": cell, "fields": with_format, "kind": STRING}
        cell = _number_cell(datetime_to_serial(moment, zone))
        cell["userEnteredFormat"] = {"numberFormat": fmt}
        return {"cell": cell, "fields": with_format, "kind": SERIAL}
    raise ValueError(f"知らない列の種類です: {rule.kind}")  # pragma: no cover


# ===== 読む =====


def decode_cell(rule: ColumnRule, unformatted: Any, formatted: Any, tz: Any) -> tuple[Any, str | None]:
    """UNFORMATTED と FORMATTED の値 → (論理の値, 型)。空は ('' か None, None)。

    読めない値は CodecError（移行は止める。照合は差として出す）。
    """
    zone = _as_tz(tz)
    v = unformatted
    if is_blank(v):
        return (None if rule.kind in ("datetime", "mixed_datetime", "int") else ""), None
    if rule.kind == "text":
        if is_number(v) and looks_like_date_display(formatted):
            return js_string(formatted), SERIAL
        return js_string(v), (SERIAL if is_number(v) else STRING)
    if rule.kind == "int":
        if is_number(v):
            return _to_int(v), SERIAL
        text = js_trim(js_string(v))
        if re.fullmatch(r"[+-]?\d+(\.\d+)?", text):
            return _to_int(float(text)), STRING
        raise CodecError(f"「{rule.header}」が数ではありません: {text!r}")
    if rule.kind in ("datetime", "mixed_datetime"):
        if is_number(v):
            return serial_to_datetime(v, zone), SERIAL
        moment = parse_jst_datetime(js_string(v))
        if moment is None:
            raise CodecError(f"「{rule.header}」が日時として読めません: {js_string(v)!r}")
        return moment, STRING
    if rule.kind == "date":
        if is_number(v):
            return serial_to_date(v), SERIAL
        text = js_string(v)
        parsed = parse_date_text(text)
        return (parsed if parsed is not None else text), STRING
    if rule.kind == "text_datetime":
        if is_number(v):
            return jst_text(serial_to_datetime(v, zone)), SERIAL
        return js_string(v), STRING
    if rule.kind == "jst_text":
        if is_number(v):
            # GAS の toJstString_: シートの時間帯で書式化した文字（壁時計のまま）を JST とみなす
            return wall_text(serial_to_wall(v)), SERIAL
        return js_string(v), STRING
    raise ValueError(f"知らない列の種類です: {rule.kind}")  # pragma: no cover


def _to_int(x: float) -> int:
    return int(math.floor(float(x) + 0.5))


# ===== 見出し =====


def header_index(header_row: list) -> dict[str, int]:
    """GAS の buildHeaderIndex_ と同じ（normalize_ した見出し → 列。同じ見出しは右が勝つ。空は無視）。"""
    idx: dict[str, int] = {}
    for i, h in enumerate(header_row or []):
        key = normalize_gas(h)
        if key:
            idx[key] = i
    return idx


def missing_headers(spec: TableSpec, header_row: list) -> list[str]:
    idx = header_index(header_row)
    return [h for h in spec.headers if h not in idx]


def cell_at(row: list, index: int) -> Any:
    return row[index] if 0 <= index < len(row) else ""
