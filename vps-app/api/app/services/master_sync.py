"""マスタ同期（決定 D・実装仕様 6-8。`sync_masters`、毎日 0:30 と手動。担当 D）。

src/syncMasters.gs の syncAllMasters を写す。マスタ元（MASTER_SOURCE_SPREADSHEET_ID。読むだけ）の 4 タブを読み、
DB スプレッドシートの同名のタブ（C 区分）へ今の GAS と同じ形で書き、同じ読み取りから DB の表も全置換する。

GAS と同じにすること（syncMasters.gs の行）:
- SYNC_MAP（12-21 行）: 作業員マスタ→作業員マスタ（見出し＋80 人）、業務NO.マスタ→業務NOマスタ、
  工番マスタ→工番マスタ、社内カレンダーマスタ→社内カレンダーマスタ。
- 見出し行の最後の値がある列まで読む（detectLastColByHeader_。書式だけ残った列を無視。見出しが全部空なら 1 列）。
- 1 列目が空の行を捨てる（compactRows_。`String(v || '').trim() !== ''`＝0・false も空。見出しは残す）。
- 捨てた後の「見出し＋先頭 80 行」で作業員を切る（元の並び）。
- Date は「JST の 0 時ちょうど → 'yyyy-MM-dd'、それ以外 → 'yyyy-MM-dd HH:mm:ss'」（normalizeDatesForWrite_・26f81a5）。
  UNFORMATTED で読み、マスタ元のファイルの時間帯の壁時計で日時に直してから JST で同じ規則の文字にする
  （日付のセルかは FORMATTED の表示で見分ける）。
- 書くのは 1 行目から、setValues と同じ USER_ENTERED。読み取りは失敗したら 2 秒待って 3 回まで（readSheetData_）。
- タブごとに失敗を握り、ほかのタブは続ける（83-85 行）。

変えたこと（実装仕様 11-1・6-8）:
- 余った行・列は値だけ消す（values:batchClear）。行・列・タブを消して作り直さない（GAS の deleteRows・
  deleteSheet/insertSheet の作り直しは写さない）。書く範囲が今の行数を超えれば、values.batchUpdate が行を足す。
- 元が読めない・見出しだけ・有効な行 0 なら、そのタブも DB の表も上書きしない（GAS は見出しだけに置き換える）。

DB の表（実装仕様 5-3）は、GAS が DB のタブを読む読み方に合わせる（USER_ENTERED で数字の文字は数になる＝
'0042' は '42'。日付の文字は日付として読まれ、GAS は 'yyyy-MM-dd' の文字で使う）。
戻り値 {ok, tabs: {名前: {ok, rows, skipped_reason}}}。sync_runs（kind='sync_masters'）に記録する。commit は呼ぶ側。
"""
from __future__ import annotations

import logging
import re
import time
from collections.abc import Callable
from typing import Any

from sqlalchemy import delete
from sqlalchemy.orm import Session

from app import timeutil
from app.jsutil import js_string, js_trim
from app.models import CompanyCalendar, Order, SyncRun, Worker
from app.services import sheet_codec as codec
from app.services.sheets_api import SheetsApi, column_letter, quote_sheet, sheet_titles, spreadsheet_timezone

log = logging.getLogger(__name__)

RUN_KIND = "sync_masters"
READ_ATTEMPTS = 3
READ_WAIT_SECONDS = 2.0

# syncMasters.gs の SYNC_MAP（src → dst、maxRows は見出しを含む行数）
SYNC_MAP: tuple[dict, ...] = (
    {"src": "作業員マスタ", "dst": "作業員マスタ", "max_rows": 81},
    {"src": "業務NO.マスタ", "dst": "業務NOマスタ", "max_rows": None},
    {"src": "工番マスタ", "dst": "工番マスタ", "max_rows": None},
    {"src": "社内カレンダーマスタ", "dst": "社内カレンダーマスタ", "max_rows": None},
)

sleep: Callable[[float], None] = time.sleep  # 試験で差し替える


# ===== GAS の読み方・書き方の写し =====


def _js_truthy_blank(v: Any) -> bool:
    """`String(v || '').trim() === ''`（0・false・null・'' は空）。"""
    if v is None or v is False or v == "" or (codec.is_number(v) and v == 0):
        return True
    return js_trim(js_string(v)) == ""


def detect_last_col(header: list) -> int:
    """detectLastColByHeader_: 見出しの最後の値がある列（1 始まり）。全部空なら 1。"""
    for c in range(len(header), 0, -1):
        if js_trim(js_string(header[c - 1])) != "":
            return c
    return 1


def date_cell_text(serial: float, source_tz) -> str:
    """normalizeDatesForWrite_: マスタ元の壁時計 → JST で 0 時なら 'yyyy-MM-dd'、ほかは 'yyyy-MM-dd HH:mm:ss'。"""
    moment = codec.serial_to_datetime(serial, source_tz).astimezone(codec.JST)
    if (moment.hour, moment.minute, moment.second) == (0, 0, 0):
        return moment.strftime("%Y-%m-%d")
    return moment.strftime("%Y-%m-%d %H:%M:%S")


def prepare(unformatted: list, formatted: list, source_tz, max_rows: int | None) -> list[list] | None:
    """元のタブ → GAS が setValues に渡す 2 次元配列（見出しつき）。元が空なら None。"""
    if not unformatted:
        return None
    last_col = detect_last_col(unformatted[0])
    data = []
    for r, row in enumerate(unformatted):
        cells = (list(row) + [""] * last_col)[:last_col]
        frow = formatted[r] if r < len(formatted) else []
        if r > 0 and _js_truthy_blank(cells[0]):
            continue
        out = []
        for c, v in enumerate(cells):
            f = frow[c] if c < len(frow) else ""
            out.append(date_cell_text(v, source_tz) if codec.is_number(v) and codec.looks_like_date_display(f) else v)
        data.append(out)
    if max_rows and len(data) > max_rows:
        data = data[:max_rows]
    return data


_NUMBER_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)\Z")


def as_gas_reads(value: Any) -> str:
    """USER_ENTERED で書いたセルを GAS が読んで文字にした形（数字の文字は数になる。日付の文字はそのまま）。"""
    if isinstance(value, str):
        s = value.strip()
        if _NUMBER_RE.match(s):
            n = float(s)
            return js_string(int(n) if n == int(n) else n)
        return value
    if isinstance(value, bool):
        return "true" if value else "false"
    return js_string(value)


def _index(header: list, label: str) -> int:
    for i, h in enumerate(header):
        if js_trim(js_string(h)) == label:
            return i
    return -1


def worker_rows(data: list[list]) -> list[dict]:
    header = [js_trim(js_string(h)) for h in data[0]]
    idx = {k: _index(header, k) for k in ("作業員コード", "氏名", "部署", "在籍フラグ")}
    g = next((i for i, h in enumerate(header) if h.startswith("Googleアカウント")), -1)
    out = []
    for n, row in enumerate(data[1:], start=1):
        values = [as_gas_reads(v) for v in row]

        def at(i: int) -> str:
            return values[i] if i >= 0 else ""

        out.append({"sort_order": n, "worker_code": at(idx["作業員コード"]), "worker_name": at(idx["氏名"]),
                    "dept": at(idx["部署"]), "active_flag": at(idx["在籍フラグ"]), "google_account": at(g),
                    "raw": {h: values[i] for i, h in enumerate(header) if h}})
    return out


def order_rows(data: list[list]) -> list[dict]:
    header = [js_trim(js_string(h)) for h in data[0]]
    cols = {"code": "工番", "customer": "受注先", "dest": "納入先", "dest_address": "納入先住所", "product": "品名",
            "quantity": "数量", "imported_at": "取込日時"}
    idx = {k: _index(header, v) for k, v in cols.items()}
    out = []
    for n, row in enumerate(data[1:], start=1):
        values = [as_gas_reads(v) for v in row]
        out.append({"sort_order": n, **{k: (values[i] if i >= 0 else "") for k, i in idx.items()}})
    return out


def calendar_rows(data: list[list]) -> list[dict]:
    out = []
    for n, row in enumerate(data[1:], start=1):
        values = [as_gas_reads(v) for v in row] + [""] * 4
        out.append({"sort_order": n, "date_text": values[0], "kind": values[1], "weekday": values[2],
                    "note": values[3]})
    return out


TABLE_FOR_DST: dict[str, tuple[Any, Callable[[list[list]], list[dict]]]] = {
    "作業員マスタ": (Worker, worker_rows),
    "工番マスタ": (Order, order_rows),
    "社内カレンダーマスタ": (CompanyCalendar, calendar_rows),
}


# ===== 読む・書く =====


def read_source(api: SheetsApi, source_id: str, tab: str) -> tuple[list, list]:
    """readSheetData_ と同じく、失敗したら 2 秒待って 3 回まで。"""
    last: Exception | None = None
    for attempt in range(READ_ATTEMPTS):
        try:
            unformatted, formatted = api.read_both(source_id, [quote_sheet(tab)])
            return unformatted[0], formatted[0]
        except Exception as error:  # noqa: BLE001 - GAS と同じく読み直す
            last = error
            if attempt + 1 < READ_ATTEMPTS:
                sleep(READ_WAIT_SECONDS)
    raise RuntimeError(f"Read failed after all retries: {last}")


def write_tab(api: SheetsApi, sid: str, dst: str, data: list[list], meta: dict) -> None:
    """1 行目から USER_ENTERED で書き、余った行・列の値だけを消す（行・列・タブは消さない）。"""
    rows, cols = len(data), len(data[0])
    q = quote_sheet(dst)
    grid = {}
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == dst:
            grid = s["properties"].get("gridProperties") or {}
    row_count = int(grid.get("rowCount") or rows)
    col_count = int(grid.get("columnCount") or cols)
    values = [["" if v is None else v for v in row] for row in data]
    api.values_batch_update(sid, [{"range": f"{q}!A1:{column_letter(cols - 1)}{rows}", "values": values}],
                            tab=dst, purpose="master_sync", input_option="USER_ENTERED")
    clear = []
    last_col = column_letter(max(col_count, cols) - 1)
    if row_count > rows:
        clear.append(f"{q}!A{rows + 1}:{last_col}")
    if col_count > cols:
        clear.append(f"{q}!{column_letter(cols)}1:{last_col}{rows}")
    if clear:
        api.values_batch_clear(sid, clear, tab=dst, purpose="master_sync")


def sync_masters(db: Session, api: SheetsApi | None = None, app_settings: Any = None) -> dict:
    """マスタ元 → DB のタブと表。戻り値 {ok, tabs:{名前:{ok, rows, skipped_reason}}}。commit は呼ぶ側。"""
    from app.config import get_settings

    settings = app_settings or get_settings()
    started = timeutil.now_utc()
    result: dict = {"ok": False, "tabs": {}}
    try:
        if api is None:
            api = SheetsApi.for_db(db, settings)
        source_id = settings.master_source_spreadsheet_id
        target_id = settings.db_spreadsheet_id
        source_meta = api.get_spreadsheet(source_id)
        source_tz = codec.ZoneInfo(spreadsheet_timezone(source_meta) or "Asia/Tokyo")
        source_titles = set(sheet_titles(source_meta))
        target_meta = api.get_spreadsheet(target_id)
        target_titles = set(sheet_titles(target_meta))
    except Exception as error:  # noqa: BLE001 - GAS の FATAL と同じく結果で返す
        result["error"] = f"{type(error).__name__}: {error}"[:500]
        _record(db, started, result)
        return result

    for m in SYNC_MAP:
        name = m["src"]
        tab = {"ok": False, "rows": 0, "skipped_reason": ""}
        result["tabs"][name] = tab
        try:
            if m["src"] not in source_titles:
                tab["skipped_reason"] = "元シートなし"
                continue
            if m["dst"] not in target_titles:
                tab["skipped_reason"] = "転記先シートなし"
                continue
            unformatted, formatted = read_source(api, source_id, m["src"])
            data = prepare(unformatted, formatted, source_tz, m["max_rows"])
            if not data:
                tab["skipped_reason"] = "空"
                continue
            if len(data) <= 1:
                tab["skipped_reason"] = "見出しだけ（有効な行 0。今のデータを残します）"
                continue
            write_tab(api, target_id, m["dst"], data, target_meta)
            target = TABLE_FOR_DST.get(m["dst"])
            if target is not None:
                model, build = target
                with db.begin_nested():
                    db.execute(delete(model))
                    db.add_all(model(**r) for r in build(data))
            tab["ok"] = True
            tab["rows"] = len(data) - 1
            tab["cols"] = len(data[0])
        except Exception as error:  # noqa: BLE001 - 1 タブの失敗でほかを止めない（GAS と同じ）
            tab["skipped_reason"] = f"ERROR {error}"[:500]
            log.warning("sync_masters %s: %s", name, error)
    result["ok"] = all(t["ok"] for t in result["tabs"].values())
    _record(db, started, result)
    log.info("sync_masters: %s", {k: (v["ok"], v["rows"], v["skipped_reason"]) for k, v in result["tabs"].items()})
    return result


def _record(db: Session, started, result: dict) -> None:
    db.add(SyncRun(kind=RUN_KIND, started_at=started, finished_at=timeutil.now_utc(), ok=bool(result.get("ok")),
                   summary_json=dict(result)))
    db.flush()


def _scheduled(session_factory) -> dict:
    with session_factory() as db:
        result = sync_masters(db)
        db.commit()
    return result


def register_worker() -> None:
    from app import worker

    worker.register_schedule("sync_masters", task=_scheduled, at=(0, 30))
