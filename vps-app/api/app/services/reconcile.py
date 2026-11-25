"""毎晩の照合（実装仕様 6-7。`sheet_reconcile`、2:10。担当 D）。

A 区分の 3 タブ（Requests・WorkLogs・BatchLogs）を UNFORMATTED_VALUE（SERIAL_NUMBER）と FORMATTED_VALUE で
全行読み、DB と 1 セルずつ比べる。

- 日時は秒の単位で比べる（GAS が書いた時刻はミリ秒を持つため）。日付・日時のセルは型（日付か文字か）と、
  表示形式（表示が日付になっているか。表示形式の無いシリアル値は戻した GAS が数として読む＝gap-4 F-1）も見る。
- 反映待ち（desired_version > synced_version）と blocked の行は比べない。
- 差があれば DB の値でそのセルを書き直し（purpose='reconcile'）、直した行・列・前後の値を sync_runs に残す。
- シートにあって DB に無い行（GAS や人が書いた行）は消さずに報告する。
- 歯止め（罠 49）: DB の requests か work_logs が 0 行、または DB の行がシートより少ないタブがあれば、1 行も書かずに止める。
- check_only=True（`python -m app.worker run sheet_reconcile --check`）は書かずに数えるだけ。

compare_tab は移行の道具（移した直後の照合）も使う。
"""
from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app import timeutil
from app.models import SYNCED_TABLES, SheetSyncState, SyncRun
from app.safety import UnsafeWriteError
from app.services import row_sync
from app.services import sheet_codec as codec
from app.services import write_gate
from app.services.sheets_api import SheetsApi, SheetsError, quote_sheet, sheet_id_by_title

log = logging.getLogger(__name__)

RUN_KIND = "sheet_reconcile"
MAX_LISTED = 200


# ===== 比べる（読むだけ。移行の道具も使う） =====


def _row(grid: list, r: int) -> list:
    return grid[r - 1] if 0 < r <= len(grid) else []


def sheet_id_key(spec: codec.TableSpec, col: dict, urow: list, tz) -> tuple:
    return tuple(row_sync._id_cell(spec, h, codec.cell_at(urow, col[h]), tz) for h in spec.id_headers)


def _display(u: Any, f: Any) -> Any:
    if codec.is_blank(u):
        return ""
    return f if isinstance(f, str) and f != "" else u


def compare_tab(spec: codec.TableSpec, unformatted: list, formatted: list, db_rows: list[dict], tz) -> dict:
    """1 つのタブの全行と DB の行を比べる。

    db_rows: [{"pk", "sheet_row", "values"(見出し → 論理の値), "kinds"(cell_kinds), "synced_kinds",
               "skip"(反映待ち・blocked なら True)}]
    戻り値: {"diffs": [...], "idMismatch": [...], "extraRows": [...], "sheetRows", "dbRows", "compared"}。
    """
    header = _row(unformatted, spec.header_row)
    missing = codec.missing_headers(spec, header)
    if missing:
        raise row_sync.LaneHalt(f"{spec.tab}: 決まりの見出し {missing} がありません")
    col = {h: i for h, i in codec.header_index(header).items() if h in spec.headers}
    diffs: list[dict] = []
    id_mismatch: list[dict] = []
    claimed = {d["sheet_row"] for d in db_rows if d.get("sheet_row")}
    compared = 0
    for d in db_rows:
        r = d.get("sheet_row")
        if not r or d.get("skip"):
            continue
        urow, frow = _row(unformatted, r), _row(formatted, r)
        key = sheet_id_key(spec, col, urow, tz)
        if not row_sync._same_id(spec, key, row_sync.db_id(spec, d["values"])):
            id_mismatch.append({"row": r, "pk": d["pk"], "sheet": [codec.js_string(k) for k in key]})
            continue
        compared += 1
        synced_kinds = d.get("synced_kinds") or {}
        for rule in spec.columns:
            c = col[rule.header]
            u, f = codec.cell_at(urow, c), codec.cell_at(frow, c)
            expected = d["values"].get(rule.header)
            entry = {"row": r, "pk": d["pk"], "header": rule.header, "sheet": _display(u, f),
                     "db": codec.canonical(rule, expected)}
            try:
                value, kind = codec.decode_cell(rule, u, f, tz)
            except codec.CodecError:
                diffs.append({**entry, "reason": "unreadable"})
                continue
            if not codec.same_value(rule, value, expected):
                diffs.append({**entry, "reason": "value"})
                continue
            if codec.is_blank(expected) or not rule.date_like or rule.kind == "jst_text":
                continue
            want = codec.expected_kind(rule, d.get("kinds"), synced_kinds)
            if want in (codec.SERIAL, codec.STRING) and kind != want:
                diffs.append({**entry, "reason": "kind"})
            elif kind == codec.SERIAL and not codec.looks_like_date_display(f):
                diffs.append({**entry, "reason": "format"})
    extra: list[dict] = []
    sheet_rows = 0
    for r in range(spec.data_start_row, len(unformatted) + 1):
        urow = _row(unformatted, r)
        key = sheet_id_key(spec, col, urow, tz)
        has_id = not row_sync._blank_id(key)
        if has_id:
            sheet_rows += 1
        has_any = any(not codec.is_blank(codec.cell_at(urow, col[h])) for h in spec.headers)
        if has_any and r not in claimed:
            extra.append({"row": r, "id": [codec.js_string(k) for k in key]})
    return {"diffs": diffs, "idMismatch": id_mismatch, "extraRows": extra, "sheetRows": sheet_rows,
            "dbRows": len(db_rows), "compared": compared, "col": col}


def db_rows_for(db: Session, table: str) -> list[dict]:
    """DB の行を compare_tab の形に（反映待ち・blocked は skip）。"""
    spec = codec.TABLES[table]
    model = SYNCED_TABLES[table]
    states = {s.pk: s for s in db.scalars(select(SheetSyncState).where(SheetSyncState.table_name == table))}
    out = []
    for row in db.scalars(select(model)):
        pk = str(row.request_id if table == "requests" else row.id)
        st = states.get(pk)
        dirty = bool(st and st.desired_version > st.synced_version)
        blocked = bool(st and st.state == "blocked")
        synced = (st.synced_values if st else {}) or {}
        out.append({"pk": pk, "sheet_row": row.sheet_row, "version": int(row.row_version),
                    "values": {c.header: codec.db_value(c, row) for c in spec.columns},
                    "kinds": dict(row.cell_kinds or {}), "synced_kinds": dict(synced.get(codec.KINDS_KEY) or {}),
                    "synced": dict(synced), "skip": dirty or blocked})
    return out


# ===== 実行 =====


def _record(session_factory, ok: bool, summary: dict) -> None:
    now = timeutil.now_utc()
    with session_factory() as db:
        db.add(SyncRun(kind=RUN_KIND, started_at=summary.get("_started", now), finished_at=now, ok=ok,
                       summary_json={k: v for k, v in summary.items() if not k.startswith("_")}))
        db.commit()


def _fix_tab(api: SheetsApi, sid: str, spec: codec.TableSpec, meta: dict, unformatted: list, result: dict,
             rows: list[dict], session_factory, tz) -> int:
    """差のあるセルを DB の値で書き直す。書き直した行の数。"""
    by_pk = {d["pk"]: d for d in rows}
    grid = {}
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == spec.tab:
            grid = s["properties"].get("gridProperties") or {}
    ctx = row_sync.TabContext(spec=spec, sheet_id=sheet_id_by_title(meta, spec.tab), col=result["col"], id_rows={},
                              last_data_row=len(unformatted), row_count=int(grid.get("rowCount") or 0),
                              col_count=int(grid.get("columnCount") or 26))
    headers_by_pk: dict[str, list[str]] = {}
    for diff in result["diffs"]:
        headers_by_pk.setdefault(diff["pk"], [])
        if diff["header"] not in headers_by_pk[diff["pk"]]:
            headers_by_pk[diff["pk"]].append(diff["header"])
    # 比べた後に変わった行（反映待ちになった・版が進んだ）は書かない
    with session_factory() as db:
        current = {}
        for pk in headers_by_pk:
            st = db.get(SheetSyncState, (spec.table, pk))
            model = SYNCED_TABLES[spec.table]
            row = db.get(model, pk if spec.table == "requests" else int(pk))
            current[pk] = (row is not None and int(row.row_version) == by_pk[pk]["version"]
                           and not (st and st.desired_version > st.synced_version))
    works = []
    for pk, headers in headers_by_pk.items():
        if not current.get(pk):
            continue
        d = by_pk[pk]
        works.append(row_sync.RowWork(table=spec.table, pk=pk, version=d["version"], values=d["values"],
                                      kinds=d["kinds"], synced=d["synced"], sheet_row=d["sheet_row"], seq=0,
                                      first_dirty_at=None, attempts=0, target_row=d["sheet_row"], headers=headers))
    if not works:
        return 0
    row_sync.load_refs(api, sid, ctx, timeutil.now_utc())
    row_sync.write_rows(api, sid, ctx, works, tz, purpose="reconcile")
    fixed = 0
    with session_factory() as db:
        for w in works:
            if w.error:
                continue
            st = db.get(SheetSyncState, (spec.table, w.pk))
            if st is not None and st.desired_version <= st.synced_version:
                st.synced_values = row_sync.merge_synced(st.synced_values, w.written or {})
            fixed += 1
        db.commit()
    return fixed


def run(session_factory: Callable[[], Session], *, check_only: bool = False, api: SheetsApi | None = None,
        app_settings: Any = None) -> dict:
    """照合を 1 回。check_only なら書かない。戻り値は sync_runs の summary と同じ形（ok を含む）。"""
    from app.config import get_settings

    settings = app_settings or get_settings()
    started = timeutil.now_utc()
    summary: dict = {"checkOnly": check_only, "tabs": {}, "_started": started}
    with session_factory() as db:
        counts = {t: int(db.scalar(select(func.count()).select_from(SYNCED_TABLES[t])) or 0) for t in codec.TABLE_ORDER}
    if counts["requests"] == 0 or counts["work_logs"] == 0:
        summary.update({"ok": False, "stopped": f"DB の requests（{counts['requests']} 行）か work_logs"
                                                f"（{counts['work_logs']} 行）が 0 行のため、照合を止めました（書きません）。"})
        _record(session_factory, False, summary)
        return {k: v for k, v in summary.items() if not k.startswith("_")}
    sid = settings.db_spreadsheet_id
    tz = settings.db_sheet_tz
    try:
        if api is None:
            api = row_sync.api_factory(session_factory, settings)
        meta = row_sync.check_timezone(api, sid, settings)
        results, reads, db_tables = {}, {}, {}
        for table in codec.TABLE_ORDER:
            spec = codec.TABLES[table]
            a1 = quote_sheet(spec.tab)
            unformatted, formatted = api.read_both(sid, [a1])
            reads[table] = (unformatted[0], formatted[0])
            with session_factory() as db:
                db_tables[table] = db_rows_for(db, table)
            results[table] = compare_tab(spec, unformatted[0], formatted[0], db_tables[table], tz)
        short = [f"{codec.TABLES[t].tab}（DB {results[t]['dbRows']} 行・シート {results[t]['sheetRows']} 行）"
                 for t in codec.TABLE_ORDER if results[t]["dbRows"] < results[t]["sheetRows"]]
        for table in codec.TABLE_ORDER:
            r = results[table]
            summary["tabs"][codec.TABLES[table].tab] = {
                "dbRows": r["dbRows"], "sheetRows": r["sheetRows"], "compared": r["compared"],
                "diffCount": len(r["diffs"]), "diffs": r["diffs"][:MAX_LISTED],
                "idMismatch": r["idMismatch"][:MAX_LISTED], "extraRows": r["extraRows"][:MAX_LISTED],
                "skipped": sum(1 for d in db_tables[table] if d["skip"]),
            }
        total = sum(len(r["diffs"]) for r in results.values())
        mismatched = sum(len(r["idMismatch"]) for r in results.values())
        summary["diffCount"] = total
        summary["idMismatchCount"] = mismatched
        if short:
            summary.update({"ok": False, "stopped": "DB の行がシートより少ないため、1 行も書かずに止めました: "
                                                    + "、".join(short)})
            _record(session_factory, False, summary)
            return {k: v for k, v in summary.items() if not k.startswith("_")}
        fixed = 0
        if not check_only and total:
            # 書く前の関所（データ安全 高-3）: GAS が移転モードでなければ書かずに報告だけ
            with session_factory() as db:
                allowed, reason = write_gate.check(db, settings)
            if not allowed:
                summary.update({"ok": False, "fixedRows": 0, "writeBlocked": reason,
                                "stopped": "書き込みの関所で止めました（差は報告だけ）: " + reason})
                _record(session_factory, False, summary)
                return {k: v for k, v in summary.items() if not k.startswith("_")}
            for table in codec.TABLE_ORDER:
                if results[table]["diffs"]:
                    fixed += _fix_tab(api, sid, codec.TABLES[table], meta, reads[table][0], results[table],
                                      db_tables[table], session_factory, tz)
        summary["fixedRows"] = fixed
        ok = (total == 0 and mismatched == 0) if check_only else mismatched == 0
        summary["ok"] = ok
    except (row_sync.LaneHalt, UnsafeWriteError, SheetsError) as error:
        summary.update({"ok": False, "stopped": str(error)[:1000]})
        _record(session_factory, False, summary)
        return {k: v for k, v in summary.items() if not k.startswith("_")}
    _record(session_factory, summary["ok"], summary)
    return {k: v for k, v in summary.items() if not k.startswith("_")}


def _scheduled(session_factory, check_only: bool = False) -> dict:
    return run(session_factory, check_only=check_only)


def register_worker() -> None:
    from app import worker

    worker.register_schedule("sheet_reconcile", task=_scheduled, at=(2, 10))


__all__ = ["compare_tab", "db_rows_for", "run", "register_worker"]
