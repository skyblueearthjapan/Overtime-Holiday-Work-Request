"""共有タブの取り込み（実装仕様 6-9・5-2。`sync_sheet_config`、SHEET_CONFIG_SYNC_MINUTES ごとと手動。担当 D）。

DB スプレッドシートの B 区分のタブを読み、DB の表を全置換する。タブには書かない（読むだけ）。

- Settings（1 行目説明・2 行目見出し・3 行目から A=key、B=value）→ settings。キーは前後の空白を取る
  （GAS の getSettings_ と同じ String(key).trim()。値はそのまま）。同じキーが 2 行以上あれば**下の行が勝つ**
  （GAS と同じ）。重なったキーは overlaps に出す（gap-5 中-2・gap-4 F-14）。
- DeptApprovers（見出し 1 行目、5 列。貯蔵品のアプリと共有＝読むだけ）→ dept_approvers。A・B・C を位置で、
  D 役職・E 通知メールは持つだけ（gap-4 F-5）。
- ApproverMap（1 行目から生のまま）→ approver_map_rows。permissions が GAS と同じく 1 行目を見出しとして読む。
- 読めない・空・有効な行 0 なら、その表は上書きしない（ApproverMap は今も見出しだけなので、行が 1 つも無いときだけ）。
- Settings の SOURCE_SSID・TEMPLATE_SSID・PDF_ROOT_FOLDER_ID・ACCUMULATION_SS_ID が .env の値と違えば mismatches に出す
  （書き込み先は .env が決める）。
- StampMap・休憩マスタ・部署マスタは取り込まない（VPS で使う所が無い。5-2）。

戻り値 {ok, counts, overlaps, mismatches, skipped}。sync_runs（kind='sync_sheet_config'）に記録する。commit は呼ぶ側
（worker の予定は自分で commit する）。
"""
from __future__ import annotations

import logging
from typing import Any

from sqlalchemy import delete
from sqlalchemy.orm import Session

from app import timeutil
from app.constants import TAB_APPROVER_MAP, TAB_DEPT_APPROVERS, TAB_SETTINGS
from app.jsutil import js_string, js_trim, normalize_gas
from app.models import ApproverMapRow, DeptApprover, Setting, SyncRun
from app.services.sheets_api import SheetsApi, quote_sheet

log = logging.getLogger(__name__)

RUN_KIND = "sync_sheet_config"


def _blank(v: Any) -> bool:
    return v is None or v == ""


def _cell(row: list, i: int) -> Any:
    return row[i] if i < len(row) else ""


def parse_settings(values: list) -> tuple[dict[str, tuple[str, int]], list[dict]]:
    """Settings の値（1 行目から）→ ({key: (value, sheet_row)}, 重なったキー)。3 行目から読む。下の行が勝つ。"""
    out: dict[str, tuple[str, int]] = {}
    rows_by_key: dict[str, list[int]] = {}
    for r in range(2, len(values)):
        row = values[r]
        key = js_trim(js_string(_cell(row, 0)))
        if not key:
            continue
        out[key] = (js_string(_cell(row, 1)), r + 1)
        rows_by_key.setdefault(key, []).append(r + 1)
    overlaps = [{"key": k, "rows": rows, "used": rows[-1]} for k, rows in rows_by_key.items() if len(rows) > 1]
    return out, overlaps


def parse_dept_approvers(values: list) -> list[dict]:
    """DeptApprovers の値（1 行目は見出し）→ データ行（A〜E。何か値のある行）。"""
    out = []
    for r in range(1, len(values)):
        row = values[r]
        cells = [js_string(_cell(row, i)) for i in range(5)]
        if not any(cells):
            continue
        out.append({"sheet_row": r + 1, "dept": cells[0], "emails": cells[1], "name": cells[2],
                    "col_d": cells[3], "col_e": cells[4]})
    return out


def valid_dept_rows(rows: list[dict]) -> int:
    """GAS の getDeptApproverMap_ が使う行（部署が空でない）の数。"""
    return sum(1 for r in rows if normalize_gas(r["dept"]))


def parse_approver_map(values: list) -> list[dict]:
    """ApproverMap の値（1 行目から生のまま。行末の空を除く）。"""
    out = []
    for r, row in enumerate(values):
        cells = list(row)
        while cells and _blank(cells[-1]):
            cells.pop()
        if not cells:
            continue
        out.append({"sheet_row": r + 1, "cells": cells})
    return out


def env_mismatches(settings_map: dict[str, str], app_settings: Any) -> list[dict]:
    """Settings の ID が .env と違う点（書き込み先は .env が決める）。"""
    expected = {
        "SOURCE_SSID": app_settings.master_source_spreadsheet_id,
        "TEMPLATE_SSID": app_settings.db_spreadsheet_id,
        "PDF_ROOT_FOLDER_ID": app_settings.pdf_folder_id,
        "ACCUMULATION_SS_ID": app_settings.ledger_spreadsheet_id,
    }
    out = []
    for key, want in expected.items():
        have = normalize_gas(settings_map.get(key, ""))
        if have != want:
            out.append({"key": key, "sheet": have, "env": want})
    return out


def _read(api: SheetsApi, sid: str, tab: str) -> list:
    return api.get_values(sid, quote_sheet(tab), "UNFORMATTED_VALUE")


def sync_sheet_config(db: Session, api: SheetsApi | None = None, app_settings: Any = None) -> dict:
    """Settings・DeptApprovers・ApproverMap を取り込む（読むだけ。commit は呼ぶ側）。"""
    from app.config import get_settings

    settings = app_settings or get_settings()
    started = timeutil.now_utc()
    result: dict = {"ok": True, "counts": {}, "overlaps": [], "mismatches": [], "skipped": {}}
    sid = settings.db_spreadsheet_id
    try:
        if api is None:
            api = SheetsApi.for_db(db, settings)
    except Exception as error:  # noqa: BLE001 - 委任が無いなど。今の表を残す
        result.update({"ok": False, "error": f"{type(error).__name__}: {error}"[:500]})
        _record(db, started, result)
        return result

    # Settings
    try:
        parsed, overlaps = parse_settings(_read(api, sid, TAB_SETTINGS))
        result["overlaps"] = overlaps
        if not parsed:
            result["skipped"]["Settings"] = "キーが 1 つもありません（今の値を残します）"
        else:
            with db.begin_nested():
                db.execute(delete(Setting))
                db.add_all(Setting(key=k, value=v, sheet_row=r) for k, (v, r) in parsed.items())
            result["counts"]["Settings"] = len(parsed)
            result["mismatches"] = env_mismatches({k: v for k, (v, _r) in parsed.items()}, settings)
    except Exception as error:  # noqa: BLE001 - タブごとに握り、ほかのタブは続ける
        result["skipped"]["Settings"] = f"読めません: {error}"[:500]

    # DeptApprovers
    try:
        rows = parse_dept_approvers(_read(api, sid, TAB_DEPT_APPROVERS))
        if valid_dept_rows(rows) == 0:
            result["skipped"]["DeptApprovers"] = "有効な行がありません（今の値を残します）"
        else:
            with db.begin_nested():
                db.execute(delete(DeptApprover))
                db.add_all(DeptApprover(**r) for r in rows)
            result["counts"]["DeptApprovers"] = len(rows)
    except Exception as error:  # noqa: BLE001
        result["skipped"]["DeptApprovers"] = f"読めません: {error}"[:500]

    # ApproverMap
    try:
        rows = parse_approver_map(_read(api, sid, TAB_APPROVER_MAP))
        if not rows:
            result["skipped"]["ApproverMap"] = "タブが空です（今の値を残します）"
        else:
            with db.begin_nested():
                db.execute(delete(ApproverMapRow))
                db.add_all(ApproverMapRow(**r) for r in rows)
            result["counts"]["ApproverMap"] = len(rows)
    except Exception as error:  # noqa: BLE001
        result["skipped"]["ApproverMap"] = f"読めません: {error}"[:500]

    result["ok"] = not result["skipped"]
    _record(db, started, result)
    if result["overlaps"] or result["mismatches"]:
        log.warning("sync_sheet_config: overlaps=%s mismatches=%s", result["overlaps"], result["mismatches"])
    return result


def _record(db: Session, started, result: dict) -> None:
    db.add(SyncRun(kind=RUN_KIND, started_at=started, finished_at=timeutil.now_utc(), ok=bool(result.get("ok")),
                   summary_json=dict(result)))
    db.flush()


def _scheduled(session_factory) -> dict:
    with session_factory() as db:
        result = sync_sheet_config(db)
        db.commit()
    return result


def register_worker() -> None:
    from app import worker
    from app.config import get_settings

    worker.register_schedule("sync_sheet_config", task=_scheduled,
                             every_minutes=get_settings().sheet_config_sync_minutes)
