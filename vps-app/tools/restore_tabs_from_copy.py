"""戻しの道具: 切り替え前の控え（tools/cutover_backup.py が作った複製）から、名前を書いたタブの値を
この環境の DB スプレッドシートへ写し戻す。何もしない確かめ（dry-run）が既定。

使い方（VPS。docs/cutover-runbook.md の B-15・B-16・C 章）:
  docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api \
      python tools/restore_tabs_from_copy.py --from-copy <控えの ID> --tabs 作業員マスタ,工番マスタ [--apply]

安全の決まり:
  - 書き先はこの環境の DB スプレッドシートだけ（safety.assert_restore_target）。控えの ID と同じ・許可リストの外は止める。
  - 控えは名前が「【切替前の控え】残業・休日出勤申請 DB」で始まるスプレッドシートだけ（台帳の控えや別のファイルを拒む）。
  - タブは --tabs に名前を書いたものだけ。`all` は A 区分と C 区分（Requests・WorkLogs・BatchLogs・マスタ 4 タブ）。
    Settings・DeptApprovers・ApproverMap は `all` に入らず、名前を書いたときだけ戻す。ほかのタブは書かない。
  - --apply が無ければ書かない（タブごとの行数・違うセルの数だけを出す）。
  - 値は UNFORMATTED_VALUE（日時はシリアル値）で読み、RAW で書く（文字は式にならない＝セキュリティ 高-1）。
    表示形式は今のタブのまま。控えに式のセルがあれば値として書くので、報告の formulaCells に数を出す。
    控えより下・右の余った値は消す（値だけ。行・列・タブは消さない・作らない）。
  - worker を止めてから使う（行の反映・照合が同じタブを書かないように）。
  - Requests に、控えに無い requestId の行（控えの後に Sebastian が直接足した行など。決定 #26）があれば、--apply を
    止める（写し戻すとその行が消えるため）。報告の rowsNotInCopy に行番号と requestId を出す。それでも戻すと決めたときだけ
    --drop-rows-not-in-copy を付ける（先にその行の中身を控えてから）。
終了コード: 0＝成功（dry-run を含む）、2＝止めた。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Optional

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "api"))

from app import safety  # noqa: E402
from app.services.sheets_api import API_BASE, column_letter, quote_sheet, sheet_titles  # noqa: E402

BACKUP_DB_PREFIX = "【切替前の控え】残業・休日出勤申請 DB"
ID_HEADERS = {"Requests": (1, "requestId")}  # 控えに無い行を確かめるタブ → (見出しの行, ID の列の見出し)
ALL_TABS = ("Requests", "WorkLogs", "BatchLogs", "作業員マスタ", "工番マスタ", "社内カレンダーマスタ", "業務NOマスタ")


class RestoreStop(RuntimeError):
    """戻しを止める理由。"""


def parse_tabs(text: str) -> tuple[list[str], set[str]]:
    """--tabs の値 → （タブの一覧, 名前を書いたタブ）。"""
    named: list[str] = []
    for part in (text or "").split(","):
        name = part.strip()
        if name and name not in named:
            named.append(name)
    if not named:
        raise RestoreStop("--tabs にタブの名前を書いてください（all は A 区分と C 区分）")
    tabs: list[str] = []
    explicit: set[str] = set()
    for name in named:
        if name == "all":
            tabs.extend(t for t in ALL_TABS if t not in tabs)
        else:
            explicit.add(name)
            if name not in tabs:
                tabs.append(name)
    return tabs, explicit


def _grid_size(meta: dict, tab: str) -> tuple[int, int]:
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == tab:
            g = s["properties"].get("gridProperties") or {}
            return int(g.get("rowCount") or 0), int(g.get("columnCount") or 0)
    return 0, 0


def _trim(grid: list[list]) -> list[list]:
    rows = [list(r) for r in grid]
    while rows and all(v in ("", None) for v in rows[-1]):
        rows.pop()
    return rows


def _pad(grid: list[list], width: int) -> list[list]:
    return [[("" if v is None else v) for v in (r + [""] * (width - len(r)))[:width]] for r in grid]


def plan_tab(api: Any, copy_id: str, target_id: str, tab: str) -> dict:
    """控えと今のタブを読み、書く値・消す範囲・違うセルの数を決める。"""
    q = quote_sheet(tab)
    src = _trim(api.get_values(copy_id, q, "UNFORMATTED_VALUE"))
    formulas = api.get_values(copy_id, q, "FORMULA")
    formula_cells = sum(1 for row in formulas for v in row if isinstance(v, str) and v.startswith("="))
    cur = _trim(api.get_values(target_id, q, "UNFORMATTED_VALUE"))
    width = max([len(r) for r in src] + [1])
    values = _pad(src, width)
    cur_width = max([len(r) for r in cur] + [0])
    cur_cells = _pad(cur, max(width, cur_width))
    diff = 0
    for r in range(max(len(values), len(cur_cells))):
        a = values[r] if r < len(values) else []
        b = cur_cells[r] if r < len(cur_cells) else []
        for c in range(max(len(a), len(b))):
            x = a[c] if c < len(a) else ""
            y = b[c] if c < len(b) else ""
            if x != y:
                diff += 1
    return {"tab": tab, "values": values, "rows": len(values), "cols": width, "currentRows": len(cur),
            "currentCols": cur_width, "diffCells": diff, "formulaCells": formula_cells,
            "rowsNotInCopy": rows_not_in_copy(tab, src, cur)}


def _ids(tab: str, grid: list[list]) -> list[tuple[int, str]]:
    header_row, header = ID_HEADERS[tab]
    head = [str(v).strip() for v in (grid[header_row - 1] if len(grid) >= header_row else [])]
    if header not in head:
        return []
    c = head.index(header)
    out = []
    for r in range(header_row, len(grid)):
        v = grid[r][c] if c < len(grid[r]) else ""
        text = str(v).strip() if v not in (None, "") else ""
        if text:
            out.append((r + 1, text))
    return out


def rows_not_in_copy(tab: str, src: list[list], cur: list[list]) -> list[dict]:
    """今のタブにあって控えに無い ID の行（写し戻すと消える行）。ID の列を持つタブ（Requests）だけ。"""
    if tab not in ID_HEADERS:
        return []
    in_copy = {i for _, i in _ids(tab, src)}
    return [{"row": r, "requestId": i} for r, i in _ids(tab, cur) if i not in in_copy]


def apply_tab(api: Any, target_id: str, plan: dict, meta: dict) -> None:
    tab = plan["tab"]
    q = quote_sheet(tab)
    rows, cols = plan["rows"], plan["cols"]
    grid_rows, grid_cols = _grid_size(meta, tab)
    if rows:
        data = [{"range": f"{q}!A1:{column_letter(cols - 1)}{rows}", "values": plan["values"]}]
        # 書く関数は SheetsApi の安全装置（目的の許可リスト）を通らないので、assert_restore_target で確かめてから呼ぶ
        api._call("POST", f"{API_BASE}/{target_id}/values:batchUpdate",
                  body={"valueInputOption": "RAW", "data": data})
    clear = []
    last_col = column_letter(max(grid_cols, cols, plan["currentCols"], 1) - 1)
    if max(grid_rows, plan["currentRows"]) > rows:
        clear.append(f"{q}!A{rows + 1}:{last_col}")
    if max(grid_cols, plan["currentCols"]) > cols and rows:
        clear.append(f"{q}!{column_letter(cols)}1:{last_col}{rows}")
    if clear:
        safety.assert_ranges_in_tab(clear, tab)
        api._call("POST", f"{API_BASE}/{target_id}/values:batchClear", body={"ranges": clear})


def run(api: Any, settings: Any, copy_id: str, tabs_arg: str, *, apply: bool,
        drop_rows_not_in_copy: bool = False) -> dict:
    copy_id = (copy_id or "").strip()
    target_id = safety.writable_spreadsheet_id(settings)
    tabs, explicit = parse_tabs(tabs_arg)
    if not copy_id:
        raise RestoreStop("--from-copy が空です")
    if copy_id == target_id or copy_id in safety.never_writable_ids(settings):
        raise RestoreStop("--from-copy に DB・台帳・マスタ元そのものは使えません（控えの ID を渡す）")
    for tab in tabs:
        try:
            safety.assert_restore_target(target_id, tab, settings, protected_listed=tab in explicit)
        except safety.UnsafeWriteError as error:
            raise RestoreStop(str(error)) from error
    copy_meta = api.get_spreadsheet(copy_id)
    title = str((copy_meta.get("properties") or {}).get("title") or "")
    if not title.startswith(BACKUP_DB_PREFIX):
        raise RestoreStop(f"控えの名前が「{BACKUP_DB_PREFIX}」で始まりません")
    target_meta = api.get_spreadsheet(target_id)
    missing = [t for t in tabs if t not in set(sheet_titles(copy_meta)) or t not in set(sheet_titles(target_meta))]
    if missing:
        raise RestoreStop(f"控えか DB に無いタブがあります（タブは作りません）: {missing}")
    report: dict = {"appEnv": settings.app_env, "apply": apply, "copyId": copy_id, "targetId": target_id, "tabs": []}
    plans = [plan_tab(api, copy_id, target_id, t) for t in tabs]
    for p in plans:
        report["tabs"].append({k: v for k, v in p.items() if k != "values"})
    lost = {p["tab"]: p["rowsNotInCopy"] for p in plans if p["rowsNotInCopy"]}
    if apply and lost and not drop_rows_not_in_copy:
        raise RestoreStop("控えに無い行があります（写し戻すと消えます。控えの後に Sebastian などが足した行）: "
                          + "; ".join(f"{t}: " + ", ".join(f"{x['row']} 行目 {x['requestId']}" for x in rows[:20])
                                      for t, rows in lost.items())
                          + "。中身を控えてから、それでも戻すときだけ --drop-rows-not-in-copy を付けます")
    if apply:
        for p in plans:
            safety.assert_restore_target(target_id, p["tab"], settings, protected_listed=p["tab"] in explicit)
            apply_tab(api, target_id, p, target_meta)
        report["applied"] = [p["tab"] for p in plans]
    return report


def main(argv: Optional[list] = None, *, api: Any = None, settings: Any = None) -> int:
    parser = argparse.ArgumentParser(description="控えからタブの値を DB スプレッドシートへ写し戻す（既定は dry-run）")
    parser.add_argument("--from-copy", required=True, help="tools/cutover_backup.py が出した DB の控えの ID")
    parser.add_argument("--tabs", required=True, help="カンマ区切りのタブの名前。all は A 区分と C 区分")
    parser.add_argument("--apply", action="store_true", help="本当に書く（無ければ確かめるだけ）")
    parser.add_argument("--drop-rows-not-in-copy", action="store_true",
                        help="控えに無い Requests の行（Sebastian の追記など）が消えてもよいときだけ付ける")
    args = parser.parse_args(argv)
    try:
        if settings is None:
            from app.config import get_settings

            settings = get_settings()
        if api is None:
            from sqlalchemy.orm import Session

            from app.db import get_engine
            from app.services.sheets_api import SheetsApi

            with Session(get_engine()) as db:
                api = SheetsApi.for_db(db, settings)
        report = run(api, settings, args.from_copy, args.tabs, apply=args.apply,
                     drop_rows_not_in_copy=args.drop_rows_not_in_copy)
    except RestoreStop as error:
        print(f"止めました: {error}", file=sys.stderr)
        return 2
    except Exception as error:  # noqa: BLE001 - 委任が無い・API の失敗など
        print(f"止めました: {type(error).__name__}: {str(error)[:300]}", file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=False, indent=1, default=str))
    if not args.apply:
        print("確かめだけです（書いていません）。書くときは --apply を付けます。", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
