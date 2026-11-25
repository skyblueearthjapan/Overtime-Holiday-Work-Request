"""本番 DB スプレッドシート → VPS の DB の移行（スプレッドシートは読むだけ。実装仕様 5 章・計画書 5 章。担当 D）。

移すタブ:
  A 区分（行単位でシートへ反映する）: Requests（見出し 1 行目）・WorkLogs（1 行目説明・見出し 2 行目）・BatchLogs
  B 区分（取り込むだけ）: Settings（見出し 2 行目、3 行目から）・DeptApprovers（5 列）・ApproverMap（1 行目から生のまま）
  報告だけ: StampMap・休憩マスタ（見出し 2 行目）の行数
  移さない: マスタ 4 タブ（切り替えで sync_masters がマスタ元から作る）・FormMap ほか

使い方（VPS。tools はイメージに入っていないので付けて動かす＝罠 5）:
  docker compose run --rm -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api \
      python tools/migrate_from_sheets.py --source sheets --dry-run --report-out /tmp/report.json
  docker compose run --rm -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api \
      python tools/migrate_from_sheets.py --source sheets [--truncate]

--source sheets は DB に保存した委任トークン（google_delegate）で読むので、--dry-run でも DB と .env が要る。

安全の決まり:
  - 読む元は本番 DB スプレッドシート（config.PROD_SPREADSHEET_ID_FIXED）だけ。ほかの ID は止める（罠 12。
    staging への移行も本番を読む＝読むだけ）。gviz は使わない。Sheets API で UNFORMATTED と FORMATTED を一緒に読む（罠 11）。
  - 列は見出しの文字で照合する（GAS の normalize_ と同じに直して）。決まりの見出しの欠け・重なりは止める（missingColumns）。
    決まりに無い見出しは droppedColumns に出す（移さない）。
  - status は submitted・approved・canceled だけ（modified があれば止める＝罠 13）。requestType は overtime・holiday だけ。
    requestId の空・重なりは止める。読めない日時・数は止める。
  - WorkLogs は Requests と 1 対 1 ではない。孤児（申請に無い ID）も移し、行の無い申請に行を作らない。
  - 各行に元の行番号（sheet_row）、型が混ざる列の行ごとの型（cell_kinds）、行の反映の状態（sheet_sync_state:
    synced_values＝今シートにある値）を入れる。移した直後の照合（reconcile.compare_tab）は差 0 になる。
  - 移行先の表（業務・B 区分・sheet_sync_state・jobs・idempotency_keys・sync_runs・sessions）のどれかに行があれば止める。
    --truncate を付けたときだけ、それらをすべて空にしてから入れる（罠 24b）。1 トランザクションで入れる。
  - 入れた後に連番を最大値の先へ進める（罠 9）。
  - Requests・WorkLogs・Settings のどれかが 0 行なら止める（--dry-run でも。罠 49）。
終了コード: 0＝成功、2＝止めた、1＝入れた後の照合が合わない。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "api"))

from app.config import PROD_SPREADSHEET_ID_FIXED  # noqa: E402
from app.constants import REQUEST_STATUSES, REQUEST_TYPES  # noqa: E402
from app.jsutil import normalize_gas  # noqa: E402
from app.services import config_sync, reconcile  # noqa: E402
from app.services import sheet_codec as codec  # noqa: E402
from app.services.sheets_api import SheetsError, quote_sheet, sheet_titles, spreadsheet_timezone  # noqa: E402

A_TABLES = codec.TABLE_ORDER
B_TABS = {"settings": "Settings", "dept_approvers": "DeptApprovers", "approver_map_rows": "ApproverMap"}
REPORT_TABS = {"StampMap": 1, "休憩マスタ": 2}  # タブ → 見出しの行
TARGET_TABLES = ("requests", "work_logs", "batch_logs", "settings", "dept_approvers", "approver_map_rows",
                 "sheet_sync_state")
# --truncate で空にする表（罠 24b: 業務の表に加え、仕事・冪等キー・同期の記録・セッション）
TRUNCATE_TABLES = TARGET_TABLES + ("jobs", "idempotency_keys", "sync_runs", "sessions")
REPORTED_SETTING_KEYS = ("SOURCE_SSID", "TEMPLATE_SSID", "PDF_ROOT_FOLDER_ID", "ACCUMULATION_SS_ID", "APP_URL",
                         "PDF_MODE")


class MigrationStop(RuntimeError):
    """移行を止める理由（安全装置）。"""


# ===== 読み込み元 =====


def check_source_id(spreadsheet_id: str) -> None:
    if spreadsheet_id != PROD_SPREADSHEET_ID_FIXED:
        raise MigrationStop(f"移行元が本番 DB スプレッドシートの ID ではありません: {spreadsheet_id}")


def read_sheets(api: Any, spreadsheet_id: str) -> dict:
    """本番 DB を読む（読むだけ）。{"tz", "tabs": {タブ: (UNFORMATTED, FORMATTED)}}。"""
    check_source_id(spreadsheet_id)
    meta = api.get_spreadsheet(spreadsheet_id)
    titles = set(sheet_titles(meta))
    tabs = [codec.TABLES[t].tab for t in A_TABLES] + list(B_TABS.values())
    missing = [t for t in tabs if t not in titles]
    if missing:
        raise MigrationStop(f"タブがありません: {', '.join(missing)}")
    wanted = tabs + [t for t in REPORT_TABS if t in titles]
    unformatted, formatted = api.read_both(spreadsheet_id, [quote_sheet(t) for t in wanted])
    return {"tz": spreadsheet_timezone(meta), "tabs": {t: (u, f) for t, u, f in zip(wanted, unformatted, formatted)}}


# ===== 読んだ値 → 表の行 =====


def _row(grid: list, r: int) -> list:
    return grid[r - 1] if 0 < r <= len(grid) else []


def decode_table(spec: codec.TableSpec, unformatted: list, formatted: list, tz) -> dict:
    header = _row(unformatted, spec.header_row)
    normalized = [normalize_gas(h) for h in header]
    dup = sorted(h for h, n in Counter(x for x in normalized if x in spec.headers).items() if n > 1)
    missing = codec.missing_headers(spec, header)
    dropped = [h for h in normalized if h and h not in spec.headers]
    col = codec.header_index(header)
    rows, errors, skipped = [], [], 0
    for r in range(spec.data_start_row, len(unformatted) + 1):
        urow, frow = _row(unformatted, r), _row(formatted, r)
        if all(codec.is_blank(v) for v in urow):
            skipped += 1
            continue
        record: dict = {"sheet_row": r, "row_version": 1, "cell_kinds": {}}
        for rule in spec.columns:
            c = col.get(rule.header, -1)
            try:
                value, kind = codec.decode_cell(rule, codec.cell_at(urow, c), codec.cell_at(frow, c), tz)
            except codec.CodecError as error:
                errors.append(f"{spec.tab} {r} 行目: {error}")
                continue
            codec.assign_db_value(rule, record, value)
            if rule.date_like and kind:
                record["cell_kinds"][rule.header] = kind
        rows.append(record)
    return {"rows": rows, "missing": missing, "duplicated": dup, "dropped": dropped, "errors": errors,
            "skippedEmptyRows": skipped, "headerCount": len([h for h in normalized if h])}


def prepare(reads: dict) -> dict:
    """読んだタブを確かめ、入れる行と報告の材料を作る。止める理由があれば MigrationStop。"""
    tz = codec.ZoneInfo(reads["tz"] or "Asia/Tokyo")
    decoded = {t: decode_table(codec.TABLES[t], *reads["tabs"][codec.TABLES[t].tab], tz) for t in A_TABLES}
    problems = []
    for t, d in decoded.items():
        tab = codec.TABLES[t].tab
        if d["missing"]:
            problems.append(f"{tab}: 決まりの見出しがありません（missingColumns）: {d['missing']}")
        if d["duplicated"]:
            problems.append(f"{tab}: 同じ見出しが重なっています: {d['duplicated']}")
        problems.extend(d["errors"][:20])
    if problems:
        raise MigrationStop("\n".join(problems))

    requests = decoded["requests"]["rows"]
    blank = [r["sheet_row"] for r in requests if not normalize_gas(r["request_id"])]
    if blank:
        raise MigrationStop(f"Requests: requestId が空の行があります（行 {blank[:20]}）")
    dup = sorted(k for k, n in Counter(r["request_id"] for r in requests).items() if n > 1)
    if dup:
        raise MigrationStop(f"Requests: 同じ requestId が重なっています: {dup[:20]}")
    bad_status = Counter(r["status"] for r in requests if r["status"] not in REQUEST_STATUSES)
    if bad_status:
        raise MigrationStop(f"Requests: 決まりに無い status があります（modified を含む。扱いをオーナーと決める）: "
                            f"{dict(bad_status)}")
    bad_type = Counter(r["request_type"] for r in requests if r["request_type"] not in REQUEST_TYPES)
    if bad_type:
        raise MigrationStop(f"Requests: 決まりに無い requestType があります: {dict(bad_type)}")

    for t in ("work_logs", "batch_logs"):
        for n, r in enumerate(decoded[t]["rows"], start=1):
            r["id"] = n

    settings_parsed, overlaps = config_sync.parse_settings(reads["tabs"]["Settings"][0])
    dept_rows = config_sync.parse_dept_approvers(reads["tabs"]["DeptApprovers"][0])
    approver_rows = config_sync.parse_approver_map(reads["tabs"]["ApproverMap"][0])
    tables = {t: decoded[t]["rows"] for t in A_TABLES}
    tables["settings"] = [{"key": k, "value": v, "sheet_row": r} for k, (v, r) in settings_parsed.items()]
    tables["dept_approvers"] = dept_rows
    tables["approver_map_rows"] = approver_rows

    empty = [name for name, n in (("Requests", len(tables["requests"])), ("WorkLogs", len(tables["work_logs"])),
                                  ("Settings", len(tables["settings"]))) if n == 0]
    if empty:
        raise MigrationStop(f"{'・'.join(empty)} が 0 行です（空のシートを移しません。罠 49）")

    warnings = []
    blank_wl = [r["sheet_row"] for r in tables["work_logs"] if not normalize_gas(r["request_id"])]
    if blank_wl:
        warnings.append(f"WorkLogs: requestId が空で中身のある行があります（そのまま移します）: 行 {blank_wl[:20]}")
    if overlaps:
        warnings.append(f"Settings: 同じキーが 2 行以上あります（下の行を使います）: {overlaps}")
    valid = config_sync.valid_dept_rows(dept_rows)
    if not valid:
        warnings.append("DeptApprovers: 有効な行がありません")
    report_tabs = {}
    for tab, header_row in REPORT_TABS.items():
        if tab in reads["tabs"]:
            grid = reads["tabs"][tab][0]
            report_tabs[tab] = sum(1 for row in grid[header_row:] if any(not codec.is_blank(v) for v in row))
    return {
        "tables": tables, "tz": reads["tz"], "warnings": warnings, "settingsOverlaps": overlaps,
        "droppedColumns": {codec.TABLES[t].tab: decoded[t]["dropped"] for t in A_TABLES if decoded[t]["dropped"]},
        "skippedEmptyRows": {codec.TABLES[t].tab: decoded[t]["skippedEmptyRows"] for t in A_TABLES},
        "headerCounts": {codec.TABLES[t].tab: decoded[t]["headerCount"] for t in A_TABLES},
        "reportOnlyTabs": report_tabs, "validDeptApprovers": valid,
    }


# ===== 照合の数（計画書 5-4。シートから読んだ行と DB から読み直した行の両方に同じ関数を使う） =====


def _month(d: Any) -> str:
    return d.strftime("%Y-%m") if isinstance(d, date) else "invalid"


def _blank(v: Any) -> bool:
    return v is None or v == ""


def summarize(tables: dict) -> dict:
    requests, logs, batches = tables["requests"], tables["work_logs"], tables["batch_logs"]
    first_log: dict[str, dict] = {}
    for w in sorted(logs, key=lambda w: (w["sheet_row"] or 0)):
        first_log.setdefault(normalize_gas(w["request_id"]), w)
    req_ids = {r["request_id"] for r in requests}
    req_month = {r["request_id"]: _month(r["target_date"]) for r in requests}

    def total(rows, key):
        return sum(int(r[key]) for r in rows if r.get(key) is not None)

    by_month: dict[str, dict] = defaultdict(lambda: {"approvedMinutes": 0, "actualMinutes": 0, "breakMinutes": 0,
                                                     "netMinutes": 0})
    for r in requests:
        by_month[_month(r["target_date"])]["approvedMinutes"] += int(r["approved_minutes"] or 0)
    for w in logs:
        m = req_month.get(w["request_id"], "orphan")
        for k, attr in (("actualMinutes", "actual_minutes"), ("breakMinutes", "break_minutes"),
                        ("netMinutes", "net_minutes")):
            by_month[m][k] += int(w[attr] or 0)

    def retro(r) -> bool:
        return (r["submitted_at"] is not None and isinstance(r["target_date"], date)
                and r["submitted_at"].astimezone(codec.JST).date() > r["target_date"])

    def wl(r):
        return first_log.get(normalize_gas(r["request_id"]))

    def net(r) -> int:
        w = wl(r)
        return int(w["net_minutes"] or 0) if w else 0

    def pdf_before_done(r) -> bool:
        w = wl(r)
        end = codec.parse_jst_datetime(w["actual_end_at"]) if w and w["actual_end_at"] else None
        return bool(r["pdf_file_id"] and r["pdf_generated_at"] and end and r["pdf_generated_at"] < end)

    kinds: dict[str, dict] = {}
    for t in A_TABLES:
        c: Counter = Counter()
        for row in tables[t]:
            for h, k in (row.get("cell_kinds") or {}).items():
                c[f"{h}:{k}"] += 1
        kinds[codec.TABLES[t].tab] = dict(sorted(c.items()))
    return {
        "counts": {t: len(tables[t]) for t in tables},
        "statusTypeCounts": dict(sorted(Counter(f"{r['status']}/{r['request_type']}" for r in requests).items())),
        "monthTypeCounts": dict(sorted(Counter(f"{_month(r['target_date'])}/{r['request_type']}"
                                               for r in requests).items())),
        "minuteSums": {"approvedMinutes": total(requests, "approved_minutes"),
                       "actualMinutes": total(logs, "actual_minutes"), "breakMinutes": total(logs, "break_minutes"),
                       "netMinutes": total(logs, "net_minutes")},
        "minuteSumsByMonth": {k: by_month[k] for k in sorted(by_month)},
        "inProgress": {
            "submitted": sum(1 for r in requests if r["status"] == "submitted"),
            "approvedWithoutActualEnd": sum(1 for r in requests if r["status"] == "approved"
                                            and not (wl(r) and wl(r)["actual_end_at"])),
            "approvedWithoutApprovedBy2": sum(1 for r in requests if r["status"] == "approved"
                                              and not r["approved_by2"]),
        },
        "pdf": {"pdfFileIdCount": sum(1 for r in requests if r["pdf_file_id"]),
                "approvedNetWithoutPdf": sum(1 for r in requests if r["status"] == "approved" and net(r) > 0
                                             and not r["pdf_file_id"]),
                "pdfBeforeCompletion": sum(1 for r in requests if pdf_before_done(r))},
        "others": {
            "approvedBy2": sum(1 for r in requests if r["approved_by2"]),
            "rejected": sum(1 for r in requests if str(r["approved_by"]).endswith("(却下)")),
            "workerEmail": sum(1 for r in requests if r["worker_email"]),
            "approvedByUnknown": sum(1 for r in requests if r["approved_by"] == "unknown"),
            "orphanWorkLogs": sum(1 for w in logs if w["request_id"] not in req_ids),
            "requestsWithoutWorkLog": sum(1 for r in requests if not wl(r)),
            "canceledWithActuals": sum(1 for r in requests if r["status"] == "canceled" and wl(r)
                                       and wl(r)["actual_end_at"]),
            "approvedUnparseableTargetDate": sum(1 for r in requests if r["status"] == "approved"
                                                 and r["target_date"] is None),
            "breakMinutes1to9": sum(1 for w in logs if w["break_minutes"] is not None and 1 <= w["break_minutes"] <= 9),
            "netOverLimit": sum(1 for r in requests if (r["request_type"] == "overtime" and net(r) > 720)
                                or (r["request_type"] == "holiday" and net(r) > 960)),
        },
        "retroactiveByMonthType": dict(sorted(Counter(f"{_month(r['target_date'])}/{r['request_type']}"
                                                      for r in requests if retro(r)).items())),
        "updatedByCounts": dict(sorted(Counter(w["updated_by"] for w in logs
                                               if w["updated_by"] in ("formSubmit", "sebastian-agent", "unknown",
                                                                      "delete", "admin-manual",
                                                                      "break-recalc")).items())),
        "batchNameCounts": dict(sorted(Counter(b["batch_name"] for b in batches).items())),
        "cellKinds": kinds,
    }


# ===== DB への書き込み =====


def _metadata_tables() -> dict:
    from app import models  # noqa: F401 - すべての表を読み込む
    from app.db import Base

    return {t: Base.metadata.tables[t] for t in TRUNCATE_TABLES}


def sync_state_rows(tables: dict) -> list[dict]:
    """移した行の行の反映の状態（今シートにある値＝synced_values。反映済み）。"""
    out = []
    for t in A_TABLES:
        spec = codec.TABLES[t]
        for row in tables[t]:
            synced = codec.snapshot(spec, row)
            synced[codec.KINDS_KEY] = dict(row.get("cell_kinds") or {})
            pk = str(row["request_id"] if t == "requests" else row["id"])
            out.append({"table_name": t, "pk": pk, "desired_version": 1, "synced_version": 1, "synced_values": synced,
                        "state": "ok", "attempts": 0, "next_try_at": None, "last_error": "", "first_dirty_at": None,
                        "last_synced_at": None, "seq": 0})
    return out


def insert_tables(engine: Any, tables: dict, *, truncate: bool) -> None:
    """1 つのトランザクションで入れる。表に行があれば止める（--truncate なら先に空にする）。"""
    from sqlalchemy import delete, func, select, text

    targets = _metadata_tables()
    with engine.begin() as conn:
        counts = {t: conn.execute(select(func.count()).select_from(targets[t])).scalar_one() for t in TRUNCATE_TABLES}
        not_empty = {t: n for t, n in counts.items() if n}
        if not_empty and not truncate:
            raise MigrationStop("移行先の DB が空ではありません（--truncate を付けると、業務の表・jobs・idempotency_keys・"
                                "sync_runs・sessions を空にしてから入れます）: "
                                + ", ".join(f"{t}={n}" for t, n in not_empty.items()))
        if truncate:
            if conn.dialect.name == "postgresql":
                conn.execute(text("TRUNCATE TABLE " + ", ".join(TRUNCATE_TABLES) + " RESTART IDENTITY"))
            else:
                for t in TRUNCATE_TABLES:
                    conn.execute(delete(targets[t]))
        for t in ("requests", "work_logs", "batch_logs", "settings", "dept_approvers", "approver_map_rows"):
            if tables[t]:
                conn.execute(targets[t].insert(), tables[t])
        states = sync_state_rows(tables)
        if states:
            conn.execute(targets["sheet_sync_state"].insert(), states)
        advance_sequences(conn)


def advance_sequences(conn: Any) -> None:
    """番号を指定して入れた列の連番を最大値の先へ進める（PostgreSQL だけ。罠 9）。"""
    from sqlalchemy import text

    if conn.dialect.name != "postgresql":
        return
    for table in ("work_logs", "batch_logs", "jobs", "sync_runs", "audit_log"):
        seq = conn.execute(text("SELECT pg_get_serial_sequence(:t, 'id')"), {"t": table}).scalar()
        if seq:
            conn.execute(text(f"SELECT setval(CAST(:s AS regclass), COALESCE((SELECT max(id) FROM {table}), 0) + 1, "
                              "false)"), {"s": seq})
    conn.execute(text("SELECT setval('sheet_sync_state_seq', COALESCE((SELECT max(seq) FROM sheet_sync_state), 0) + 1, "
                      "false)"))


def read_back(engine: Any) -> dict:
    from sqlalchemy import select

    targets = _metadata_tables()
    with engine.connect() as conn:
        return {t: [dict(r._mapping) for r in conn.execute(select(targets[t]))]
                for t in ("requests", "work_logs", "batch_logs", "settings", "dept_approvers", "approver_map_rows")}


def verify_in_db(engine: Any, tables: dict, reads: dict) -> tuple[list, dict]:
    """入れた後に DB から読み直し、照合の数と、シートとの 1 セルずつの照合（差 0）を確かめる。"""
    from sqlalchemy.orm import Session

    problems = []
    db_tables = read_back(engine)
    expected, actual = summarize(tables), summarize(db_tables)
    for key in expected:
        if expected[key] != actual[key]:
            problems.append(f"{key}: シート {json.dumps(expected[key], ensure_ascii=False, default=str)} / "
                            f"DB {json.dumps(actual[key], ensure_ascii=False, default=str)}")
    tz = codec.ZoneInfo(reads["tz"] or "Asia/Tokyo")
    diffs = {}
    with Session(engine) as db:
        for t in A_TABLES:
            spec = codec.TABLES[t]
            u, f = reads["tabs"][spec.tab]
            result = reconcile.compare_tab(spec, u, f, reconcile.db_rows_for(db, t), tz)
            n = len(result["diffs"]) + len(result["idMismatch"]) + len(result["extraRows"])
            diffs[spec.tab] = n
            if n:
                problems.append(f"{spec.tab}: 移した直後の照合で差があります: {result['diffs'][:5]} "
                                f"{result['idMismatch'][:5]} {result['extraRows'][:5]}")
    return problems, diffs


# ===== 実行 =====


def run(reads: dict, engine: Optional[Any], *, dry_run: bool, truncate: bool = False,
        expected_tz: str | None = None) -> dict:
    if expected_tz and reads["tz"] != expected_tz:
        raise MigrationStop(f"DB スプレッドシートの時間帯（{reads['tz']}）が DB_SHEET_TIMEZONE（{expected_tz}）と違います")
    prepared = prepare(reads)
    tables = prepared["tables"]
    settings = {r["key"]: r["value"] for r in tables["settings"]}
    report = {
        "dryRun": dry_run,
        "sheetTimeZone": reads["tz"],
        "summary": summarize(tables),
        "warnings": prepared["warnings"],
        "droppedColumns": prepared["droppedColumns"],
        "missingColumns": {},
        "headerCounts": prepared["headerCounts"],
        "skippedEmptyRows": prepared["skippedEmptyRows"],
        "reportOnlyTabs": prepared["reportOnlyTabs"],
        "settingsOverlaps": prepared["settingsOverlaps"],
        "settingsForEnv": {k: settings.get(k, "") for k in REPORTED_SETTING_KEYS},
    }
    if not dry_run:
        if engine is None:
            raise MigrationStop("移行先の DB（--database-url か DATABASE_URL）が未設定です。")
        insert_tables(engine, tables, truncate=truncate)
        report["verifyProblems"], report["postMigrationDiffs"] = verify_in_db(engine, tables, reads)
    return report


def _json_default(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)


def _sheets_api(engine: Any) -> Any:
    from sqlalchemy.orm import Session

    from app.services.sheets_api import SheetsApi

    with Session(engine) as db:
        return SheetsApi.for_db(db)


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="本番 DB スプレッドシート → DB の移行（読むだけ）")
    parser.add_argument("--source", choices=["sheets"], required=True)
    parser.add_argument("--spreadsheet-id", default=PROD_SPREADSHEET_ID_FIXED)
    parser.add_argument("--database-url", default=os.environ.get("DATABASE_URL", ""))
    parser.add_argument("--dry-run", action="store_true", help="確かめと照合の報告だけ。DB に書かない")
    parser.add_argument("--truncate", action="store_true",
                        help="移行先の表と jobs・idempotency_keys・sync_runs・sessions を空にしてから入れる")
    parser.add_argument("--report-out", help="報告の JSON の書き出し先")
    args = parser.parse_args(argv)
    try:
        check_source_id(args.spreadsheet_id)  # トークンを取る前に止める
        if not args.database_url:
            raise MigrationStop("移行先の DB（--database-url か DATABASE_URL）が未設定です"
                                "（委任トークンを DB から読むので --dry-run でも要ります）。")
        from app.config import get_settings
        from app.db import make_engine

        expected_tz = get_settings().db_sheet_timezone
        engine = make_engine(args.database_url)
        reads = read_sheets(_sheets_api(engine), args.spreadsheet_id)
        report = run(reads, None if args.dry_run else engine, dry_run=args.dry_run, truncate=args.truncate,
                     expected_tz=expected_tz)
    except (MigrationStop, SheetsError) as error:
        print(f"止めました: {error}", file=sys.stderr)
        return 2
    except Exception as error:  # 委任が無い（DelegateNotAuthorized）など
        print(f"止めました: {type(error).__name__}: {error}", file=sys.stderr)
        return 2
    text = json.dumps(report, ensure_ascii=False, indent=1, default=_json_default)
    if args.report_out:
        Path(args.report_out).write_text(text, encoding="utf-8")
    print(text)
    return 1 if report.get("verifyProblems") else 0


if __name__ == "__main__":
    sys.exit(main())
