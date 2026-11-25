"""シートの取り込み（決定 #26。担当 D2）: 外から Requests へ直接足された行を DB へ入れる。

Sebastian（同じ VPS の /root/.hermes で動く AI。こちらからは変えない）は、本番 DB スプレッドシートの Requests へ
Sheets API の values.append（USER_ENTERED・INSERT_ROWS）で新しい申請の行を直接足す。GAS の時代は何もしなくても
その行が全画面・集計に出た。切り替えの後は VPS の DB が正なので、取り込まないと VPS の画面に出ない。

流れ（worker が 60 秒ごと `sheet_intake`、管理者の「運用」欄の ops_runSheetIntake、Sebastian の完了が DB に無い
申請を指したとき）:
  1. 関所: production は GAS が移転モード（gas_ping の migrated=true）のときだけ。切り替えの前（GAS が正）と
     戻した後は取り込まない。staging はステージングの写しを読むだけなので、いつでもよい。
  2. 鍵: 同時に 1 つだけ（プロセスの中は threading.Lock、PostgreSQL は pg_try_advisory_lock）。取れなければ何もしない。
  3. 時間帯を確かめ（row_sync.check_timezone）、requestId の列と行の数を読む（row_sync.load_tab。安い読み取り）。
  4. DB に無い requestId の行と、どの申請の行でもない requestId の空の行だけを、UNFORMATTED と FORMATTED で読む
     （移行の道具と同じ。罠 11）。
  5. 移行の道具と同じ決まり（sheet_codec.decode_row）で読み、確かめて、申請として入れる。sheet_row＝その行、
     sheet_sync_state は「反映済み」（synced_values＝今シートにある値）なので、行の反映は書き直さない。後で承認などを
     すれば、行の反映が requestId で行を確かめて、同じ行の変わった列だけを書く。
  6. requestId の空・重なり・決まりに無い値・読めない値の行は入れずに記録する（管理者の欄の lastRuns.sheet_intake の
     skipped）。それらの行へは書かない（DB に無い行は row_sync・照合のどちらも書かない）。

シートへは書かない（読むだけ）。WorkLogs の行は作らない: GAS は Sebastian が足した行に何もしなかった（ensureWorkLogRow_ は
api_submitRequest だけ）。実績を記録するときに store.update_worklog が行を作る（updateWorkLog_ と同じ）。
"""
from __future__ import annotations

import logging
import threading
import time
from collections import Counter
from collections.abc import Callable
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Any, Iterator

from sqlalchemy import delete, select, text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session, sessionmaker

from app import timeutil
from app.constants import REQUEST_STATUSES, REQUEST_TYPES
from app.jsutil import js_string, normalize_gas
from app.models import Request, SheetSyncState, SyncRun
from app.services import row_sync
from app.services import sheet_codec as codec
from app.services import write_gate
from app.services.sheets_api import SheetsApi, SheetsError, column_letter, quote_sheet

log = logging.getLogger(__name__)

RUN_KIND = "sheet_intake"  # 手で動かした回（ops・Sebastian）の sync_runs
SCHEDULE_KIND = "schedule:sheet_intake"
INGEST_KIND = "sheet_intake_ingest"  # 取り込んだ申請の記録（消さない）
EVERY_MINUTES = 1
ADVISORY_LOCK = 739_310_002  # pg_try_advisory_lock の鍵（worker の GAS のレーンは 739_310_001）
READ_CHUNK = 100  # 1 回の batchGet で読む行の数
MAX_LISTED = 100
KEEP_SCHEDULE_RUNS = timedelta(days=3)  # 60 秒ごとの枠の記録はこれより古いものを消す

_local_lock = threading.Lock()

# 入れるときに必ず要る列（読めない・空なら入れずに記録する）
MANDATORY_TEXT = (("dept", "dept"), ("worker_code", "workerCode"))


class IntakeBusy(RuntimeError):
    """ほかの取り込みが動いている。"""


def _settings(app_settings: Any):
    if app_settings is not None:
        return app_settings
    from app.config import get_settings

    return get_settings()


def intake_allowed(db: Session, settings: Any, now: datetime | None = None) -> tuple[bool, str]:
    """取り込んでよいか。production は GAS が移転モードのときだけ（write_gate と同じ判定）。staging はいつでも。"""
    if getattr(settings, "is_staging", False):
        return True, ""
    ok, reason = write_gate.check(db, settings, now)
    if ok:
        return True, ""
    return False, "本番の GAS が移転モードではない（切り替えの前か戻した後）ため、取り込みません: " + reason


@contextmanager
def intake_lock(session_factory: Callable[[], Session], wait_seconds: float = 0.0) -> Iterator[None]:
    """同時に 1 つだけ。取れなければ IntakeBusy。"""
    deadline = time.monotonic() + max(0.0, wait_seconds)
    while not _local_lock.acquire(blocking=False):
        if time.monotonic() >= deadline:
            raise IntakeBusy("ほかの取り込みが動いています")
        time.sleep(0.2)
    lock_db: Session | None = None
    try:
        lock_db = session_factory()
        if lock_db.get_bind().dialect.name == "postgresql":
            while not lock_db.execute(text("SELECT pg_try_advisory_lock(:k)"), {"k": ADVISORY_LOCK}).scalar():
                lock_db.rollback()
                if time.monotonic() >= deadline:
                    raise IntakeBusy("ほかのプロセスの取り込みが動いています")
                time.sleep(0.5)
            lock_db.commit()
            try:
                yield
            finally:
                lock_db.execute(text("SELECT pg_advisory_unlock(:k)"), {"k": ADVISORY_LOCK})
                lock_db.commit()
        else:
            lock_db.close()
            lock_db = None
            yield
    finally:
        if lock_db is not None:
            lock_db.close()
        _local_lock.release()


# ===== 読む =====


def _known(db: Session) -> tuple[set[str], set[int]]:
    """DB の申請の requestId（normalize_gas）と、申請が持つ行番号。"""
    ids, rows = set(), set()
    for rid, sheet_row in db.execute(select(Request.request_id, Request.sheet_row)):
        ids.add(normalize_gas(rid))
        if sheet_row:
            rows.add(int(sheet_row))
    return ids, rows


def find_candidates(ctx: row_sync.TabContext, known_ids: set[str], claimed_rows: set[int]) -> tuple[list[int], Counter]:
    """読む行（DB に無い requestId の行と、どの申請の行でもない requestId の空の行）と、requestId の数。"""
    counts: Counter = Counter()
    for values in ctx.id_rows.values():
        rid = js_string(values[0]) if values and values[0] is not None else ""
        if rid:
            counts[rid] += 1
    rows = []
    for r, values in sorted(ctx.id_rows.items()):
        rid = js_string(values[0]) if values and values[0] is not None else ""
        if rid:
            if rid not in known_ids:
                rows.append(r)
        elif r <= ctx.last_data_row and r not in claimed_rows:
            rows.append(r)
    # requestId の列より下にだけ値がある行（列が短い）
    height = max(ctx.id_rows) if ctx.id_rows else ctx.spec.header_row
    for r in range(height + 1, ctx.last_data_row + 1):
        if r not in claimed_rows:
            rows.append(r)
    return rows, counts


def read_rows(api: SheetsApi, sid: str, ctx: row_sync.TabContext, rows: list[int]) -> dict[int, tuple[list, list]]:
    """行ごとに UNFORMATTED と FORMATTED（罠 11）。"""
    q = quote_sheet(ctx.spec.tab)
    last_col = column_letter(max(max(ctx.col.values()), ctx.col_count - 1, 0))
    out: dict[int, tuple[list, list]] = {}
    for i in range(0, len(rows), READ_CHUNK):
        chunk = rows[i:i + READ_CHUNK]
        ranges = [f"{q}!A{r}:{last_col}{r}" for r in chunk]
        unformatted, formatted = api.read_both(sid, ranges)
        for k, r in enumerate(chunk):
            u = unformatted[k][0] if k < len(unformatted) and unformatted[k] else []
            f = formatted[k][0] if k < len(formatted) and formatted[k] else []
            out[r] = (u, f)
    return out


def check_record(record: dict, settings: Any) -> str:
    """入れてよいか（空なら入れる）。だめなら理由。"""
    rid = normalize_gas(record.get("request_id"))
    if not rid:
        return "requestId が空です"
    if record.get("request_type") not in REQUEST_TYPES:
        return f"requestType が決まりにありません: {js_string(record.get('request_type'))!r}"
    if record.get("status") not in REQUEST_STATUSES:
        return f"status が決まりにありません: {js_string(record.get('status'))!r}"
    if not isinstance(record.get("target_date"), date):
        return f"targetDate が日付として読めません: {js_string(record.get('target_date_text'))!r}"
    if record.get("submitted_at") is None:
        return "submittedAt が空です"
    for attr, header in MANDATORY_TEXT:
        if not normalize_gas(record.get(attr)):
            return f"{header} が空です"
    if len(js_string(record.get("request_id"))) > 64:
        return "requestId が長すぎます（64 字まで）"
    if getattr(settings, "is_staging", False) and (record.get("pdf_file_id") or record.get("pdf_folder_id")):
        # ステージングの GAS が本番の PDF に触れないよう（移行の道具の blank_pdf_ids_for_staging と同じ考え）
        return "ステージングでは pdfFileId・pdfFolderId のある行を取り込みません"
    return ""


def _skip(row: int, rid: str, reason: str) -> dict:
    return {"row": row, "requestId": rid, "reason": reason}


# ===== 入れる =====


def _insert(session_factory: Callable[[], Session], records: list[dict]) -> tuple[list[dict], list[dict]]:
    """申請と sheet_sync_state（反映済み）を入れる。入れた行と、入れなかった行（もうあった）。"""
    spec = codec.REQUESTS
    done, skipped = [], []
    with session_factory() as db:
        for rec in records:
            rid = rec["request_id"]
            if db.get(Request, rid) is not None:
                skipped.append(_skip(rec["sheet_row"], rid, "もう DB にあります（ほかの回が先に入れた）"))
                continue
            try:
                with db.begin_nested():
                    db.add(Request(**rec))
                    db.flush()
                    db.add(SheetSyncState(table_name="requests", pk=rid, desired_version=1, synced_version=1,
                                          synced_values=codec.synced_values_for(spec, rec), state="ok", attempts=0,
                                          next_try_at=None, last_error="", first_dirty_at=None,
                                          last_synced_at=timeutil.now_utc(), seq=0))
                    db.flush()
            except IntegrityError:
                skipped.append(_skip(rec["sheet_row"], rid, "もう DB にあります（同時に入った）"))
                continue
            done.append(rec)
        if done:
            db.add(SyncRun(kind=INGEST_KIND, started_at=timeutil.now_utc(), finished_at=timeutil.now_utc(), ok=True,
                           summary_json={"ingested": [{"requestId": r["request_id"], "row": r["sheet_row"]}
                                                      for r in done]}))
        db.commit()
    return done, skipped


def _purge_old_runs(session_factory: Callable[[], Session], now: datetime) -> None:
    with session_factory() as db:
        db.execute(delete(SyncRun).where(SyncRun.kind == SCHEDULE_KIND, SyncRun.started_at < now - KEEP_SCHEDULE_RUNS))
        db.commit()


# ===== 1 回分 =====


def run(session_factory: Callable[[], Session], *, api: SheetsApi | None = None, app_settings: Any = None,
        wait_seconds: float = 0.0) -> dict:
    """取り込みを 1 回。戻り値 {ok, ingested:[requestId], skipped:[{row, requestId, reason}], ...}。シートは読むだけ。"""
    settings = _settings(app_settings)
    started = timeutil.now_utc()
    with session_factory() as db:
        allowed, reason = intake_allowed(db, settings, started)
    if not allowed:
        return {"ok": False, "active": False, "stopped": reason, "ingested": [], "skipped": []}
    try:
        with intake_lock(session_factory, wait_seconds):
            result = _run_locked(session_factory, api, settings)
    except IntakeBusy as busy:
        return {"ok": True, "busy": True, "message": str(busy), "ingested": [], "skipped": []}
    result["durationMs"] = int((timeutil.now_utc() - started).total_seconds() * 1000)
    try:
        _purge_old_runs(session_factory, started)
    except Exception:  # noqa: BLE001 - 記録の掃除の失敗で取り込みを失敗にしない
        log.exception("sheet_intake: purge failed")
    return result


def _run_locked(session_factory: Callable[[], Session], api: SheetsApi | None, settings: Any) -> dict:
    spec = codec.REQUESTS
    sid = settings.db_spreadsheet_id
    tz = settings.db_sheet_tz
    try:
        if api is None:
            api = row_sync.api_factory(session_factory, settings)
        meta = row_sync.check_timezone(api, sid, settings)
        ctx = row_sync.load_tab(api, sid, spec, meta, tz, need_tail=True)
        with session_factory() as db:
            known_ids, claimed_rows = _known(db)
        rows, counts = find_candidates(ctx, known_ids, claimed_rows)
        base = {"ok": True, "active": True, "sheetRows": sum(counts.values()), "dbRows": len(known_ids),
                "rowCount": ctx.row_count, "candidates": len(rows)}
        if not rows:
            return {**base, "ingested": [], "skipped": []}
        cells = read_rows(api, sid, ctx, rows)
    except (row_sync.LaneHalt, SheetsError) as error:
        log.warning("sheet_intake: %s", error)
        return {"ok": False, "active": True, "stopped": str(error)[:1000], "ingested": [], "skipped": []}
    except Exception as error:  # noqa: BLE001 - 委任が無いなど。次の回にもう一度（書いていないので安全）
        log.exception("sheet_intake: read failed")
        return {"ok": False, "active": True, "stopped": f"{type(error).__name__}: {error}"[:1000], "ingested": [],
                "skipped": []}

    skipped: list[dict] = []
    records: list[dict] = []
    for r in rows:
        urow, frow = cells.get(r, ([], []))
        if all(codec.is_blank(v) for v in urow):
            continue  # 空の行（行の反映が取ったばかりの行より下の空き・消された行）
        record, errors = codec.decode_row(spec, ctx.col, urow, frow, tz, r)
        rid = js_string(record.get("request_id"))
        if errors:
            skipped.append(_skip(r, rid, "読めない値があります: " + " / ".join(errors)[:500]))
            continue
        key = normalize_gas(rid)
        if key and counts.get(key, 0) > 1:
            skipped.append(_skip(r, rid, f"同じ requestId の行が {counts[key]} 行あります"))
            continue
        reason = check_record(record, settings)
        if reason:
            skipped.append(_skip(r, rid, reason))
            continue
        records.append(record)
    done, already = _insert(session_factory, records) if records else ([], [])
    for rec in done:
        log.info("sheet_intake: ingested %s (row %s)", rec["request_id"], rec["sheet_row"])
    for s in skipped:
        log.warning("sheet_intake: skipped row %s (%s): %s", s["row"], s["requestId"], s["reason"])
    return {**base, "ingested": [r["request_id"] for r in done][:MAX_LISTED], "ingestedCount": len(done),
            "skipped": (skipped + already)[:MAX_LISTED], "skippedCount": len(skipped) + len(already)}


# ===== 呼び口 =====


def _record_manual(session_factory: Callable[[], Session], started: datetime, result: dict, source: str) -> None:
    with session_factory() as db:
        db.add(SyncRun(kind=RUN_KIND, started_at=started, finished_at=timeutil.now_utc(),
                       ok=result.get("ok") is not False, summary_json={**result, "source": source}))
        db.commit()


def run_now(db: Session, app_settings: Any = None) -> dict:
    """管理者の「運用」欄（ops_runSheetIntake）。結果を sync_runs（kind='sheet_intake'）にも残す。"""
    factory = sessionmaker(bind=db.get_bind(), expire_on_commit=False)
    started = timeutil.now_utc()
    result = run(factory, app_settings=app_settings, wait_seconds=10.0)
    _record_manual(factory, started, result, "ops")
    return result


def ensure_request(db: Session, request_id: str, app_settings: Any = None) -> bool:
    """Sebastian の完了が DB に無い申請を指したとき、その場で 1 回取り込む（GAS ならシートの行がすぐ見えたため）。

    取り込めた（DB にある）なら True。失敗しても例外にしない（呼ぶ側はいつもどおり「見つかりません」を返す）。
    """
    rid = normalize_gas(request_id)
    if not rid or db.get(Request, rid) is not None:
        return bool(rid)
    try:
        factory = sessionmaker(bind=db.get_bind(), expire_on_commit=False)
        started = timeutil.now_utc()
        result = run(factory, app_settings=app_settings, wait_seconds=15.0)
        if result.get("ingested"):
            _record_manual(factory, started, result, "sebastian")
    except Exception:  # noqa: BLE001 - 詳細はログだけ
        log.exception("sheet_intake: on-demand intake for %s failed", rid)
        return False
    db.expire_all()
    return db.get(Request, rid) is not None


def _scheduled(session_factory: Callable[[], Session]) -> dict:
    return run(session_factory)


def register_worker() -> None:
    from app import worker

    worker.register_schedule("sheet_intake", task=_scheduled, every_minutes=EVERY_MINUTES)


def status(db: Session) -> dict:
    """管理者の欄に出す形（最後の回）。"""
    rows = db.scalars(select(SyncRun).where(SyncRun.kind.in_((RUN_KIND, SCHEDULE_KIND)),
                                            SyncRun.finished_at.is_not(None))
                      .order_by(SyncRun.started_at.desc(), SyncRun.id.desc()).limit(10)).all()
    # 鍵が取れずに何もしなかった回（busy）は飛ばし、最後に実際に読んだ回を出す
    row = next((r for r in rows if not (r.summary_json or {}).get("busy")), rows[0] if rows else None)
    if row is None:
        return {"lastRunAt": "", "ok": None, "active": None, "skipped": [], "stopped": ""}
    s = row.summary_json or {}
    return {"lastRunAt": timeutil.to_jst_text(row.finished_at or row.started_at, "yyyy-MM-dd HH:mm:ss"),
            "ok": bool(row.ok), "active": s.get("active"), "skipped": list(s.get("skipped") or []),
            "stopped": s.get("stopped") or "", "ingested": list(s.get("ingested") or [])}


__all__ = ["run", "run_now", "ensure_request", "register_worker", "status", "intake_allowed", "find_candidates"]
