"""朝バッチ（実装仕様 7-5・gap-3 G3-H1。担当 M）。

GAS の morningBatchNoMail_（mail.gs:1454-1468）を写す:
  batchGeneratePdfs_(昨日)（pdfExport.gs:570-600）→ logBatchResult_('morningBatchNoMail', 昨日, 結果)
  → appendToAccumulationSS_(昨日)

毎朝 6:50 JST の予定 `morning_batch` が仕事 `morning_batch`（lane='gas'、priority=5）を 1 本積み、GAS のレーンが
次の順に進める（outbox に積んで先へ進まない。進み具合は jobs.result_json に持つ）:

1. 対象日の approved の申請を DB から拾う（listApprovedRequestsByDate_ と同じ条件・(部署+氏名) の並び）。
2. 1 件ずつ: pdfFileId があれば skip+1。無ければ internal.batchPdf を呼び、created／already なら ok+1、
   error なら fail+1 と `<requestId> (<dept>/<workerName>): <message>`。通信の失敗が 3 回続いたら（30 秒おき）
   fail、message は `GAS呼び出し失敗: <内容>`、送り直しの仕事（internal.batchPdf）を積む。
   1 件ごとにレーンを放す（RetryLater）ので、priority の高い仕事（Sebastian=10）が先に流れる。
   6 件以上なら 1 件ごとに 500ms 以上空ける。
3. batch_logs に 1 行（実行日時=今、バッチ名、対象日、成功・スキップ・失敗、エラー詳細は \\n でつなぐ）、mark_dirty。
4. internal.ledgerRewriteDay(対象日)（台帳が開けない・見出しが違うときは failed にして管理者の欄へ）。
5. 7:05 に、Requests と BatchLogs の行の反映が済んだかを確かめる（予定 morning_batch_check）。

管理者の ops_batchPdfForDate(ymd) は enqueue_manual_batch（1〜3 だけ。バッチ名 manualBatchPdfs。台帳は書かない）。
"""
from __future__ import annotations

import logging
import re
import unicodedata
import uuid
from collections.abc import Callable
from datetime import date, datetime, timedelta

from sqlalchemy import select
from sqlalchemy.orm import Session

from app import jobs, models, sheet_queue, timeutil
from app.jobs import JobContext, PermanentJobError, RetryLater
from app.jsutil import normalize_gas
from app.services import gas_jobs

log = logging.getLogger(__name__)

KIND = "morning_batch"
PRIORITY = 5  # 通常（0）より先、Sebastian（10）より後
BATCH_MORNING = "morningBatchNoMail"
BATCH_MANUAL = "manualBatchPdfs"
RUN_KIND = {BATCH_MORNING: "morning_batch", BATCH_MANUAL: "manual_batch_pdfs"}  # sync_runs.kind
CHECK_KIND = "morning_batch_check"
TRANSPORT_TRIES = 3
TRANSPORT_WAIT_SECONDS = 30.0
ITEM_SPACING_SECONDS = 0.5
ITEM_SPACING_MIN_ITEMS = 6  # GAS: if (items.length > 5) Utilities.sleep(500)
_YMD = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# 今の朝バッチで「台帳を開けない」は直さない限り続く（7-5 の 4: failed にして管理者へ）
LEDGER_PERMANENT_ERRORS = ("ledger_unavailable", "ledger_header_mismatch", "bad_request", "unauthorized")


def ja_sort_key(text: str) -> tuple[int, ...]:
    """localeCompare(…, 'ja') の近似（ICU が無いため）。全角半角をそろえ、カタカナをひらがなに寄せ、大文字小文字を無視。
    並びは処理とエラー詳細の順にだけ効く。"""
    s = unicodedata.normalize("NFKC", text).casefold()
    return tuple((ord(ch) - 0x60) if 0x30A1 <= ord(ch) <= 0x30F6 else ord(ch) for ch in s)


def _valid_ymd(ymd: str) -> date:
    if not _YMD.match(str(ymd or "")):
        raise ValueError(f"日付は yyyy-MM-dd で指定してください: {ymd}")
    return datetime.strptime(ymd, "%Y-%m-%d").date()


def list_items(db: Session, ymd: str) -> list[dict]:
    """対象日の approved の申請（listApprovedRequestsByDate_ と同じ条件と並び）。"""
    target = _valid_ymd(ymd)
    R = models.Request
    rows = db.scalars(select(R).where(R.status == "approved")).all()
    items = []
    for r in rows:
        if r.target_date is not None:
            if r.target_date != target:
                continue
        elif gas_jobs.request_target_ymd(r) != ymd:
            continue
        dept, name = normalize_gas(r.dept), normalize_gas(r.worker_name)
        items.append({"requestId": r.request_id, "dept": dept, "workerName": name})
    items.sort(key=lambda it: ja_sort_key(it["dept"] + it["workerName"]))
    return items


# ===== 積む =====


def _existing_batch_job(db: Session, ymd: str, batch_name: str, states: tuple[str, ...]) -> int | None:
    J = models.Job
    for row in db.scalars(select(J).where(J.kind == KIND, J.state.in_(states)).order_by(J.id)):
        args = gas_jobs._loads(row.args_json) or {}
        if args.get("targetDate") == ymd and args.get("batchName") == batch_name:
            return int(row.id)
    return None


def enqueue_batch(db: Session, ymd: str, batch_name: str, *, ledger: bool) -> int:
    _valid_ymd(ymd)
    existing = _existing_batch_job(db, ymd, batch_name, ("pending", "running"))
    if existing is not None:
        return existing
    args = {"targetDate": ymd, "batchName": batch_name, "ledger": bool(ledger), "runId": "mb-" + uuid.uuid4().hex[:20]}
    return jobs.enqueue(db, KIND, args, lane="gas", priority=PRIORITY)


def enqueue_manual_batch(db: Session, ymd: str) -> int:
    """管理者の「指定日の PDF 一括」（1〜3。BatchLogs に manualBatchPdfs の 1 行。台帳は書かない）。commit しない。"""
    return enqueue_batch(db, ymd, BATCH_MANUAL, ledger=False)


def run_morning_batch(session_factory: Callable[[], Session]) -> dict:
    """毎朝 6:50 JST の予定。対象日（JST の昨日）の朝バッチの仕事を積む（同じ日のものが既にあれば積まない）。"""
    ymd = (timeutil.today_jst() - timedelta(days=1)).strftime("%Y-%m-%d")
    with session_factory() as db:
        existing = _existing_batch_job(db, ymd, BATCH_MORNING, ("pending", "running", "done"))
        if existing is not None:
            return {"ok": True, "jobId": existing, "targetDate": ymd, "already": True}
        job_id = enqueue_batch(db, ymd, BATCH_MORNING, ledger=True)
        db.commit()
    return {"ok": True, "jobId": job_id, "targetDate": ymd}


# ===== 仕事の中身 =====


def _save(ctx: JobContext, state: dict) -> None:
    row = ctx.db.get(models.Job, ctx.job_id)
    row.result_json = gas_jobs._dumps(state)
    ctx.commit()


def _load_state(ctx: JobContext) -> dict:
    row = ctx.db.get(models.Job, ctx.job_id)
    state = gas_jobs._loads(row.result_json if row is not None else "")
    if isinstance(state, dict) and state.get("phase"):
        return state
    args = ctx.args
    return {"phase": "init", "targetDate": args.get("targetDate"), "batchName": args.get("batchName", BATCH_MORNING),
            "ledger": bool(args.get("ledger")), "runId": args.get("runId") or ("mb-" + str(ctx.job_id))}


def _error_line(item: dict, message: str) -> str:
    return f"{item['requestId']} ({item['dept']}/{item['workerName']}): {message}"


def _process_item(ctx: JobContext, state: dict, item: dict) -> None:
    """1 件を処理して state の数を進める（RetryLater で待つときは state を保存してから上げる）。"""
    from app.services.gas_client import GasTransportError

    db = ctx.db
    req = db.get(models.Request, item["requestId"])
    if req is not None and req.pdf_file_id:
        state["skip"] += 1
        return
    if req is None:
        state["fail"] += 1
        state["errors"].append(_error_line(item, "requestIdが見つかりません。"))
        return
    pending = gas_jobs.unsynced_rows(db, req.request_id, with_work_logs=False)
    if pending:
        _save(ctx, state)
        raise RetryLater(gas_jobs.SYNC_WAIT_SECONDS, "row_not_synced（行の反映待ち）: " + ", ".join(pending))
    snapshot = gas_jobs.pdf_snapshot(req)
    job_id = f"{state['runId']}-i{state['index']}"
    failure = ""
    try:
        resp = gas_jobs.call_gas(ctx.app_settings, "internal.batchPdf", job_id=job_id, request_id=req.request_id,
                                 expect={"status": req.status}, args={"targetDate": state["targetDate"]})
    except GasTransportError as exc:
        resp = None
        failure = str(exc)
    if resp is not None:
        pdf = resp.get("pdf") if isinstance(resp.get("pdf"), dict) else {}
        status = pdf.get("status")
        if status in ("created", "already"):
            merged = dict(pdf)
            for key in ("pdfFileId", "pdfFolderId", "pdfGeneratedAt"):
                if not merged.get(key) and resp.get(key):
                    merged[key] = resp[key]
            gas_jobs.record_pdf_result(db, req.request_id, merged, expected_current=snapshot)
            state["ok"] += 1
            state["tries"] = 0
            return
        if status == "error":
            state["fail"] += 1
            state["errors"].append(_error_line(item, str(pdf.get("error") or resp.get("message") or "")))
            state["tries"] = 0
            return
        if resp.get("ok") is True:
            # PDF を作らなかった（承認済みでなくなった など）。今の GAS では generatePdf* の例外＝失敗
            state["fail"] += 1
            state["errors"].append(_error_line(item, str(pdf.get("reason") or "skipped")))
            state["tries"] = 0
            return
        if resp.get("retryable") is not True:
            state["fail"] += 1
            state["errors"].append(_error_line(item, gas_jobs._error_text(resp)))
            state["tries"] = 0
            return
        failure = gas_jobs._error_text(resp)  # busy・internal・row_not_synced は通信の失敗と同じに数える
    state["tries"] = int(state.get("tries") or 0) + 1
    if state["tries"] < TRANSPORT_TRIES:
        state["lastFailure"] = failure
        _save(ctx, state)
        raise RetryLater(TRANSPORT_WAIT_SECONDS, f"morning_batch: GAS呼び出し失敗（{state['tries']} 回目）: {failure}")
    state["fail"] += 1
    state["errors"].append(_error_line(item, f"GAS呼び出し失敗: {failure}"))
    state["tries"] = 0
    resend = gas_jobs.enqueue_batch_pdf(db, req.request_id, target_date=state["targetDate"],
                                        rewrite_ledger=bool(state.get("ledger")))
    state.setdefault("resent", []).append({"requestId": req.request_id, "jobId": resend})


def _write_batch_log(ctx: JobContext, state: dict) -> int:
    db = ctx.db
    row = models.BatchLog(executed_at=timeutil.now_utc(), batch_name=state["batchName"],
                          target_date=_valid_ymd(state["targetDate"]), target_date_text=state["targetDate"],
                          success=int(state["ok"]), skip=int(state["skip"]), fail=int(state["fail"]),
                          error_detail="\n".join(state["errors"]))
    db.add(row)
    db.flush()
    sheet_queue.mark_dirty(db, "batch_logs", str(row.id))
    return int(row.id)


def _finish_run(ctx: JobContext, state: dict, ok: bool, error: str = "") -> None:
    now = timeutil.now_utc()
    summary = {k: state.get(k) for k in ("targetDate", "batchName", "ok", "skip", "fail", "errors", "batchLogId",
                                         "ledgerResult", "resent")}
    summary["counts"] = {"ok": state.get("ok"), "skip": state.get("skip"), "fail": state.get("fail")}
    summary["jobId"] = ctx.job_id
    if error:
        summary["error"] = error
    ctx.db.add(models.SyncRun(kind=RUN_KIND.get(state.get("batchName"), "morning_batch"), started_at=now,
                              finished_at=now, ok=ok, summary_json=summary))


def handle_batch(ctx: JobContext) -> dict:
    state = _load_state(ctx)
    if state["phase"] == "init":
        items = list_items(ctx.db, state["targetDate"])
        state.update(phase="pdfs", items=items, index=0, ok=0, skip=0, fail=0, errors=[], tries=0)
        _save(ctx, state)
    if state["phase"] == "pdfs":
        items = state["items"]
        if state["index"] < len(items):
            gas_jobs.require_gas_ping(ctx.db, ctx.now)
            _process_item(ctx, state, items[state["index"]])
            state["index"] += 1
            state.pop("lastFailure", None)
            _save(ctx, state)
            if state["index"] < len(items):
                # 1 件ごとにレーンを放す（priority の高い仕事を先に流す）。6 件以上なら 500ms 空ける
                wait = ITEM_SPACING_SECONDS if len(items) >= ITEM_SPACING_MIN_ITEMS else 0.0
                raise RetryLater(wait, f"morning_batch: {state['index']}/{len(items)}")
        state["phase"] = "log"
    if state["phase"] == "log":
        state["batchLogId"] = _write_batch_log(ctx, state)
        state["phase"] = "ledger" if state.get("ledger") else "done"
        _save(ctx, state)
    if state["phase"] == "ledger":
        gas_jobs.require_gas_ping(ctx.db, ctx.now)
        resp = gas_jobs.call_gas(ctx.app_settings, "internal.ledgerRewriteDay", job_id=f"{state['runId']}-ledger",
                                 args={"targetDate": state["targetDate"]})
        state["ledgerResult"] = resp.get("ledger") if isinstance(resp.get("ledger"), dict) else {"ok": resp.get("ok")}
        if resp.get("ok") is not True:
            text = gas_jobs._error_text(resp)
            state["ledgerError"] = text
            _save(ctx, state)
            if resp.get("retryable") is True and str(resp.get("error")) not in LEDGER_PERMANENT_ERRORS:
                raise gas_jobs.GasJobError(text)
            raise PermanentJobError(text)
        state.pop("ledgerError", None)
        state["phase"] = "done"
        _save(ctx, state)
    _finish_run(ctx, state, ok=True)
    return {k: v for k, v in state.items() if k != "items"} | {"itemCount": len(state.get("items") or [])}


def on_exhausted(ctx: JobContext, message: str) -> None:
    row = ctx.db.get(models.Job, ctx.job_id)
    state = gas_jobs._loads(row.result_json if row is not None else "") or {}
    if not isinstance(state, dict):
        state = {}
    state.setdefault("targetDate", ctx.args.get("targetDate"))
    state.setdefault("batchName", ctx.args.get("batchName"))
    _finish_run(ctx, state, ok=False, error=message[:1000])
    gas_jobs.on_exhausted(ctx, message)


# ===== 7:05 の確かめ（7-5 の 5） =====


def check_morning_sync(session_factory: Callable[[], Session]) -> dict:
    """今朝の朝バッチの BatchLogs の行と、対象日の Requests の行の反映が済んだか。済んでいなければ ok=False。"""
    ymd = (timeutil.today_jst() - timedelta(days=1)).strftime("%Y-%m-%d")
    problems: list[str] = []
    with session_factory() as db:
        job_id = _existing_batch_job(db, ymd, BATCH_MORNING, ("pending", "running", "done", "failed", "skipped"))
        if job_id is None:
            problems.append("朝バッチの仕事がありません")
        else:
            job = db.get(models.Job, job_id)
            state = gas_jobs._loads(job.result_json) or {}
            if job.state != "done":
                problems.append(f"朝バッチが終わっていません（{job.state}）")
            log_id = state.get("batchLogId") if isinstance(state, dict) else None
            if log_id and not sheet_queue.is_synced(db, "batch_logs", str(log_id),
                                                    sheet_queue.current_version(db, "batch_logs", str(log_id))):
                problems.append("BatchLogs の行がまだシートに反映されていません")
        for item in list_items(db, ymd):
            if gas_jobs.unsynced_rows(db, item["requestId"], with_work_logs=False):
                problems.append(f"Requests の行がまだシートに反映されていません: {item['requestId']}")
        now = timeutil.now_utc()
        db.add(models.SyncRun(kind=CHECK_KIND, started_at=now, finished_at=now, ok=not problems,
                              summary_json={"targetDate": ymd, "problems": problems}))
        db.commit()
    if problems:
        log.error("朝バッチの 7:05 の確かめ: %s", problems)
    return {"ok": not problems, "targetDate": ymd, "problems": problems}


def register_worker() -> None:
    from app import worker

    jobs.register_kind(KIND, handle_batch, lane="gas", on_exhausted=on_exhausted)
    worker.register_schedule("morning_batch", task=run_morning_batch, at=(6, 50))
    worker.register_schedule(CHECK_KIND, task=check_morning_sync, at=(7, 5))

