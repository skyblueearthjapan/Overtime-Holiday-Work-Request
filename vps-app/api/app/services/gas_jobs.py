"""GAS の後処理の仕事（決定 B・実装仕様 7-3・7-4。担当 M）。

W の業務関数は、業務の書き込みと同じトランザクションで enqueue_* を呼ぶ（commit は呼び出し側）。worker の GAS の
レーン（全体で同時に 1 本。app/jobs.py）が、積んだ順に 1 件ずつ GAS の内部用の入口（services/gas_client）を呼ぶ。

仕事の種類（jobs.kind）と、実際に呼ぶ GAS の action（G の src/Internal.gs に合わせた）:

| kind（7-4）                | GAS の action          | 中身（今の関数の順序・条件。G が写した） |
|----------------------------|------------------------|------------------------------------------|
| internal.afterApprove      | internal.afterApprove  | actualEndAt があるときだけ PDF → 成功したときだけ台帳 |
| internal.afterComplete     | internal.afterComplete | approved なら PDF（screen/admin は wlOverride、sebastian は XLOOKUP）→ いつも台帳 |
| internal.regenerate        | internal.regenerate    | 前の PDF をゴミ箱へ → PDF → secondary/edit は台帳、admin は台帳なし |
| internal.ledgerUpsert      | vps.upsertLedger       | 台帳だけ（PDF の無い二次承認の単票）。G に internal.ledgerUpsert の別名が無い |
| internal.afterDelete       | internal.afterDelete   | 前の PDF をゴミ箱へ → 台帳の行を消す |
| internal.batchPdf          | internal.batchPdf      | 朝バッチで通信に失敗した件の送り直し（台帳は書かない） |
| internal.ledgerRewriteDay  | internal.ledgerRewriteDay | その日の台帳の書き直し（朝バッチの送り直しが後で成功したとき） |
| internal.ledgerRebuild     | internal.ledgerRebuild | 台帳の作り直し（管理者）。G に未実装（bad_request で failed になる） |

動かす前の確かめ（7-3）:
  1. requires（積んだ時点の版）は jobs.JobRunner が待つ。ここでは「今の版」まで行の反映が済んだかも確かめ、
     まだなら RetryLater（試行回数は増やさない）。24 時間を過ぎても済まなければ failed。
  2. status が canceled なら afterApprove・afterComplete・regenerate・ledgerUpsert・batchPdf は skipped。
  3. expect は DB の今の値（シートは反映済み）。
  4. 最後の gas_ping が 10 分以内に合格していなければ動かさない（1 分後に）。
結果の記録: pdf.status が created／already なら record_pdf_result（6-5）。trash が済めば old_pdf_file_id_pending を空に。
再試行: retryable:true（busy・pdf_failed・row_not_synced・internal・ledger_unavailable）と通信の失敗・時間切れは
jobs の間隔（1・2・4・8・16・30 分、その後 30 分ごと、24 時間まで）。retryable:false は failed。
冪等: GAS へ送る jobId は積んだときに決めて guard_json.gasJobId に持つ（送り直しても同じ jobId。GAS が同じ返事を返す）。
"""
from __future__ import annotations

import json
import logging
import threading
import time
import uuid
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app import jobs, models, safety, sheet_queue, timeutil
from app.jobs import JobContext, PermanentJobError, RetryLater, SkipJob
from app.jsutil import normalize_gas

log = logging.getLogger(__name__)

# ===== 定数 =====

KIND_AFTER_APPROVE = "internal.afterApprove"
KIND_AFTER_COMPLETE = "internal.afterComplete"
KIND_REGENERATE = "internal.regenerate"
KIND_LEDGER_UPSERT = "internal.ledgerUpsert"
KIND_AFTER_DELETE = "internal.afterDelete"
KIND_BATCH_PDF = "internal.batchPdf"
KIND_LEDGER_REWRITE_DAY = "internal.ledgerRewriteDay"
KIND_LEDGER_REBUILD = "internal.ledgerRebuild"

COMPLETE_SOURCES = ("screen", "admin", "sebastian")
REGENERATE_REASONS = ("secondary", "edit", "admin")

GAS_PING_KIND = "gas_ping"  # sync_runs.kind
GAS_PING_MAX_AGE = timedelta(minutes=10)
GAS_PING_WAIT_SECONDS = 60.0
SYNC_WAIT_SECONDS = 5.0
SYNC_WAIT_DEADLINE = timedelta(hours=24)
BATCH_SPACING_SECONDS = 0.5  # 一括の操作で 6 件以上積んだまとまりの呼び出しの間（GAS の Utilities.sleep(500)）
BATCH_SPACING_MIN_JOBS = 6
EXHAUSTED_KIND = "gas_job_failed"  # sync_runs.kind（管理者の欄に出す）

SEBASTIAN_TIMEOUT_MESSAGE = "PDF作成の結果待ちが時間切れになりました（作成は続いています）"

# PDF の 3 列（シートの見出し → requests の列）。synced_values もこの見出しで持つ（6-5）
PDF_COLUMNS = (("pdfFileId", "pdf_file_id"), ("pdfFolderId", "pdf_folder_id"), ("pdfGeneratedAt", "pdf_generated_at"))


class GasJobError(RuntimeError):
    """GAS が ok:false（retryable:true）を返した。jobs の間隔で送り直す。"""


# ===== 差し替えの口（試験で偽物を入れる） =====


def _default_client_factory(settings: Any) -> Any:
    from app.services.gas_client import GasClient

    return GasClient.from_settings(settings)


_client_factory: Callable[[Any], Any] = _default_client_factory
_sleep: Callable[[float], None] = time.sleep
_monotonic: Callable[[], float] = time.monotonic


def set_client_factory(factory: Callable[[Any], Any] | None) -> Callable[[Any], Any]:
    """GasClient を作る関数を差し替える（試験）。None で元に戻す。前の関数を返す。"""
    global _client_factory
    previous = _client_factory
    _client_factory = factory or _default_client_factory
    return previous


def _settings(ctx_settings: Any = None) -> Any:
    if ctx_settings is not None:
        return ctx_settings
    from app.config import get_settings

    return get_settings()


def get_client(settings: Any = None) -> Any:
    return _client_factory(_settings(settings))


def new_gas_job_id() -> str:
    """GAS へ送る jobId（[A-Za-z0-9._:-]{1,100}）。cutover で jobs を空にしても重ならないよう UUID。"""
    return "ot-" + uuid.uuid4().hex


# ===== 小道具 =====


def _dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, default=str, separators=(",", ":"))


def _loads(text: str | None) -> Any:
    try:
        return json.loads(text or "")
    except ValueError:
        return None


def parse_jst_text(text: Any) -> datetime | None:
    """'yyyy-MM-dd HH:mm:ss'（JST）か 'yyyy-MM-dd HH:mm'・'yyyy-MM-dd' を UTC の datetime に。読めなければ None。"""
    s = str(text or "").strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d"):
        try:
            naive = datetime.strptime(s, fmt)
        except ValueError:
            continue
        return naive.replace(tzinfo=timeutil.JST).astimezone(timeutil.UTC)
    return None


def request_target_ymd(req: models.Request) -> str:
    """申請の対象日（'yyyy-MM-dd'）。日付の列が空なら元の文字から読む。"""
    if req.target_date is not None:
        return req.target_date.strftime("%Y-%m-%d")
    text = str(req.target_date_text or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def _work_logs_for(db: Session, request_id: str) -> list[models.WorkLog]:
    W = models.WorkLog
    return list(db.scalars(select(W).where(W.request_id == request_id).order_by(W.sheet_row.is_(None), W.sheet_row, W.id)))


def _sync_targets(db: Session, request_id: str, with_work_logs: bool) -> list[tuple[str, str]]:
    rows = [("requests", request_id)]
    if with_work_logs:
        rows += [("work_logs", str(w.id)) for w in _work_logs_for(db, request_id)]
    return rows


def requires_for_request(db: Session, request_id: str, *, with_work_logs: bool) -> list[dict]:
    """requires を作る（10-2 の形）。sheet_sync_state の無い行（一度も印を付けていない＝シートと同じ）は待たない。"""
    out = []
    for table, pk in _sync_targets(db, request_id, with_work_logs):
        if db.get(models.SheetSyncState, (table, pk)) is None:
            continue
        out.append({"table": table, "pk": pk, "version": sheet_queue.current_version(db, table, pk)})
    return out


def unsynced_rows(db: Session, request_id: str, *, with_work_logs: bool = True) -> list[str]:
    """今の版まで行の反映が済んでいない行（'requests:<id>' の形）。空なら済んでいる。"""
    pending = []
    for table, pk in _sync_targets(db, request_id, with_work_logs):
        state = db.get(models.SheetSyncState, (table, pk))
        if state is None:
            continue
        if int(state.synced_version) < sheet_queue.current_version(db, table, pk):
            pending.append(f"{table}:{pk}")
    return pending


def last_gas_ping(db: Session) -> models.SyncRun | None:
    R = models.SyncRun
    return db.scalars(select(R).where(R.kind == GAS_PING_KIND).order_by(R.started_at.desc(), R.id.desc()).limit(1)).first()


def gas_ping_ok(db: Session, now: datetime | None = None) -> tuple[bool, str]:
    """最後の gas_ping が 10 分以内に合格しているか（7-3 の 4）。"""
    now = now or timeutil.now_utc()
    row = last_gas_ping(db)
    if row is None:
        return False, "gas_ping がまだありません"
    if not row.ok:
        problems = (row.summary_json or {}).get("problems") or []
        return False, "gas_ping が不合格です: " + " / ".join(str(p) for p in problems)[:500]
    at = row.finished_at or row.started_at
    if now - at > GAS_PING_MAX_AGE:
        return False, "gas_ping が 10 分以上ありません"
    return True, ""


def require_gas_ping(db: Session, now: datetime | None = None) -> None:
    ok, reason = gas_ping_ok(db, now)
    if not ok:
        raise RetryLater(GAS_PING_WAIT_SECONDS, reason)


def require_synced(db: Session, request_id: str, *, created_at: datetime | None, now: datetime,
                   with_work_logs: bool = True) -> None:
    """申請の行（と WorkLogs の行）が今の版までシートに反映済みでなければ待つ（7-3 の 1）。"""
    pending = unsynced_rows(db, request_id, with_work_logs=with_work_logs)
    if not pending:
        return
    reason = "row_not_synced（行の反映待ち）: " + ", ".join(pending)
    if created_at is not None and now - created_at > SYNC_WAIT_DEADLINE:
        raise PermanentJobError(reason + "（24 時間を過ぎました）")
    raise RetryLater(SYNC_WAIT_SECONDS, reason)


# ===== PDF の結果の記録（6-5） =====


def _pdf_generated_at(pdf: Mapping[str, Any]) -> datetime | None:
    epoch = pdf.get("pdfGeneratedAtEpochMs")
    if isinstance(epoch, (int, float)) and not isinstance(epoch, bool) and epoch > 0:
        return datetime.fromtimestamp(epoch / 1000.0, tz=timeutil.UTC)
    return parse_jst_text(pdf.get("pdfGeneratedAt"))


def _synced_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.astimezone(timeutil.UTC).isoformat()
    return value if value is not None else ""


def pdf_snapshot(req: models.Request | None) -> tuple | None:
    if req is None:
        return None
    return (req.pdf_file_id or "", req.pdf_folder_id or "", req.pdf_generated_at)


def _lock_request(db: Session, request_id: str) -> models.Request | None:
    R = models.Request
    return db.scalars(select(R).where(R.request_id == request_id).with_for_update()).first()


def _note_sheet_has(db: Session, request_id: str, values: Mapping[str, Any]) -> None:
    """シートにこの値が既にあると記録する（synced_values の 3 列だけ）。"""
    state = db.get(models.SheetSyncState, ("requests", request_id))
    if state is None:
        return
    merged = dict(state.synced_values or {})
    merged.update({k: _synced_value(v) for k, v in values.items()})
    state.synced_values = merged


def record_pdf_result(db: Session, request_id: str, pdf: Mapping[str, Any], *,
                      expected_current: tuple | None = None) -> bool:
    """GAS がシートに書いた PDF の 3 列を DB に入れる（同じトランザクション。commit しない）。入れたら True。

    - sheet_queue.mark_dirty で版を進め、行が反映済みだったときは synced_version も同じだけ進め、synced_values の
      3 列を同じ値にする（シートに既にある値なので、行の反映が書き直さない）。
    - expected_current（GAS を呼んだ時点の DB の 3 列）を渡すと、その後の操作（編集・二次承認・削除など）が 3 列を
      変えていたら DB を上書きしない（新しい値を守る）。そのときは synced_values だけ直して行の反映に DB の値を
      書き戻させ、DB が空なら作られた PDF の ID を old_pdf_file_id_pending に残す（後の仕事がゴミ箱へ移す）。
    """
    file_id = normalize_gas(pdf.get("pdfFileId"))
    if not file_id:
        return False
    folder_id = normalize_gas(pdf.get("pdfFolderId") or pdf.get("folderId"))
    generated_at = _pdf_generated_at(pdf)
    req = _lock_request(db, request_id)
    if req is None:
        return False
    sheet_values = {"pdfFileId": file_id, "pdfFolderId": folder_id, "pdfGeneratedAt": generated_at}
    current = pdf_snapshot(req)
    if expected_current is not None and current != tuple(expected_current) and current[0] != file_id:
        log.info("record_pdf_result: %s は後の操作で PDF の列が変わったので上書きしません", request_id)
        state = db.get(models.SheetSyncState, ("requests", request_id))
        if not req.pdf_file_id and not req.old_pdf_file_id_pending:
            req.old_pdf_file_id_pending = file_id
        sheet_queue.mark_dirty(db, "requests", request_id)
        if state is not None:
            _note_sheet_has(db, request_id, sheet_values)
        return False
    if current == (file_id, folder_id, generated_at):
        _note_sheet_has(db, request_id, sheet_values)
        return True
    state = db.get(models.SheetSyncState, ("requests", request_id))
    was_clean = state is not None and int(state.desired_version) <= int(state.synced_version)
    req.pdf_file_id, req.pdf_folder_id, req.pdf_generated_at = file_id, folder_id, generated_at
    if req.old_pdf_file_id_pending == file_id:
        req.old_pdf_file_id_pending = ""
    version = sheet_queue.mark_dirty(db, "requests", request_id)
    state = db.get(models.SheetSyncState, ("requests", request_id))
    if state is not None:
        if was_clean:
            state.synced_version = version
            state.desired_version = version
            state.last_synced_at = timeutil.now_utc()
        _note_sheet_has(db, request_id, sheet_values)
    db.flush()
    return True


def _apply_trash(db: Session, request_id: str, trash: Any, old_pdf_file_id: str) -> None:
    if not isinstance(trash, Mapping) or not request_id:
        return
    status = str(trash.get("status") or "")
    done = status == "trashed" or (status == "skipped" and trash.get("reason") == "sheet_has_newer_pdf")
    if not done:
        return
    req = db.get(models.Request, request_id)
    if req is None or not req.old_pdf_file_id_pending:
        return
    fid = normalize_gas(trash.get("fileId")) or old_pdf_file_id
    if req.old_pdf_file_id_pending in (fid, old_pdf_file_id):
        req.old_pdf_file_id_pending = ""


# ===== 一括の間隔（7-3 の最後） =====

_batch_last_call: dict[str, float] = {}
_batch_lock = threading.Lock()


def _batch_spacing(db: Session, batch_id: str) -> None:
    if not batch_id:
        return
    J = models.Job
    size = db.scalar(select(func.count()).select_from(J).where(J.batch_id == batch_id)) or 0
    if size < BATCH_SPACING_MIN_JOBS:
        return
    with _batch_lock:
        last = _batch_last_call.get(batch_id)
    if last is not None:
        wait = BATCH_SPACING_SECONDS - (_monotonic() - last)
        if wait > 0:
            _sleep(wait)


def _mark_batch_call(batch_id: str) -> None:
    if batch_id:
        with _batch_lock:
            _batch_last_call[batch_id] = _monotonic()
            if len(_batch_last_call) > 500:
                for key in list(_batch_last_call)[:250]:
                    _batch_last_call.pop(key, None)


# ===== 仕事の種類 =====


@dataclass(frozen=True)
class GasKind:
    kind: str
    gas_action: str
    skip_if_canceled: bool = True
    with_work_logs: bool = False


KINDS: dict[str, GasKind] = {
    KIND_AFTER_APPROVE: GasKind(KIND_AFTER_APPROVE, "internal.afterApprove", with_work_logs=True),
    KIND_AFTER_COMPLETE: GasKind(KIND_AFTER_COMPLETE, "internal.afterComplete", with_work_logs=True),
    KIND_REGENERATE: GasKind(KIND_REGENERATE, "internal.regenerate"),
    KIND_LEDGER_UPSERT: GasKind(KIND_LEDGER_UPSERT, "vps.upsertLedger"),
    KIND_AFTER_DELETE: GasKind(KIND_AFTER_DELETE, "internal.afterDelete", skip_if_canceled=False, with_work_logs=True),
    KIND_BATCH_PDF: GasKind(KIND_BATCH_PDF, "internal.batchPdf"),
    KIND_LEDGER_REWRITE_DAY: GasKind(KIND_LEDGER_REWRITE_DAY, "internal.ledgerRewriteDay", skip_if_canceled=False),
    KIND_LEDGER_REBUILD: GasKind(KIND_LEDGER_REBUILD, "internal.ledgerRebuild", skip_if_canceled=False),
}


def _old_pdf_file_id(args: Mapping[str, Any], req: models.Request | None) -> str:
    old = normalize_gas(args.get("oldPdfFileId"))
    if not old and req is not None:
        old = normalize_gas(req.old_pdf_file_id_pending)
    return old


def build_call(kind: str, req: models.Request | None, args: Mapping[str, Any]) -> tuple[dict | None, dict]:
    """GAS へ送る expect と args（DB の今の値で作る）。"""
    if req is None:
        if kind in (KIND_LEDGER_REWRITE_DAY,):
            return None, {"targetDate": str(args.get("targetDate") or "")}
        return None, {}
    expect: dict[str, Any] = {"status": req.status}
    out: dict[str, Any] = {}
    if kind == KIND_AFTER_COMPLETE:
        out["source"] = args.get("source", "")
        out["statusAtEnqueue"] = args.get("statusAtEnqueue", "")
        override = args.get("wlOverride")
        if isinstance(override, Mapping) and args.get("source") != "sebastian":
            out["workLog"] = dict(override)  # G の名前（今の workLogOverride と同じ形）
            out["wlOverride"] = dict(override)  # 仕様 7-2 の名前
            if override.get("actualEndAt"):
                expect["actualEndAt"] = override["actualEndAt"]
    elif kind == KIND_REGENERATE:
        out["reason"] = args.get("reason", "")
        out["oldPdfFileId"] = _old_pdf_file_id(args, req)
        if out["reason"] == "secondary":
            expect["approvedBy2"] = req.approved_by2
    elif kind == KIND_AFTER_DELETE:
        out["oldPdfFileId"] = _old_pdf_file_id(args, req)
        out["requestType"] = args.get("requestType") or req.request_type
        out["targetDate"] = args.get("targetDate") or request_target_ymd(req)
    elif kind == KIND_BATCH_PDF:
        out["targetDate"] = args.get("targetDate") or request_target_ymd(req)
    return expect, out


def call_gas(ctx_settings: Any, action: str, *, job_id: str, request_id: str = "", expect: dict | None = None,
             args: dict | None = None) -> dict:
    """GAS を 1 回呼ぶ（通信の失敗は GasTransportError のまま上げる）。"""
    client = get_client(ctx_settings)
    return client.call(action, job_id=job_id, request_id=request_id, expect=expect, args=args)


def _gas_job_id(ctx: JobContext) -> str:
    gid = str((ctx.guard or {}).get("gasJobId") or "")
    if gid:
        return gid
    created = int(ctx.created_at.timestamp()) if ctx.created_at else 0
    return f"ot-job-{ctx.job_id}-{created}"


def _store_job_result(db: Session, job_id: int, result: Mapping[str, Any]) -> None:
    row = db.get(models.Job, job_id)
    if row is not None:
        row.result_json = _dumps(dict(result))


def _error_text(resp: Mapping[str, Any]) -> str:
    code = str(resp.get("error") or "error")
    message = str(resp.get("message") or "")
    return f"{code}: {message}" if message and message != code else code


def handle_job(ctx: JobContext) -> dict:
    """GAS の仕事 1 件（jobs.register_kind のハンドラー）。"""
    spec = KINDS[ctx.kind]
    db = ctx.db
    req: models.Request | None = None
    if ctx.request_id:
        req = db.get(models.Request, ctx.request_id)
        if req is None:
            raise SkipJob("request_not_found")
        if spec.skip_if_canceled and req.status == "canceled":
            raise SkipJob("canceled")
        require_synced(db, ctx.request_id, created_at=ctx.created_at, now=ctx.now, with_work_logs=spec.with_work_logs)
    require_gas_ping(db, ctx.now)
    expect, gas_args = build_call(ctx.kind, req, ctx.args)
    snapshot = pdf_snapshot(req)
    _batch_spacing(db, ctx.batch_id)
    try:
        resp = call_gas(ctx.app_settings, spec.gas_action, job_id=_gas_job_id(ctx), request_id=ctx.request_id,
                        expect=expect, args=gas_args)
    finally:
        _mark_batch_call(ctx.batch_id)
    return apply_response(ctx, req, resp, snapshot, old_pdf_file_id=str(gas_args.get("oldPdfFileId") or ""))


def apply_response(ctx: JobContext, req: models.Request | None, resp: Mapping[str, Any], snapshot: tuple | None,
                   *, old_pdf_file_id: str = "") -> dict:
    """GAS の返事を記録する。失敗でも PDF ができていれば先に記録してから上げる（7-1）。"""
    db = ctx.db
    resp = dict(resp)
    pdf = resp.get("pdf")
    if req is not None and isinstance(pdf, Mapping) and pdf.get("status") in ("created", "already"):
        merged = dict(pdf)
        for key in ("pdfFileId", "pdfFolderId", "pdfGeneratedAt", "pdfGeneratedAtEpochMs"):
            if not merged.get(key) and resp.get(key):
                merged[key] = resp[key]
        record_pdf_result(db, req.request_id, merged, expected_current=snapshot)
    if req is not None:
        _apply_trash(db, req.request_id, resp.get("trash"), old_pdf_file_id)
    if resp.get("ok") is True:
        _after_success(ctx, resp)
        return resp
    _store_job_result(db, ctx.job_id, resp)
    ctx.commit()  # PDF の記録と返事を残してから失敗にする
    if resp.get("retryable") is True:
        raise GasJobError(_error_text(resp))
    raise PermanentJobError(_error_text(resp))


def _after_success(ctx: JobContext, resp: Mapping[str, Any]) -> None:
    """朝バッチの送り直しが成功したら、その日の台帳の書き直しをもう一度積む（7-5 の 6）。"""
    if ctx.kind != KIND_BATCH_PDF or not ctx.args.get("rewriteLedger"):
        return
    pdf = resp.get("pdf") if isinstance(resp.get("pdf"), Mapping) else {}
    if pdf.get("status") not in ("created", "already"):
        return
    ymd = str(ctx.args.get("targetDate") or "")
    if ymd:
        enqueue_ledger_rewrite_day(ctx.db, ymd)


def on_exhausted(ctx: JobContext, message: str) -> None:
    """failed になった仕事を管理者の欄に出す（sync_runs kind='gas_job_failed'）。"""
    now = timeutil.now_utc()
    ctx.db.add(models.SyncRun(kind=EXHAUSTED_KIND, started_at=now, finished_at=now, ok=False, summary_json={
        "jobId": ctx.job_id, "kind": ctx.kind, "requestId": ctx.request_id, "attempts": ctx.attempt,
        "error": message[:1000], "batchId": ctx.batch_id}))
    log.error("GAS の仕事が失敗のまま止まりました: job=%s kind=%s request=%s %s", ctx.job_id, ctx.kind,
              ctx.request_id, message)


# ===== 積む（10-2。W が業務のトランザクションの中で呼ぶ。commit しない） =====


def _request_or_error(db: Session, request_id: str) -> models.Request:
    req = db.get(models.Request, request_id)
    if req is None:
        raise ValueError(f"申請がありません: {request_id}")
    return req


def _pending_same(db: Session, kind: str, request_id: str = "", args_json: str | None = None) -> int | None:
    J = models.Job
    q = select(J).where(J.kind == kind, J.request_id == request_id, J.state == "pending").order_by(J.id)
    for row in db.scalars(q):
        if args_json is None or row.args_json == args_json:
            return int(row.id)
    return None


def _enqueue(db: Session, kind: str, args: dict, *, request_id: str = "", priority: int = 0,
             requires: list | tuple = (), guard: dict | None = None, batch_id: str = "") -> int:
    full_guard = {"gasJobId": new_gas_job_id()}
    full_guard.update(guard or {})
    return jobs.enqueue(db, kind, args, lane="gas", request_id=request_id, priority=priority, requires=requires,
                        guard=full_guard, batch_id=batch_id)


def enqueue_after_approve(db: Session, request_id: str) -> int:
    """一次承認（WorkLogs に actualEndAt があるとき W が呼ぶ）。同じ申請の待ちの仕事があればそれを返す。"""
    req = _request_or_error(db, request_id)
    existing = _pending_same(db, KIND_AFTER_APPROVE, request_id)
    if existing is not None:
        return existing
    return _enqueue(db, KIND_AFTER_APPROVE, {}, request_id=request_id,
                    requires=requires_for_request(db, request_id, with_work_logs=True),
                    guard={"statusAtEnqueue": req.status})


def enqueue_after_complete(db: Session, request_id: str, *, source: str, status_at_enqueue: str,
                           wl_override: dict | None = None, priority: int = 0) -> int:
    """残業・休日の完了（screen）、総務の手動入力（admin）、Sebastian の完了（sebastian、priority=10）。"""
    if source not in COMPLETE_SOURCES:
        raise ValueError(f"source は {COMPLETE_SOURCES} のどれかです: {source}")
    _request_or_error(db, request_id)
    args: dict[str, Any] = {"source": source, "statusAtEnqueue": status_at_enqueue}
    if wl_override is not None and source != "sebastian":
        args["wlOverride"] = dict(wl_override)
    return _enqueue(db, KIND_AFTER_COMPLETE, args, request_id=request_id, priority=priority,
                    requires=requires_for_request(db, request_id, with_work_logs=True),
                    guard={"statusAtEnqueue": status_at_enqueue})


def enqueue_regenerate(db: Session, request_id: str, *, reason: str, old_pdf_file_id: str = "",
                       batch_id: str = "") -> int:
    """PDF の作り直し（secondary＝二次承認、edit＝編集、admin＝管理者）。W は先に DB の 3 列を空にしておく。"""
    if reason not in REGENERATE_REASONS:
        raise ValueError(f"reason は {REGENERATE_REASONS} のどれかです: {reason}")
    req = _request_or_error(db, request_id)
    old = old_pdf_file_id or ""
    if old and not req.old_pdf_file_id_pending:
        req.old_pdf_file_id_pending = old
    return _enqueue(db, KIND_REGENERATE, {"reason": reason, "oldPdfFileId": old}, request_id=request_id,
                    requires=requires_for_request(db, request_id, with_work_logs=False),
                    guard={"statusAtEnqueue": req.status, "approvedBy2": req.approved_by2}, batch_id=batch_id)


def enqueue_ledger_upsert(db: Session, request_id: str, *, batch_id: str = "") -> int:
    """台帳だけ（PDF の無い二次承認の単票）。"""
    req = _request_or_error(db, request_id)
    return _enqueue(db, KIND_LEDGER_UPSERT, {}, request_id=request_id,
                    requires=requires_for_request(db, request_id, with_work_logs=False),
                    guard={"statusAtEnqueue": req.status}, batch_id=batch_id)


def enqueue_after_delete(db: Session, request_id: str, *, request_type: str, target_date: str,
                         old_pdf_file_id: str) -> int:
    """削除（PDF の有無に関係なく。台帳の行を消し、前の PDF をゴミ箱へ）。"""
    req = _request_or_error(db, request_id)
    if old_pdf_file_id and not req.old_pdf_file_id_pending:
        req.old_pdf_file_id_pending = old_pdf_file_id
    args = {"oldPdfFileId": old_pdf_file_id or "", "requestType": request_type, "targetDate": target_date}
    return _enqueue(db, KIND_AFTER_DELETE, args, request_id=request_id,
                    requires=requires_for_request(db, request_id, with_work_logs=True),
                    guard={"statusAtEnqueue": req.status})


def enqueue_ledger_rebuild(db: Session) -> int:
    """台帳の作り直し（管理者の ops_rebuildLedger）。待ちのものがあればそれを返す。"""
    existing = _pending_same(db, KIND_LEDGER_REBUILD)
    if existing is not None:
        return existing
    return _enqueue(db, KIND_LEDGER_REBUILD, {})


def enqueue_batch_pdf(db: Session, request_id: str, *, target_date: str, rewrite_ledger: bool) -> int:
    """朝バッチで通信に失敗した件の送り直し（7-5 の 2）。"""
    _request_or_error(db, request_id)
    return _enqueue(db, KIND_BATCH_PDF, {"targetDate": target_date, "rewriteLedger": bool(rewrite_ledger)},
                    request_id=request_id, requires=requires_for_request(db, request_id, with_work_logs=False))


def enqueue_ledger_rewrite_day(db: Session, target_date: str) -> int:
    """その日の台帳の書き直し（同じ日の待ちの仕事があればそれを返す）。"""
    args = {"targetDate": target_date}
    existing = _pending_same(db, KIND_LEDGER_REWRITE_DAY, "", _dumps(args))
    if existing is not None:
        return existing
    return _enqueue(db, KIND_LEDGER_REWRITE_DAY, args)


# ===== 結果を待つ（Sebastian。4-4） =====


def _job_outcome(row: models.Job) -> dict | None:
    result = _loads(row.result_json)
    if row.state == "done":
        out = dict(result) if isinstance(result, dict) else {"ok": True}
        out["jobState"] = "done"
        return out
    if row.state in ("failed", "skipped"):
        out = dict(result) if isinstance(result, dict) else {}
        out.setdefault("ok", False)
        out.setdefault("error", row.last_error or row.state)
        out["jobState"] = row.state
        return out
    if row.state == "pending" and int(row.attempts or 0) >= 1 and isinstance(result, dict) and result.get("ok") is False:
        # 1 回目が失敗して送り直しを待っている（仕事は続く）。今の GAS と同じく失敗をすぐ返す
        out = dict(result)
        out["jobState"] = "retrying"
        return out
    return None


def wait_for_result(session_factory: Callable[[], Session], job_id: int, timeout_seconds: float) -> dict | None:
    """仕事の結果を待つ。done なら GAS の返事（jobState='done'）、failed／skipped・1 回目の失敗なら
    {ok:false, error, …, jobState}。timeout_seconds を過ぎたら None（仕事は続く）。"""
    deadline = _monotonic() + max(0.0, float(timeout_seconds))
    interval = 0.25
    while True:
        with session_factory() as db:
            row = db.get(models.Job, int(job_id))
            if row is None:
                return {"ok": False, "error": "job_not_found", "jobState": "missing"}
            outcome = _job_outcome(row)
        if outcome is not None:
            return outcome
        remaining = deadline - _monotonic()
        if remaining <= 0:
            return None
        _sleep(min(interval, remaining))
        interval = min(interval * 1.5, 1.0)


def sebastian_pdf_view(outcome: Mapping[str, Any] | None) -> dict:
    """wait_for_result の結果を、今の Sebastian の返事の pdf の形にする（4-4）。"""
    if outcome is None:
        return {"error": SEBASTIAN_TIMEOUT_MESSAGE}
    pdf = outcome.get("pdf") if isinstance(outcome.get("pdf"), Mapping) else {}
    status = pdf.get("status")
    file_id = normalize_gas(pdf.get("pdfFileId") or outcome.get("pdfFileId"))
    if status == "created":
        return {"ok": True, "pdfFileId": file_id, "pdfName": pdf.get("pdfName") or outcome.get("pdfName") or "",
                "folderId": normalize_gas(pdf.get("folderId") or outcome.get("pdfFolderId"))}
    if status == "already":
        return {"already": True, "pdfFileId": file_id}
    if status == "error":
        return {"error": str(pdf.get("error") or outcome.get("message") or "PDF error")}
    if outcome.get("jobState") == "skipped":
        return {"error": "PDF は作りませんでした（" + str(outcome.get("error") or "skipped") + "）"}
    return {"error": str(outcome.get("message") or outcome.get("error") or pdf.get("reason") or "PDF error")}


def wait_for_sebastian_pdf(session_factory: Callable[[], Session], job_id: int, settings: Any = None) -> dict:
    """SEBASTIAN_PDF_WAIT_SECONDS だけ待って pdf の形で返す（W の Sebastian の入口が使う）。"""
    seconds = float(getattr(_settings(settings), "sebastian_pdf_wait_seconds", 40))
    return sebastian_pdf_view(wait_for_result(session_factory, job_id, seconds))


# ===== gas_ping（5 分ごと。9-1・9-3） =====


def normalize_ping(ping: Mapping[str, Any]) -> dict:
    """G の ping（dbId・dbTimeZone）を仕様 7-2 の名前（spreadsheetId・spreadsheetTimeZone）にそろえる。"""
    out = dict(ping)
    if not out.get("spreadsheetId") and out.get("dbId"):
        out["spreadsheetId"] = out["dbId"]
    if not out.get("spreadsheetTimeZone") and out.get("dbTimeZone"):
        out["spreadsheetTimeZone"] = out["dbTimeZone"]
    return out


def run_gas_ping(session_factory: Callable[[], Session], settings: Any = None) -> dict:
    """internal.ping を呼び、この環境の許可リストと照合して sync_runs（kind='gas_ping'）に残す。"""
    from app.services.gas_client import GasTransportError

    s = _settings(settings)
    started = timeutil.now_utc()
    summary: dict[str, Any] = {}
    try:
        ping = normalize_ping(call_gas(s, "internal.ping", job_id="ping-" + uuid.uuid4().hex))
        problems = safety.verify_gas_ping(ping, s)
        gas_settings = ping.get("settings") if isinstance(ping.get("settings"), Mapping) else {}
        summary = {"problems": problems, "version": ping.get("version") or ping.get("apiVersion"),
                   "migrationMode": ping.get("migrationMode", ping.get("migrated")),
                   "pdfMode": str(gas_settings.get("PDF_MODE") or ""),
                   "spreadsheetTimeZone": ping.get("spreadsheetTimeZone") or ""}
    except GasTransportError as exc:
        problems = [f"GAS に届きませんでした: {exc}"]
        summary = {"problems": problems}
    except Exception as exc:  # noqa: BLE001 - 設定の誤りなども不合格として残す
        problems = [f"{type(exc).__name__}: {exc}"]
        summary = {"problems": problems}
    ok = not problems
    with session_factory() as db:
        db.add(models.SyncRun(kind=GAS_PING_KIND, started_at=started, finished_at=timeutil.now_utc(), ok=ok,
                              summary_json=summary))
        db.commit()
    if not ok:
        log.error("gas_ping 不合格（GAS のレーンを止めます）: %s", problems)
    return {"ok": ok, **summary}


# ===== 登録 =====


def register_worker() -> None:
    from app import worker

    for kind in KINDS:
        jobs.register_kind(kind, handle_job, lane="gas", on_exhausted=on_exhausted)
    worker.register_schedule("gas_ping", task=run_gas_ping, every_minutes=5, run_at_startup=True)


__all__ = [
    "GasJobError", "KINDS", "SEBASTIAN_TIMEOUT_MESSAGE", "enqueue_after_approve", "enqueue_after_complete",
    "enqueue_after_delete", "enqueue_batch_pdf", "enqueue_ledger_rebuild", "enqueue_ledger_rewrite_day",
    "enqueue_ledger_upsert", "enqueue_regenerate", "gas_ping_ok", "handle_job", "parse_jst_text",
    "record_pdf_result", "register_worker", "run_gas_ping", "sebastian_pdf_view", "set_client_factory",
    "wait_for_result", "wait_for_sebastian_pdf",
]

