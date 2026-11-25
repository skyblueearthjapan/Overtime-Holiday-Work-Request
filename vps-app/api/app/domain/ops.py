"""管理者だけの関数（実装仕様 4-2）。rpc が呼ぶ前に permissions.is_admin を確かめる（ADMIN_ONLY_FUNCTIONS）。

担当 D・M の仕事は 10-2 の名前で呼ぶ（モジュールがまだ無ければ「まだ用意されていません」の 400）。
休憩の再計算（ops_recalcBreaks）は worklog.gs:712-806 の recalcWorkLogsBreakMinutes を写す。
"""
from __future__ import annotations

import re
from datetime import timedelta

from app import auth, jobs, sheet_queue
from app.jsutil import js_round, js_string, normalize_gas
from app.rpc import RpcContext, RpcError

from . import gasjobs, store
from .breaks import calc_break_by_clock_time
from .jsdate import js_truthy


def ops_getStatus(ctx: RpcContext):
    return gasjobs.service("ops_status").get_ops_status(ctx.db)


def ops_retryJob(ctx: RpcContext, job_id=None):
    try:
        jid = int(job_id)
    except (TypeError, ValueError) as exc:
        raise RpcError("仕事の番号が正しくありません。") from exc
    if not jobs.retry_failed(ctx.db, jid):
        raise RpcError("失敗した仕事が見つかりません。")
    return {"ok": True}


def ops_syncSheetConfigNow(ctx: RpcContext):
    return gasjobs.service("config_sync").sync_sheet_config(ctx.db)


def ops_runMasterSync(ctx: RpcContext):
    return gasjobs.service("master_sync").sync_masters(ctx.db)


def ops_regeneratePdf(ctx: RpcContext, request_id=None):
    """管理者の作り直し（7-4）: 3 列を空にして old_pdf_file_id_pending、internal.regenerate（reason=admin）。"""
    gas_jobs = gasjobs.service("gas_jobs")
    req = store.lock_request(ctx.db, request_id)
    if req is None:
        raise RpcError("申請が見つかりません。")
    old_pdf = ""
    if normalize_gas(req.pdf_file_id):
        old_pdf = store.clear_pdf_columns(req)
        store.mark_request(ctx.db, req)
    job_id = gas_jobs.enqueue_regenerate(ctx.db, req.request_id, reason="admin", old_pdf_file_id=old_pdf)
    return {"ok": True, "jobId": job_id}


_YMD = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")


def ops_batchPdfForDate(ctx: RpcContext, ymd=None):
    text = js_string(ymd) if ymd is not None else ""
    if not _YMD.match(text):
        raise RpcError("日付が不正です（yyyy-MM-dd）")
    job_id = gasjobs.service("morning_batch").enqueue_manual_batch(ctx.db, text)
    return {"ok": True, "jobId": job_id}


def ops_rebuildLedger(ctx: RpcContext):
    job_id = gasjobs.service("gas_jobs").enqueue_ledger_rebuild(ctx.db)
    return {"ok": True, "jobId": job_id}


def recalc_breaks(db, now, updated_by: str = "break-recalc") -> int:
    """recalcWorkLogsBreakMinutes: 開始・終了がそろう行の休憩・net を時刻窓方式で計算し直し、変わった行だけ書く。"""
    changed = 0
    stamp = store.jst_text(now)
    for wl in store.worklogs_in_sheet_order(db):
        if not wl.actual_start_at or not wl.actual_end_at:
            continue
        start = store.parse_wl_datetime(wl.actual_start_at)
        end = store.parse_wl_datetime(wl.actual_end_at)
        if start is None or end is None:
            continue
        if end <= start:
            continue
        actual = int(wl.actual_minutes or 0)
        if not actual:
            actual = max(0, js_round(((end - start) // timedelta(milliseconds=1)) / 60000))
        new_break = calc_break_by_clock_time(start, end)
        new_net = max(0, actual - new_break)
        if new_break != int(wl.break_minutes or 0) or new_net != int(wl.net_minutes or 0):
            wl.break_minutes = new_break
            wl.net_minutes = new_net
            wl.updated_at = stamp
            wl.updated_by = updated_by
            store.set_cell_kind(wl, "updatedAt", "string")
            db.flush()
            sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
            changed += 1
    return changed


def ops_recalcBreaks(ctx: RpcContext):
    return {"ok": True, "changed": recalc_breaks(ctx.db, ctx.now)}


def ops_revokeSessions(ctx: RpcContext, email=None):
    if not js_truthy(email) or not isinstance(email, str):
        raise RpcError("メールアドレスを指定してください。")
    return {"ok": True, "revoked": auth.revoke_sessions(ctx.db, email)}


HANDLERS = {
    "ops_getStatus": ops_getStatus,
    "ops_retryJob": ops_retryJob,
    "ops_syncSheetConfigNow": ops_syncSheetConfigNow,
    "ops_runMasterSync": ops_runMasterSync,
    "ops_regeneratePdf": ops_regeneratePdf,
    "ops_batchPdfForDate": ops_batchPdfForDate,
    "ops_rebuildLedger": ops_rebuildLedger,
    "ops_recalcBreaks": ops_recalcBreaks,
    "ops_revokeSessions": ops_revokeSessions,
}
