"""二次承認（4-3 の 19・20・21）。approval.gs:593-883 の写し。

- 一覧は読むだけ（GAS の見出しの自動追加は写さない＝Q17）。
- 単票: PDF があれば 3 列を空にして internal.regenerate（reason=secondary）、無ければ internal.ledgerUpsert
  （単票はいつも台帳を書く）。返事の pdfRegenerated はいつも false（PDF は後で作る）。
- 一括: 件数の確かめが先、権限が後（approval.gs:776-779）。新しく二次承認したものだけ積む。PDF の作り直しは
  1 つのまとまり（batch_id）にして、6 件以上なら M のレーンが 500ms ずつ空ける。
"""
from __future__ import annotations

import functools
import uuid

from app import timeutil
from app.jsutil import js_sort_key, normalize_gas
from app.rpc import RpcContext, RpcError

from . import gasjobs, store
from .jsdate import MISSING, days_ago, is_retroactive, js_new_date


def _lt(a: str, b: str) -> bool:
    return js_sort_key(a) < js_sort_key(b)


def api_getSecondaryApprovalList(ctx: RpcContext, date_ymd=MISSING):
    moment = js_new_date(date_ymd)
    if moment is None:
        raise RpcError("日付が不正です（yyyy-MM-dd）")
    ymd = timeutil.to_jst_text(moment, "yyyy-MM-dd")
    today = timeutil.to_jst_text(ctx.now, "yyyy-MM-dd")
    out: list[dict] = []
    for req in store.requests_in_sheet_order(ctx.db):
        status = normalize_gas(req.status)
        if not status or status == "canceled":
            continue
        approved_by2 = normalize_gas(req.approved_by2)
        if approved_by2:
            continue
        target = store.target_ymd(req)
        if target is None:
            continue
        if target > ymd:
            continue
        out.append({
            "requestId": normalize_gas(req.request_id),
            "requestType": normalize_gas(req.request_type),
            "status": status,
            "dept": normalize_gas(req.dept),
            "workerName": normalize_gas(req.worker_name),
            "targetDate": target,
            "isOverdue": target < ymd,
            "approvedMinutes": int(req.approved_minutes or 0),
            "submittedAt": store.jst_text(req.submitted_at, "yyyy-MM-dd HH:mm"),
            "approvedBy": normalize_gas(req.approved_by),
            "approvedAt": store.jst_text(req.approved_at, "yyyy-MM-dd HH:mm"),
            "approvedBy2": approved_by2,
            "approvedAt2": store.jst_text(req.approved_at2, "yyyy-MM-dd HH:mm"),
            "pdfFileId": normalize_gas(req.pdf_file_id),
            "isRetroactive": is_retroactive(req.submitted_at, target),
            "daysAgo": days_ago(target, today),
        })

    def cmp(a: dict, b: dict) -> int:
        a_over, b_over = (0 if a["isOverdue"] else 1), (0 if b["isOverdue"] else 1)
        if a_over != b_over:
            return a_over - b_over
        a_ready = 0 if (a["status"] == "approved" and not a["approvedBy2"]) else 1
        b_ready = 0 if (b["status"] == "approved" and not b["approvedBy2"]) else 1
        if a_ready != b_ready:
            return a_ready - b_ready
        if a["targetDate"] != b["targetDate"]:
            return -1 if _lt(a["targetDate"], b["targetDate"]) else 1
        if a["dept"] != b["dept"]:
            return -1 if _lt(a["dept"], b["dept"]) else 1
        return -1 if _lt(a["workerName"], b["workerName"]) else 1 if _lt(b["workerName"], a["workerName"]) else 0

    out.sort(key=functools.cmp_to_key(cmp))
    return {"date": ymd, "items": out}


def api_secondaryApprove(ctx: RpcContext, request_id=None):
    email = ctx.require_admin()
    db = ctx.db
    req = store.lock_request(db, request_id)
    if req is None:
        if store.request_count(db) == 0:
            raise RpcError("Requestsが空です。")
        raise RpcError("requestIdが見つかりません。")
    status = normalize_gas(req.status)
    if status != "approved":
        raise RpcError("一次承認が完了していません（現在: " + status + "）。")
    if normalize_gas(req.approved_by2):
        return {"ok": True, "message": "既に二次承認済みです。"}

    req.approved_by2 = email
    req.approved_at2 = ctx.now
    had_pdf = bool(normalize_gas(req.pdf_file_id))
    old_pdf = store.clear_pdf_columns(req) if had_pdf else ""
    store.mark_request(db, req)
    if had_pdf:
        gasjobs.regenerate(db, req.request_id, reason="secondary", old_pdf_file_id=old_pdf)
    else:
        gasjobs.ledger_upsert(db, req.request_id)
    return {"ok": True, "requestId": req.request_id, "approvedBy2": email,
            "approvedAt2": store.jst_text(ctx.now), "pdfRegenerated": False}


def api_secondaryApproveBatch(ctx: RpcContext, request_ids=None):
    if not request_ids:
        return {"ok": True, "results": []}
    if not isinstance(request_ids, list):
        request_ids = [request_ids]
    if len(request_ids) > 50:
        raise RpcError("一括承認は50件までです。")
    email = ctx.require_admin()
    db = ctx.db

    target: dict[str, bool] = {}
    for rid in request_ids:
        target[rid if isinstance(rid, str) else str(rid)] = True
    locked = store.lock_requests(db, list(target))
    if store.request_count(db) == 0:
        return {"ok": True, "results": []}

    results: list[dict] = []
    pdf_targets: list[tuple[str, str]] = []
    no_pdf: list[str] = []
    stamp = store.jst_text(ctx.now)
    for req in store.requests_in_sheet_order(db):
        rid = normalize_gas(req.request_id)
        if not rid or not target.get(rid):
            continue
        req = locked.get(req.request_id, req)
        status = normalize_gas(req.status)
        if normalize_gas(req.approved_by2):
            results.append({"requestId": rid, "ok": True, "message": "既に二次承認済み"})
        elif status != "approved":
            results.append({"requestId": rid, "ok": False, "error": "一次承認未完了"})
        else:
            req.approved_by2 = email
            req.approved_at2 = ctx.now
            if normalize_gas(req.pdf_file_id):
                pdf_targets.append((req.request_id, store.clear_pdf_columns(req)))
            else:
                no_pdf.append(req.request_id)
            store.mark_request(db, req)
            results.append({"requestId": rid, "ok": True, "approvedBy2": email, "approvedAt2": stamp})
        del target[rid]
    for missing in target:
        results.append({"requestId": missing, "ok": False, "error": "見つかりません"})

    # GAS: PDF の作り直し（6 件以上は 500ms ずつ）→ 新しく二次承認した分の台帳。作り直しの仕事が台帳も書く。
    batch_id = str(uuid.uuid4()) if pdf_targets else ""
    for rid, old_pdf in pdf_targets:
        gasjobs.regenerate(db, rid, reason="secondary", old_pdf_file_id=old_pdf, batch_id=batch_id)
    for rid in no_pdf:
        gasjobs.ledger_upsert(db, rid)
    return {"ok": True, "results": results}


HANDLERS = {
    "api_getSecondaryApprovalList": api_getSecondaryApprovalList,
    "api_secondaryApprove": api_secondaryApprove,
    "api_secondaryApproveBatch": api_secondaryApproveBatch,
}
