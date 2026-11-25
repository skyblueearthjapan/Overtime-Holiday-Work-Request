"""一次承認（4-3 の 13・15・16・17・18）。approval.gs:280-551 の写し。

- 承認は申請の行を FOR UPDATE で読んでから状態を見る。2 つのタブで押しても、2 回目は「既に承認済みです。」で
  仕事は 1 回だけ積まれる。WorkLogs の actualEndAt（下の行が勝つ）があれば internal.afterApprove を積む。
- 一括は ID を並べ替えてからロックする。結果は GAS と同じく Requests の行の順（その後に見つからなかった ID）。
- 却下は submitted のときだけ。台帳も PDF も触らない。
"""
from __future__ import annotations

from typing import Any

from app import permissions
from app.jsutil import normalize_gas
from app.rpc import RpcContext, RpcError

from . import gasjobs, store
from .summaries import build_monthly_summary, recompute_people_summary


def _find_for_approval(ctx: RpcContext, request_id: Any):
    req = store.lock_request(ctx.db, request_id)
    if req is None:
        if store.request_count(ctx.db) == 0:
            raise RpcError("Requestsが空です。")
        raise RpcError("requestIdが見つかりません。")
    return req


def api_approveRequest(ctx: RpcContext, request_id=None):
    email = ctx.user.email or "unknown"
    req = _find_for_approval(ctx, request_id)
    dept = normalize_gas(req.dept)
    status = normalize_gas(req.status)
    if not permissions.can_approve_dept(ctx.db, email, dept, ctx.settings):
        raise RpcError("この申請の承認権限がありません。")
    if status == "approved":
        return {"ok": True, "message": "既に承認済みです。"}
    if status == "canceled":
        raise RpcError("キャンセル済みです。")

    req.status = "approved"
    req.approved_at = ctx.now
    req.approved_by = email
    store.mark_request(ctx.db, req)
    if store.worklog_view(ctx.db, req.request_id).actualEndAt:
        gasjobs.after_approve(ctx.db, req.request_id)
    return {"ok": True, "requestId": req.request_id, "approvedBy": email, "approvedAt": store.jst_text(ctx.now)}


def api_rejectRequest(ctx: RpcContext, request_id=None):
    email = ctx.user.email or "unknown"
    req = _find_for_approval(ctx, request_id)
    dept = normalize_gas(req.dept)
    status = normalize_gas(req.status)
    if not permissions.can_approve_dept(ctx.db, email, dept, ctx.settings):
        raise RpcError("この申請の承認権限がありません。")
    if status != "submitted":
        raise RpcError("未承認状態でのみ却下可能です（現在: " + status + "）。")
    req.status = "canceled"
    req.approved_by = email + "(却下)"
    req.approved_at = ctx.now
    store.mark_request(ctx.db, req)
    return {"ok": True, "requestId": req.request_id}


def api_approveRequestsBatch(ctx: RpcContext, request_ids=None):
    if not request_ids:
        return {"ok": True, "results": []}
    if not isinstance(request_ids, list):
        request_ids = [request_ids]
    if len(request_ids) > 50:
        raise RpcError("一括承認は50件までです。")
    email = ctx.user.email or "unknown"
    db = ctx.db

    target: dict[str, bool] = {}
    for rid in request_ids:
        target[rid if isinstance(rid, str) else str(rid)] = True
    locked = store.lock_requests(db, list(target))
    if store.request_count(db) == 0:
        return {"ok": True, "results": []}

    results: list[dict] = []
    approved_ids: list[str] = []
    approved_at = store.jst_text(ctx.now)
    for req in store.requests_in_sheet_order(db):
        rid = normalize_gas(req.request_id)
        if not rid or not target.get(rid):
            continue
        req = locked.get(req.request_id, req)
        dept = normalize_gas(req.dept)
        status = normalize_gas(req.status)
        if status == "approved":
            results.append({"requestId": rid, "ok": True, "message": "既に承認済み"})
        elif status == "canceled":
            results.append({"requestId": rid, "ok": False, "error": "キャンセル済み"})
        elif not permissions.can_approve_dept(db, email, dept, ctx.settings):
            results.append({"requestId": rid, "ok": False, "error": "承認権限なし"})
        else:
            req.status = "approved"
            req.approved_at = ctx.now
            req.approved_by = email
            store.mark_request(db, req)
            approved_ids.append(req.request_id)
            results.append({"requestId": rid, "ok": True, "approvedBy": email, "approvedAt": approved_at})
        del target[rid]

    if approved_ids:
        wl_map = store.worklog_map(db)
        for rid in approved_ids:
            if wl_map.get(normalize_gas(rid), store.EMPTY_WL).actualEndAt:
                gasjobs.after_approve(db, rid)

    for missing in target:
        results.append({"requestId": missing, "ok": False, "error": "見つかりません"})
    return {"ok": True, "results": results}


def api_getApproverDepts(ctx: RpcContext):
    email = ctx.user.email
    profiles = permissions.approver_profiles(ctx.db)
    if ctx.is_admin:
        return {"isAdmin": True, "approvers": profiles, "currentEmail": email}
    mine = [p for p in profiles if p["email"] == email]
    if not mine:
        raise RpcError("承認者権限がありません。DeptApprovers または ApproverMap に登録してください。")
    return {"isAdmin": False, "approvers": mine, "currentEmail": email}


def api_approverMonthlySummary(ctx: RpcContext, year_month=None, depts_or_dept=None):
    if isinstance(depts_or_dept, list):
        dept_list = [d for d in (normalize_gas(x) for x in depts_or_dept) if d]
    else:
        dept_list = [normalize_gas(depts_or_dept)]
    if not dept_list or not dept_list[0]:
        raise RpcError("部署が指定されていません。")
    for dept in dept_list:
        if not permissions.can_approve_dept(ctx.db, ctx.user.email, dept, ctx.settings):
            raise RpcError("部署「" + dept + "」の月次サマリーを閲覧する権限がありません。")
    if len(dept_list) == 1:
        return build_monthly_summary(ctx, year_month, dept_list[0])

    data = build_monthly_summary(ctx, year_month, "ALL")
    allowed = set(dept_list)
    data["people"] = [p for p in data["people"] if p["dept"] in allowed]
    data["watch"] = [w for w in data["watch"] if w["dept"] in allowed]
    data["deptFilter"] = ", ".join(dept_list)
    recompute_people_summary(data)  # charts.dept は作り直さない（H14）
    return data


HANDLERS = {
    "api_approveRequest": api_approveRequest,
    "api_rejectRequest": api_rejectRequest,
    "api_approveRequestsBatch": api_approveRequestsBatch,
    "api_getApproverDepts": api_getApproverDepts,
    "api_approverMonthlySummary": api_approverMonthlySummary,
}
