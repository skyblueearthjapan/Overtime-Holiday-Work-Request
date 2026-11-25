"""一覧の画面（4-3 の 6・14）。worklog.gs:404-533（TOP）と approval.gs:132-276（承認者画面）の写し。

表示の条件は 01 §2.7 のとおり（今のまま）:
- TOP: canceled・空は出さない。完了済みかつ二次承認済みは当日だけ。期限超過（過去・approved・30 日以内・未完了）は
  いつも出す。残業はそれ以外は当日だけ（過去の submitted の残業は出ない＝Q5）。休日は submitted・approved の全件。
- 承認者: 空だけ除く（canceled の当日分も出す＝Q4）。残業の submitted は今日以前、休日の submitted は全件。
  approved・canceled は当日だけ。
「今日」はサーバーの JST の日付（8-1）。
"""
from __future__ import annotations

import functools
from datetime import timedelta
from typing import Any

from app import permissions, timeutil
from app.jsutil import js_sort_key, normalize_gas
from app.rpc import RpcContext

from . import store
from .jsdate import days_ago, is_retroactive, js_truthy, label_ja, label_md_dow


def _lt(a: str, b: str) -> bool:
    return js_sort_key(a) < js_sort_key(b)


def _cmp_str(a: str, b: str) -> int:
    return -1 if _lt(a, b) else 1 if _lt(b, a) else 0


def build_dashboard(ctx: RpcContext) -> dict:
    db = ctx.db
    wl_map = store.worklog_map(db)
    now_jst = timeutil.to_jst(ctx.now)
    today = now_jst.date().isoformat()
    overdue_limit = (now_jst.date() - timedelta(days=30)).isoformat()
    overtime: list[dict] = []
    holiday: list[dict] = []

    for req in store.requests_in_sheet_order(db):
        status = normalize_gas(req.status)
        if not status or status == "canceled":
            continue
        target = store.target_ymd(req)
        if target is None:
            continue
        request_type = normalize_gas(req.request_type)
        rid = normalize_gas(req.request_id)
        wl = wl_map.get(rid, store.EMPTY_WL)
        start_at, end_at = str(wl.actualStartAt or ""), str(wl.actualEndAt or "")
        approved_by2 = normalize_gas(req.approved_by2)

        is_overdue = False
        if target < today and status == "approved" and target >= overdue_limit:
            if request_type == "overtime" and not end_at:
                is_overdue = True
            elif request_type == "holiday" and (not start_at or not end_at):
                is_overdue = True

        item = {
            "requestId": rid or "",
            "requestType": request_type or "",
            "status": status or "",
            "dept": normalize_gas(req.dept) or "",
            "workerCode": normalize_gas(req.worker_code) or "",
            "workerName": normalize_gas(req.worker_name) or "",
            "targetDate": target or "",
            "targetDateLabel": "",
            "approvedMinutes": int(req.approved_minutes or 0),
            "actualStartAt": start_at,
            "actualEndAt": end_at,
            "actualMinutes": wl.actualMinutes,
            "breakMinutes": wl.breakMinutes,
            "netMinutes": wl.netMinutes,
            "pdfFileId": normalize_gas(req.pdf_file_id) or "",
            "approvedBy2": approved_by2,
            "isOverdue": is_overdue,
            "reason": req.reason or "",
        }

        if wl.actualEndAt and approved_by2 and target != today:
            continue
        if is_overdue:
            item["targetDateLabel"] = label_md_dow(target)
            (overtime if request_type == "overtime" else holiday).append(item)
            continue
        if request_type == "overtime" and target == today:
            overtime.append(item)
        elif request_type == "holiday" and status in ("submitted", "approved"):
            item["targetDateLabel"] = label_md_dow(target)
            holiday.append(item)

    def ot_cmp(a: dict, b: dict) -> int:
        if a["dept"] != b["dept"]:
            return -1 if _lt(a["dept"], b["dept"]) else 1
        return _cmp_str(a["workerName"], b["workerName"])

    def hol_cmp(a: dict, b: dict) -> int:
        if a["targetDate"] != b["targetDate"]:
            return -1 if _lt(a["targetDate"], b["targetDate"]) else 1
        return ot_cmp(a, b)

    overtime.sort(key=functools.cmp_to_key(ot_cmp))
    holiday.sort(key=functools.cmp_to_key(hol_cmp))
    return {"today": today, "overtime": overtime, "holiday": holiday}


def api_getDashboard(ctx: RpcContext):
    try:
        return build_dashboard(ctx)
    except Exception as exc:  # noqa: BLE001 - GAS と同じく 200 で {error}
        return {"today": "", "overtime": [], "holiday": [], "error": str(exc)}


def api_getApproverDashboard(ctx: RpcContext, depts_or_dept=None):
    try:
        email = ctx.user.email or ""
        admin = (not email) or ctx.is_admin
        if not js_truthy(depts_or_dept) or (isinstance(depts_or_dept, list) and len(depts_or_dept) == 0):
            return {"today": "", "overtime": [], "holiday": [], "isAdmin": admin, "error": "部署が指定されていません。"}
        dept_list = ([normalize_gas(d) for d in depts_or_dept] if isinstance(depts_or_dept, list)
                     else [normalize_gas(depts_or_dept)])
        if not admin:
            for dept in dept_list:
                if not permissions.can_approve_dept(ctx.db, email, dept, ctx.settings):
                    return {"today": "", "overtime": [], "holiday": [], "isAdmin": False,
                            "error": "部署「" + dept + "」の承認権限がありません。"}
        allowed = set(dept_list)
        today = timeutil.to_jst(ctx.now).date().isoformat()
        overtime: list[dict] = []
        holiday: list[dict] = []
        for req in store.requests_in_sheet_order(ctx.db):
            status = normalize_gas(req.status)
            if not status:
                continue
            dept = normalize_gas(req.dept)
            if dept not in allowed:
                continue
            target = store.target_ymd(req)
            if target is None:
                continue
            request_type = normalize_gas(req.request_type)
            item: dict[str, Any] = {
                "requestId": normalize_gas(req.request_id) or "",
                "requestType": request_type or "",
                "status": status or "",
                "dept": dept or "",
                "workerName": normalize_gas(req.worker_name) or "",
                "targetDate": target or "",
                "targetDateLabel": label_ja(target, today),
                "approvedMinutes": int(req.approved_minutes or 0),
                "submittedAt": store.jst_text(req.submitted_at),
                "isOverdue": False,
                "isRetroactive": is_retroactive(req.submitted_at, target),
                "daysAgo": days_ago(target, today),
            }
            if request_type == "overtime":
                if status == "submitted":
                    if target <= today:
                        item["isOverdue"] = target < today
                        overtime.append(item)
                elif target == today:
                    overtime.append(item)
            elif request_type == "holiday":
                if status == "submitted":
                    item["isOverdue"] = target < today
                    holiday.append(item)
                elif target == today:
                    holiday.append(item)

        def cmp(a: dict, b: dict) -> int:
            if a["status"] != b["status"]:
                if a["status"] == "submitted":
                    return -1
                if b["status"] == "submitted":
                    return 1
            if a["targetDate"] != b["targetDate"]:
                return -1 if _lt(a["targetDate"], b["targetDate"]) else 1
            if a["dept"] != b["dept"]:
                return -1 if _lt(a["dept"], b["dept"]) else 1
            return _cmp_str(a["workerName"], b["workerName"])

        overtime.sort(key=functools.cmp_to_key(cmp))
        holiday.sort(key=functools.cmp_to_key(cmp))
        return {"today": today, "overtime": overtime, "holiday": holiday, "isAdmin": admin,
                "selectedDepts": dept_list}
    except Exception as exc:  # noqa: BLE001 - H2 の癖のまま（isAdmin:true）
        return {"today": "", "overtime": [], "holiday": [], "isAdmin": True, "error": str(exc)}


HANDLERS = {
    "api_getDashboard": api_getDashboard,
    "api_getApproverDashboard": api_getApproverDashboard,
}
