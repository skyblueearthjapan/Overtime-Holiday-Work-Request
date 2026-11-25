"""TOP と承認者画面の一覧（4-3 の 6・14。worklog.gs:404-533・approval.gs:132-276、01 §2.7）。"""
from __future__ import annotations

from datetime import date

from app.domain import dashboards
from test_domain_helpers import ADMIN, APPROVER, USER, add_request, add_wl, ctx, jst, set_dept_approvers

NOW = jst(2026, 11, 26, 10)  # 木曜


def ids(items):
    return [i["requestId"] for i in items]


def test_top_dashboard_conditions(db):
    today_ot = add_request(db, target=date(2026, 11, 26), dept="B", worker_name="b")
    today_ot2 = add_request(db, target=date(2026, 11, 26), dept="A", worker_name="z", status="approved")
    add_request(db, target=date(2026, 11, 25))  # 過去の submitted の残業は出ない（Q5）
    overdue = add_request(db, target=date(2026, 10, 27), status="approved")  # 30 日前ちょうど → 期限超過
    add_request(db, target=date(2026, 10, 26), status="approved")  # 31 日前 → 出ない
    add_request(db, target=date(2026, 11, 26), status="canceled")
    done_2nd = add_request(db, target=date(2026, 11, 25), status="approved", approved_by2=ADMIN)
    add_wl(db, done_2nd.request_id, end="2026-11-25 19:00:00")
    done_2nd_today = add_request(db, target=date(2026, 11, 26), status="approved", approved_by2=ADMIN)
    add_wl(db, done_2nd_today.request_id, start="2026-11-26 17:20:00", end="2026-11-26 19:00:00",
           actual=100, brk=0, net=100)
    hol_future = add_request(db, request_type="holiday", target=date(2027, 1, 9))
    hol_old_appr = add_request(db, request_type="holiday", status="approved", target=date(2026, 1, 10))
    hol_overdue = add_request(db, request_type="holiday", status="approved", target=date(2026, 11, 22))
    add_wl(db, hol_overdue.request_id, start="2026-11-22 08:00:00")

    res = dashboards.api_getDashboard(ctx(db, now=NOW))
    assert res["today"] == "2026-11-26" and "error" not in res
    # 部署 → 氏名（JS の <）。同じなら元の行の順（安定）
    assert ids(res["overtime"]) == [today_ot2.request_id, today_ot.request_id, overdue.request_id,
                                    done_2nd_today.request_id]
    assert ids(res["holiday"]) == [hol_old_appr.request_id, hol_overdue.request_id, hol_future.request_id]
    by = {i["requestId"]: i for i in res["overtime"] + res["holiday"]}
    assert by[overdue.request_id]["isOverdue"] is True and by[overdue.request_id]["targetDateLabel"] == "10/27(火)"
    assert by[today_ot.request_id]["targetDateLabel"] == ""
    assert by[hol_future.request_id]["targetDateLabel"] == "1/9(土)"
    assert by[hol_overdue.request_id]["isOverdue"] is True
    item = by[done_2nd_today.request_id]
    assert item == {
        "requestId": done_2nd_today.request_id, "requestType": "overtime", "status": "approved", "dept": "製造1課",
        "workerCode": "A001", "workerName": "山田太郎", "targetDate": "2026-11-26", "targetDateLabel": "",
        "approvedMinutes": 60, "actualStartAt": "2026-11-26 17:20:00", "actualEndAt": "2026-11-26 19:00:00",
        "actualMinutes": 100, "breakMinutes": 0, "netMinutes": 100, "pdfFileId": "", "approvedBy2": ADMIN,
        "isOverdue": False, "reason": "急ぎ作業を要するため",
    }


def test_top_dashboard_skips_unparseable_target_and_reports_errors(db, monkeypatch):
    add_request(db, target=None, target_text="不明")
    assert dashboards.api_getDashboard(ctx(db, now=NOW))["overtime"] == []

    def boom(db):
        raise RuntimeError("WorkLogsに requestId 列がありません。")

    monkeypatch.setattr(dashboards.store, "worklog_map", boom)
    assert dashboards.api_getDashboard(ctx(db, now=NOW)) == {
        "today": "", "overtime": [], "holiday": [], "error": "WorkLogsに requestId 列がありません。"}


def test_approver_dashboard_conditions(db):
    set_dept_approvers(db, [("製造1課", APPROVER, "上司")])
    old_sub = add_request(db, target=date(2025, 12, 1), submitted_at=jst(2025, 12, 3, 8))
    today_sub = add_request(db, target=date(2026, 11, 26), worker_name="b")
    add_request(db, target=date(2026, 11, 27))  # 未来の残業 submitted は出ない
    today_cancel = add_request(db, target=date(2026, 11, 26), status="canceled", worker_name="a")  # Q4
    add_request(db, target=date(2026, 11, 25), status="approved")
    hol_future = add_request(db, request_type="holiday", target=date(2026, 12, 5))
    add_request(db, request_type="holiday", target=date(2026, 12, 5), status="approved")
    add_request(db, dept="製造2課", target=date(2026, 11, 26))

    res = dashboards.api_getApproverDashboard(ctx(db, email=APPROVER, now=NOW), ["製造1課"])
    assert res["isAdmin"] is False and res["selectedDepts"] == ["製造1課"] and res["today"] == "2026-11-26"
    assert ids(res["overtime"]) == [old_sub.request_id, today_sub.request_id, today_cancel.request_id]
    first = res["overtime"][0]
    assert first == {
        "requestId": old_sub.request_id, "requestType": "overtime", "status": "submitted", "dept": "製造1課",
        "workerName": "山田太郎", "targetDate": "2025-12-01", "targetDateLabel": "2025/12/1(月)",
        "approvedMinutes": 60, "submittedAt": "2025-12-03 08:00:00", "isOverdue": True, "isRetroactive": True,
        "daysAgo": 360,
    }
    assert ids(res["holiday"]) == [hol_future.request_id]
    assert res["holiday"][0]["isOverdue"] is False and res["holiday"][0]["targetDateLabel"] == "12/5(土)"


def test_approver_dashboard_errors_and_string_dept(db):
    set_dept_approvers(db, [("製造1課", APPROVER, "上司")])
    assert dashboards.api_getApproverDashboard(ctx(db, email=APPROVER), []) == {
        "today": "", "overtime": [], "holiday": [], "isAdmin": False, "error": "部署が指定されていません。"}
    assert dashboards.api_getApproverDashboard(ctx(db, email=ADMIN))["isAdmin"] is True
    assert dashboards.api_getApproverDashboard(ctx(db, email=APPROVER), ["製造1課", "製造2課"]) == {
        "today": "", "overtime": [], "holiday": [], "isAdmin": False, "error": "部署「製造2課」の承認権限がありません。"}
    assert dashboards.api_getApproverDashboard(ctx(db, email=USER), "製造1課")["error"] == "部署「製造1課」の承認権限がありません。"
    res = dashboards.api_getApproverDashboard(ctx(db, email=ADMIN), " 製造９課 ")
    assert res["isAdmin"] is True and res["selectedDepts"] == ["製造9課"] and "error" not in res
