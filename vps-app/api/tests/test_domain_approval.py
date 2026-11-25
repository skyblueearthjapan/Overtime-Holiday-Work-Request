"""一次承認（4-3 の 13・15・16・17・18。approval.gs:280-551）。"""
from __future__ import annotations

from datetime import date

import pytest

from app.domain import approval
from app.rpc import RpcError
from test_domain_helpers import (
    ADMIN,
    APPROVER,
    USER,
    add_request,
    add_wl,
    ctx,
    dirty,
    jobs_of,
    jst,
    set_dept_approvers,
)

NOW = jst(2026, 11, 26, 12, 34, 56)


@pytest.fixture
def approvers(db):
    set_dept_approvers(db, [("製造1課", "Boss@lineworks-local.info, other@lineworks-local.info", "上司"),
                            ("製造2課", "x@lineworks-local.info", "")])


def approve(db, rid, email=APPROVER):
    return approval.api_approveRequest(ctx(db, email=email, now=NOW), rid)


def test_approve_without_worklog_end_does_not_enqueue(db, approvers):
    req = add_request(db)
    res = approve(db, req.request_id)
    assert res == {"ok": True, "requestId": req.request_id, "approvedBy": APPROVER,
                   "approvedAt": "2026-11-26 12:34:56"}
    db.refresh(req)
    assert req.status == "approved" and req.approved_by == APPROVER and req.approved_at == NOW
    assert ("requests", req.request_id) in dirty(db)
    assert jobs_of(db) == []


def test_approve_completed_enqueues_after_approve_once_even_from_two_tabs(db, approvers):
    req = add_request(db)
    add_wl(db, req.request_id, start="2026-11-26 17:20:00", end="2026-11-26 19:30:00", actual=130, brk=10, net=120)
    approve(db, req.request_id)
    assert approve(db, req.request_id) == {"ok": True, "message": "既に承認済みです。"}
    jobs = jobs_of(db)
    assert [(j["kind"], j["requestId"]) for j in jobs] == [("internal.afterApprove", req.request_id)]
    assert jobs[0]["guard"] == {"status": "approved"}


def test_approve_uses_last_worklog_row_for_done_check(db, approvers):
    req = add_request(db)
    add_wl(db, req.request_id, end="2026-11-26 19:00:00")
    add_wl(db, req.request_id, end="")  # 下の行が勝つ（buildWorkLogsMapByRequestId_）
    approve(db, req.request_id)
    assert jobs_of(db) == []


def test_admin_can_approve_any_dept_and_self_approval_allowed(db, approvers):
    req = add_request(db, dept="知らない部署", worker_email=ADMIN)
    assert approve(db, req.request_id, email=ADMIN)["approvedBy"] == ADMIN


def test_approve_errors(db, approvers):
    with pytest.raises(RpcError) as e:
        approve(db, "x")
    assert e.value.message == "Requestsが空です。"
    req = add_request(db, dept="製造2課")
    with pytest.raises(RpcError) as e:
        approve(db, "x")
    assert e.value.message == "requestIdが見つかりません。"
    with pytest.raises(RpcError) as e:
        approve(db, req.request_id)
    assert e.value.message == "この申請の承認権限がありません。"
    canceled = add_request(db, status="canceled")
    with pytest.raises(RpcError) as e:
        approve(db, canceled.request_id)
    assert e.value.message == "キャンセル済みです。"


def test_permission_is_checked_before_already_approved(db, approvers):
    req = add_request(db, dept="製造2課", status="approved")
    with pytest.raises(RpcError) as e:
        approve(db, req.request_id)
    assert e.value.message == "この申請の承認権限がありません。"


# ===== 却下 =====


def test_reject(db, approvers):
    req = add_request(db, pdf_file_id="")
    res = approval.api_rejectRequest(ctx(db, email=APPROVER, now=NOW), req.request_id)
    assert res == {"ok": True, "requestId": req.request_id}
    db.refresh(req)
    assert req.status == "canceled" and req.approved_by == APPROVER + "(却下)" and req.approved_at == NOW
    assert jobs_of(db) == []  # 台帳・PDF は触らない
    with pytest.raises(RpcError) as e:
        approval.api_rejectRequest(ctx(db, email=APPROVER), req.request_id)
    assert e.value.message == "未承認状態でのみ却下可能です（現在: canceled）。"
    ok = add_request(db, status="approved")
    with pytest.raises(RpcError) as e:
        approval.api_rejectRequest(ctx(db, email=APPROVER), ok.request_id)
    assert e.value.message == "未承認状態でのみ却下可能です（現在: approved）。"
    other = add_request(db, dept="製造2課")
    with pytest.raises(RpcError) as e:
        approval.api_rejectRequest(ctx(db, email=APPROVER), other.request_id)
    assert e.value.message == "この申請の承認権限がありません。"


# ===== 一括 =====


def test_batch_results_in_sheet_order_and_jobs(db, approvers):
    a = add_request(db)
    b = add_request(db, status="approved")
    c = add_request(db, status="canceled")
    d = add_request(db, dept="製造2課")
    e = add_request(db)
    add_wl(db, e.request_id, end="2026-11-26 19:00:00")
    ids = [e.request_id, "missing-1", d.request_id, c.request_id, b.request_id, a.request_id, a.request_id]
    res = approval.api_approveRequestsBatch(ctx(db, email=APPROVER, now=NOW), ids)
    stamp = "2026-11-26 12:34:56"
    assert res == {"ok": True, "results": [
        {"requestId": a.request_id, "ok": True, "approvedBy": APPROVER, "approvedAt": stamp},
        {"requestId": b.request_id, "ok": True, "message": "既に承認済み"},
        {"requestId": c.request_id, "ok": False, "error": "キャンセル済み"},
        {"requestId": d.request_id, "ok": False, "error": "承認権限なし"},
        {"requestId": e.request_id, "ok": True, "approvedBy": APPROVER, "approvedAt": stamp},
        {"requestId": "missing-1", "ok": False, "error": "見つかりません"},
    ]}
    assert [(j["kind"], j["requestId"]) for j in jobs_of(db)] == [("internal.afterApprove", e.request_id)]


def test_batch_limits(db, approvers):
    assert approval.api_approveRequestsBatch(ctx(db, email=APPROVER), []) == {"ok": True, "results": []}
    assert approval.api_approveRequestsBatch(ctx(db, email=APPROVER), None) == {"ok": True, "results": []}
    with pytest.raises(RpcError) as e:
        approval.api_approveRequestsBatch(ctx(db, email=APPROVER), [f"id{i}" for i in range(51)])
    assert e.value.message == "一括承認は50件までです。"
    assert approval.api_approveRequestsBatch(ctx(db, email=APPROVER), ["a"]) == {"ok": True, "results": []}


# ===== 承認者の部署・月次 =====


def test_get_approver_depts(db, approvers):
    res = approval.api_getApproverDepts(ctx(db, email=ADMIN))
    assert res["isAdmin"] is True and res["currentEmail"] == ADMIN
    assert {p["email"] for p in res["approvers"]} == {APPROVER, "other@lineworks-local.info", "x@lineworks-local.info"}
    res = approval.api_getApproverDepts(ctx(db, email=APPROVER))
    assert res == {"isAdmin": False, "currentEmail": APPROVER,
                   "approvers": [{"email": APPROVER, "name": "上司", "depts": ["製造1課"]}]}
    with pytest.raises(RpcError) as e:
        approval.api_getApproverDepts(ctx(db, email=USER))
    assert e.value.message == "承認者権限がありません。DeptApprovers または ApproverMap に登録してください。"


def test_approver_monthly_summary_errors_and_multi_dept(db, approvers):
    with pytest.raises(RpcError) as e:
        approval.api_approverMonthlySummary(ctx(db, email=APPROVER), "2026-11", [])
    assert e.value.message == "部署が指定されていません。"
    with pytest.raises(RpcError) as e:
        approval.api_approverMonthlySummary(ctx(db, email=APPROVER), "2026-11", "")
    assert e.value.message == "部署が指定されていません。"
    with pytest.raises(RpcError) as e:
        approval.api_approverMonthlySummary(ctx(db, email=APPROVER), "2026-11", ["製造1課", "製造2課"])
    assert e.value.message == "部署「製造2課」の月次サマリーを閲覧する権限がありません。"

    for dept, name, net in (("製造1課", "A", 100), ("製造2課", "B", 200), ("製造3課", "C", 300)):
        r = add_request(db, dept=dept, worker_name=name, target=date(2026, 11, 2))
        add_wl(db, r.request_id, net=net)
    one = approval.api_approverMonthlySummary(ctx(db, email=ADMIN, now=NOW), "2026-11", "製造1課")
    assert one["deptFilter"] == "製造1課" and [p["workerName"] for p in one["people"]] == ["A"]
    multi = approval.api_approverMonthlySummary(ctx(db, email=ADMIN, now=NOW), "2026-11", ["製造1課", " 製造2課 "])
    assert multi["deptFilter"] == "製造1課, 製造2課"
    assert [p["workerName"] for p in multi["people"]] == ["B", "A"]
    assert multi["kpi"]["totalPeople"] == 2 and multi["kpi"]["totalNetMinutes"] == 300
    assert multi["charts"]["people"]["labels"] == ["B", "A"]
    assert multi["charts"]["dept"]["labels"] == ["製造1課", "製造2課", "製造3課"]  # H14: 作り直さない
