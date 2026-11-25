"""二次承認（4-3 の 19・20・21。approval.gs:593-883）。"""
from __future__ import annotations

from datetime import date

import pytest

from app.domain import secondary
from app.rpc import RpcError, RpcForbidden
from test_domain_helpers import ADMIN, USER, add_request, ctx, dirty, jobs_of, jst

NOW = jst(2026, 11, 26, 9, 15, 30)


def test_list_filters_formats_and_sorts(db):
    a = add_request(db, status="approved", target=date(2026, 11, 26), dept="B", worker_name="b",
                    submitted_at=jst(2026, 11, 26, 7, 5, 9), approved_at=jst(2026, 11, 26, 8, 1, 2), approved_by="x")
    over = add_request(db, status="submitted", target=date(2026, 11, 20), submitted_at=jst(2026, 11, 22, 6))
    add_request(db, status="approved", target=date(2026, 11, 27))  # 未来は出さない
    add_request(db, status="canceled", target=date(2026, 11, 26))
    add_request(db, status="approved", target=date(2026, 11, 26), approved_by2=ADMIN)  # 二次承認済みは出さない
    ready = add_request(db, status="approved", target=date(2026, 11, 26), dept="A", worker_name="a")
    sub = add_request(db, status="submitted", target=date(2026, 11, 26), dept="A", worker_name="a")
    res = secondary.api_getSecondaryApprovalList(ctx(db, email=ADMIN, now=NOW), "2026-11-26")
    assert res["date"] == "2026-11-26"
    assert [i["requestId"] for i in res["items"]] == [over.request_id, ready.request_id, a.request_id,
                                                      sub.request_id]
    item = next(i for i in res["items"] if i["requestId"] == a.request_id)
    assert item == {
        "requestId": a.request_id, "requestType": "overtime", "status": "approved", "dept": "B",
        "workerName": "b", "targetDate": "2026-11-26", "isOverdue": False, "approvedMinutes": 60,
        "submittedAt": "2026-11-26 07:05", "approvedBy": "x", "approvedAt": "2026-11-26 08:01", "approvedBy2": "",
        "approvedAt2": "", "pdfFileId": "", "isRetroactive": False, "daysAgo": 0,
    }
    o = res["items"][0]
    assert o["isOverdue"] is True and o["isRetroactive"] is True and o["daysAgo"] == 6


def test_list_days_ago_is_from_real_today(db):
    add_request(db, status="approved", target=date(2026, 11, 20))
    res = secondary.api_getSecondaryApprovalList(ctx(db, email=ADMIN, now=NOW), "2026-11-21")
    assert res["items"][0]["daysAgo"] == 6 and res["items"][0]["isOverdue"] is True


def test_list_bad_date(db):
    with pytest.raises(RpcError) as e:
        secondary.api_getSecondaryApprovalList(ctx(db, email=ADMIN), "x")
    assert e.value.message == "日付が不正です（yyyy-MM-dd）"
    with pytest.raises(RpcError):
        secondary.api_getSecondaryApprovalList(ctx(db, email=ADMIN))


def test_approve_with_pdf_regenerates(db):
    req = add_request(db, status="approved", pdf_file_id="pdf-9")
    res = secondary.api_secondaryApprove(ctx(db, email=ADMIN, now=NOW), req.request_id)
    assert res == {"ok": True, "requestId": req.request_id, "approvedBy2": ADMIN,
                   "approvedAt2": "2026-11-26 09:15:30", "pdfRegenerated": False}
    db.refresh(req)
    assert req.approved_by2 == ADMIN and req.approved_at2 == NOW and req.pdf_file_id == ""
    assert req.old_pdf_file_id_pending == "pdf-9" and ("requests", req.request_id) in dirty(db)
    assert [(j["kind"], j["args"]) for j in jobs_of(db)] == [
        ("internal.regenerate", {"reason": "secondary", "oldPdfFileId": "pdf-9"})]
    assert secondary.api_secondaryApprove(ctx(db, email=ADMIN), req.request_id) == {
        "ok": True, "message": "既に二次承認済みです。"}
    assert len(jobs_of(db)) == 1


def test_approve_without_pdf_upserts_ledger(db):
    req = add_request(db, status="approved")
    secondary.api_secondaryApprove(ctx(db, email=ADMIN), req.request_id)
    assert [j["kind"] for j in jobs_of(db)] == ["internal.ledgerUpsert"]


def test_approve_errors(db):
    with pytest.raises(RpcError) as e:
        secondary.api_secondaryApprove(ctx(db, email=ADMIN), "x")
    assert e.value.message == "Requestsが空です。"
    req = add_request(db, status="submitted", approved_by2=ADMIN)
    with pytest.raises(RpcError) as e:
        secondary.api_secondaryApprove(ctx(db, email=ADMIN), req.request_id)  # 状態が先、二次承認済みが後
    assert e.value.message == "一次承認が完了していません（現在: submitted）。"
    with pytest.raises(RpcError) as e:
        secondary.api_secondaryApprove(ctx(db, email=ADMIN), "x")
    assert e.value.message == "requestIdが見つかりません。"
    with pytest.raises(RpcForbidden):
        secondary.api_secondaryApprove(ctx(db, email=USER), req.request_id)


def test_batch(db):
    pdfs = [add_request(db, status="approved", pdf_file_id=f"p{i}") for i in range(6)]
    plain = add_request(db, status="approved")
    done = add_request(db, status="approved", approved_by2="z@x")
    sub = add_request(db, status="submitted")
    ids = [sub.request_id, done.request_id, plain.request_id, "nope"] + [p.request_id for p in pdfs]
    res = secondary.api_secondaryApproveBatch(ctx(db, email=ADMIN, now=NOW), ids)
    stamp = "2026-11-26 09:15:30"
    expected = [{"requestId": p.request_id, "ok": True, "approvedBy2": ADMIN, "approvedAt2": stamp} for p in pdfs]
    expected += [{"requestId": plain.request_id, "ok": True, "approvedBy2": ADMIN, "approvedAt2": stamp},
                 {"requestId": done.request_id, "ok": True, "message": "既に二次承認済み"},
                 {"requestId": sub.request_id, "ok": False, "error": "一次承認未完了"},
                 {"requestId": "nope", "ok": False, "error": "見つかりません"}]
    assert res == {"ok": True, "results": expected}
    jobs = jobs_of(db)
    regen = [j for j in jobs if j["kind"] == "internal.regenerate"]
    assert [j["requestId"] for j in regen] == [p.request_id for p in pdfs]
    assert len({j["batchId"] for j in regen}) == 1 and regen[0]["batchId"]
    assert [(j["kind"], j["requestId"]) for j in jobs if j["kind"] != "internal.regenerate"] == [
        ("internal.ledgerUpsert", plain.request_id)]


def test_batch_count_is_checked_before_admin(db):
    assert secondary.api_secondaryApproveBatch(ctx(db, email=USER), []) == {"ok": True, "results": []}
    with pytest.raises(RpcError) as e:
        secondary.api_secondaryApproveBatch(ctx(db, email=USER), [str(i) for i in range(51)])
    assert e.value.message == "一括承認は50件までです。" and not isinstance(e.value, RpcForbidden)
    with pytest.raises(RpcForbidden):
        secondary.api_secondaryApproveBatch(ctx(db, email=USER), ["a"])
