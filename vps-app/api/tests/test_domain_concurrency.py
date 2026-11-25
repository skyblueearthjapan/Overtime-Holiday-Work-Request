"""行ロックと二重押し（実装仕様 5-5・9-1 #5、gap-1 G3、罠 17・18）。

- 承認と完了がどちらの順で起きても、PDF を作る仕事（afterApprove か、statusAtEnqueue=approved の afterComplete）が
  ちょうど 1 つ積まれる。SQLite では順番の 2 通りを確かめ、PostgreSQL（postgres マーカー）では本当に同時に走らせる。
- 一括の処理は ID を並べ替えてからロックする。
- 2 つのタブで同じ申請を承認しても仕事は 1 つ。同じ冪等キーの二重送信で申請は 1 件。
"""
from __future__ import annotations

import os
import threading
from datetime import date

import pytest
from sqlalchemy import func, select

from app import rpc
from app.auth import user_from_email
from app.domain import approval, register, store, worklog
from app.models import Request, WorkLog
from app.rpc import RpcContext
from test_domain_helpers import ADMIN, APPROVER, USER, add_request, ctx, jobs_of, jst, set_dept_approvers

NOW = jst(2026, 11, 26, 19, 30)


def pdf_jobs(db):
    return [j for j in jobs_of(db) if j["kind"] == "internal.afterApprove"
            or (j["kind"] == "internal.afterComplete" and j["args"]["statusAtEnqueue"] == "approved")]


@pytest.mark.parametrize("order", ["approve_then_complete", "complete_then_approve"])
def test_approve_and_complete_always_make_one_pdf_job(db, order):
    set_dept_approvers(db, [("製造1課", APPROVER, "上司")])
    req = add_request(db, target=date(2026, 11, 26))
    steps = {
        "approve": lambda: approval.api_approveRequest(ctx(db, email=APPROVER, now=NOW), req.request_id),
        "complete": lambda: worklog.api_markOvertimeDone(ctx(db, now=NOW), req.request_id),
    }
    for name in order.split("_then_"):
        steps[name]()
        db.commit()
    assert len(pdf_jobs(db)) == 1


def test_batch_locks_ids_in_sorted_order(db, monkeypatch):
    ids = [add_request(db, request_id=f"0000000{i}-0000-4000-8000-000000000000").request_id for i in (3, 1, 2)]
    seen: list[str] = []
    real = store.lock_request

    def spy(db_, rid):
        seen.append(rid)
        return real(db_, rid)

    monkeypatch.setattr(store, "lock_request", spy)
    approval.api_approveRequestsBatch(ctx(db, email=ADMIN), ids + ["zzz", "aaa"])
    assert seen == sorted(ids + ["zzz", "aaa"])


@pytest.fixture
def http(client, db, login_as):
    register.register_all(rpc.register)
    return client


def test_two_tabs_approve_once(http, db, login_as):
    set_dept_approvers(db, [("製造1課", APPROVER, "上司")])
    req = add_request(db)
    db.add(WorkLog(request_id=req.request_id, actual_end_at="2026-11-26 19:00:00"))
    db.commit()
    login_as(APPROVER)
    first = http.post("/api/rpc/api_approveRequest", json=[req.request_id],
                      headers={"X-Idempotency-Key": "tab-one-key-0001"})
    second = http.post("/api/rpc/api_approveRequest", json=[req.request_id],
                       headers={"X-Idempotency-Key": "tab-two-key-0002"})
    assert first.status_code == 200 and first.json()["result"]["approvedBy"] == APPROVER
    assert second.json() == {"ok": True, "result": {"ok": True, "message": "既に承認済みです。"}}
    db.expire_all()
    assert [j["kind"] for j in jobs_of(db)] == ["internal.afterApprove"]


def test_double_submit_same_key_makes_one_request(http, db, login_as):
    login_as(USER)
    body = [{"requestType": "overtime", "dept": "製造1課", "workerLabel": "A001 山田", "reason": "急ぎ作業を要するため",
             "hours": "1.0"}]
    headers = {"X-Idempotency-Key": "submit-key-000001"}
    a = http.post("/api/rpc/api_submitRequest", json=body, headers=headers)
    b = http.post("/api/rpc/api_submitRequest", json=body, headers=headers)
    assert a.status_code == b.status_code == 200 and a.json() == b.json()
    assert db.scalar(select(func.count()).select_from(Request)) == 1


def test_business_error_over_http_is_400_and_soft_error_is_200(http, db, login_as):
    login_as(USER)
    r = http.post("/api/rpc/api_submitRequest", json=[{"requestType": "x"}])
    assert r.status_code == 200 and r.json() == {"ok": True, "result": {"ok": False, "error": "申請種別が不正です: x"}}
    r = http.post("/api/rpc/api_deleteRequest", json=["nope"])
    assert r.status_code == 400 and r.json() == {"ok": False, "error": "申請が見つかりません。"}
    r = http.post("/api/rpc/api_markOvertimeDone", json=["nope"])  # 省略した引数は送らない
    assert r.status_code == 400 and r.json()["error"] == "申請が見つかりません。"


@pytest.mark.postgres
@pytest.mark.skipif(not os.environ.get("TEST_DATABASE_URL"), reason="PostgreSQL の結合試験だけ")
def test_concurrent_approve_and_complete_postgres(session_factory):
    with session_factory() as s:
        set_dept_approvers(s, [("製造1課", APPROVER, "上司")])
        rid = add_request(s, target=date(2026, 11, 26)).request_id
    barrier = threading.Barrier(2)
    errors: list[BaseException] = []

    def run(kind):
        with session_factory() as s:
            try:
                barrier.wait()
                if kind == "approve":
                    c = RpcContext(db=s, user=user_from_email(APPROVER), settings=ctx(s).settings, now=NOW)
                    approval.api_approveRequest(c, rid)
                else:
                    c = RpcContext(db=s, user=user_from_email(USER), settings=ctx(s).settings, now=NOW)
                    worklog.api_markOvertimeDone(c, rid)
                s.commit()
            except BaseException as exc:  # noqa: BLE001
                errors.append(exc)

    threads = [threading.Thread(target=run, args=(k,)) for k in ("approve", "complete")]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert not errors
    with session_factory() as s:
        assert len(pdf_jobs(s)) == 1
