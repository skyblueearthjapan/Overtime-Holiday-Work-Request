"""管理者だけの関数（実装仕様 4-2）。M・D のモジュールは偽物を差し込んで確かめる。"""
from __future__ import annotations

import sys
import types

import pytest

from app import jobs, rpc
from app.auth import create_session
from app.domain import ops, register
from app.models import Job, UserSession
from app.rpc import RpcError
from test_domain_helpers import ADMIN, USER, add_request, add_wl, ctx, dirty, jst

NOW = jst(2026, 11, 27, 9)


@pytest.fixture
def fake_services(monkeypatch):
    calls: list = []
    gas_jobs = types.ModuleType("app.services.gas_jobs")
    gas_jobs.enqueue_regenerate = lambda db, rid, **kw: calls.append(("regenerate", rid, kw)) or 11
    gas_jobs.enqueue_ledger_rebuild = lambda db: calls.append(("rebuild",)) or 12
    morning = types.ModuleType("app.services.morning_batch")
    morning.enqueue_manual_batch = lambda db, ymd: calls.append(("batch", ymd)) or 13
    status = types.ModuleType("app.services.ops_status")
    status.get_ops_status = lambda db: {"rowSync": {"pending": 0}}
    config_sync = types.ModuleType("app.services.config_sync")
    config_sync.sync_sheet_config = lambda db: {"ok": True, "counts": {}, "overlaps": [], "mismatches": []}
    master_sync = types.ModuleType("app.services.master_sync")
    master_sync.sync_masters = lambda db: {"ok": True, "tabs": {}}
    for name, mod in (("gas_jobs", gas_jobs), ("morning_batch", morning), ("ops_status", status),
                      ("config_sync", config_sync), ("master_sync", master_sync)):
        monkeypatch.setitem(sys.modules, f"app.services.{name}", mod)
    return calls


def test_ops_call_owner_entry_points(db, fake_services):
    c = ctx(db, email=ADMIN, now=NOW)
    assert ops.ops_getStatus(c) == {"rowSync": {"pending": 0}}
    assert ops.ops_syncSheetConfigNow(c)["ok"] is True
    assert ops.ops_runMasterSync(c) == {"ok": True, "tabs": {}}
    assert ops.ops_batchPdfForDate(c, "2026-11-25") == {"ok": True, "jobId": 13}
    assert ops.ops_rebuildLedger(c) == {"ok": True, "jobId": 12}
    req = add_request(db, status="approved", pdf_file_id="old-pdf")
    assert ops.ops_regeneratePdf(c, req.request_id) == {"ok": True, "jobId": 11}
    db.refresh(req)
    assert req.pdf_file_id == "" and req.old_pdf_file_id_pending == "old-pdf"
    assert ("requests", req.request_id) in dirty(db)
    assert ("regenerate", req.request_id, {"reason": "admin", "old_pdf_file_id": "old-pdf"}) in fake_services
    with pytest.raises(RpcError) as e:
        ops.ops_regeneratePdf(c, "nope")
    assert e.value.message == "申請が見つかりません。"
    with pytest.raises(RpcError) as e:
        ops.ops_batchPdfForDate(c, "2026/11/25")
    assert e.value.message == "日付が不正です（yyyy-MM-dd）"


def test_ops_not_ready_without_modules(db, monkeypatch):
    for name in ("gas_jobs", "morning_batch", "ops_status", "config_sync", "master_sync"):
        monkeypatch.setitem(sys.modules, f"app.services.{name}", None)  # import すると ModuleNotFoundError
    c = ctx(db, email=ADMIN)
    for fn, args in ((ops.ops_getStatus, ()), (ops.ops_syncSheetConfigNow, ()), (ops.ops_runMasterSync, ()),
                     (ops.ops_rebuildLedger, ()), (ops.ops_batchPdfForDate, ("2026-11-25",)),
                     (ops.ops_regeneratePdf, ("x",))):
        with pytest.raises(RpcError) as e:
            fn(c, *args)
        assert e.value.message.startswith("この機能はまだ用意されていません。（services.")


def test_retry_job(db):
    jid = jobs.enqueue(db, "internal.ledgerUpsert", {}, request_id="r1")
    job = db.get(Job, jid)
    job.state = "failed"
    db.commit()
    assert ops.ops_retryJob(ctx(db, email=ADMIN), jid) == {"ok": True}
    assert db.get(Job, jid).state == "pending"
    with pytest.raises(RpcError) as e:
        ops.ops_retryJob(ctx(db, email=ADMIN), jid)
    assert e.value.message == "失敗した仕事が見つかりません。"
    with pytest.raises(RpcError):
        ops.ops_retryJob(ctx(db, email=ADMIN), "abc")


def test_revoke_sessions(db, settings):
    from fastapi.responses import Response

    create_session(db, Response(), "victim@lineworks-local.info", settings)
    create_session(db, Response(), "victim@lineworks-local.info", settings)
    db.commit()
    assert ops.ops_revokeSessions(ctx(db, email=ADMIN), "Victim@lineworks-local.info") == {"ok": True, "revoked": 2}
    assert db.query(UserSession).count() == 0
    with pytest.raises(RpcError):
        ops.ops_revokeSessions(ctx(db, email=ADMIN), "")


def test_recalc_breaks_only_changed_rows(db):
    ok = add_wl(db, "r1", start="2026-11-26 17:20:00", end="2026-11-26 19:30:00", actual=130, brk=10, net=120)
    wrong = add_wl(db, "r2", start="2026-11-26 17:20:00", end="2026-11-26 19:21:00", actual=121, brk=1, net=120)
    zero_actual = add_wl(db, "r3", start="2026-11-26 08:00:00", end="2026-11-26 12:00:00", actual=None, brk=None,
                         net=None)
    backwards = add_wl(db, "r4", start="2026-11-26 19:00:00", end="2026-11-26 18:00:00", actual=0, brk=5, net=0)
    add_wl(db, "r5", start="", end="2026-11-26 18:00:00", brk=3)
    res = ops.ops_recalcBreaks(ctx(db, email=ADMIN, now=NOW))
    assert res == {"ok": True, "changed": 2}
    for wl in (ok, wrong, zero_actual, backwards):
        db.refresh(wl)
    assert (wrong.break_minutes, wrong.net_minutes, wrong.updated_by) == (0, 121, "break-recalc")
    assert wrong.updated_at == "2026-11-27 09:00:00" and wrong.cell_kinds == {"updatedAt": "string"}
    assert (zero_actual.break_minutes, zero_actual.net_minutes, zero_actual.actual_minutes) == (10, 230, None)
    assert ok.updated_by == "formSubmit" and backwards.break_minutes == 5
    assert dirty(db) == {("work_logs", str(wrong.id)), ("work_logs", str(zero_actual.id))}


def test_ops_are_admin_only_over_http(client, db, login_as, fake_services):
    register.register_all(rpc.register)
    login_as(USER)
    resp = client.post("/api/rpc/ops_getStatus", json=[])
    assert resp.status_code == 403 and resp.json()["error"] == "総務部（管理者）権限がありません。"
    login_as(ADMIN)
    resp = client.post("/api/rpc/ops_getStatus", json=[])
    assert resp.status_code == 200 and resp.json() == {"ok": True, "result": {"rowSync": {"pending": 0}}}
