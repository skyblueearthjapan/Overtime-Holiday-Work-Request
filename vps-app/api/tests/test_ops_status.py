"""管理者の運用欄（services/ops_status.get_ops_status。実装仕様 3-10・4-2）。"""
from __future__ import annotations

import json
from datetime import timedelta

import pytest
from gas_fakes import Clock, FakeGas, add_ping, add_request, drain, sync_all

from app import models, sheet_queue, timeutil
from app.services import gas_jobs, morning_batch, ops_status


@pytest.fixture
def clock(monkeypatch):
    c = Clock()
    monkeypatch.setattr(timeutil, "now_utc", c)
    return c


@pytest.fixture
def gas():
    fake = FakeGas()
    gas_jobs.set_client_factory(lambda settings: fake)
    yield fake
    gas_jobs.set_client_factory(None)


def test_empty_status_is_json(db, clock):
    status = ops_status.get_ops_status(db)
    json.dumps(status)
    assert status["rowSync"] == {"pending": 0, "oldestPendingSeconds": 0, "delayed": 0, "blocked": []}
    assert status["gasJobs"]["counts"]["failed"] == 0
    assert status["gasPing"]["ok"] is False and status["gasPing"]["lanePaused"] is True
    assert status["lastRuns"]["morning_batch"] is None


def test_status_shows_failures_blocked_and_delays(session_factory, db, clock, gas):
    gas_jobs.register_worker()
    morning_batch.register_worker()
    add_ping(db, clock())
    add_request(db, "R1")
    add_request(db, "R2")
    with session_factory() as s:
        a = gas_jobs.enqueue_regenerate(s, "R1", reason="edit")
        b = gas_jobs.enqueue_ledger_upsert(s, "R1")
        s.commit()
    gas.responses["internal.regenerate"] = {"ok": False, "error": "ledger_header_mismatch",
                                            "message": "見出しが違います", "retryable": False}
    drain(session_factory, clock, max_steps=5)

    # 行の反映: R2 を 6 分前から待たせ、別の行を blocked に
    with session_factory() as s:
        s.get(models.Request, "R2").approved_minutes = 30
        sheet_queue.mark_dirty(s, "requests", "R2")
        st = s.get(models.SheetSyncState, ("requests", "R2"))
        st.first_dirty_at = clock() - timedelta(minutes=6)
        st1 = s.get(models.SheetSyncState, ("requests", "R1"))
        st1.state, st1.last_error = "blocked", "requestId がシートに見つかりません"
        s.add(models.SyncRun(kind="sheet_reconcile", started_at=clock(), finished_at=clock(), ok=True,
                             summary_json={"fixedRows": [{"row": 5}, {"row": 9}]}))
        s.add(models.Setting(key="PDF_MODE", value="direct で直接書込方式", sheet_row=4))
        s.commit()

    status = ops_status.get_ops_status(db)
    json.dumps(status, ensure_ascii=False)
    gj = status["gasJobs"]
    assert gj["counts"]["failed"] == 1
    assert gj["failed"] == [{"id": a, "requestId": "R1", "action": "internal.regenerate",
                             "reason": "PermanentJobError: ledger_header_mismatch: 見出しが違います", "attempts": 1,
                             "state": "failed", "createdAt": gj["failed"][0]["createdAt"],
                             "runAfter": gj["failed"][0]["runAfter"]}]
    assert [j["id"] for j in gj["blocked"]] == [b]
    assert status["rowSync"]["pending"] == 1 and status["rowSync"]["delayed"] == 1
    assert status["rowSync"]["oldestPendingSeconds"] == 360
    assert status["rowSync"]["blocked"][0]["pk"] == "R1"
    assert status["gasPing"]["ok"] is True and status["gasPing"]["fresh"] is True
    assert status["reconcileFixedRows"] == 2
    assert status["exhausted"][0]["kind"] == "internal.regenerate"
    assert status["pdfMode"] == "direct で直接書込方式"
    assert status["now"] == timeutil.to_jst_text(clock())


def test_status_shows_waiting_reason_and_last_batch(session_factory, db, clock, gas):
    gas_jobs.register_worker()
    morning_batch.register_worker()
    add_request(db, "R1")  # gas_ping が無いので待つ
    with session_factory() as s:
        gas_jobs.enqueue_ledger_upsert(s, "R1")
        s.commit()
    drain(session_factory, clock, max_steps=1)
    status = ops_status.get_ops_status(db)
    assert "gas_ping" in status["gasJobs"]["waiting"][0]["reason"]

    add_ping(db, clock())
    sync_all(db)
    morning_batch.run_morning_batch(session_factory)
    drain(session_factory, clock, sync=lambda s: add_ping(s, clock()))
    last = ops_status.get_ops_status(db)["lastRuns"]["morning_batch"]
    assert last["ok"] is True and last["summary"]["counts"] == {"ok": 1, "skip": 0, "fail": 0}
