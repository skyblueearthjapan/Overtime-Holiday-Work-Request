"""GAS の後処理の仕事（services/gas_jobs.py。実装仕様 7-3・7-4・6-5・12-2 の M）。

7-4 の表の行ごとに積む action と args・requires を待つ・canceled なら skipped・record_pdf_result で 3 列と
synced_values が同じになる・再試行の間隔と 24 時間・retryable:false で failed・同じ申請の順番・batch_id の 500ms・
gas_ping が合わないと止まる・直列（同時に 2 本動かない）・jobId の再生・Sebastian の待ち。
"""
from __future__ import annotations

import json
from datetime import timedelta

import pytest
from gas_fakes import (T0, Clock, FakeGas, add_ping, add_request, created_pdf, drain, good_ping, job, sync_all)

from app import jobs, models, sheet_queue, timeutil, worker
from app.jobs import JobRunner
from app.services import gas_jobs
from app.services.gas_client import ACTIONS, GasTransportError

WL_OVERRIDE = {"actualStartAt": "2026-11-25 17:20:00", "actualEndAt": "2026-11-25 19:00:00", "actualMinutes": 100,
               "breakMinutes": 10, "netMinutes": 90}


@pytest.fixture
def clock(monkeypatch):
    c = Clock()
    monkeypatch.setattr(timeutil, "now_utc", c)
    return c


@pytest.fixture
def gas(monkeypatch):
    fake = FakeGas()
    gas_jobs.set_client_factory(lambda settings: fake)
    sleeps: list[float] = []
    monkeypatch.setattr(gas_jobs, "_sleep", lambda s: sleeps.append(s))
    monkeypatch.setattr(gas_jobs, "_monotonic", lambda: 1000.0)
    gas_jobs._batch_last_call.clear()
    fake.sleeps = sleeps
    yield fake
    gas_jobs.set_client_factory(None)


@pytest.fixture
def env(session_factory, db, clock, gas):
    gas_jobs.register_worker()
    add_ping(db, clock())
    return session_factory


def run(sf, clock):
    """ping を新しく保ちながら流しきる。"""
    def fresh(s):
        add_ping(s, clock())
    return drain(sf, clock, sync=fresh)


def enq(sf, fn, *args, **kw):
    with sf() as s:
        job_id = fn(s, *args, **kw)
        s.commit()
    return job_id


# ===== 7-4 の表: 積む仕事と GAS へ送るもの =====


def test_after_approve_calls_gas_and_records_pdf(env, db, gas, clock):
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00")
    job_id = enq(env, gas_jobs.enqueue_after_approve, "R1")
    row = job(db, job_id)
    assert row.kind == "internal.afterApprove" and row.lane == "gas" and row.request_id == "R1"
    reqs = json.loads(row.requires_json)
    assert {r["table"] for r in reqs} == {"requests", "work_logs"}
    guard = json.loads(row.guard_json)
    assert guard["gasJobId"].startswith("ot-") and guard["statusAtEnqueue"] == "approved"

    assert run(env, clock) == ["done"]
    assert gas.actions() == ["internal.afterApprove"]
    call = gas.calls[-1]
    assert call["requestId"] == "R1" and call["expect"] == {"status": "approved"} and call["jobId"] == guard["gasJobId"]
    req = db.get(models.Request, "R1")
    db.refresh(req)
    assert req.pdf_file_id == "PDF-R1-1" and req.pdf_folder_id == "FOLDER-1"
    assert timeutil.to_jst_text(req.pdf_generated_at) == "2026-11-26 09:00:05"
    st = db.get(models.SheetSyncState, ("requests", "R1"))
    db.refresh(st)
    # GAS がシートに書いた値なので、行の反映は書き直さない（版は同じだけ進み、synced_values も同じ）
    assert st.synced_version == st.desired_version == req.row_version
    assert st.synced_values["pdfFileId"] == "PDF-R1-1" and st.synced_values["pdfFolderId"] == "FOLDER-1"
    assert st.synced_values["pdfGeneratedAt"] == req.pdf_generated_at.isoformat()
    assert job(db, job_id).state == "done"


def test_after_approve_dedupes_pending(env, db):
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00")
    a = enq(env, gas_jobs.enqueue_after_approve, "R1")
    b = enq(env, gas_jobs.enqueue_after_approve, "R1")
    assert a == b


@pytest.mark.parametrize("source,priority", [("screen", 0), ("admin", 0), ("sebastian", 10)])
def test_after_complete_args(env, db, gas, clock, source, priority):
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00")
    job_id = enq(env, gas_jobs.enqueue_after_complete, "R1", source=source, status_at_enqueue="approved",
                 wl_override=WL_OVERRIDE, priority=priority)
    assert job(db, job_id).priority == priority
    run(env, clock)
    call = gas.calls[-1]
    assert call["action"] == "internal.afterComplete"
    assert call["args"]["source"] == source and call["args"]["statusAtEnqueue"] == "approved"
    if source == "sebastian":
        # Sebastian はいつも XLOOKUP 方式・override なし（api.gs:105）
        assert "workLog" not in call["args"] and "wlOverride" not in call["args"]
        assert call["expect"] == {"status": "approved"}
    else:
        assert call["args"]["workLog"] == WL_OVERRIDE and call["args"]["wlOverride"] == WL_OVERRIDE
        assert call["expect"] == {"status": "approved", "actualEndAt": "2026-11-25 19:00:00"}


def test_after_complete_rejects_unknown_source(env, db):
    add_request(db, "R1")
    with pytest.raises(ValueError):
        gas_jobs.enqueue_after_complete(db, "R1", source="mail", status_at_enqueue="approved")


def test_secondary_regenerate_args_and_trash_clears_pending(env, db, gas, clock):
    add_request(db, "R1", approved_by2="boss@lineworks-local.info")
    with env() as s:
        r = s.get(models.Request, "R1")
        r.pdf_file_id, r.pdf_folder_id, r.pdf_generated_at = "", "", None  # W が 3 列を空にする
        r.old_pdf_file_id_pending = "OLD-PDF"
        sheet_queue.mark_dirty(s, "requests", "R1")
        gas_jobs.enqueue_regenerate(s, "R1", reason="secondary", old_pdf_file_id="OLD-PDF")
        s.commit()
    gas.responses["internal.regenerate"] = {**created_pdf("R1"), "trash": {"status": "trashed", "fileId": "OLD-PDF"}}
    # 行の反映（3 列を空にした版）が済むまで GAS を呼ばない
    assert drain(env, clock, max_steps=3) == []
    assert gas.actions() == []
    sync_all(db)
    add_ping(db, clock())
    run(env, clock)
    call = gas.calls[-1]
    assert call["action"] == "internal.regenerate"
    assert call["args"] == {"reason": "secondary", "oldPdfFileId": "OLD-PDF"}
    assert call["expect"] == {"status": "approved", "approvedBy2": "boss@lineworks-local.info"}
    req = db.get(models.Request, "R1")
    db.refresh(req)
    assert req.old_pdf_file_id_pending == "" and req.pdf_file_id == "PDF-R1-1"


@pytest.mark.parametrize("reason", ["edit", "admin"])
def test_regenerate_edit_admin(env, db, gas, clock, reason):
    add_request(db, "R1")
    enq(env, gas_jobs.enqueue_regenerate, "R1", reason=reason, old_pdf_file_id="OLD")
    run(env, clock)
    assert gas.calls[-1]["action"] == "internal.regenerate"
    assert gas.calls[-1]["args"] == {"reason": reason, "oldPdfFileId": "OLD"}
    with pytest.raises(ValueError):
        gas_jobs.enqueue_regenerate(db, "R1", reason="other")


def test_ledger_upsert_uses_g_action(env, db, gas, clock):
    add_request(db, "R1", approved_by2="boss@lineworks-local.info")
    job_id = enq(env, gas_jobs.enqueue_ledger_upsert, "R1")
    assert job(db, job_id).kind == "internal.ledgerUpsert"
    run(env, clock)
    # G の Internal.gs には internal.ledgerUpsert の別名が無いので vps.upsertLedger を呼ぶ
    assert gas.actions() == ["vps.upsertLedger"]
    assert "vps.upsertLedger" in ACTIONS


def test_after_delete_runs_even_when_canceled(env, db, gas, clock):
    add_request(db, "R1", status="canceled", request_type="holiday")
    job_id = enq(env, gas_jobs.enqueue_after_delete, "R1", request_type="holiday", target_date="2026-11-25",
                 old_pdf_file_id="OLD")
    run(env, clock)
    assert job(db, job_id).state == "done"
    call = gas.calls[-1]
    assert call["action"] == "internal.afterDelete"
    assert call["args"] == {"oldPdfFileId": "OLD", "requestType": "holiday", "targetDate": "2026-11-25"}
    assert call["expect"] == {"status": "canceled"}


@pytest.mark.parametrize("enqueue", [
    lambda s: gas_jobs.enqueue_after_approve(s, "R1"),
    lambda s: gas_jobs.enqueue_after_complete(s, "R1", source="screen", status_at_enqueue="approved"),
    lambda s: gas_jobs.enqueue_regenerate(s, "R1", reason="edit"),
    lambda s: gas_jobs.enqueue_ledger_upsert(s, "R1"),
    lambda s: gas_jobs.enqueue_batch_pdf(s, "R1", target_date="2026-11-25", rewrite_ledger=False),
])
def test_canceled_request_is_skipped(env, db, gas, clock, enqueue):
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00")
    job_id = enq(env, enqueue)
    with env() as s:
        s.get(models.Request, "R1").status = "canceled"
        sheet_queue.mark_dirty(s, "requests", "R1")
        s.commit()
    sync_all(db)
    run(env, clock)
    assert job(db, job_id).state == "skipped" and job(db, job_id).last_error == "canceled"
    assert gas.actions() == []


def test_ledger_rebuild_and_rewrite_day(env, db, gas, clock):
    a = enq(env, gas_jobs.enqueue_ledger_rebuild)
    assert enq(env, gas_jobs.enqueue_ledger_rebuild) == a
    gas.responses["internal.ledgerRebuild"] = {"ok": False, "error": "bad_request",
                                               "message": "unknown action: internal.ledgerRebuild", "retryable": False}
    enq(env, gas_jobs.enqueue_ledger_rewrite_day, "2026-11-25")
    run(env, clock)
    assert gas.actions() == ["internal.ledgerRebuild", "internal.ledgerRewriteDay"]
    assert gas.calls[-1]["args"] == {"targetDate": "2026-11-25"}
    # G に ledgerRebuild が無いことは failed として管理者の欄に出る
    assert job(db, a).state == "failed" and "unknown action" in job(db, a).last_error


# ===== 動かす前の確かめ =====


def test_waits_for_sync_without_counting_attempts(env, db, gas, clock):
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00", synced=False)
    job_id = enq(env, gas_jobs.enqueue_after_approve, "R1")
    runner = JobRunner(env, lane="gas", now=clock)
    assert runner.run_once()["done"] == 0
    assert job(db, job_id).attempts == 0 and gas.actions() == []
    sync_all(db)
    clock.advance(seconds=6)
    add_ping(db, clock())
    assert runner.run_once()["done"] == 1


def test_waits_when_row_dirtied_after_enqueue(env, db, gas, clock):
    """requires の版は済んでいても、今の版が未反映なら待つ（expect と合わせるため）。"""
    add_request(db, "R1")
    job_id = enq(env, gas_jobs.enqueue_ledger_upsert, "R1")
    with env() as s:
        s.get(models.Request, "R1").approved_minutes = 90
        sheet_queue.mark_dirty(s, "requests", "R1")
        s.commit()
    runner = JobRunner(env, lane="gas", now=clock)
    assert runner.run_once() == {"done": 0, "skipped": 0, "retry": 0, "waiting": 1, "failed": 0}
    row = job(db, job_id)
    assert row.attempts == 0 and "row_not_synced" in row.last_error
    sync_all(db)
    clock.advance(seconds=6)
    add_ping(db, clock())
    assert runner.run_once()["done"] == 1


def test_sync_wait_fails_after_24_hours(env, db, gas, clock):
    add_request(db, "R1")
    job_id = enq(env, gas_jobs.enqueue_ledger_upsert, "R1")
    with env() as s:
        s.get(models.Request, "R1").approved_minutes = 90
        sheet_queue.mark_dirty(s, "requests", "R1")
        s.commit()
    clock.advance(hours=25)
    add_ping(db, clock())
    JobRunner(env, lane="gas", now=clock).run_once()
    assert job(db, job_id).state == "failed"
    assert db.query(models.SyncRun).filter_by(kind="gas_job_failed").count() == 1


@pytest.mark.parametrize("ping", ["none", "failed", "stale"])
def test_gas_ping_gate(session_factory, db, clock, gas, ping):
    gas_jobs.register_worker()
    if ping == "failed":
        add_ping(db, clock(), ok=False)
    elif ping == "stale":
        add_ping(db, clock() - timedelta(minutes=11))
    add_request(db, "R1")
    job_id = enq(session_factory, gas_jobs.enqueue_ledger_upsert, "R1")
    JobRunner(session_factory, lane="gas", now=clock).run_once()
    row = job(db, job_id)
    assert gas.actions() == [] and row.state == "pending" and row.attempts == 0
    assert row.run_after == clock() + timedelta(seconds=60) and "gas_ping" in row.last_error


# ===== 再試行・失敗・冪等 =====


@pytest.mark.parametrize("failure", [
    {"ok": False, "error": "busy", "message": "script lock not available", "retryable": True},
    {"ok": False, "error": "row_not_synced", "message": "sheet row does not match expect: status", "retryable": True},
    {"ok": False, "error": "internal", "message": "boom", "retryable": True},
    {"ok": False, "error": "pdf_failed", "message": "テンプレートが開けません", "retryable": True,
     "pdf": {"status": "error", "error": "テンプレートが開けません"}},
    GasTransportError("GAS の呼び出しが時間切れになりました", timeout=True),
])
def test_retryable_errors_back_off_and_reuse_job_id(env, db, gas, clock, failure):
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00")
    job_id = enq(env, gas_jobs.enqueue_after_approve, "R1")
    gas.responses["internal.afterApprove"] = [failure, failure, None]
    runner = JobRunner(env, lane="gas", now=clock)
    assert runner.run_once()["retry"] == 1
    row = job(db, job_id)
    assert row.state == "pending" and row.attempts == 1 and row.run_after == clock() + timedelta(minutes=1)
    clock.advance(minutes=1)
    add_ping(db, clock())
    assert runner.run_once()["retry"] == 1
    assert job(db, job_id).run_after == clock() + timedelta(minutes=2)
    clock.advance(minutes=2)
    add_ping(db, clock())
    assert runner.run_once()["done"] == 1
    ids = {c["jobId"] for c in gas.calls if c["action"] == "internal.afterApprove"}
    assert len(ids) == 1  # 同じ jobId で送り直す（GAS が同じ返事を返す＝二重にならない）


def test_replayed_response_is_recorded_once(env, db, gas, clock):
    """時間切れの後の送り直しに GAS が replayed:true で同じ返事を返しても、記録は 1 回分。"""
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00")
    job_id = enq(env, gas_jobs.enqueue_after_approve, "R1")
    gas.responses["internal.afterApprove"] = [GasTransportError("timeout", timeout=True),
                                              {**created_pdf("R1"), "replayed": True}]
    run(env, clock)
    assert job(db, job_id).state == "done"
    req = db.get(models.Request, "R1")
    db.refresh(req)
    assert req.pdf_file_id == "PDF-R1-1"
    st = db.get(models.SheetSyncState, ("requests", "R1"))
    db.refresh(st)
    assert st.synced_version == st.desired_version


def test_non_retryable_fails_and_is_visible(env, db, gas, clock):
    add_request(db, "R1")
    job_id = enq(env, gas_jobs.enqueue_ledger_upsert, "R1")
    gas.responses["vps.upsertLedger"] = {"ok": False, "error": "ledger_header_mismatch",
                                         "message": "残業_2026年度 の 3 列目の見出しが違います", "retryable": False}
    run(env, clock)
    row = job(db, job_id)
    assert row.state == "failed" and "ledger_header_mismatch" in row.last_error
    assert json.loads(row.result_json)["error"] == "ledger_header_mismatch"
    ex = db.query(models.SyncRun).filter_by(kind="gas_job_failed").one()
    assert ex.summary_json["requestId"] == "R1" and ex.summary_json["kind"] == "internal.ledgerUpsert"


def test_retries_stop_after_24_hours(env, db, gas, clock):
    add_request(db, "R1")
    job_id = enq(env, gas_jobs.enqueue_ledger_upsert, "R1")
    gas.responses["vps.upsertLedger"] = lambda **kw: {"ok": False, "error": "busy", "retryable": True}
    run(env, clock)
    row = job(db, job_id)
    assert row.state == "failed"
    # 1・2・4・8・16・30 分、その後 30 分ごとに 24 時間まで
    n = len(gas.actions())
    assert n == 52  # 7 回（61 分）＋ 30 分ごとに 45 回
    assert clock() - T0 <= timedelta(hours=24, minutes=1)


def test_failure_with_pdf_records_pdf_first(env, db, gas, clock):
    add_request(db, "R1")
    job_id = enq(env, gas_jobs.enqueue_regenerate, "R1", reason="edit")
    resp = created_pdf("R1")
    resp.update(ok=False, error="internal", message="ledger exploded", retryable=True)
    gas.responses["internal.regenerate"] = [resp, None]
    JobRunner(env, lane="gas", now=clock).run_once()
    assert job(db, job_id).state == "pending"
    req = db.get(models.Request, "R1")
    db.refresh(req)
    assert req.pdf_file_id == "PDF-R1-1"


def test_same_request_order_and_failed_blocks(env, db, gas, clock):
    add_request(db, "R1")
    add_request(db, "R2")
    a = enq(env, gas_jobs.enqueue_regenerate, "R1", reason="edit")
    b = enq(env, gas_jobs.enqueue_ledger_upsert, "R1")
    c = enq(env, gas_jobs.enqueue_ledger_upsert, "R2")
    gas.responses["internal.regenerate"] = {"ok": False, "error": "bad_request", "retryable": False}
    run(env, clock)
    assert job(db, a).state == "failed"
    assert job(db, b).state == "pending"  # 前の仕事が failed なので止まる（台帳を古い値で上書きしない）
    assert job(db, c).state == "done"  # 別の申請は先へ進む
    assert [j["id"] for j in jobs.list_blocked(db)] == [b]


def test_lane_is_serial(env, db, gas, clock):
    """GAS を呼んでいる最中に別の worker が取り出しても何も取れない（同時に 2 本動かない）。"""
    for rid in ("R1", "R2", "R3"):
        add_request(db, rid)
        enq(env, gas_jobs.enqueue_ledger_upsert, rid)
    other = JobRunner(env, lane="gas", now=clock)
    grabbed = []
    gas.during = lambda action, rid: grabbed.append(other.claim())
    run(env, clock)
    assert gas.actions() == ["vps.upsertLedger"] * 3
    assert grabbed == [[], [], []] and gas.max_in_flight == 1


def test_batch_spacing_500ms(env, db, gas, clock):
    for i in range(6):
        add_request(db, f"R{i}")
        enq(env, gas_jobs.enqueue_ledger_upsert, f"R{i}", batch_id="B6")
    for i in range(5):
        add_request(db, f"S{i}")
        enq(env, gas_jobs.enqueue_ledger_upsert, f"S{i}", batch_id="B5")
    run(env, clock)
    assert len(gas.actions()) == 11
    assert gas.sleeps == [0.5] * 5  # 6 件のまとまりだけ、2 件目から 500ms 空ける


# ===== record_pdf_result（6-5） =====


def test_record_pdf_result_keeps_pending_changes(db, clock):
    add_request(db, "R1")
    req = db.get(models.Request, "R1")
    req.approved_minutes = 30
    sheet_queue.mark_dirty(db, "requests", "R1")  # 別の列の反映待ち
    db.commit()
    st = db.get(models.SheetSyncState, ("requests", "R1"))
    synced_before = st.synced_version
    assert gas_jobs.record_pdf_result(db, "R1", {"pdfFileId": "P9", "folderId": "F9",
                                                 "pdfGeneratedAt": "2026-11-26 10:00:00"})
    db.commit()
    db.refresh(st)
    assert st.synced_version == synced_before  # 未反映の版を「済み」にしない
    assert st.desired_version == req.row_version
    assert st.synced_values["pdfFileId"] == "P9"


def test_record_pdf_result_does_not_clobber_newer(db, clock):
    add_request(db, "R1")
    req = db.get(models.Request, "R1")
    snapshot = gas_jobs.pdf_snapshot(req)
    # GAS が PDF を作っている間に二次承認が 3 列を空に……ではなく別の PDF に変わった（新しい値）
    req.pdf_file_id, req.pdf_folder_id = "NEWER", "F2"
    sheet_queue.mark_dirty(db, "requests", "R1")
    db.commit()
    assert not gas_jobs.record_pdf_result(db, "R1", {"pdfFileId": "STALE", "folderId": "F1",
                                                     "pdfGeneratedAt": "2026-11-26 09:00:00"},
                                          expected_current=snapshot)
    db.commit()
    db.refresh(req)
    assert req.pdf_file_id == "NEWER"
    st = db.get(models.SheetSyncState, ("requests", "R1"))
    assert st.synced_values["pdfFileId"] == "STALE"  # シートにあるのは STALE → 行の反映が NEWER を書き戻す
    assert st.desired_version > st.synced_version


def test_record_pdf_result_stashes_stale_pdf_when_cleared(db, clock):
    add_request(db, "R1", pdf_file_id="OLD")
    req = db.get(models.Request, "R1")
    snapshot = gas_jobs.pdf_snapshot(req)
    req.pdf_file_id, req.pdf_folder_id, req.pdf_generated_at = "", "", None
    sheet_queue.mark_dirty(db, "requests", "R1")
    db.commit()
    assert not gas_jobs.record_pdf_result(db, "R1", {"pdfFileId": "STALE"}, expected_current=snapshot)
    db.commit()
    db.refresh(req)
    assert req.pdf_file_id == "" and req.old_pdf_file_id_pending == "STALE"


def test_record_pdf_result_epoch_ms(db, clock):
    add_request(db, "R1")
    gas_jobs.record_pdf_result(db, "R1", {"pdfFileId": "P1", "pdfGeneratedAtEpochMs": 1795655405123})
    db.commit()
    req = db.get(models.Request, "R1")
    assert req.pdf_generated_at.timestamp() == pytest.approx(1795655405.123)


# ===== Sebastian の待ち（4-4） =====


def test_wait_for_result_done_and_view(env, db, gas, clock, monkeypatch):
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00")
    job_id = enq(env, gas_jobs.enqueue_after_complete, "R1", source="sebastian", status_at_enqueue="approved",
                 priority=10)
    assert gas_jobs.wait_for_result(env, job_id, 0) is None
    assert gas_jobs.sebastian_pdf_view(None) == {"error": "PDF作成の結果待ちが時間切れになりました（作成は続いています）"}
    run(env, clock)
    outcome = gas_jobs.wait_for_result(env, job_id, 5)
    assert outcome["jobState"] == "done"
    assert gas_jobs.sebastian_pdf_view(outcome) == {"ok": True, "pdfFileId": "PDF-R1-1", "pdfName": "R1.pdf",
                                                    "folderId": "FOLDER-1"}
    assert gas_jobs.wait_for_sebastian_pdf(env, job_id)["pdfFileId"] == "PDF-R1-1"


def test_wait_for_result_shapes(env, db, gas, clock):
    add_request(db, "R1")
    add_request(db, "R2")
    a = enq(env, gas_jobs.enqueue_after_complete, "R1", source="sebastian", status_at_enqueue="approved")
    b = enq(env, gas_jobs.enqueue_after_complete, "R2", source="sebastian", status_at_enqueue="approved")
    gas.responses["internal.afterComplete"] = [
        {"ok": True, "pdf": {"status": "already", "pdfFileId": "P-OLD"}},
        {"ok": False, "error": "pdf_failed", "message": "テンプレ無し", "retryable": True,
         "pdf": {"status": "error", "error": "テンプレ無し"}},
    ]
    runner = JobRunner(env, lane="gas", now=clock)
    runner.run_once()
    runner.run_once()
    assert gas_jobs.sebastian_pdf_view(gas_jobs.wait_for_result(env, a, 0)) == {"already": True,
                                                                                "pdfFileId": "P-OLD"}
    outcome = gas_jobs.wait_for_result(env, b, 0)
    assert outcome["jobState"] == "retrying"  # 仕事は続く（後で送り直す）
    assert gas_jobs.sebastian_pdf_view(outcome) == {"error": "テンプレ無し"}


def test_wait_for_result_polls_until_timeout(env, db, gas, monkeypatch):
    add_request(db, "R1")
    job_id = enq(env, gas_jobs.enqueue_ledger_upsert, "R1")
    now = [0.0]
    monkeypatch.setattr(gas_jobs, "_monotonic", lambda: now[0])
    monkeypatch.setattr(gas_jobs, "_sleep", lambda s: now.__setitem__(0, now[0] + s))
    assert gas_jobs.wait_for_result(env, job_id, 40) is None
    assert now[0] == pytest.approx(40)


# ===== gas_ping =====


def test_gas_ping_accepts_g_field_names(session_factory, db, clock, gas, settings):
    result = gas_jobs.run_gas_ping(session_factory, settings)
    assert result["ok"] is True and result["problems"] == []
    row = gas_jobs.last_gas_ping(db)
    assert row.ok and gas_jobs.gas_ping_ok(db)[0]


def test_gas_ping_mismatch_stops_lane(session_factory, db, clock, gas, settings):
    bad = good_ping()
    bad.pop("scriptId")  # G の vpsPing_ は scriptId を返さない（仕様 7-2 との違い）
    bad["dbId"] = "1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ"
    gas.responses["internal.ping"] = bad
    result = gas_jobs.run_gas_ping(session_factory, settings)
    assert result["ok"] is False
    assert any("scriptId" in p for p in result["problems"])
    assert any("本番" in p for p in result["problems"])
    assert gas_jobs.gas_ping_ok(db)[0] is False


def test_gas_ping_transport_error(session_factory, db, clock, gas, settings):
    gas.responses["internal.ping"] = GasTransportError("connect failed")
    assert gas_jobs.run_gas_ping(session_factory, settings)["ok"] is False
    assert gas_jobs.last_gas_ping(db).ok is False


def test_register_worker(session_factory):
    gas_jobs.register_worker()
    assert set(jobs.registered_kinds("gas")) == set(gas_jobs.KINDS)
    sched = worker.schedules()["gas_ping"]
    assert sched.every_minutes == 5 and sched.run_at_startup


def test_resend_success_enqueues_ledger_rewrite(env, db, gas, clock):
    add_request(db, "R1")
    enq(env, gas_jobs.enqueue_batch_pdf, "R1", target_date="2026-11-25", rewrite_ledger=True)
    run(env, clock)
    assert gas.actions() == ["internal.batchPdf", "internal.ledgerRewriteDay"]
    assert gas.calls[-1]["args"] == {"targetDate": "2026-11-25"}


def test_all_g_actions_known_to_client():
    for spec in gas_jobs.KINDS.values():
        assert spec.gas_action in ACTIONS
