"""朝バッチ（services/morning_batch.py。実装仕様 7-5・12-2 の M）。

成功・スキップ・失敗の数、エラー詳細の形、BatchLogs の 1 行の値と型、ledgerRewriteDay が最後、6 件以上の待ち、
Sebastian を先に流す、送り直しで ledgerRewriteDay をもう一度、手動の PDF 一括、台帳が開けないとき、7:05 の確かめ。
"""
from __future__ import annotations

import json
from datetime import date, datetime, timedelta, timezone

import pytest
from gas_fakes import Clock, FakeGas, add_ping, add_request, drain, job, sync_all

from app import jobs, models, timeutil, worker
from app.jobs import JobRunner
from app.services import gas_jobs, morning_batch
from app.services.gas_client import GasTransportError

MORNING = datetime(2026, 11, 25, 21, 50, tzinfo=timezone.utc)  # 2026-11-26 06:50 JST
YMD = "2026-11-25"


@pytest.fixture
def clock(monkeypatch):
    c = Clock(MORNING)
    monkeypatch.setattr(timeutil, "now_utc", c)
    return c


@pytest.fixture
def gas(monkeypatch):
    fake = FakeGas()
    gas_jobs.set_client_factory(lambda settings: fake)
    monkeypatch.setattr(gas_jobs, "_sleep", lambda s: None)
    yield fake
    gas_jobs.set_client_factory(None)


@pytest.fixture
def env(session_factory, db, clock, gas):
    gas_jobs.register_worker()
    morning_batch.register_worker()
    add_ping(db, clock())
    return session_factory


def run(sf, clock, **kw):
    return drain(sf, clock, sync=lambda s: add_ping(s, clock()), **kw)


def batch_rows(db):
    db.expire_all()
    return db.query(models.BatchLog).all()


def test_schedule_registration(env):
    sched = worker.schedules()
    assert sched["morning_batch"].at == (6, 50)
    assert sched["morning_batch_check"].at == (7, 5)
    assert "morning_batch" in jobs.registered_kinds("gas")


def test_full_sequence(env, db, gas, clock):
    add_request(db, "A", dept="製造部", worker_name="あべ", pdf_file_id="EXIST")  # 作成済み → スキップ
    add_request(db, "B", dept="製造部", worker_name="いとう")  # 作る
    add_request(db, "C", dept="製造部", worker_name="うえだ")  # PDF の例外
    add_request(db, "D", dept="製造部", worker_name="えんどう")  # 通信の失敗が 3 回
    add_request(db, "X", target=date(2026, 11, 24))  # 別の日
    add_request(db, "Y", status="submitted")  # 未承認
    add_request(db, "Z", status="canceled")
    result = morning_batch.run_morning_batch(env)
    assert result["targetDate"] == YMD
    assert morning_batch.run_morning_batch(env)["jobId"] == result["jobId"]  # 同じ日は 1 本だけ

    # D は 3 回とも失敗させ、送り直しの仕事では成功させる
    d_calls = []

    def batch_pdf(**kw):
        if kw["request_id"] == "D":
            d_calls.append(kw["job_id"])
            if len(d_calls) <= 3:
                return GasTransportError("GAS に届きませんでした（internal.batchPdf）: ConnectError")
        if kw["request_id"] == "C":
            return {"ok": False, "error": "pdf_failed", "message": "テンプレートが見つかりません", "retryable": True,
                    "pdf": {"status": "error", "error": "テンプレートが見つかりません"}}
        return gas.default("internal.batchPdf", kw["request_id"])

    gas.responses["internal.batchPdf"] = batch_pdf
    run(env, clock)

    seq = [(c["action"], c["requestId"]) for c in gas.calls if c["action"] != "internal.ping"]
    assert seq[:6] == [("internal.batchPdf", "B"), ("internal.batchPdf", "C"), ("internal.batchPdf", "D"),
                       ("internal.batchPdf", "D"), ("internal.batchPdf", "D"), ("internal.ledgerRewriteDay", "")]
    # 送り直し（priority 0）は朝バッチの後。成功したので、その日の台帳の書き直しをもう一度
    assert seq[6:] == [("internal.batchPdf", "D"), ("internal.ledgerRewriteDay", "")]
    assert len(set(d_calls[:3])) == 1  # 朝バッチの中の送り直しは同じ jobId
    assert d_calls[3] != d_calls[0]
    assert all(c["args"] == {"targetDate": YMD} for c in gas.calls if c["action"] != "internal.ping")

    rows = batch_rows(db)
    assert len(rows) == 1
    row = rows[0]
    assert row.batch_name == "morningBatchNoMail" and row.target_date == date(2026, 11, 25)
    assert row.target_date_text == YMD
    assert (row.success, row.skip, row.fail) == (1, 1, 2)
    assert row.error_detail == ("C (製造部/うえだ): テンプレートが見つかりません\n"
                                "D (製造部/えんどう): GAS呼び出し失敗: GAS に届きませんでした（internal.batchPdf）: ConnectError")
    assert row.executed_at is not None
    st = db.get(models.SheetSyncState, ("batch_logs", str(row.id)))
    assert st is not None and st.desired_version > st.synced_version  # 行の反映の印
    b = db.get(models.Request, "B")
    db.refresh(b)
    assert b.pdf_file_id == "PDF-B-1"

    assert job(db, result["jobId"]).state == "done"
    run_row = db.query(models.SyncRun).filter_by(kind="morning_batch").one()
    assert run_row.ok and run_row.summary_json["counts"] == {"ok": 1, "skip": 1, "fail": 2}


def test_items_sorted_by_dept_and_name(env, db, gas, clock):
    add_request(db, "R1", dept="製造部", worker_name="たなか")
    add_request(db, "R2", dept="営業部", worker_name="すずき")
    add_request(db, "R3", dept="製造部", worker_name="アオキ")
    items = morning_batch.list_items(db, YMD)
    # 製造部 はひらがな・カタカナを同じに並べる（localeCompare 'ja' の近似）
    assert [i["requestId"] for i in items] == ["R2", "R3", "R1"]


def test_text_target_date_is_matched(env, db):
    add_request(db, "R1")
    r = db.get(models.Request, "R1")
    r.target_date, r.target_date_text = None, "2026/11/25"
    db.commit()
    assert [i["requestId"] for i in morning_batch.list_items(db, YMD)] == ["R1"]


def test_six_items_wait_500ms_and_yield(env, db, gas, clock):
    for i in range(6):
        add_request(db, f"R{i}", worker_name=f"n{i}")
    job_id = morning_batch.run_morning_batch(env)["jobId"]
    runner = JobRunner(env, lane="gas", now=clock)
    assert runner.run_once()["waiting"] == 1
    row = job(db, job_id)
    assert row.state == "pending" and row.run_after == clock() + timedelta(milliseconds=500)
    assert json.loads(row.result_json)["index"] == 1


def test_five_items_do_not_wait(env, db, gas, clock):
    for i in range(5):
        add_request(db, f"R{i}", worker_name=f"n{i}")
    job_id = morning_batch.run_morning_batch(env)["jobId"]
    JobRunner(env, lane="gas", now=clock).run_once()
    assert job(db, job_id).run_after == clock()


def test_sebastian_runs_between_items(env, db, gas, clock):
    for i in range(3):
        add_request(db, f"R{i}", worker_name=f"n{i}")
    add_request(db, "S", target=date(2026, 11, 26), work_log_end="2026-11-26 06:00:00")
    morning_batch.run_morning_batch(env)
    state = {"done": False}

    def after_step():
        if not state["done"] and gas.actions():
            state["done"] = True
            with env() as s:
                gas_jobs.enqueue_after_complete(s, "S", source="sebastian", status_at_enqueue="approved", priority=10)
                s.commit()

    run(env, clock, on_step=after_step)
    seq = [(c["action"], c["requestId"]) for c in gas.calls if c["action"] != "internal.ping"]
    assert seq == [("internal.batchPdf", "R0"), ("internal.afterComplete", "S"), ("internal.batchPdf", "R1"),
                   ("internal.batchPdf", "R2"), ("internal.ledgerRewriteDay", "")]


def test_manual_batch_has_no_ledger(env, db, gas, clock):
    add_request(db, "R1", target=date(2026, 11, 20))
    with env() as s:
        job_id = morning_batch.enqueue_manual_batch(s, "2026-11-20")
        s.commit()
    with pytest.raises(ValueError):
        morning_batch.enqueue_manual_batch(db, "2026/11/20")
    run(env, clock)
    assert gas.actions() == ["internal.batchPdf"]
    row = batch_rows(db)[0]
    assert row.batch_name == "manualBatchPdfs" and row.target_date_text == "2026-11-20" and row.success == 1
    assert job(db, job_id).state == "done"
    assert db.query(models.SyncRun).filter_by(kind="manual_batch_pdfs").one().ok


def test_empty_day_still_logs_and_rewrites(env, db, gas, clock):
    morning_batch.run_morning_batch(env)
    run(env, clock)
    assert gas.actions() == ["internal.ledgerRewriteDay"]
    row = batch_rows(db)[0]
    assert (row.success, row.skip, row.fail, row.error_detail) == (0, 0, 0, "")


def test_ledger_unavailable_fails_then_resumes_at_ledger(env, db, gas, clock):
    add_request(db, "R1")
    gas.responses["internal.ledgerRewriteDay"] = [
        {"ok": False, "error": "ledger_unavailable", "message": "Unexpected error while getting the method",
         "retryable": True}, None]
    job_id = morning_batch.run_morning_batch(env)["jobId"]
    run(env, clock)
    row = job(db, job_id)
    assert row.state == "failed" and "ledger_unavailable" in row.last_error
    failed_run = db.query(models.SyncRun).filter_by(kind="morning_batch").one()
    assert failed_run.ok is False and "ledger_unavailable" in failed_run.summary_json["error"]
    assert db.query(models.SyncRun).filter_by(kind="gas_job_failed").count() == 1
    # 管理者が「失敗した仕事の再実行」→ PDF は作り直さず、台帳だけもう一度
    with env() as s:
        assert jobs.retry_failed(s, job_id)
        s.commit()
    run(env, clock)
    assert gas.actions() == ["internal.batchPdf", "internal.ledgerRewriteDay", "internal.ledgerRewriteDay"]
    assert job(db, job_id).state == "done"
    assert len(batch_rows(db)) == 1


def test_waits_for_row_sync_and_gas_ping(env, db, gas, clock):
    add_request(db, "R1", synced=False)
    job_id = morning_batch.run_morning_batch(env)["jobId"]
    runner = JobRunner(env, lane="gas", now=clock)
    runner.run_once()
    row = job(db, job_id)
    assert row.state == "pending" and row.attempts == 0 and "row_not_synced" in row.last_error
    assert gas.actions() == []
    sync_all(db)
    clock.advance(minutes=11)  # gas_ping が古い
    runner.run_once()
    assert gas.actions() == [] and "gas_ping" in job(db, job_id).last_error
    run(env, clock)
    assert gas.actions() == ["internal.batchPdf", "internal.ledgerRewriteDay"]


def test_check_morning_sync(env, db, gas, clock):
    add_request(db, "R1")
    morning_batch.run_morning_batch(env)
    run(env, clock)
    clock.at = datetime(2026, 11, 25, 22, 5, tzinfo=timezone.utc)  # 07:05 JST
    result = morning_batch.check_morning_sync(env)
    assert result["ok"] is False  # BatchLogs と R1（PDF の列）の反映がまだ
    sync_all(db)
    assert morning_batch.check_morning_sync(env)["ok"] is True
    assert db.query(models.SyncRun).filter_by(kind="morning_batch_check").count() == 2
