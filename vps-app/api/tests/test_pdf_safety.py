"""データ安全 高-2・中-1: GAS がシートに書いた PDF の 3 列を、照合が空で上書きしない。

- GAS が skipped（has_pdf）で既存の pdfFileId を返したら DB に記録する（仕事・朝バッチの両方）。
- 照合は、シートの PDF の列に値があり DB が空なら書き直さず報告する（pdfKeptOnSheet）。
- 照合は、GAS の仕事が待ち・実行中の申請を比べない。
"""
from __future__ import annotations

from datetime import date, datetime, timezone

import pytest
from gas_fakes import Clock, FakeGas, add_ping, add_request, drain

from app import jobs, models, timeutil
from app.services import gas_jobs, morning_batch, reconcile, row_sync
from conftest import freeze_now
from sheet_fixtures import R2, REQUEST_HEADERS, STAGING_SID, make_fake, migrate

HAS_PDF = {"ok": True, "requestId": "R1",
           "pdf": {"status": "skipped", "reason": "has_pdf"},
           "pdfFileId": "PDF-EXISTING", "pdfFolderId": "FOLDER-9", "pdfGeneratedAt": "2026-11-26 09:00:05",
           "ledger": {"status": "upserted"}, "trash": {"status": "skipped"}}


@pytest.fixture
def clock(monkeypatch):
    c = Clock(datetime(2026, 11, 25, 21, 50, tzinfo=timezone.utc))
    monkeypatch.setattr(timeutil, "now_utc", c)
    return c


@pytest.fixture
def gas(monkeypatch):
    fake = FakeGas()
    gas_jobs.set_client_factory(lambda settings: fake)
    monkeypatch.setattr(gas_jobs, "_sleep", lambda s: None)
    gas_jobs._batch_last_call.clear()
    yield fake
    gas_jobs.set_client_factory(None)


@pytest.fixture
def env(session_factory, db, clock, gas):
    gas_jobs.register_worker()
    morning_batch.register_worker()
    add_ping(db, clock())
    return session_factory


def run(sf, clock):
    return drain(sf, clock, sync=lambda s: add_ping(s, clock()))


def test_has_pdf_skip_from_a_job_is_recorded(env, db, gas, clock):
    add_request(db, "R1", work_log_end="2026-11-25 19:00:00")
    gas.responses["internal.afterApprove"] = [dict(HAS_PDF)]
    with env() as s:
        gas_jobs.enqueue_after_approve(s, "R1")
        s.commit()
    assert run(env, clock) == ["done"]
    req = db.get(models.Request, "R1")
    db.refresh(req)
    assert req.pdf_file_id == "PDF-EXISTING" and req.pdf_folder_id == "FOLDER-9"
    st = db.get(models.SheetSyncState, ("requests", "R1"))
    db.refresh(st)
    assert st.synced_values["pdfFileId"] == "PDF-EXISTING"  # シートに既にある値（行の反映は書き直さない）


def test_skipped_without_file_id_is_not_recorded():
    assert gas_jobs.pdf_result_to_record({"pdf": {"status": "skipped", "reason": "not_approved"}}) is False
    assert gas_jobs.pdf_result_to_record(HAS_PDF) is True
    assert gas_jobs.pdf_result_to_record({"pdf": {"status": "created", "pdfFileId": "X"}}) is True


def test_has_pdf_skip_in_morning_batch_is_recorded_and_counted_as_skip(env, db, gas, clock):
    add_request(db, "R1")
    gas.responses["internal.batchPdf"] = [dict(HAS_PDF)]
    morning_batch.run_morning_batch(env)
    run(env, clock)
    req = db.get(models.Request, "R1")
    db.refresh(req)
    assert req.pdf_file_id == "PDF-EXISTING"
    db.expire_all()
    row = db.query(models.BatchLog).one()
    assert (row.success, row.skip, row.fail) == (0, 1, 0)


# ===== 照合 =====


NOW = datetime(2026, 11, 26, 17, 10, tzinfo=timezone.utc)


@pytest.fixture
def rec(engine, session_factory, settings, monkeypatch):
    row_sync.reset_lane()
    freeze_now(monkeypatch, NOW)
    fake = make_fake()
    migrate(fake, engine, settings)
    fake.calls.clear()

    def go(check_only=False):
        return reconcile.run(session_factory, check_only=check_only, api=fake.api(settings), app_settings=settings)

    yield type("Rec", (), {"fake": fake, "tab": fake.books[STAGING_SID].tabs["Requests"], "run": staticmethod(go)})
    row_sync.reset_lane()


def _row_of(tab, request_id):
    col = REQUEST_HEADERS.index("requestId")
    for r in range(1, 200):
        if tab.get(r, col).value == request_id:
            return r
    raise AssertionError(request_id)


def test_reconcile_never_blanks_pdf_columns_on_sheet(rec, db):
    req = db.get(models.Request, R2)
    req.pdf_file_id, req.pdf_folder_id, req.pdf_generated_at = "", "", None
    db.commit()
    r = _row_of(rec.tab, R2)
    rec.tab.get(r, REQUEST_HEADERS.index("pdfFileId")).value = "PDF-ON-SHEET"
    rec.tab.get(r, REQUEST_HEADERS.index("workerName")).value = "人が直した名前"
    result = rec.run()
    assert rec.tab.get(r, REQUEST_HEADERS.index("pdfFileId")).value == "PDF-ON-SHEET"  # 残る
    assert rec.tab.get(r, REQUEST_HEADERS.index("workerName")).value == req.worker_name  # ほかの差は直す
    assert [d["header"] for d in result["pdfKeptOnSheet"]] == ["pdfFileId"]


def test_reconcile_skips_requests_with_running_gas_jobs(rec, db):
    r = _row_of(rec.tab, R2)
    rec.tab.get(r, REQUEST_HEADERS.index("workerName")).value = "GAS の仕事の途中"
    db.add(models.Job(kind="internal.afterApprove", lane="gas", request_id=R2, state="running",
                      created_at=NOW, updated_at=NOW))
    db.commit()
    result = rec.run()
    assert result["diffCount"] == 0 and result["tabs"]["Requests"]["skipped"] >= 1
    assert rec.tab.get(r, REQUEST_HEADERS.index("workerName")).value == "GAS の仕事の途中"
