"""データ安全 高-3: シートへ書く前の関所（services/write_gate）。

本番は「最後に合格した gas_ping が N 分以内・GAS が移転モード（migrated=true）」でなければ、行の反映・照合・
マスタ同期はシートへ書かない。ステージングは、ステージングの GAS への合格した ping が N 分以内なら書ける。
"""
from __future__ import annotations

import types
from datetime import datetime, timedelta, timezone

import pytest
from sqlalchemy import select

from app import sheet_queue
from app.models import Request, SyncRun
from app.services import master_sync, ops_status, reconcile, row_sync, write_gate
from conftest import freeze_now
from sheet_fixtures import R4, REQUEST_HEADERS, STAGING_SID, make_fake, migrate

pytestmark = pytest.mark.real_write_gate

NOW = datetime(2026, 11, 26, 3, 0, tzinfo=timezone.utc)
PROD = types.SimpleNamespace(is_staging=False, sheet_write_ping_max_minutes=10)
STAGING = types.SimpleNamespace(is_staging=True, sheet_write_ping_max_minutes=10)


def add_ping(db, *, at=NOW, ok=True, migrated=True):
    db.add(SyncRun(kind="gas_ping", started_at=at, finished_at=at, ok=ok,
                   summary_json={"problems": [] if ok else ["x"], "migrationMode": migrated}))
    db.commit()


@pytest.fixture(autouse=True)
def _lane():
    row_sync.reset_lane()
    yield
    row_sync.reset_lane()


def test_no_ping_blocks(db):
    ok, reason = write_gate.evaluate(db, PROD, NOW)
    assert not ok and "gas_ping" in reason


def test_production_requires_fresh_ok_ping_with_migrated_true(db):
    add_ping(db, at=NOW - timedelta(minutes=3), migrated=True)
    assert write_gate.evaluate(db, PROD, NOW) == (True, "")
    assert not write_gate.evaluate(db, PROD, NOW + timedelta(minutes=8))[0]  # 11 分前は古い


@pytest.mark.parametrize("migrated", [False, None, "true"])
def test_production_blocks_when_gas_not_in_migration_mode(db, migrated):
    add_ping(db, migrated=migrated)
    ok, reason = write_gate.evaluate(db, PROD, NOW)
    assert not ok and "移転モード" in reason


def test_failed_latest_ping_falls_back_to_last_ok_one(db):
    add_ping(db, at=NOW - timedelta(minutes=4), migrated=True)
    add_ping(db, at=NOW - timedelta(minutes=1), ok=False, migrated=None)
    assert write_gate.evaluate(db, PROD, NOW)[0] is True
    add_ping(db, at=NOW, ok=True, migrated=False)  # 戻した（移転モードを解いた）
    assert write_gate.evaluate(db, PROD, NOW)[0] is False


def test_staging_needs_fresh_staging_ping_but_not_migration_mode(db):
    assert not write_gate.evaluate(db, STAGING, NOW)[0]
    add_ping(db, migrated=False)
    assert write_gate.evaluate(db, STAGING, NOW) == (True, "")


def test_row_sync_halts_without_writing_and_admin_sees_it(engine, session_factory, settings, db, monkeypatch):
    freeze_now(monkeypatch, NOW)
    fake = make_fake()
    migrate(fake, engine, settings)
    fake.calls.clear()
    req = db.get(Request, R4)
    req.status = "approved"
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    result = row_sync.run_once(session_factory, api=fake.api(settings), app_settings=settings)
    assert result["ok"] is False and "書き込みの関所" in result["halted"] and result["written"] == 0
    assert fake.writes() == []
    assert row_sync.status(db)["halted"] is True
    st = ops_status.get_ops_status(db)
    assert st["sheetWriteGate"]["blocked"] is True and st["rowSync"]

    # 合格した ping が来れば、確かめ直しの後に書く
    add_ping(db, migrated=False)  # ステージング
    freeze_now(monkeypatch, NOW + timedelta(seconds=row_sync.GATE_RECHECK_SECONDS + 1))
    result = row_sync.run_once(session_factory, api=fake.api(settings), app_settings=settings)
    assert result["written"] == 1
    tab = fake.books[STAGING_SID].tabs["Requests"]
    assert tab.get(4, REQUEST_HEADERS.index("status(submitted/approved/canceled)")).value == "approved"


def test_reconcile_reports_but_does_not_write_when_gate_is_closed(engine, session_factory, settings, db, monkeypatch):
    freeze_now(monkeypatch, NOW)
    fake = make_fake()
    migrate(fake, engine, settings)
    fake.calls.clear()
    tab = fake.books[STAGING_SID].tabs["Requests"]
    tab.get(2, REQUEST_HEADERS.index("workerName")).value = "人が直した名前"
    result = reconcile.run(session_factory, api=fake.api(settings), app_settings=settings)
    assert result["ok"] is False and result["diffCount"] == 1 and result["fixedRows"] == 0 and result["writeBlocked"]
    assert fake.writes() == []
    assert tab.get(2, REQUEST_HEADERS.index("workerName")).value == "人が直した名前"


def test_master_sync_does_not_write_when_gate_is_closed(db, settings, monkeypatch):
    monkeypatch.setattr(master_sync, "sleep", lambda s: None)
    freeze_now(monkeypatch, NOW)
    fake = make_fake()
    result = master_sync.sync_masters(db, api=fake.api(settings), app_settings=settings)
    assert result["ok"] is False and result["writeBlocked"]
    assert fake.writes() == []
    assert db.scalars(select(SyncRun).where(SyncRun.kind == "sync_masters")).first() is not None
