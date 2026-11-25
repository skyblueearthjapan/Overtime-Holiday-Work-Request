"""worker の予定（sync_runs の枠・起動の直後に走らせ直さない・run の CLI）と purge。"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest
from sqlalchemy import select

from app import timeutil, worker
from app.models import IdempotencyKey, Job, SyncRun, UserSession
from app.timeutil import JST


def jst(*args):
    return datetime(*args, tzinfo=JST).astimezone(timezone.utc)


def test_register_schedule_validation():
    with pytest.raises(ValueError):
        worker.register_schedule("x", task=lambda sf: {}, at=(1, 0), every_minutes=5)
    with pytest.raises(ValueError):
        worker.register_schedule("x", task=lambda sf: {})
    with pytest.raises(ValueError):
        worker.register_schedule("x", task=lambda sf: {}, at=(24, 0))


def test_daily_slot_window():
    worker.register_schedule("sync_masters", task=lambda sf: {}, at=(0, 30))
    sched = worker.schedules()["sync_masters"]
    assert worker.slot_start(sched, jst(2026, 11, 26, 0, 29)) is None
    assert worker.slot_start(sched, jst(2026, 11, 26, 0, 30)) == jst(2026, 11, 26, 0, 30)
    assert worker.slot_start(sched, jst(2026, 11, 26, 1, 29)) == jst(2026, 11, 26, 0, 30)
    assert worker.slot_start(sched, jst(2026, 11, 26, 1, 30)) is None
    assert worker.slot_start(sched, jst(2026, 11, 26, 10, 0)) is None  # 枠の外で起動しても動かない


def test_every_minutes_slot():
    worker.register_schedule("sync_sheet_config", task=lambda sf: {}, every_minutes=5)
    sched = worker.schedules()["sync_sheet_config"]
    assert worker.slot_start(sched, jst(2026, 11, 26, 9, 7, 30)) == jst(2026, 11, 26, 9, 5)
    assert worker.slot_start(sched, jst(2026, 11, 26, 9, 10)) == jst(2026, 11, 26, 9, 10)


def test_slot_runs_once_and_not_again_after_restart(session_factory, monkeypatch):
    calls = []
    worker.register_schedule("sheet_reconcile", task=lambda sf: calls.append(1) or {"ok": True, "fixed": 0},
                             at=(2, 10))
    now = jst(2026, 11, 26, 2, 15)
    monkeypatch.setattr(timeutil, "now_utc", lambda: now)
    s1 = worker.Scheduler(session_factory)
    assert s1.tick(now) == ["sheet_reconcile"]
    assert s1.tick(now + timedelta(minutes=5)) == []  # 同じ枠では 1 回だけ
    # worker を起動し直しても、同じ枠では動き直さない
    s2 = worker.Scheduler(session_factory)
    assert s2.tick(now + timedelta(minutes=10)) == []
    assert calls == [1]
    # 次の日の枠では動く
    assert s2.tick(jst(2026, 11, 27, 2, 10)) == ["sheet_reconcile"]
    with session_factory() as s:
        rows = s.scalars(select(SyncRun).where(SyncRun.kind == "schedule:sheet_reconcile")).all()
        assert len(rows) == 2 and all(r.ok for r in rows)
        assert rows[0].summary_json == {"ok": True, "fixed": 0}


def test_failure_is_recorded_and_waits_for_next_slot(session_factory, monkeypatch):
    def boom(sf):
        raise RuntimeError("sheets down")

    worker.register_schedule("sync_masters", task=boom, at=(0, 30))
    now = jst(2026, 11, 26, 0, 31)
    monkeypatch.setattr(timeutil, "now_utc", lambda: now)
    s = worker.Scheduler(session_factory)
    assert s.tick(now) == ["sync_masters"]
    assert s.tick(now + timedelta(minutes=1)) == []
    with session_factory() as db:
        row = db.scalars(select(SyncRun)).one()
        assert row.ok is False and "sheets down" in row.summary_json["error"]


def test_ok_false_result_is_recorded_as_failure(session_factory):
    worker.register_schedule("sync_sheet_config", task=lambda sf: {"ok": False, "reason": "empty"}, every_minutes=5)
    row = worker.run_task(session_factory, "sync_sheet_config")
    assert row.ok is False and row.summary_json == {"ok": False, "reason": "empty"}


def test_run_at_startup_runs_even_if_slot_taken(session_factory, monkeypatch):
    calls = []
    worker.register_schedule("gas_ping", task=lambda sf: calls.append(1) or {"ok": True}, every_minutes=5,
                             run_at_startup=True)
    now = jst(2026, 11, 26, 9, 1)
    monkeypatch.setattr(timeutil, "now_utc", lambda: now)
    worker.Scheduler(session_factory).tick(now)
    s = worker.Scheduler(session_factory)  # 起動し直し
    assert s.tick(now + timedelta(seconds=30)) == ["gas_ping"]  # 起動時にも 1 回
    assert s.tick(now + timedelta(minutes=1)) == []
    assert s.tick(now + timedelta(minutes=4)) == ["gas_ping"]
    assert calls == [1, 1, 1]


def test_run_task_passes_check_only_when_accepted(session_factory):
    seen = {}

    def reconcile(sf, *, check_only=False):
        seen["check_only"] = check_only
        return {"ok": True}

    worker.register_schedule("sheet_reconcile", task=reconcile, at=(2, 10))
    worker.register_schedule("purge", task=lambda sf: {"ok": True}, at=(3, 30))
    worker.run_task(session_factory, "sheet_reconcile", options={"check_only": True})
    assert seen == {"check_only": True}
    worker.run_task(session_factory, "purge", options={"check_only": True})  # 受けない関数には渡さない


def test_builtin_purge_schedule(session_factory):
    worker.register_builtin_schedules()
    sched = worker.schedules()["purge"]
    assert sched.at == (3, 30)
    now = timeutil.now_utc()
    with session_factory() as s:
        s.add_all([
            IdempotencyKey(user_email="a@x", key="old-key-0001", fn="f", state="done",
                           created_at=now - timedelta(hours=25)),
            IdempotencyKey(user_email="a@x", key="new-key-0001", fn="f", state="done", created_at=now),
            UserSession(token_hash="a" * 64, email="a@x", created_at=now - timedelta(days=91),
                        expires_at=now - timedelta(days=1)),
            UserSession(token_hash="b" * 64, email="b@x", created_at=now, expires_at=now + timedelta(days=90)),
            Job(kind="k", lane="gas", state="done", created_at=now - timedelta(days=100),
                updated_at=now - timedelta(days=91)),
            Job(kind="k", lane="gas", state="failed", created_at=now - timedelta(days=100),
                updated_at=now - timedelta(days=91)),
        ])
        s.commit()
    result = worker.purge(session_factory)
    assert result == {"ok": True, "idempotencyKeys": 1, "sessions": 1, "doneJobs": 1}
    with session_factory() as s:
        assert [j.state for j in s.scalars(select(Job))] == ["failed"]  # failed は消さない


def test_load_registrations_skips_missing_modules(monkeypatch):
    import sys
    import types

    calls = []
    fake = types.ModuleType("app.services.row_sync")
    fake.register_worker = lambda: calls.append("row_sync")
    monkeypatch.setitem(sys.modules, "app.services.row_sync", fake)
    loaded = worker.load_registrations()
    assert "app.services.row_sync" in loaded and calls == ["row_sync"]


def test_gas_lane_body_runs_one_job(session_factory):
    from app import jobs

    ran = []
    jobs.register_kind("internal.ping", lambda ctx: ran.append(ctx.job_id), lane="gas")
    with session_factory() as s:
        jobs.enqueue(s, "internal.ping", {})
        jobs.enqueue(s, "internal.ping", {})
        s.commit()
    w = worker.Worker(session_factory)
    w.gas_lane_body()
    assert len(ran) == 1
    w.gas_lane_body()
    assert len(ran) == 2


def test_skip_daily_on_startup_skips_current_daily_slot_only(session_factory, monkeypatch):
    """データ安全 中-5: SKIP_DAILY_ON_STARTUP=1 で起動したら、枠の中の毎日の予定はその回だけ飛ばす。"""
    calls = []
    worker.register_schedule("sheet_reconcile", task=lambda sf: calls.append("rec") or {"ok": True}, at=(2, 10))
    worker.register_schedule("gas_ping", task=lambda sf: calls.append("ping") or {"ok": True}, every_minutes=5,
                             run_at_startup=True)
    now = jst(2026, 11, 26, 2, 15)
    monkeypatch.setattr(timeutil, "now_utc", lambda: now)
    monkeypatch.setenv("SKIP_DAILY_ON_STARTUP", "1")
    s = worker.Scheduler(session_factory, skip_daily_on_startup=worker.skip_daily_on_startup_env())
    assert s.tick(now) == ["gas_ping"]
    assert s.tick(now + timedelta(minutes=5)) == ["gas_ping"]  # 同じ枠では照合は動かない
    assert worker.Scheduler(session_factory).tick(now + timedelta(minutes=6)) == ["gas_ping"]  # 起動し直しても
    assert "sheet_reconcile" in s.tick(jst(2026, 11, 27, 2, 10))  # 次の日は動く
    with session_factory() as db:
        skipped = db.scalars(select(SyncRun).where(SyncRun.kind == "schedule:sheet_reconcile")).first()
        assert skipped.summary_json == {"state": "skipped", "reason": "SKIP_DAILY_ON_STARTUP"}
    monkeypatch.setenv("SKIP_DAILY_ON_STARTUP", "0")
    assert worker.skip_daily_on_startup_env() is False
