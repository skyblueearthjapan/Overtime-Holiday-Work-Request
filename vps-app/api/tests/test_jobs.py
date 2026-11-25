"""後回しの仕事（app/jobs.py）: レーン・優先度・同じ申請の順番・前が failed なら止まる・requires で待つと試行回数が
増えない・リース切れの拾い直し・再試行の間隔と 24 時間・GAS のレーンは全体で 1 本。"""
from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone

import pytest

from app import jobs, models, sheet_queue
from app.jobs import JobRunner, PermanentJobError, RetryLater, SkipJob

T0 = datetime(2026, 11, 26, 0, 0, tzinfo=timezone.utc)


class Clock:
    def __init__(self, at=T0):
        self.at = at

    def __call__(self):
        return self.at

    def advance(self, **kw):
        self.at = self.at + timedelta(**kw)


@pytest.fixture
def clock(monkeypatch):
    from app import timeutil

    c = Clock()
    monkeypatch.setattr(timeutil, "now_utc", c)
    return c


def runner(session_factory, clock, lane="gas"):
    return JobRunner(session_factory, lane=lane, now=clock)


def job(db, job_id):
    db.expire_all()
    return db.get(models.Job, job_id)


def enqueue(session_factory, kind="internal.afterApprove", **kw):
    with session_factory() as s:
        job_id = jobs.enqueue(s, kind, kw.pop("args", {"x": 1}), **kw)
        s.commit()
    return job_id


def test_enqueue_stores_everything(db, clock):
    job_id = jobs.enqueue(db, "internal.afterComplete", {"source": "screen"}, request_id="R1", priority=10,
                          requires=[{"table": "requests", "pk": "R1", "version": 3}], guard={"status": "approved"},
                          batch_id="B1")
    db.commit()
    row = job(db, job_id)
    assert row.lane == "gas" and row.state == "pending" and row.priority == 10 and row.batch_id == "B1"
    assert json.loads(row.args_json) == {"source": "screen"}
    assert json.loads(row.requires_json) == [{"table": "requests", "pk": "R1", "version": 3}]
    assert json.loads(row.guard_json) == {"status": "approved"}
    assert row.run_after == T0 and row.attempts == 0
    with pytest.raises(ValueError):
        jobs.enqueue(db, "x", {}, lane="mail")
    with pytest.raises(ValueError):
        jobs.enqueue(db, "x", {}, requires=[{"table": "requests"}])


def test_unregistered_kind_stays_pending(session_factory, db, clock):
    job_id = enqueue(session_factory)
    assert runner(session_factory, clock).run_once() == {"done": 0, "skipped": 0, "retry": 0, "waiting": 0,
                                                         "failed": 0}
    assert job(db, job_id).state == "pending"


def test_done_stores_result_and_context(session_factory, db, clock):
    seen = {}

    def handler(ctx):
        seen.update(kind=ctx.kind, args=ctx.args, request_id=ctx.request_id, attempt=ctx.attempt, guard=ctx.guard,
                    lane=ctx.lane)
        return {"ok": True, "pdf": {"status": "created"}}

    jobs.register_kind("internal.afterApprove", handler, lane="gas")
    job_id = enqueue(session_factory, request_id="R1", guard={"status": "approved"})
    assert runner(session_factory, clock).run_once()["done"] == 1
    row = job(db, job_id)
    assert row.state == "done" and row.finished_at == T0 and row.leased_until is None
    assert json.loads(row.result_json) == {"ok": True, "pdf": {"status": "created"}}
    assert seen == {"kind": "internal.afterApprove", "args": {"x": 1}, "request_id": "R1", "attempt": 1,
                    "guard": {"status": "approved"}, "lane": "gas"}


def test_lanes_are_separate(session_factory, db, clock):
    ran = []
    jobs.register_kind("internal.ping", lambda ctx: ran.append("gas"), lane="gas")
    jobs.register_kind("cleanup", lambda ctx: ran.append("default"), lane="default")
    enqueue(session_factory, "internal.ping")
    enqueue(session_factory, "cleanup", lane="default")
    runner(session_factory, clock, "default").run_once()
    assert ran == ["default"]
    runner(session_factory, clock, "gas").run_once()
    assert ran == ["default", "gas"]
    with pytest.raises(ValueError):
        jobs.register_kind("x", lambda ctx: None, lane="mail")


def test_gas_lane_runs_one_at_a_time_in_priority_order(session_factory, db, clock):
    ran = []
    jobs.register_kind("internal.afterComplete", lambda ctx: ran.append(ctx.request_id), lane="gas")
    enqueue(session_factory, "internal.afterComplete", request_id="A")
    enqueue(session_factory, "internal.afterComplete", request_id="B")
    enqueue(session_factory, "internal.afterComplete", request_id="SEB", priority=10)
    r = runner(session_factory, clock)
    assert len(r.claim(limit=10)) == 1  # GAS のレーンはいつも 1 件
    # 動いている gas の仕事があれば、ほかの runner も取り出さない（全体で 1 本）
    assert runner(session_factory, clock).claim() == []


def test_priority_then_id(session_factory, db, clock):
    ran = []
    jobs.register_kind("internal.afterComplete", lambda ctx: ran.append(ctx.request_id), lane="gas")
    for rid, prio in (("A", 0), ("B", 0), ("SEB", 10)):
        enqueue(session_factory, "internal.afterComplete", request_id=rid, priority=prio)
    r = runner(session_factory, clock)
    for _ in range(3):
        r.run_once()
    assert ran == ["SEB", "A", "B"]  # Sebastian（優先度 10）が先、同じなら積んだ順


def test_same_request_runs_in_order_and_waits_for_previous(session_factory, db, clock):
    ran = []
    jobs.register_kind("internal.afterApprove", lambda ctx: ran.append(ctx.job_id), lane="gas")
    jobs.register_kind("internal.regenerate", lambda ctx: ran.append(ctx.job_id), lane="gas")
    first = enqueue(session_factory, "internal.afterApprove", request_id="R1")
    second = enqueue(session_factory, "internal.regenerate", request_id="R1", priority=10)
    r = runner(session_factory, clock)
    r.run_once()
    r.run_once()
    assert ran == [first, second]  # 優先度が高くても、同じ申請の前の仕事が先


def test_failed_previous_job_blocks_same_request(session_factory, db, clock):
    ran = []

    def fail(ctx):
        raise PermanentJobError("ledger_header_mismatch")

    jobs.register_kind("internal.afterApprove", fail, lane="gas")
    jobs.register_kind("internal.regenerate", lambda ctx: ran.append(ctx.request_id), lane="gas")
    jobs.register_kind("internal.ledgerUpsert", lambda ctx: ran.append(ctx.request_id), lane="gas")
    first = enqueue(session_factory, "internal.afterApprove", request_id="R1")
    second = enqueue(session_factory, "internal.regenerate", request_id="R1")
    other = enqueue(session_factory, "internal.ledgerUpsert", request_id="R2")
    r = runner(session_factory, clock)
    assert r.run_once()["failed"] == 1
    assert job(db, first).state == "failed" and "ledger_header_mismatch" in job(db, first).last_error
    for _ in range(3):
        r.run_once()
    assert ran == ["R2"]  # 別の申請は先へ進む
    assert job(db, second).state == "pending" and job(db, other).state == "done"
    assert [j["id"] for j in jobs.list_blocked(db)] == [second]
    assert [j["id"] for j in jobs.list_failed(db)] == [first]
    # 管理者が再実行すると、前の仕事から順に流れる
    jobs.register_kind("internal.afterApprove", lambda ctx: ran.append("retry-" + ctx.request_id), lane="gas")
    assert jobs.retry_failed(db, first) is True
    db.commit()
    r.run_once()
    r.run_once()
    assert ran == ["R2", "retry-R1", "R1"]
    assert jobs.retry_failed(db, first) is False  # failed でなければ戻さない


def test_skipped_previous_job_lets_next_run(session_factory, db, clock):
    def skip(ctx):
        raise SkipJob("canceled")

    ran = []
    jobs.register_kind("internal.afterApprove", skip, lane="gas")
    jobs.register_kind("internal.afterDelete", lambda ctx: ran.append(ctx.request_id), lane="gas")
    first = enqueue(session_factory, "internal.afterApprove", request_id="R1")
    enqueue(session_factory, "internal.afterDelete", request_id="R1")
    r = runner(session_factory, clock)
    assert r.run_once()["skipped"] == 1
    assert job(db, first).state == "skipped" and job(db, first).last_error == "canceled"
    r.run_once()
    assert ran == ["R1"]


def _make_request(session_factory, rid="R1"):
    with session_factory() as s:
        s.add(models.Request(request_id=rid, request_type="overtime", status="approved"))
        s.flush()
        version = sheet_queue.mark_dirty(s, "requests", rid)
        s.commit()
    return version


def test_requires_waits_without_counting_attempts(session_factory, db, clock):
    version = _make_request(session_factory)
    ran = []
    jobs.register_kind("internal.afterApprove", lambda ctx: ran.append(1), lane="gas")
    job_id = enqueue(session_factory, request_id="R1", requires=[{"table": "requests", "pk": "R1", "version": version}])
    r = runner(session_factory, clock)
    for _ in range(3):
        assert r.run_once()["done"] == 0
        assert job(db, job_id).run_after == clock() + jobs.REQUIRES_WAIT
        clock.advance(seconds=5)
    row = job(db, job_id)
    assert row.attempts == 0 and row.state == "pending" and ran == []
    # 行の反映が済むと動く
    with session_factory() as s:
        state = s.get(models.SheetSyncState, ("requests", "R1"))
        state.synced_version = version
        s.commit()
    assert r.run_once()["done"] == 1
    assert job(db, job_id).attempts == 1


def test_retry_later_does_not_count_attempts(session_factory, db, clock):
    calls = []

    def handler(ctx):
        calls.append(ctx.attempt)
        if len(calls) < 3:
            raise RetryLater(60, "gas_ping not ok")
        return {"ok": True}

    jobs.register_kind("internal.afterApprove", handler, lane="gas")
    job_id = enqueue(session_factory, request_id="R1")
    r = runner(session_factory, clock)
    assert r.run_once()["waiting"] == 1
    row = job(db, job_id)
    assert row.attempts == 0 and row.run_after == clock() + timedelta(seconds=60) and row.last_error == "gas_ping not ok"
    assert r.run_once()["done"] == 0  # まだ時刻が来ていない
    clock.advance(seconds=60)
    r.run_once()
    clock.advance(seconds=60)
    assert r.run_once()["done"] == 1
    assert calls == [1, 1, 1]


def test_retry_backoff_then_failed_after_24_hours(session_factory, db, clock):
    exhausted = []

    def boom(ctx):
        raise RuntimeError("GAS busy")

    jobs.register_kind("internal.afterApprove", boom, lane="gas",
                       on_exhausted=lambda ctx, msg: exhausted.append((ctx.job_id, msg)))
    job_id = enqueue(session_factory, request_id="R1")
    r = runner(session_factory, clock)
    delays = []
    while True:
        outcome = r.run_once()
        row = job(db, job_id)
        if outcome["failed"]:
            break
        assert outcome["retry"] == 1
        delays.append(row.run_after - clock())
        clock.at = row.run_after
    minutes = [int(d.total_seconds() // 60) for d in delays]
    assert minutes[:6] == [1, 2, 4, 8, 16, 30]
    assert set(minutes[6:]) == {30}
    assert clock() <= T0 + timedelta(hours=24)
    assert row.state == "failed" and "GAS busy" in row.last_error
    assert exhausted and exhausted[0][0] == job_id


def test_permanent_error_fails_immediately(session_factory, db, clock):
    def bad(ctx):
        raise PermanentJobError("bad_request")

    jobs.register_kind("internal.afterApprove", bad, lane="gas")
    job_id = enqueue(session_factory)
    assert runner(session_factory, clock).run_once()["failed"] == 1
    assert job(db, job_id).attempts == 1


def test_handler_error_rolls_back_its_writes(session_factory, db, clock):
    def write_then_fail(ctx):
        ctx.db.add(models.Setting(key="X", value="1"))
        ctx.db.flush()
        raise RuntimeError("fail")

    jobs.register_kind("internal.afterApprove", write_then_fail, lane="gas")
    enqueue(session_factory)
    runner(session_factory, clock).run_once()
    db.expire_all()
    assert db.get(models.Setting, "X") is None


def test_expired_lease_is_picked_up_again(session_factory, db, clock):
    jobs.register_kind("internal.afterApprove", lambda ctx: None, lane="gas")
    job_id = enqueue(session_factory, request_id="R1")
    r = runner(session_factory, clock)
    claimed = r.claim()  # 取り出した後に worker が止まった（実行しない）
    assert [c.id for c in claimed] == [job_id]
    assert job(db, job_id).state == "running"
    assert r.claim() == []  # リースの間は拾わない
    clock.advance(minutes=11)  # gas のリース（10 分）が切れた
    again = r.claim()
    assert [c.id for c in again] == [job_id] and again[0].attempts == 2
    r.run_job(again[0])
    assert job(db, job_id).state == "done"


def test_default_lane_takes_several_but_one_per_request(session_factory, db, clock):
    ran = []
    jobs.register_kind("cleanup", lambda ctx: ran.append(ctx.request_id), lane="default")
    for rid in ("A", "A", "B", ""):
        enqueue(session_factory, "cleanup", lane="default", request_id=rid)
    claimed = runner(session_factory, clock, "default").claim(limit=10)
    assert sorted(c.request_id for c in claimed) == ["", "A", "B"]


def test_counts_and_purge(session_factory, db, clock):
    jobs.register_kind("cleanup", lambda ctx: None, lane="default")
    enqueue(session_factory, "cleanup", lane="default")
    enqueue(session_factory, "internal.ping")
    runner(session_factory, clock, "default").run_once()
    assert jobs.counts_by_state(db)["done"] == 1 and jobs.counts_by_state(db, "gas")["pending"] == 1
    assert jobs.purge_done(db, now=clock() + timedelta(days=89)) == 0
    assert jobs.purge_done(db, now=clock() + timedelta(days=91)) == 1
