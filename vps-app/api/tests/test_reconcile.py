"""services/reconcile: 移した直後は差 0・秒の単位・反映待ちの行は比べない・差を直して記録・DB が少なければ書かない・
--check は書かない・表示形式の無いシリアル値を見つける・シートにだけある行は消さずに報告。"""
from __future__ import annotations

from datetime import datetime, timezone

import pytest
from sqlalchemy import delete, select

from app import sheet_queue
from app.models import Request, SyncRun, WorkLog
from app.services import reconcile, row_sync
from conftest import freeze_now
from fake_sheets import Cell, serial
from sheet_fixtures import R1, R2, R4, REQUEST_HEADERS, STAGING_SID, make_fake, migrate

NOW = datetime(2026, 11, 26, 17, 10, tzinfo=timezone.utc)


@pytest.fixture(autouse=True)
def _lane():
    row_sync.reset_lane()
    yield
    row_sync.reset_lane()


@pytest.fixture
def env(engine, session_factory, settings, monkeypatch):
    freeze_now(monkeypatch, NOW)
    fake = make_fake()
    migrate(fake, engine, settings)
    fake.calls.clear()

    def run(check_only=False):
        return reconcile.run(session_factory, check_only=check_only, api=fake.api(settings), app_settings=settings)

    return type("Env", (), {"fake": fake, "book": fake.books[STAGING_SID], "run": staticmethod(run)})


def cell(book, row, header):
    return book.tabs["Requests"].get(row - 1, REQUEST_HEADERS.index(header))


def test_no_differences_right_after_migration(env, db):
    result = env.run(check_only=True)
    assert result["ok"] is True and result["diffCount"] == 0 and result["idMismatchCount"] == 0
    assert result["tabs"]["Requests"]["compared"] == 4 and result["tabs"]["WorkLogs"]["compared"] == 4
    assert env.fake.writes() == []
    run = db.scalars(select(SyncRun).where(SyncRun.kind == "sheet_reconcile")).one()
    assert run.ok and run.summary_json["checkOnly"] is True


def test_seconds_granularity(env):
    c = cell(env.book, 2, "submittedAt")
    c.value = serial("2026-03-02 08:15:30.900")  # DB は .250。同じ秒
    assert env.run(check_only=True)["diffCount"] == 0
    c.value = serial("2026-03-02 08:15:32")
    result = env.run(check_only=True)
    assert result["diffCount"] == 1 and result["ok"] is False
    assert result["tabs"]["Requests"]["diffs"][0]["header"] == "submittedAt"


def test_fixes_hand_edits_and_records(env, db):
    cell(env.book, 3, "workerName").value = "人が直した名前"
    cell(env.book, 5, "approvedMinutes").value = 999
    result = env.run()
    assert result["ok"] is True and result["diffCount"] == 2 and result["fixedRows"] == 2
    assert cell(env.book, 3, "workerName").value == "試験 二郎"
    assert cell(env.book, 5, "approvedMinutes").value == 90
    diffs = {(d["row"], d["header"]): d for d in result["tabs"]["Requests"]["diffs"]}
    assert diffs[(3, "workerName")]["sheet"] == "人が直した名前" and diffs[(3, "workerName")]["db"] == "試験 二郎"
    for req in env.fake.batch_requests():
        assert list(req) == ["updateCells"]
    assert env.run(check_only=True)["diffCount"] == 0
    runs = db.scalars(select(SyncRun).where(SyncRun.kind == "sheet_reconcile").order_by(SyncRun.id)).all()
    assert runs[0].summary_json["tabs"]["Requests"]["diffCount"] == 2


def test_check_only_does_not_write(env):
    cell(env.book, 3, "workerName").value = "人が直した名前"
    result = env.run(check_only=True)
    assert result["diffCount"] == 1 and result["ok"] is False
    assert env.fake.writes() == []
    assert cell(env.book, 3, "workerName").value == "人が直した名前"


def test_rows_waiting_for_sync_are_not_compared(env, db):
    r4 = db.get(Request, R4)
    r4.worker_name = "変えた（まだ反映していない）"
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    result = env.run()
    assert result["diffCount"] == 0 and result["tabs"]["Requests"]["skipped"] == 1
    assert env.fake.writes() == []


def test_serial_without_date_format_is_found_and_fixed(env):
    """表示形式の無いシリアル値は、戻した GAS が数として読む（gap-4 F-1）。照合で見つけて表示形式を当てる。"""
    c = cell(env.book, 2, "approvedAt")
    c.fmt = None
    assert c.formatted().startswith("46083.")
    result = env.run()
    assert [d["reason"] for d in result["tabs"]["Requests"]["diffs"]] == ["format"]
    assert c.fmt and c.fmt["type"] == "DATE_TIME" and c.formatted() == "3/2/2026 9:00:00"


def test_kind_difference_is_reported(env):
    c = cell(env.book, 3, "targetDate")  # 移行の時は文字
    c.value, c.fmt = serial("2026-03-07"), {"type": "DATE", "pattern": "yyyy-mm-dd"}
    result = env.run()
    assert [d["reason"] for d in result["tabs"]["Requests"]["diffs"]] == ["kind"]
    assert c.value == "2026-03-07"


def test_stops_without_writing_when_db_has_fewer_rows(env, db):
    tab = env.book.tabs["Requests"]
    tab.rows.append([Cell("eeeeeeee-0000-4000-8000-000000000001"), Cell("overtime"), Cell("submitted")])
    cell(env.book, 3, "workerName").value = "人が直した名前"
    result = env.run()
    assert result["ok"] is False and "少ない" in result["stopped"]
    assert env.fake.writes() == []


@pytest.mark.parametrize("model", [Request, WorkLog])
def test_stops_when_main_table_is_empty(env, db, model):
    db.execute(delete(model))
    db.commit()
    result = env.run()
    assert result["ok"] is False and "0 行" in result["stopped"]
    assert env.fake.calls == []


def test_extra_rows_are_reported_not_deleted(env):
    tab = env.book.tabs["Requests"]
    tab.rows.append([Cell(""), Cell("overtime"), Cell("submitted"), Cell("人のメモ")])
    result = env.run()
    assert result["tabs"]["Requests"]["extraRows"] == [{"row": 6, "id": [""]}]
    assert tab.row_values(6)[3] == "人のメモ"


def test_id_mismatch_is_reported(env):
    cell(env.book, 3, "requestId").value = "moved"
    result = env.run(check_only=True)
    assert result["tabs"]["Requests"]["idMismatch"][0]["pk"] == R2 and result["ok"] is False


def test_timezone_mismatch_stops(env):
    env.book.tz = "Europe/London"
    result = env.run()
    assert result["ok"] is False and "時間帯" in result["stopped"]


def test_registered_schedule_and_check_option(session_factory, monkeypatch):
    from app import worker

    reconcile.register_worker()
    sched = worker.schedules()["sheet_reconcile"]
    assert sched.at == (2, 10)
    seen = {}
    monkeypatch.setattr(reconcile, "run", lambda sf, check_only=False: seen.setdefault("check", check_only) or {})
    worker.run_task(session_factory, "sheet_reconcile", options={"check_only": True})
    assert seen["check"] is True


def test_db_rows_for_marks_blocked_as_skipped(env, db):
    from app.models import SheetSyncState

    st = db.get(SheetSyncState, ("requests", R1))
    st.state = "blocked"
    db.commit()
    rows = {d["pk"]: d for d in reconcile.db_rows_for(db, "requests")}
    assert rows[R1]["skip"] is True and rows[R2]["skip"] is False


def test_rows_moved_between_read_and_write_are_not_written(env, db, monkeypatch):
    """データ安全 低-1: 比べた後・書く前に行が並べ替えられたら、その行には書かない。"""
    cell(env.book, 3, "workerName").value = "人が直した名前"
    original = reconcile.row_sync.load_tab

    def swapped(api, sid, spec, meta, tz, *, need_tail):
        tab = env.book.tabs["Requests"]
        if spec.tab == "Requests":  # 書く直前に人が 2 行目と 3 行目を入れ替えた
            tab.rows[1], tab.rows[2] = tab.rows[2], tab.rows[1]
        return original(api, sid, spec, meta, tz, need_tail=need_tail)

    monkeypatch.setattr(reconcile.row_sync, "load_tab", swapped)
    result = env.run()
    assert result["diffCount"] == 1 and result["fixedRows"] == 0
    assert env.fake.batch_requests() == []
