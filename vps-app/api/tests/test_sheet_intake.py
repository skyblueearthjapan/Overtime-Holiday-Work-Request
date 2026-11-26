"""services/sheet_intake（決定 #26）: Sebastian が Requests へ直接足した行を DB へ取り込む。

- 取り込んだ行は「反映済み」で入り、行の反映は書き直さない。承認すると同じ行の変わった列だけを書く。
- Sebastian の追記（INSERT_ROWS）と VPS の新しい行（values.append OVERWRITE）が同時でも、互いの行を上書きしない。
- requestId の空・重なり・読めない値の行は入れずに記録し、その行へは書かない。
- 何度動かしても同じ（冪等）。同時に 2 つは動かない（鍵）。production は GAS が移転モードのときだけ。
"""
from __future__ import annotations

import urllib.parse
from datetime import date, datetime, timedelta, timezone

import pytest
from sqlalchemy import func, select

from app import sheet_queue, worker
from app.models import Request, SheetSyncState, SyncRun, WorkLog
from app.services import ops_status, reconcile, row_sync, sheet_intake, write_gate
from app.services.sheets_api import API_BASE
from conftest import freeze_now
from fake_sheets import Cell
from sheet_fixtures import R4, REQUEST_HEADERS, STAGING_SID, make_fake, migrate

NOW = datetime(2026, 11, 26, 3, 0, tzinfo=timezone.utc)
S1 = "5eba5eba-0000-4000-8000-000000000001"
S2 = "5eba5eba-0000-4000-8000-000000000002"
V1 = "7a7a7a7a-0000-4000-8000-000000000001"


@pytest.fixture(autouse=True)
def _lane():
    row_sync.reset_lane()
    yield
    row_sync.reset_lane()


def seb_row(rid: str, **kw) -> list:
    """Sebastian が作る行（見出しの名前から。ほかの列は空）。"""
    d = {"requestId": rid, "requestType(overtime/holiday)": "overtime", "status(submitted/approved/canceled)": "submitted",
         "dept": "機械加工", "workerCode": "1001", "workerName": "試験 一郎",
         "workerEmail": "kikai-shared@lineworks-local.info", "targetDate": "2026-11-26",
         "submittedAt": "11/26/2026 9:05:00", "approvedMinutes": 60, "reason": "急ぎ作業を要するため",
         "workContent": "加工 1/2", "reasonDetail": "", "workNo1": "6401", "orderNo1": "LW00001",
         "customer1": "架空工業", "product1": "試作機"}
    d.update(kw)
    return [d.get(h, "") for h in REQUEST_HEADERS]


def sebastian_append(fake, row: list) -> None:
    """Sebastian と同じ呼び方: values().append(range='Requests!A:A', USER_ENTERED, INSERT_ROWS)。"""
    url = f"{API_BASE}/{STAGING_SID}/values/{urllib.parse.quote('Requests!A:A', safe='')}:append"
    resp = fake.request("POST", url, params={"valueInputOption": "USER_ENTERED", "insertDataOption": "INSERT_ROWS"},
                        json={"values": [row]})
    assert resp.status_code == 200, resp.text
    fake.calls.pop()  # VPS の呼び出しの記録に混ぜない


def vps_writes(fake) -> list:
    return [(m, u) for m, u, _ in fake.writes()]


@pytest.fixture
def env(engine, session_factory, settings, monkeypatch):
    freeze_now(monkeypatch, NOW)
    fake = make_fake()
    migrate(fake, engine, settings)
    fake.calls.clear()
    book = fake.books[STAGING_SID]
    monkeypatch.setattr(row_sync, "api_factory", lambda sf, s: fake.api(settings))

    def intake(**kw):
        return sheet_intake.run(session_factory, app_settings=settings, **kw)

    def sync():
        return row_sync.run_once(session_factory, api=fake.api(settings), app_settings=settings)

    return type("Env", (), {"fake": fake, "book": book, "intake": staticmethod(intake), "sync": staticmethod(sync),
                            "sf": session_factory, "settings": settings})


def cell(book, row, header):
    return book.tabs["Requests"].get(row - 1, REQUEST_HEADERS.index(header))


# ===== 取り込み =====


def test_ingests_sebastian_row_read_only_and_synced(env, db):
    sebastian_append(env.fake, seb_row(S1))
    result = env.intake()
    assert result["ok"] and result["ingested"] == [S1] and result["skipped"] == []
    assert vps_writes(env.fake) == []  # シートは読むだけ
    req = db.get(Request, S1)
    assert req.sheet_row == 6 and req.status == "submitted" and req.request_type == "overtime"
    assert req.target_date == date(2026, 11, 26) and req.worker_code == "1001" and req.approved_minutes == 60
    assert req.submitted_at == datetime(2026, 11, 26, 0, 5, tzinfo=timezone.utc)  # 9:05 JST
    assert req.work_content == "加工 1/2" and req.order_no1 == "LW00001"
    assert req.cell_kinds == {"targetDate": "serial", "submittedAt": "serial"}
    st = db.get(SheetSyncState, ("requests", S1))
    assert st.desired_version == st.synced_version == 1 and st.state == "ok"
    assert st.synced_values["status(submitted/approved/canceled)"] == "submitted"
    # GAS と同じく WorkLogs の行は作らない（実績の記録のときに作る）
    assert db.scalar(select(func.count()).select_from(WorkLog).where(WorkLog.request_id == S1)) == 0
    assert db.scalars(select(SyncRun).where(SyncRun.kind == sheet_intake.INGEST_KIND)).one().summary_json == {
        "ingested": [{"requestId": S1, "row": 6}]}
    # 行の反映は書き直さない。照合も差 0
    env.fake.calls.clear()
    env.sync()
    assert vps_writes(env.fake) == []
    check = reconcile.run(env.sf, check_only=True, api=env.fake.api(env.settings), app_settings=env.settings)
    assert check["diffCount"] == 0 and check["idMismatchCount"] == 0 and check["tabs"]["Requests"]["notInDb"] == []


def test_intake_is_idempotent(env, db):
    sebastian_append(env.fake, seb_row(S1))
    assert env.intake()["ingested"] == [S1]
    again = env.intake()
    assert again["ok"] and again["ingested"] == [] and again["candidates"] == 0
    assert db.scalar(select(func.count()).select_from(Request).where(Request.request_id == S1)) == 1
    # 同じ行を 2 度入れようとしても、2 度目は「もうある」で入らない
    rec = {"request_id": S1, "sheet_row": 6, "row_version": 1, "cell_kinds": {}}
    done, skipped = sheet_intake._insert(env.sf, [rec])
    assert done == [] and skipped[0]["reason"].startswith("もう DB にあります")


def test_ingested_row_approved_in_app_is_written_in_place(env, db):
    sebastian_append(env.fake, seb_row(S1))
    env.intake()
    before = {h: (cell(env.book, 6, h).value, cell(env.book, 6, h).fmt) for h in REQUEST_HEADERS}
    req = db.get(Request, S1)
    req.status = "approved"
    req.approved_at = datetime(2026, 11, 26, 2, 0, tzinfo=timezone.utc)
    req.approved_by = "approver-a@lineworks-local.info"
    sheet_queue.mark_dirty(db, "requests", S1)
    db.commit()
    env.fake.calls.clear()
    assert env.sync()["written"] == 1
    tab = env.book.tabs["Requests"]
    assert tab.last_row() == 6  # 新しい行を作らない
    assert cell(env.book, 6, "status(submitted/approved/canceled)").value == "approved"
    assert cell(env.book, 6, "approvedBy").value == "approver-a@lineworks-local.info"
    assert cell(env.book, 6, "approvedAt").formatted() == "11/26/2026 11:00:00"
    changed = {"status(submitted/approved/canceled)", "approvedAt", "approvedBy"}
    for h in REQUEST_HEADERS:
        if h not in changed:
            assert (cell(env.book, 6, h).value, cell(env.book, 6, h).fmt) == before[h], h
    assert not any(u.endswith(":append") for _, u in vps_writes(env.fake))


def test_ingested_row_is_found_again_after_rows_shift(env, db):
    """人やほかのシステムが上に行を挿入して行がずれても、requestId で探し直して同じ申請の行に書く。"""
    sebastian_append(env.fake, seb_row(S1))
    env.intake()
    tab = env.book.tabs["Requests"]
    tab.rows.insert(1, [Cell("")])  # 2 行目に空の行が入った（下がずれる）
    req = db.get(Request, S1)
    req.reason = "業者対応のため"
    sheet_queue.mark_dirty(db, "requests", S1)
    db.commit()
    assert env.sync()["written"] == 1
    assert cell(env.book, 7, "requestId").value == S1 and cell(env.book, 7, "reason").value == "業者対応のため"
    db.expire_all()
    assert db.get(Request, S1).sheet_row == 7


# ===== Sebastian と VPS が同時に足す =====


def _new_vps_request(db, rid=V1):
    db.add(Request(request_id=rid, request_type="overtime", status="submitted", dept="機械加工", worker_code="1002",
                   worker_name="試験 二郎", target_date=date(2026, 11, 26), submitted_at=NOW, approved_minutes=30,
                   reason="顧客対応のため"))
    db.flush()
    sheet_queue.mark_dirty(db, "requests", rid)
    db.commit()


def _once(fake, predicate, action):
    state = {"done": False}

    def hook(method, url, params, body):
        if not state["done"] and predicate(method, url):
            state["done"] = True
            action()

    fake.before_request.append(hook)
    return state


def _assert_both_rows_intact(env, seb_row_no, vps_row_no):
    assert cell(env.book, seb_row_no, "requestId").value == S1
    assert cell(env.book, seb_row_no, "workerEmail").value == "kikai-shared@lineworks-local.info"
    assert cell(env.book, seb_row_no, "approvedMinutes").value == 60
    assert cell(env.book, seb_row_no, "workerCode").value == 1001
    assert cell(env.book, vps_row_no, "requestId").value == V1
    assert cell(env.book, vps_row_no, "workerCode").value == "1002"
    assert cell(env.book, vps_row_no, "approvedMinutes").value == 30
    assert cell(env.book, vps_row_no, "workerEmail").value in (None, "")


def test_sebastian_appends_just_before_vps_takes_its_row(env, db):
    """VPS が最後の行を読んだ後、行を取る前に Sebastian が足した: VPS は次の行を取る（上書きしない）。"""
    _new_vps_request(db)
    hook = _once(env.fake, lambda m, u: m == "POST" and u.endswith(":append"),
                 lambda: sebastian_append(env.fake, seb_row(S1)))
    assert env.sync()["written"] == 1 and hook["done"]
    _assert_both_rows_intact(env, 6, 7)
    assert db.get(Request, V1).sheet_row == 7
    assert env.intake()["ingested"] == [S1]
    db.expire_all()
    assert db.get(Request, S1).sheet_row == 6


def test_sebastian_appends_between_vps_append_and_cell_write(env, db):
    """VPS が行を取った後、中身を書く前に Sebastian が足した: Sebastian は次の行へ入る（どちらも消えない）。"""
    _new_vps_request(db)
    hook = _once(env.fake, lambda m, u: m == "POST" and u.endswith(":batchUpdate") and "/values" not in u,
                 lambda: sebastian_append(env.fake, seb_row(S1)))
    assert env.sync()["written"] == 1 and hook["done"]
    _assert_both_rows_intact(env, 7, 6)
    assert db.get(Request, V1).sheet_row == 6
    result = env.intake()
    assert result["ingested"] == [S1]
    db.expire_all()
    assert db.get(Request, S1).sheet_row == 7
    # その後に両方を承認しても、それぞれの行に書く
    for rid in (S1, V1):
        r = db.get(Request, rid)
        r.status = "approved"
        sheet_queue.mark_dirty(db, "requests", rid)
    db.commit()
    assert env.sync()["written"] == 2
    assert cell(env.book, 6, "status(submitted/approved/canceled)").value == "approved"
    assert cell(env.book, 7, "status(submitted/approved/canceled)").value == "approved"
    _assert_both_rows_intact(env, 7, 6)


def test_vps_row_not_yet_written_is_not_ingested(env, db):
    """VPS が行だけ取って中身を書く前（requestId だけの行）でも、DB にある申請なので取り込まない。"""
    _new_vps_request(db)
    env.fake.fail_write_rows.add(("Requests", 6))
    assert env.sync()["failed"] == 1
    assert env.book.tabs["Requests"].row_values(6) == [V1]
    result = env.intake()
    assert result["ingested"] == [] and result["skipped"] == []


# ===== 取り込まない行 =====


def test_invalid_rows_are_skipped_recorded_and_never_written(env, db):
    tab = env.book.tabs["Requests"]
    sebastian_append(env.fake, seb_row(""))  # requestId が空
    sebastian_append(env.fake, seb_row(S2))  # 重なり（2 行）
    sebastian_append(env.fake, seb_row(S2, workerName="別の人"))
    sebastian_append(env.fake, seb_row("5eba5eba-0000-4000-8000-000000000003",
                                       **{"status(submitted/approved/canceled)": "modified"}))
    sebastian_append(env.fake, seb_row("5eba5eba-0000-4000-8000-000000000004", targetDate="来週"))
    sebastian_append(env.fake, seb_row("5eba5eba-0000-4000-8000-000000000005", approvedMinutes="六十"))
    sebastian_append(env.fake, seb_row(S1))  # 正しい行は入る
    snapshot = [[(c.value, c.fmt) for c in row] for row in tab.rows]
    result = env.intake()
    assert result["ingested"] == [S1]
    reasons = {(s["row"], s["requestId"]): s["reason"] for s in result["skipped"]}
    assert set(reasons) == {(6, ""), (7, S2), (8, S2), (9, "5eba5eba-0000-4000-8000-000000000003"),
                            (10, "5eba5eba-0000-4000-8000-000000000004"),
                            (11, "5eba5eba-0000-4000-8000-000000000005")}
    assert reasons[(6, "")] == "requestId が空です"
    assert "2 行" in reasons[(7, S2)]
    assert "status" in reasons[(9, "5eba5eba-0000-4000-8000-000000000003")]
    assert "targetDate" in reasons[(10, "5eba5eba-0000-4000-8000-000000000004")]
    assert "読めない" in reasons[(11, "5eba5eba-0000-4000-8000-000000000005")]
    assert db.scalar(select(func.count()).select_from(Request).where(Request.request_id.like("5eba5eba%"))) == 1
    # 管理者の欄に出る
    with env.sf() as s:
        s.add(SyncRun(kind=sheet_intake.SCHEDULE_KIND, started_at=NOW, finished_at=NOW, ok=True,
                      summary_json=result))
        s.commit()
        st = ops_status.get_ops_status(s)
    assert len(st["sheetIntake"]["skipped"]) == 6 and st["sheetIntake"]["ok"] is True
    # 行の反映・照合を動かしても、取り込まなかった行には書かない
    r4 = db.get(Request, R4)
    r4.reason = "x"
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    env.sync()
    reconcile.run(env.sf, api=env.fake.api(env.settings), app_settings=env.settings)
    after = [[(c.value, c.fmt) for c in row] for row in tab.rows]
    assert after[5:11] == snapshot[5:11]


def test_blank_rows_and_gaps_are_ignored(env):
    tab = env.book.tabs["Requests"]
    tab.rows.append([Cell(""), Cell("")])  # 空の行
    sebastian_append(env.fake, seb_row(S1))
    result = env.intake()
    assert result["ingested"] == [S1] and result["skipped"] == []


# ===== 鍵・関所・登録 =====


def test_concurrent_intake_is_refused_by_lock(env, db):
    sebastian_append(env.fake, seb_row(S1))
    nested: list = []
    _once(env.fake, lambda m, u: m == "GET" and u.endswith("values:batchGet") and len(nested) == 0,
          lambda: nested.append(env.intake()))
    result = env.intake()
    assert nested and nested[0]["busy"] is True and nested[0]["ingested"] == []
    assert result["ingested"] == [S1]
    assert db.scalar(select(func.count()).select_from(Request).where(Request.request_id == S1)) == 1


def test_lock_held_elsewhere_returns_busy(env):
    assert sheet_intake._local_lock.acquire(blocking=False)
    try:
        assert env.intake()["busy"] is True
    finally:
        sheet_intake._local_lock.release()


def test_production_runs_only_when_gas_is_migrated(db, monkeypatch):
    prod = type("S", (), {"is_staging": False})()
    monkeypatch.setattr(write_gate, "check", lambda *a, **k: (False, "migrated が true ではない"))
    ok, reason = sheet_intake.intake_allowed(db, prod)
    assert ok is False and "移転モード" in reason
    monkeypatch.setattr(write_gate, "check", lambda *a, **k: (True, ""))
    assert sheet_intake.intake_allowed(db, prod) == (True, "")
    staging = type("S", (), {"is_staging": True})()
    monkeypatch.setattr(write_gate, "check", lambda *a, **k: (False, "x"))
    assert sheet_intake.intake_allowed(db, staging) == (True, "")


def test_gate_closed_reads_nothing(env, monkeypatch):
    monkeypatch.setattr(sheet_intake, "intake_allowed", lambda db, s, now=None: (False, "閉じている"))
    sebastian_append(env.fake, seb_row(S1))
    result = env.intake()
    assert result["ok"] is False and result["active"] is False and env.fake.calls == []


def test_timezone_mismatch_stops_without_ingesting(env, db):
    env.book.tz = "UTC"
    sebastian_append(env.fake, seb_row(S1))
    result = env.intake()
    assert result["ok"] is False and "時間帯" in result["stopped"] and db.get(Request, S1) is None


def test_registered_every_minute():
    sheet_intake.register_worker()
    sched = worker.schedules()["sheet_intake"]
    assert sched.every_minutes == 1 and sched.at is None
    assert "app.services.sheet_intake" in worker.REGISTRATION_MODULES


def test_old_schedule_runs_are_purged(env, db):
    with env.sf() as s:
        s.add(SyncRun(kind=sheet_intake.SCHEDULE_KIND, started_at=NOW - timedelta(days=4), ok=True, summary_json={}))
        s.add(SyncRun(kind=sheet_intake.SCHEDULE_KIND, started_at=NOW - timedelta(hours=1), ok=True, summary_json={}))
        s.commit()
    env.intake()
    kinds = db.scalars(select(SyncRun.started_at).where(SyncRun.kind == sheet_intake.SCHEDULE_KIND)).all()
    assert len(kinds) == 1


# ===== 呼び口（運用の欄・Sebastian） =====


def test_run_now_records_and_shows_in_ops_status(env, db):
    sebastian_append(env.fake, seb_row(S1))
    result = sheet_intake.run_now(db, app_settings=env.settings)
    assert result["ingested"] == [S1]
    run = db.scalars(select(SyncRun).where(SyncRun.kind == sheet_intake.RUN_KIND)).one()
    assert run.ok and run.summary_json["source"] == "ops"
    st = ops_status.get_ops_status(db)
    assert st["sheetIntake"]["ingested"] == [S1] and st["lastRuns"]["sheet_intake"]["ok"] is True


def test_ops_function_is_admin_write(env, db):
    from app import rpc
    from app.domain import ops

    assert "ops_runSheetIntake" in rpc.OPS_FUNCTIONS and "ops_runSheetIntake" in rpc.WRITE_FUNCTIONS
    assert "ops_runSheetIntake" in rpc.ADMIN_ONLY_FUNCTIONS and ops.HANDLERS["ops_runSheetIntake"]


def test_ensure_request_ingests_on_demand(env, db):
    sebastian_append(env.fake, seb_row(S1))
    assert db.get(Request, S1) is None
    assert sheet_intake.ensure_request(db, S1, app_settings=env.settings) is True
    assert db.get(Request, S1).sheet_row == 6
    assert sheet_intake.ensure_request(db, "no-such-id", app_settings=env.settings) is False


def test_sebastian_completion_right_after_append(env, client, db, monkeypatch):
    """Sebastian が行を足した直後に完了を送っても、その場で取り込んで受け付ける（GAS と同じく見つかる）。"""
    from test_sebastian import call

    monkeypatch.setattr(sheet_intake, "_settings", lambda s: env.settings)
    sebastian_append(env.fake, seb_row(S1))
    res = call(client, "completeOvertime", requestId=S1)
    assert res.get("ok") is True and res["requestId"] == S1, res
    db.expire_all()
    assert db.get(Request, S1).sheet_row == 6


# ===== values.append の安全装置（行の反映が Requests の新しい行を取るときだけ） =====


def test_append_values_is_limited_to_requests_ids(env):
    from app.safety import UnsafeWriteError

    api = env.fake.api(env.settings)
    with pytest.raises(UnsafeWriteError):
        api.append_values(STAGING_SID, "WorkLogs", [["x"]])
    with pytest.raises(UnsafeWriteError):
        api.append_values(STAGING_SID, "作業員マスタ", [["x"]], purpose="master_sync")
    with pytest.raises(UnsafeWriteError):
        api.append_values(STAGING_SID, "Requests", [[1.5]])
    with pytest.raises(UnsafeWriteError):
        api.append_values("some-other-sheet", "Requests", [["x"]])
    assert vps_writes(env.fake) == []
    assert api.append_values(STAGING_SID, "Requests", [["=1+1"]]) == 6
    assert env.book.tabs["Requests"].row_values(6) == ["=1+1"]  # RAW なので式にならない
    assert type(env.book.tabs["Requests"].get(5, 0).value) is str


def test_status_skips_busy_runs(env, db):
    with env.sf() as s:
        s.add(SyncRun(kind=sheet_intake.SCHEDULE_KIND, started_at=NOW - timedelta(minutes=2), finished_at=NOW, ok=True,
                      summary_json={"ok": True, "active": True, "ingested": [], "skipped": [{"row": 9}]}))
        s.add(SyncRun(kind=sheet_intake.SCHEDULE_KIND, started_at=NOW - timedelta(minutes=1), finished_at=NOW, ok=True,
                      summary_json={"ok": True, "busy": True, "ingested": [], "skipped": []}))
        s.commit()
    assert sheet_intake.status(db)["skipped"] == [{"row": 9}]


# ===== 決定 #27: Sebastian の submittedAt（JST の壁時計の文字を GMT のシートへ USER_ENTERED）を直す =====


def test_submitted_at_jst_text_on_gmt_sheet_is_fixed_and_written_back(engine, session_factory, monkeypatch):
    from app.domain import dashboards
    from conftest import make_settings
    from fake_sheets import serial
    from test_domain_helpers import ctx

    freeze_now(monkeypatch, NOW)
    gmt = make_settings(db_sheet_timezone="Etc/GMT")
    fake = make_fake()
    book = fake.books[STAGING_SID]
    book.tz = "Etc/GMT"
    migrate(fake, engine, gmt)
    before_existing = session_factory().get(Request, R4).submitted_at  # 既存の行は変えない
    sebastian_append(fake, seb_row(S1, submittedAt="11/26/2026 8:43:31"))
    original = {h: (cell(book, 6, h).value, cell(book, 6, h).fmt) for h in REQUEST_HEADERS}
    assert original["submittedAt"][0] == serial("2026-11-26 08:43:31")  # JST の壁時計を GMT とみなした値
    result = sheet_intake.run(session_factory, api=fake.api(gmt), app_settings=gmt)
    assert result["ingested"] == [S1]
    with session_factory() as db:
        req = db.get(Request, S1)
        assert req.submitted_at == datetime(2026, 11, 25, 23, 43, 31, tzinfo=timezone.utc)  # 08:43:31 JST
        assert db.get(Request, R4).submitted_at == before_existing
        st = db.get(SheetSyncState, ("requests", S1))
        assert st.desired_version == 2 and st.synced_version == 1
        res = dashboards.api_getApproverDashboard(ctx(db, email="approver-a@lineworks-local.info", settings=gmt),
                                                  "機械加工")
        rows = [r for r in res["overtime"] if r["requestId"] == S1]
        assert len(rows) == 1 and rows[0]["submittedAt"] == "2026-11-26 08:43:31", res
        from app.domain import store
        assert store.jst_text(req.submitted_at) == "2026-11-26 08:43:31"
    fake.calls.clear()
    out = row_sync.run_once(session_factory, api=fake.api(gmt), app_settings=gmt)
    assert out["written"] == 1
    cols = [b["updateCells"]["start"]["columnIndex"] for b in fake.batch_requests() if "updateCells" in b]
    assert cols == [REQUEST_HEADERS.index("submittedAt")]  # ほかの列は書かない
    c = cell(book, 6, "submittedAt")
    assert c.value == serial("2026-11-25 23:43:31") and c.fmt and c.fmt["type"] == "DATE_TIME"
    for h in REQUEST_HEADERS:
        if h != "submittedAt":
            assert (cell(book, 6, h).value, cell(book, 6, h).fmt) == original[h], h
    # もう一度取り込んでも、反映しても何も変わらない
    assert sheet_intake.run(session_factory, api=fake.api(gmt), app_settings=gmt)["ingested"] == []
    fake.calls.clear()
    assert row_sync.run_once(session_factory, api=fake.api(gmt), app_settings=gmt)["written"] == 0
