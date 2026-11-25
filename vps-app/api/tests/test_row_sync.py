"""services/row_sync: 差分だけ書く・新しい行は末尾で作った順・ID の確かめと探し直し・見つからなければ止める・見出しの違いで
止まる・時間帯の違いで止まる・GAS が書いた 3 列を上書きしない・削除は値の書き換え・行やタブを消す要求を出さない。"""
from __future__ import annotations

from datetime import date, datetime, timedelta, timezone

import pytest
from sqlalchemy import select

from app import safety, sheet_queue
from app.models import BatchLog, Request, SheetSyncState, SyncRun, WorkLog
from app.services import row_sync
from app.services import sheet_codec as codec
from conftest import freeze_now
from fake_sheets import Cell
from sheet_fixtures import R1, R2, R4, REQUEST_HEADERS, STAGING_SID, WORKLOG_HEADERS, make_fake, migrate

NOW = datetime(2026, 11, 26, 3, 0, tzinfo=timezone.utc)


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
    book = fake.books[STAGING_SID]

    def run(**kw):
        return row_sync.run_once(session_factory, api=fake.api(settings), app_settings=settings, **kw)

    return type("Env", (), {"fake": fake, "book": book, "run": staticmethod(run), "sf": session_factory})


def req_cell(book, row, header):
    return book.tabs["Requests"].get(row - 1, REQUEST_HEADERS.index(header))


def wl_cell(book, row, header):
    return book.tabs["WorkLogs"].get(row - 1, WORKLOG_HEADERS.index(header))


def touched_columns(fake, tab_id):
    out = []
    for req in fake.batch_requests():
        body = req.get("updateCells")
        if body and body["start"]["sheetId"] == tab_id:
            out.append((body["start"]["rowIndex"] + 1, body["start"]["columnIndex"]))
    return out


def assert_only_allowed_requests(fake):
    for req in fake.batch_requests():
        (kind,) = req.keys()
        assert kind in ("updateCells", "appendDimension")
    for method, url, _ in fake.writes():
        assert url.endswith(":batchUpdate") and STAGING_SID in url and "values:" not in url


def test_nothing_dirty_reads_only_timezone_once(env):
    assert env.run()["written"] == 0
    assert all(m == "GET" for m, *_ in env.fake.calls)
    env.fake.calls.clear()
    assert env.run()["written"] == 0
    assert env.fake.calls == []  # 1 時間は読み直さない


def test_writes_only_changed_columns(env, db):
    r4 = db.get(Request, R4)
    r4.status = "approved"
    r4.approved_at = datetime(2026, 11, 26, 2, 0, tzinfo=timezone.utc)
    r4.approved_by = "approver-a@lineworks-local.info"
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    result = env.run()
    assert result["written"] == 1 and result["ok"]
    sid = env.book.tabs["Requests"].sheet_id
    cols = sorted(c for _, c in touched_columns(env.fake, sid))
    assert cols == sorted(REQUEST_HEADERS.index(h) for h in ("status(submitted/approved/canceled)", "approvedAt",
                                                              "approvedBy"))
    assert req_cell(env.book, 5, "status(submitted/approved/canceled)").value == "approved"
    cell = req_cell(env.book, 5, "approvedAt")
    assert isinstance(cell.value, float) and codec.looks_like_date_display(cell.formatted())
    assert cell.formatted() == "11/26/2026 11:00:00"  # 同じ列の表示形式を写す（JST の壁時計）
    st = db.get(SheetSyncState, ("requests", R4))
    db.refresh(st)
    assert st.synced_version == st.desired_version == 2 and st.state == "ok" and st.first_dirty_at is None
    assert st.synced_values["status(submitted/approved/canceled)"] == "approved"
    assert_only_allowed_requests(env.fake)
    # もう一度動かしても書かない
    env.fake.calls.clear()
    assert env.run()["written"] == 0 and env.fake.writes() == []


def test_new_rows_go_to_the_end_in_creation_order(env, db):
    for rid, n in (("aaaaaaaa-0000-4000-8000-000000000001", 1), ("aaaaaaaa-0000-4000-8000-000000000002", 2)):
        db.add(Request(request_id=rid, request_type="overtime", status="submitted", dept="機械加工",
                       worker_code="1001", worker_name="試験 一郎", target_date=date(2026, 11, 26),
                       submitted_at=NOW, approved_minutes=60 * n, reason="急ぎ作業を要するため"))
        db.flush()
        sheet_queue.mark_dirty(db, "requests", rid)
        wl = WorkLog(request_id=rid, updated_at="2026-11-26 12:00:00", updated_by="formSubmit",
                     cell_kinds={"updatedAt": "serial"})
        db.add(wl)
        db.flush()
        sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
    db.commit()
    result = env.run()
    assert result["written"] == 4
    rows = env.book.tabs["Requests"]
    assert rows.row_values(6)[0] == "aaaaaaaa-0000-4000-8000-000000000001"
    assert rows.row_values(7)[0] == "aaaaaaaa-0000-4000-8000-000000000002"
    assert req_cell(env.book, 6, "approvedMinutes").value == 60
    assert req_cell(env.book, 7, "targetDate").formatted() == "2026-11-26"  # 最後の行と同じ日付のセル
    assert req_cell(env.book, 6, "submittedAt").formatted() == "11/26/2026 12:00:00"
    assert req_cell(env.book, 6, "pdfFileId").value is None
    # WorkLogs は 1 行目の説明・2 行目の見出しに触れず、最後のデータ行（6 行目）の次へ
    assert env.book.tabs["WorkLogs"].row_values(1)[0].startswith("※実績")
    assert wl_cell(env.book, 7, "requestId").value == "aaaaaaaa-0000-4000-8000-000000000001"
    upd = wl_cell(env.book, 7, "updatedAt")
    assert isinstance(upd.value, float) and codec.looks_like_date_display(upd.formatted())  # 下書き行は日時のセル
    assert wl_cell(env.book, 7, "updatedBy").value == "formSubmit"
    assert [r.sheet_row for r in db.scalars(select(Request).where(Request.sheet_row > 5).order_by(Request.sheet_row))] \
        == [6, 7]
    assert_only_allowed_requests(env.fake)


def test_appends_rows_when_grid_is_full(env, db):
    env.book.tabs["Requests"].row_count = 5
    db.add(Request(request_id="bbbbbbbb-0000-4000-8000-000000000001", request_type="holiday", status="submitted",
                   target_date=date(2026, 11, 28), submitted_at=NOW))
    db.flush()
    sheet_queue.mark_dirty(db, "requests", "bbbbbbbb-0000-4000-8000-000000000001")
    db.commit()
    assert env.run()["written"] == 1
    appends = [r["appendDimension"] for r in env.fake.batch_requests() if "appendDimension" in r]
    assert appends == [{"sheetId": env.book.tabs["Requests"].sheet_id, "dimension": "ROWS", "length": 1}]
    assert env.book.tabs["Requests"].row_values(6)[0] == "bbbbbbbb-0000-4000-8000-000000000001"


def test_adopts_a_row_written_before_a_crash(env, db):
    """書いた後に DB の記録の前で落ちても、同じ ID の行があればそこへ書く（2 行にしない）。"""
    rid = "cccccccc-0000-4000-8000-000000000001"
    env.book.tabs["Requests"].cell(5, 0).value = rid
    db.add(Request(request_id=rid, request_type="overtime", status="submitted", target_date=date(2026, 11, 26),
                   submitted_at=NOW))
    db.flush()
    sheet_queue.mark_dirty(db, "requests", rid)
    db.commit()
    assert env.run()["written"] == 1
    assert env.book.tabs["Requests"].last_row() == 6
    assert db.get(Request, rid).sheet_row == 6


def test_relocates_when_rows_moved(env, db):
    """人が上に行を挿入して位置がずれても、ID の列から探し直して正しい行に書く。"""
    tab = env.book.tabs["Requests"]
    tab.rows.insert(1, [Cell("人が足した行")])
    r2 = db.get(Request, R2)
    r2.approved_by2 = "admin-z@lineworks-local.info"
    sheet_queue.mark_dirty(db, "requests", R2)
    db.commit()
    assert env.run()["written"] == 1
    assert req_cell(env.book, 4, "approvedBy2").value == "admin-z@lineworks-local.info"
    assert req_cell(env.book, 3, "approvedBy2").value == "admin-a@lineworks-local.info"  # 元の位置（今は R1）に書かない
    db.refresh(r2)
    assert r2.sheet_row == 4


def test_blocks_when_row_not_found(env, db):
    tab = env.book.tabs["Requests"]
    tab.cell(2, 0).value = "someone-else"  # R2 の行の ID が違い、ほかにも無い
    r2 = db.get(Request, R2)
    r2.reason_detail = "変えた"
    sheet_queue.mark_dirty(db, "requests", R2)
    db.commit()
    result = env.run()
    assert result["blocked"] == 1 and result["written"] == 0
    st = db.get(SheetSyncState, ("requests", R2))
    db.refresh(st)
    assert st.state == "blocked" and "見つかりません" in st.last_error
    assert env.fake.writes() == []
    # 管理者が直して unblock すれば書ける
    tab.cell(2, 0).value = R2
    assert row_sync.unblock(db, "requests", R2)
    db.commit()
    assert env.run()["written"] == 1
    assert row_sync.status(db)["blocked"] == []


def test_header_mismatch_halts_lane(env, db, monkeypatch):
    env.book.tabs["WorkLogs"].cell(1, WORKLOG_HEADERS.index("netMinutes")).value = "net"
    wl = db.scalars(select(WorkLog).where(WorkLog.request_id == R1)).one()
    wl.updated_by = "x"
    sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
    db.commit()
    result = env.run()
    assert result["ok"] is False and "netMinutes" in result["halted"]
    assert env.fake.writes() == []
    runs = db.scalars(select(SyncRun).where(SyncRun.kind == "row_sync")).all()
    assert len(runs) == 1 and runs[0].summary_json["halted"] is True
    assert row_sync.status(db)["halted"] is True
    # 止めている間は読みもしない
    env.fake.calls.clear()
    assert env.run()["halted"]
    assert env.fake.calls == []
    # 5 分後に確かめ直し、直っていれば再開して記録する
    env.book.tabs["WorkLogs"].cell(1, WORKLOG_HEADERS.index("netMinutes")).value = "netMinutes"
    freeze_now(monkeypatch, NOW + timedelta(minutes=6))
    assert env.run()["written"] == 1
    assert row_sync.status(db)["halted"] is False


def test_timezone_mismatch_halts_lane(env, db):
    env.book.tz = "America/New_York"
    r4 = db.get(Request, R4)
    r4.reason = "顧客対応のため（変更）"
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    result = env.run()
    assert "時間帯" in result["halted"]
    assert env.fake.writes() == []


def test_safety_refusal_halts_lane(env, db, monkeypatch):
    monkeypatch.setitem(safety.PURPOSE_TABS, "row_sync", frozenset({"BatchLogs"}))
    r4 = db.get(Request, R4)
    r4.reason = "変更"
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    result = env.run()
    assert result["ok"] is False and "Requests" in result["halted"]
    assert env.fake.writes() == []
    st = db.get(SheetSyncState, ("requests", R4))
    db.refresh(st)
    assert st.attempts == 0 and st.synced_version < st.desired_version  # 行の失敗として数えない


def test_does_not_overwrite_pdf_columns_written_by_gas(env, db):
    """GAS が PDF の 3 列をシートに直接書き、M が DB と synced_values を同じ値にした後、別の列が変わっても 3 列を書かない。"""
    r2 = db.get(Request, R2)
    # GAS がシートに書いた（DB の版は進めない）
    req_cell(env.book, 3, "pdfFileId").value = "gas-pdf-2"
    req_cell(env.book, 3, "pdfFolderId").value = "gas-folder-2"
    # M の record_pdf_result の形: DB の 3 列と synced_values を同じ値に
    r2.pdf_file_id, r2.pdf_folder_id = "gas-pdf-2", "gas-folder-2"
    r2.pdf_generated_at = datetime(2026, 11, 26, 2, 30, tzinfo=timezone.utc)
    st = db.get(SheetSyncState, ("requests", R2))
    synced = dict(st.synced_values)
    synced.update(codec.snapshot(codec.REQUESTS, r2))
    st.synced_values = synced
    # 二次承認（W）
    r2.approved_by2 = "admin-a@lineworks-local.info"
    sheet_queue.mark_dirty(db, "requests", R2)
    db.commit()
    req_cell(env.book, 3, "pdfFileId").value = "gas-pdf-2-newer"  # その後に GAS がまた書いた
    assert env.run()["written"] == 1
    cols = {c for _, c in touched_columns(env.fake, env.book.tabs["Requests"].sheet_id)}
    assert cols == {REQUEST_HEADERS.index("approvedBy2")}
    assert req_cell(env.book, 3, "pdfFileId").value == "gas-pdf-2-newer"


def test_clearing_pdf_columns_writes_blanks(env, db):
    r1 = db.get(Request, R1)
    r1.pdf_file_id, r1.pdf_folder_id, r1.pdf_generated_at = "", "", None
    r1.approved_by2, r1.approved_at2 = "", None
    sheet_queue.mark_dirty(db, "requests", R1)
    db.commit()
    assert env.run()["written"] == 1
    for h in ("pdfFileId", "pdfFolderId", "pdfGeneratedAt", "approvedBy2", "approvedAt2"):
        assert req_cell(env.book, 2, h).value is None


def test_delete_is_a_value_rewrite(env, db):
    """削除は status=canceled と WorkLogs の 5 列を空にする書き換え（行を消さない）。"""
    r1 = db.get(Request, R1)
    r1.status = "canceled"
    sheet_queue.mark_dirty(db, "requests", R1)
    wl = db.scalars(select(WorkLog).where(WorkLog.request_id == R1)).one()
    wl.actual_start_at = wl.actual_end_at = ""
    wl.actual_minutes = wl.break_minutes = wl.net_minutes = None
    wl.updated_at, wl.updated_by = "2026-11-26 12:00:00", "delete"
    wl.cell_kinds = {**wl.cell_kinds, "updatedAt": "string"}
    sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
    db.commit()
    before_rows = env.book.tabs["WorkLogs"].last_row()
    assert env.run()["written"] == 2
    assert req_cell(env.book, 2, "status(submitted/approved/canceled)").value == "canceled"
    for h in ("actualStartAt", "actualEndAt", "actualMinutes", "breakMinutes", "netMinutes"):
        assert wl_cell(env.book, 3, h).value is None
    upd = wl_cell(env.book, 3, "updatedAt")
    assert upd.value == "2026-11-26 12:00:00" and upd.fmt == codec.TEXT_FORMAT
    assert env.book.tabs["WorkLogs"].last_row() == before_rows
    assert_only_allowed_requests(env.fake)


def test_completion_writes_jst_text_with_text_format(env, db):
    rid = R4
    wl = WorkLog(request_id=rid, actual_start_at="2026-03-11 17:20:00", actual_end_at="2026-03-11 19:00:00",
                 actual_minutes=100, break_minutes=10, net_minutes=90, updated_at="2026-03-11 19:00:01",
                 updated_by="machine-01@lineworks-local.info", cell_kinds={"updatedAt": "string"})
    db.add(wl)
    db.flush()
    sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
    db.commit()
    assert env.run()["written"] == 1
    end = wl_cell(env.book, 7, "actualEndAt")
    assert end.value == "2026-03-11 19:00:00" and end.fmt == codec.TEXT_FORMAT
    assert wl_cell(env.book, 7, "netMinutes").value == 90


def test_batchlogs_new_row(env, db):
    bl = BatchLog(executed_at=datetime(2026, 11, 25, 21, 50, 0, 123000, tzinfo=timezone.utc),
                  batch_name="morningBatchNoMail", target_date=date(2026, 11, 25), success=3, skip=0, fail=1,
                  error_detail="a\nb")
    db.add(bl)
    db.flush()
    sheet_queue.mark_dirty(db, "batch_logs", str(bl.id))
    db.commit()
    assert env.run()["written"] == 1
    tab = env.book.tabs["BatchLogs"]
    assert tab.get(3, 0).formatted() == "11/26/2026 6:50:00"
    # 対象日は auto＝最後のデータ行と同じ型（控えの最後の行は文字）
    assert tab.row_values(4)[1:] == ["morningBatchNoMail", "2026-11-25", 3, 0, 1, "a\nb"]
    # M が型を決めれば（cell_kinds）その型で書く。表示形式は同じ列の日付のセルを写す
    bl.cell_kinds = {"対象日": "serial"}
    sheet_queue.mark_dirty(db, "batch_logs", str(bl.id))
    db.commit()
    assert env.run()["written"] == 1
    assert tab.get(3, 2).value == codec.date_to_serial(date(2026, 11, 25))
    assert tab.get(3, 2).formatted() == "2026-11-25"
    # 書いた後の ID の確かめ（実行日時とバッチ名）で、次の書き込みも同じ行へ
    bl.fail = 0
    sheet_queue.mark_dirty(db, "batch_logs", str(bl.id))
    db.commit()
    assert env.run()["written"] == 1
    assert tab.get(3, 5).value == 0


def test_one_row_failure_retries_with_backoff_and_others_proceed(env, db, monkeypatch):
    env.fake.fail_write_rows.add(("Requests", 3))
    for rid in (R2, R4):
        r = db.get(Request, rid)
        r.reason = "変更 " + rid[:4]
        sheet_queue.mark_dirty(db, "requests", rid)
    db.commit()
    result = env.run()
    assert result["written"] == 1 and result["failed"] == 1
    st = db.get(SheetSyncState, ("requests", R2))
    db.refresh(st)
    assert st.state == "retry" and st.attempts == 1 and st.next_try_at == NOW + timedelta(seconds=10)
    assert req_cell(env.book, 5, "reason").value.startswith("変更")
    env.fake.calls.clear()
    assert env.run()["written"] == 0  # 10 秒たつまで試さない
    env.fake.fail_write_rows.clear()
    freeze_now(monkeypatch, NOW + timedelta(seconds=11))
    assert env.run()["written"] == 1
    db.refresh(st)
    assert st.state == "ok" and st.attempts == 0


def test_retry_delays():
    assert [row_sync._retry_delay(n) for n in range(1, 8)] == [10, 30, 60, 120, 300, 300, 300]


def test_new_rows_wait_for_an_earlier_failed_new_row(env, db, monkeypatch):
    ids = ["dddddddd-0000-4000-8000-00000000000" + str(i) for i in (1, 2)]
    for rid in ids:
        db.add(Request(request_id=rid, request_type="overtime", status="submitted", target_date=date(2026, 11, 26),
                       submitted_at=NOW))
        db.flush()
        sheet_queue.mark_dirty(db, "requests", rid)
    db.commit()
    env.fake.fail_write_rows.add(("Requests", 6))
    result = env.run()
    assert result["failed"] == 1 and result["written"] == 0
    assert db.get(Request, ids[0]).sheet_row is None and db.get(Request, ids[1]).sheet_row is None
    assert env.book.tabs["Requests"].last_row() == 5  # 後ろの新しい行だけ先に書かない
    env.fake.fail_write_rows.clear()
    with env.sf() as s:
        works, _ = row_sync.select_work(s, 50, NOW + timedelta(seconds=5))
    assert [w.pk for w in works if w.pk in ids] == []  # 先の行が待ちの間は、後ろも選ばない
    freeze_now(monkeypatch, NOW + timedelta(seconds=11))
    assert env.run()["written"] == 2
    rows = env.book.tabs["Requests"]
    assert [rows.row_values(6)[0], rows.row_values(7)[0]] == ids


def test_status_reports_lag(env, db, monkeypatch):
    r4 = db.get(Request, R4)
    r4.reason = "x"
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    freeze_now(monkeypatch, NOW + timedelta(minutes=6))
    s = row_sync.status(db)
    assert s["pending"] == 1 and s["late"] is True and s["oldestDirtySeconds"] == 360


def test_kind_change_alone_is_written(env, db):
    """値が同じでも、W が型を変えたら（下書きの日時のセル → 完了の文字）書き直す。"""
    wl = db.scalars(select(WorkLog).where(WorkLog.request_id == R1)).one()
    wl.cell_kinds = {**wl.cell_kinds, "updatedAt": "serial"}
    sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
    db.commit()
    assert env.run()["written"] == 1
    assert isinstance(wl_cell(env.book, 3, "updatedAt").value, float)


def test_worker_resolves_run_once():
    from app import worker

    assert worker._resolve_row_sync() is row_sync.run_once


def test_transient_failure_backs_off_without_counting_row_attempts(env, db, monkeypatch):
    r4 = db.get(Request, R4)
    r4.reason = "x"
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    env.fake.fail_queue.extend([503] * 5)  # SheetsApi の再試行（4 回）も尽きる
    result = env.run()
    assert result["ok"] is False and "503" in result["error"]
    env.fake.calls.clear()
    assert env.run()["error"] == "一時的な失敗の後の待ち" and env.fake.calls == []
    st = db.get(SheetSyncState, ("requests", R4))
    db.refresh(st)
    assert st.attempts == 0 and st.state == "ok"
    freeze_now(monkeypatch, NOW + timedelta(seconds=31))
    assert env.run()["written"] == 1


def test_missing_delegate_does_not_crash(env, session_factory, settings, monkeypatch):
    def boom(sf, s):
        raise RuntimeError("委任がありません")

    monkeypatch.setattr(row_sync, "api_factory", boom)
    result = row_sync.run_once(session_factory, app_settings=settings)
    assert result["ok"] is False and "委任" in result["error"]


def test_merge_synced_keeps_columns_fixed_meanwhile():
    current = {"pdfFileId": "gas-new", "reason": "old", "__kinds__": {"approvedAt": "serial", "modifiedAt": "string"}}
    written = {"reason": "new", "__kinds__": {"modifiedAt": None}}
    merged = row_sync.merge_synced(current, written)
    assert merged == {"pdfFileId": "gas-new", "reason": "new", "__kinds__": {"approvedAt": "serial"}}


def test_rows_written_before_a_halt_are_recorded(env, db):
    """Requests を書いた後に WorkLogs の見出しの違いで止まっても、書いた Requests の行は記録する（2 重に足さない）。"""
    env.book.tabs["WorkLogs"].cell(1, WORKLOG_HEADERS.index("updatedBy")).value = "who"
    rid = "ffffffff-0000-4000-8000-000000000001"
    db.add(Request(request_id=rid, request_type="overtime", status="submitted", target_date=date(2026, 11, 26),
                   submitted_at=NOW))
    db.flush()
    sheet_queue.mark_dirty(db, "requests", rid)
    wl = WorkLog(request_id=rid, updated_at="2026-11-26 12:00:00", updated_by="formSubmit")
    db.add(wl)
    db.flush()
    sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
    db.commit()
    result = env.run()
    assert result["halted"] and result["written"] == 1
    assert db.get(Request, rid).sheet_row == 6
    st = db.get(SheetSyncState, ("work_logs", str(wl.id)))
    db.refresh(st)
    assert st.synced_version < st.desired_version


def test_with_real_record_pdf_result_pdf_columns_are_not_rewritten(env, db):
    """担当 M の gas_jobs.record_pdf_result（synced_values を見出しの名前・UTC の ISO で入れる）と組み合わせても、
    GAS がシートに書いた PDF の 3 列を行の反映が書き直さない。"""
    from app.services import gas_jobs

    req_cell(env.book, 3, "pdfFileId").value = "gas-pdf-2"
    req_cell(env.book, 3, "pdfFolderId").value = "gas-folder-2"
    assert gas_jobs.record_pdf_result(db, R2, {"pdfFileId": "gas-pdf-2", "pdfFolderId": "gas-folder-2",
                                               "pdfGeneratedAt": "2026-11-26 11:30:00"})
    db.commit()
    assert env.run()["written"] == 0  # 版は進んだが、シートと同じ値なので何も書かない
    assert env.fake.writes() == []
    r2 = db.get(Request, R2)
    r2.approved_by2 = "admin-a@lineworks-local.info"
    sheet_queue.mark_dirty(db, "requests", R2)
    db.commit()
    assert env.run()["written"] == 1
    cols = {c for _, c in touched_columns(env.fake, env.book.tabs["Requests"].sheet_id)}
    assert cols == {REQUEST_HEADERS.index("approvedBy2")}
