"""tools/migrate_from_sheets.py: 本番 ID 以外を拒む・空でなければ止める・--truncate・modified で止める・孤児と行の無い申請・
sheet_row と cell_kinds と synced_values・移した直後の照合が差 0・0 行で止める・dry-run の報告。偽の Sheets API で動かす。"""
from __future__ import annotations

from datetime import date, datetime, timezone

import pytest
from sqlalchemy import func, select

from app.config import MASTER_SOURCE_SPREADSHEET_ID_FIXED, PROD_SPREADSHEET_ID_FIXED
from app.models import (ApproverMapRow, BatchLog, DeptApprover, IdempotencyKey, Job, Request, Setting, SheetSyncState,
                        SyncRun, UserSession, WorkLog)
from sheet_fixtures import ORPHAN, R1, R2, R3, R4, STAGING_SID, load_migrate_tool, make_fake, migrate


@pytest.fixture
def tool():
    return load_migrate_tool()


def _reads(tool, fake, settings):
    return tool.read_sheets(fake.api(settings), PROD_SPREADSHEET_ID_FIXED)


@pytest.mark.parametrize("sid", [STAGING_SID, MASTER_SOURCE_SPREADSHEET_ID_FIXED, "other"])
def test_refuses_any_other_spreadsheet(tool, settings, sid):
    fake = make_fake()
    with pytest.raises(tool.MigrationStop):
        tool.read_sheets(fake.api(settings), sid)
    assert fake.calls == []  # 読む前に止める


def test_main_refuses_other_id_before_any_call(tool, capsys):
    assert tool.main(["--source", "sheets", "--spreadsheet-id", STAGING_SID, "--database-url", "sqlite://"]) == 2
    assert "本番 DB スプレッドシート" in capsys.readouterr().err


def test_never_writes_to_the_source(tool, settings, engine):
    fake = make_fake()
    migrate(fake, engine, settings)
    assert fake.writes() == []
    assert all("values:batchGet" in u or u.endswith(PROD_SPREADSHEET_ID_FIXED) for _m, u, _p, _b in fake.calls)


def test_dry_run_report(tool, settings, engine, db):
    fake = make_fake()
    report = tool.run(_reads(tool, fake, settings), None, dry_run=True)
    s = report["summary"]
    assert s["counts"]["requests"] == 4 and s["counts"]["work_logs"] == 4 and s["counts"]["batch_logs"] == 2
    assert s["statusTypeCounts"] == {"approved/holiday": 1, "approved/overtime": 1, "canceled/overtime": 1,
                                     "submitted/overtime": 1}
    assert s["monthTypeCounts"] == {"2026-03/holiday": 1, "2026-03/overtime": 3}
    assert s["minuteSums"] == {"approvedMinutes": 660, "actualMinutes": 1150, "breakMinutes": 130, "netMinutes": 1020}
    assert s["minuteSumsByMonth"]["orphan"]["netMinutes"] == 480
    assert s["pdf"]["pdfFileIdCount"] == 1 and s["pdf"]["approvedNetWithoutPdf"] == 1
    assert s["others"]["orphanWorkLogs"] == 1 and s["others"]["requestsWithoutWorkLog"] == 1
    assert s["others"]["rejected"] == 1 and s["others"]["approvedBy2"] == 1
    assert s["inProgress"]["submitted"] == 1
    # R2: 提出 03-05 > 対象 03-07 ではない。R3: 提出 03-10 > 対象 03-09 → 過去申請
    assert s["retroactiveByMonthType"] == {"2026-03/overtime": 1}
    assert s["updatedByCounts"] == {"delete": 1, "sebastian-agent": 1, "unknown": 1}
    assert s["batchNameCounts"] == {"morningBatchNoMail": 2}
    assert s["cellKinds"]["Requests"]["targetDate:string"] == 1 and s["cellKinds"]["Requests"]["targetDate:serial"] == 3
    assert s["cellKinds"]["WorkLogs"]["updatedAt:serial"] == 1
    assert report["settingsForEnv"]["SOURCE_SSID"] == MASTER_SOURCE_SPREADSHEET_ID_FIXED
    assert report["settingsOverlaps"][0]["key"] == "ADMIN_EMAILS"
    assert report["reportOnlyTabs"] == {"StampMap": 1, "休憩マスタ": 1}
    assert report["droppedColumns"] == {}
    # dry-run は DB に書かない
    assert db.scalar(select(func.count()).select_from(Request)) == 0


def test_migrates_rows_with_positions_kinds_and_sync_state(tool, settings, engine, db):
    fake = make_fake()
    report = migrate(fake, engine, settings)
    assert report["verifyProblems"] == []
    assert report["postMigrationDiffs"] == {"Requests": 0, "WorkLogs": 0, "BatchLogs": 0}
    r1 = db.get(Request, R1)
    assert r1.sheet_row == 2 and r1.worker_code == "1001" and r1.target_date == date(2026, 3, 2)
    # 2026-03-02 08:15:30.250 JST
    assert r1.submitted_at == datetime(2026, 3, 1, 23, 15, 30, 250000, tzinfo=timezone.utc)
    assert r1.cell_kinds["submittedAt"] == "serial" and r1.work_content == "加工 1/2"
    r2 = db.get(Request, R2)
    assert r2.cell_kinds["targetDate"] == "string" and r2.target_date == date(2026, 3, 7)
    assert r2.cell_kinds["submittedAt"] == "string"
    assert r2.submitted_at == datetime(2026, 3, 5, 3, 0, tzinfo=timezone.utc)
    r3 = db.get(Request, R3)
    assert r3.cell_kinds["modifiedAt"] == "string" and r3.status == "canceled"
    # WorkLogs: 孤児も移し、行の無い申請（R4）に行を作らない。行番号は 3 行目から
    logs = db.scalars(select(WorkLog).order_by(WorkLog.id)).all()
    assert [(w.id, w.request_id, w.sheet_row) for w in logs] == [(1, R1, 3), (2, ORPHAN, 4), (3, R2, 5), (4, R3, 6)]
    assert logs[1].actual_start_at == "2026-02-10 08:00:00" and logs[1].cell_kinds["actualStartAt"] == "serial"
    assert logs[0].actual_minutes == 70 and logs[3].actual_minutes is None
    assert not db.scalars(select(WorkLog).where(WorkLog.request_id == R4)).all()
    bl = db.scalars(select(BatchLog).order_by(BatchLog.id)).all()
    assert bl[0].target_date == date(2026, 3, 3) and bl[1].cell_kinds["対象日"] == "string"
    assert bl[1].error_detail == "x: エラー\ny: エラー"
    st = db.get(SheetSyncState, ("requests", R1))
    assert st.desired_version == st.synced_version == r1.row_version == 1
    assert st.synced_values["pdfFileId"] == "pdf-file-1" and st.synced_values["__kinds__"]["approvedAt"] == "serial"
    assert db.scalar(select(func.count()).select_from(SheetSyncState)) == 10
    # B 区分
    assert db.get(Setting, "ADMIN_EMAILS").value == "owner@lineworks-local.info,admin-a@lineworks-local.info"
    assert db.get(Setting, "GENERAL_AFFAIRS_EMAILS").value == "soumu@lineworks-local.info"
    assert db.get(DeptApprover, 3).col_e == "notify-b@example.co.jp"
    assert [r.sheet_row for r in db.scalars(select(ApproverMapRow))] == [1, 2]


def test_refuses_when_target_not_empty_and_truncate_clears_runtime_tables(tool, settings, engine, db):
    fake = make_fake()
    migrate(fake, engine, settings)
    with pytest.raises(tool.MigrationStop, match="空ではありません"):
        migrate(fake, engine, settings)
    now = datetime(2026, 11, 26, tzinfo=timezone.utc)
    db.add(Job(kind="x", created_at=now, updated_at=now))
    db.add(IdempotencyKey(user_email="a", key="k", fn="f", created_at=now))
    db.add(SyncRun(kind="sync_masters", started_at=now))
    db.add(UserSession(token_hash="h", email="a", created_at=now, expires_at=now))
    db.commit()
    report = migrate(fake, engine, settings, truncate=True)
    assert report["verifyProblems"] == []
    for model in (Job, IdempotencyKey, SyncRun, UserSession):
        assert db.scalar(select(func.count()).select_from(model)) == 0
    assert db.scalar(select(func.count()).select_from(Request)) == 4


def test_refuses_when_only_runtime_tables_have_rows(tool, settings, engine, db):
    now = datetime(2026, 11, 26, tzinfo=timezone.utc)
    db.add(Job(kind="x", created_at=now, updated_at=now))
    db.commit()
    with pytest.raises(tool.MigrationStop, match="jobs=1"):
        migrate(make_fake(), engine, settings)


def _set(fake, tab, row, header, value):
    t = fake.books[PROD_SPREADSHEET_ID_FIXED].tabs[tab]
    headers = [c.value for c in t.rows[0 if tab != "WorkLogs" else 1]]
    t.cell(row - 1, headers.index(header)).value = value


def test_stops_on_modified_status(tool, settings, engine):
    fake = make_fake()
    _set(fake, "Requests", 5, "status(submitted/approved/canceled)", "modified")
    with pytest.raises(tool.MigrationStop, match="modified"):
        tool.run(_reads(tool, fake, settings), engine, dry_run=True)


@pytest.mark.parametrize("change,match", [
    (("Requests", 3, "requestId", R1), "重なって"),
    (("Requests", 3, "requestId", ""), "requestId が空"),
    (("Requests", 3, "requestType(overtime/holiday)", "night"), "requestType"),
    (("Requests", 3, "submittedAt", "きのう"), "日時として読めません"),
    (("WorkLogs", 3, "actualMinutes", "abc"), "数ではありません"),
])
def test_stops_on_bad_values(tool, settings, engine, change, match):
    fake = make_fake()
    _set(fake, *change)
    with pytest.raises(tool.MigrationStop, match=match):
        tool.run(_reads(tool, fake, settings), engine, dry_run=True)


def test_missing_column_stops_and_extra_column_is_reported(tool, settings, engine):
    fake = make_fake()
    t = fake.books[PROD_SPREADSHEET_ID_FIXED].tabs["Requests"]
    t.cell(0, 40).value = "memoByHuman"
    report = tool.run(_reads(tool, fake, settings), None, dry_run=True)
    assert report["droppedColumns"] == {"Requests": ["memoByHuman"]}
    t.cell(0, 38).value = "modifiedBy?"
    with pytest.raises(tool.MigrationStop, match="missingColumns"):
        tool.run(_reads(tool, fake, settings), None, dry_run=True)


def test_header_row_2_for_worklogs_is_required(tool, settings):
    fake = make_fake()
    wl = fake.books[PROD_SPREADSHEET_ID_FIXED].tabs["WorkLogs"]
    wl.rows[0], wl.rows[1] = wl.rows[1], wl.rows[0]  # 見出しが 1 行目にある形は止める
    with pytest.raises(tool.MigrationStop, match="WorkLogs"):
        tool.run(_reads(tool, fake, settings), None, dry_run=True)


@pytest.mark.parametrize("tab", ["Requests", "WorkLogs", "Settings"])
def test_stops_when_main_table_would_be_empty(tool, settings, engine, tab):
    fake = make_fake()
    t = fake.books[PROD_SPREADSHEET_ID_FIXED].tabs[tab]
    keep = {"Requests": 1, "WorkLogs": 2, "Settings": 2}[tab]
    t.rows = t.rows[:keep]
    with pytest.raises(tool.MigrationStop, match="0 行"):
        tool.run(_reads(tool, fake, settings), engine, dry_run=True)


def test_timezone_must_match_env(tool, settings):
    fake = make_fake()
    fake.books[PROD_SPREADSHEET_ID_FIXED].tz = "America/New_York"
    with pytest.raises(tool.MigrationStop, match="時間帯"):
        tool.run(_reads(tool, fake, settings), None, dry_run=True, expected_tz="Asia/Tokyo")


def test_sheet_timezone_is_used_for_serials(tool, settings, engine, db):
    """DB のファイルの時間帯が東京でなくても、シリアル値はその壁時計として読む。"""
    fake = make_fake()
    fake.books[PROD_SPREADSHEET_ID_FIXED].tz = "Etc/UTC"
    migrate(fake, engine, settings)
    assert db.get(Request, R1).submitted_at == datetime(2026, 3, 2, 8, 15, 30, 250000, tzinfo=timezone.utc)


def test_next_insert_after_migration_does_not_collide(tool, settings, engine, db):
    """連番を進めているので、移行の後に足す行の主キーがぶつからない（罠 9。SQLite は最大＋1）。"""
    migrate(make_fake(), engine, settings)
    w = WorkLog(request_id=R4)
    db.add(w)
    db.flush()
    assert w.id == 5


# ===== データ安全 高-1: ステージングへの移行は本番の PDF の ID を空にする =====


def test_staging_migration_blanks_production_pdf_ids(tool, settings, engine, db, session_factory, monkeypatch):
    from app.services import row_sync
    from sheet_fixtures import REQUEST_HEADERS, STAGING_SID

    fake = make_fake()
    report = migrate(fake, engine, settings, app_env="staging")
    assert report["verifyProblems"] == [] and report["stagingPdfIdsBlanked"] >= 1
    r1 = db.get(Request, R1)
    assert r1.pdf_file_id == "" and r1.pdf_folder_id == ""
    assert r1.source_pdf_file_id == "pdf-file-1" and r1.source_pdf_folder_id
    st = db.get(SheetSyncState, ("requests", R1))
    assert st.desired_version == r1.row_version == 2 and st.synced_version == 1
    assert st.synced_values["pdfFileId"] == "pdf-file-1"  # シートの今の値
    # 行の反映がステージングのシートの 2 列も空にする
    row_sync.reset_lane()
    result = row_sync.run_once(session_factory, api=fake.api(settings), app_settings=settings)
    assert result["ok"] and result["written"] >= 1
    tab = fake.books[STAGING_SID].tabs["Requests"]
    assert tab.get(1, REQUEST_HEADERS.index("pdfFileId")).value in (None, "")
    assert tab.get(1, REQUEST_HEADERS.index("pdfFolderId")).value in (None, "")
    row_sync.reset_lane()


def test_production_migration_keeps_pdf_ids(tool, settings, engine, db):
    report = migrate(make_fake(), engine, settings, app_env="production")
    assert "stagingPdfIdsBlanked" not in report
    r1 = db.get(Request, R1)
    assert r1.pdf_file_id == "pdf-file-1" and r1.source_pdf_file_id == ""
