"""services/config_sync: Settings の下の行が勝つ・キーの空白・空なら上書きしない・DeptApprovers の 5 列を持つ・
ApproverMap は 1 行目から生のまま・.env との食い違い・タブに書かない。"""
from __future__ import annotations

from sqlalchemy import select

from app import permissions
from app.models import ApproverMapRow, DeptApprover, Setting, SyncRun
from app.services import config_sync
from sheet_fixtures import STAGING_SID, make_fake


def _sync(db, settings, fake=None):
    fake = fake or make_fake()
    result = config_sync.sync_sheet_config(db, api=fake.api(settings), app_settings=settings)
    db.commit()
    return result, fake


def test_imports_settings_with_lower_row_winning_and_trimmed_keys(db, settings):
    result, fake = _sync(db, settings)
    assert result["ok"] is True
    values = {s.key: s.value for s in db.scalars(select(Setting))}
    assert values["ADMIN_EMAILS"] == "owner@lineworks-local.info,admin-a@lineworks-local.info"
    assert values["GENERAL_AFFAIRS_EMAILS"] == "soumu@lineworks-local.info"
    assert values["STAMP_HOLIDAY_FILE_ID"] == "stamp-file"
    assert db.get(Setting, "ADMIN_EMAILS").sheet_row == 11
    assert result["overlaps"] == [{"key": "ADMIN_EMAILS", "rows": [3, 11], "used": 11}]
    assert result["counts"] == {"Settings": 8, "DeptApprovers": 2, "ApproverMap": 2}
    # 管理者の判定に効く（GENERAL_AFFAIRS_EMAILS はキーの空白を取ったので効く）
    assert permissions.is_admin(db, "soumu@lineworks-local.info", settings)
    assert not permissions.is_admin(db, "old-admin@lineworks-local.info", settings)
    assert fake.writes() == []  # B 区分のタブには書かない


def test_reports_env_mismatches(db, settings):
    result, _ = _sync(db, settings)
    assert result["mismatches"] == [{"key": "ACCUMULATION_SS_ID", "sheet": "wrong-ledger-id",
                                     "env": settings.ledger_spreadsheet_id}]
    run = db.scalars(select(SyncRun).where(SyncRun.kind == "sync_sheet_config")).one()
    assert run.summary_json["mismatches"][0]["key"] == "ACCUMULATION_SS_ID"


def test_dept_approvers_keep_five_columns(db, settings):
    _sync(db, settings)
    rows = db.scalars(select(DeptApprover).order_by(DeptApprover.sheet_row)).all()
    assert [(r.sheet_row, r.dept, r.name, r.col_d, r.col_e) for r in rows] == [
        (2, "機械加工", "承認 A", "主任", "notify-a@example.co.jp"),
        (3, "組立・塗装", "承認 B", "係長", "notify-b@example.co.jp")]
    assert permissions.dept_approver_map(db)["組立・塗装"] == ["approver-b@lineworks-local.info",
                                                               "approver-c@lineworks-local.info"]


def test_approver_map_is_raw_from_row_1(db, settings):
    _sync(db, settings)
    rows = db.scalars(select(ApproverMapRow).order_by(ApproverMapRow.sheet_row)).all()
    assert rows[0].cells == ["※部署×Googleアカウントで承認画面の入室可否と承認権限を管理。"]
    assert rows[1].cells == ["部署", "承認者メール", "role(approver/admin)", "有効フラグ"]


def test_empty_or_unreadable_tabs_do_not_overwrite(db, settings):
    _sync(db, settings)
    fake = make_fake()
    book = fake.books[STAGING_SID]
    book.tabs["Settings"].rows = book.tabs["Settings"].rows[:2]  # 見出しだけ
    book.tabs["DeptApprovers"].rows = book.tabs["DeptApprovers"].rows[:1]
    del book.tabs["ApproverMap"]  # 読めない
    result, _ = _sync(db, settings, fake)
    assert result["ok"] is False
    assert set(result["skipped"]) == {"Settings", "DeptApprovers", "ApproverMap"}
    assert db.get(Setting, "ADMIN_EMAILS") is not None
    assert len(db.scalars(select(DeptApprover)).all()) == 2
    assert len(db.scalars(select(ApproverMapRow)).all()) == 2


def test_one_tab_failure_does_not_stop_others(db, settings):
    fake = make_fake()
    del fake.books[STAGING_SID].tabs["DeptApprovers"]
    result, _ = _sync(db, settings, fake)
    assert result["counts"]["Settings"] == 8 and "DeptApprovers" in result["skipped"]


def test_removed_key_disappears(db, settings):
    _sync(db, settings)
    fake = make_fake()
    rows = fake.books[STAGING_SID].tabs["Settings"].rows
    fake.books[STAGING_SID].tabs["Settings"].rows = [r for r in rows if r[0].value != "PDF_MODE"]
    _sync(db, settings, fake)
    assert db.get(Setting, "PDF_MODE") is None


def test_registered_every_configured_minutes():
    from app import worker

    config_sync.register_worker()
    assert worker.schedules()["sync_sheet_config"].every_minutes == 5


def test_scheduled_task_commits(session_factory, settings, monkeypatch):
    fake = make_fake()
    real = config_sync.sync_sheet_config
    monkeypatch.setattr(config_sync, "sync_sheet_config",
                        lambda db: real(db, api=fake.api(settings), app_settings=settings))
    config_sync._scheduled(session_factory)
    with session_factory() as db:
        assert db.get(Setting, "SOURCE_SSID") is not None
