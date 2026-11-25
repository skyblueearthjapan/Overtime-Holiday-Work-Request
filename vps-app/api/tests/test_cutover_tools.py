"""tools/cutover_backup.py（Drive の複製で控え）と tools/restore_tabs_from_copy.py（控えから写し戻す）。"""
from __future__ import annotations

import copy
import importlib.util
import json
import sys
import types
from datetime import datetime
from pathlib import Path

import httpx
import pytest

from app import config
from conftest import make_settings
from fake_sheets import FakeBook
from sheet_fixtures import STAGING_SID, make_fake

TOOLS = Path(__file__).resolve().parents[2] / "tools"


def _load(name):
    path = TOOLS / f"{name}.py"
    if not path.exists():
        pytest.skip(f"tools/{name}.py がありません")
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


# ===== cutover_backup =====


class FakeDrive:
    """Drive の files.copy と Sheets の行数の読み取り（ID・名前だけ）。"""

    def __init__(self, files):
        self.files = files  # id -> {tab: rows}
        self.calls = []
        self.n = 0
        self.after_copy = None

    def handler(self, request: httpx.Request) -> httpx.Response:
        self.calls.append((request.method, str(request.url).split("?")[0]))
        assert request.headers["authorization"] == "Bearer TOKEN"
        path = request.url.path
        if request.method == "POST" and path.endswith("/copy"):
            src = path.split("/")[-2]
            self.n += 1
            new = f"COPY-{self.n}"
            body = json.loads(request.content)
            self.files[new] = copy.deepcopy(self.files[src])
            if self.after_copy:
                self.after_copy(new, body["name"])
            return httpx.Response(200, json={"id": new, "name": body["name"]})
        sid = path.split("/")[3]
        if path.endswith("values:batchGet"):
            tabs = list(self.files[sid])
            return httpx.Response(200, json={"valueRanges": [{"values": [["x"] * self.files[sid][t]]} for t in tabs]})
        return httpx.Response(200, json={"sheets": [{"properties": {"title": t}} for t in self.files[sid]]})


def _http(tool, drive):
    return tool.Http(lambda: "TOKEN", transport=httpx.MockTransport(drive.handler))


def test_backup_copies_only_this_environments_db_and_ledger():
    tool = _load("cutover_backup")
    s = make_settings()
    drive = FakeDrive({s.staging_spreadsheet_id: {"Requests": 10, "WorkLogs": 12},
                       s.staging_ledger_spreadsheet_id: {"2026年度": 40}})
    report = tool.run(_http(tool, drive), s, now=datetime(2026, 11, 26, 21, 30))
    assert report["ok"] is True
    assert [b["sourceId"] for b in report["backups"]] == [s.staging_spreadsheet_id, s.staging_ledger_spreadsheet_id]
    assert report["backups"][0]["name"] == "【切替前の控え】残業・休日出勤申請 DB 2026-11-26 2130"
    copies = [u for m, u in drive.calls if m == "POST"]
    assert len(copies) == 2 and all(u.endswith("/copy") for u in copies)
    assert not any(config.PROD_SPREADSHEET_ID_FIXED in u for _, u in drive.calls)  # ステージングは本番に触れない


def test_backup_production_uses_fixed_ids_and_refuses_others():
    tool = _load("cutover_backup")
    prod = types.SimpleNamespace(is_staging=False, app_env="production",
                                 prod_spreadsheet_id=config.PROD_SPREADSHEET_ID_FIXED,
                                 prod_ledger_spreadsheet_id=config.PROD_LEDGER_SPREADSHEET_ID_FIXED)
    assert [fid for _, fid in tool.allowed_sources(prod)] == [config.PROD_SPREADSHEET_ID_FIXED,
                                                              config.PROD_LEDGER_SPREADSHEET_ID_FIXED]
    with pytest.raises(tool.BackupStop):
        tool.check_source(config.MASTER_SOURCE_SPREADSHEET_ID_FIXED, prod)
    with pytest.raises(tool.BackupStop):
        tool.check_source(config.PROD_SPREADSHEET_ID_FIXED, make_settings())  # ステージングから本番は複製しない
    bad = types.SimpleNamespace(is_staging=True, staging_spreadsheet_id=config.PROD_SPREADSHEET_ID_FIXED,
                                staging_ledger_spreadsheet_id="x")
    with pytest.raises(tool.BackupStop):
        tool.allowed_sources(bad)


def test_backup_reports_ng_when_row_counts_differ(capsys):
    tool = _load("cutover_backup")
    s = make_settings()
    drive = FakeDrive({s.staging_spreadsheet_id: {"Requests": 10}, s.staging_ledger_spreadsheet_id: {"t": 1}})

    def shorten(new_id, name):
        if " DB " in name:
            drive.files[new_id] = {"Requests": 9}

    drive.after_copy = shorten
    code = tool.main([], http=_http(tool, drive), settings=s)
    out = capsys.readouterr()
    assert code == 1 and '"ok": false' in out.out and "NG" in out.err
    assert "TOKEN" not in out.out + out.err


# ===== restore_tabs_from_copy =====

COPY_ID = "copy-of-db-0001"


@pytest.fixture
def restore_env(settings):
    tool = _load("restore_tabs_from_copy")
    fake = make_fake()
    db_book = fake.books[STAGING_SID]
    backup = FakeBook(sid=COPY_ID, title="【切替前の控え】残業・休日出勤申請 DB 2026-11-26 2130")
    backup.tabs = copy.deepcopy(db_book.tabs)
    fake.books[COPY_ID] = backup
    return tool, fake, db_book, backup


def test_restore_dry_run_does_not_write(restore_env, settings):
    tool, fake, db_book, backup = restore_env
    db_book.tabs["工番マスタ"].cell(1, 1).value = "壊れた値"
    fake.calls.clear()
    report = tool.run(fake.api(settings), settings, COPY_ID, "工番マスタ", apply=False)
    assert report["apply"] is False and report["tabs"][0]["diffCells"] == 1
    assert fake.writes() == []


def test_restore_apply_writes_values_raw_and_clears_surplus(restore_env, settings):
    tool, fake, db_book, backup = restore_env
    tab = db_book.tabs["工番マスタ"]
    original = backup.tabs["工番マスタ"].cell(1, 1).value
    tab.cell(1, 1).value = "壊れた値"
    extra_row = len(backup.tabs["工番マスタ"].rows) + 1
    tab.cell(extra_row, 0).value = "余りの行"
    backup.tabs["工番マスタ"].cell(2, 1).value = "=1+1"  # 式に見える文字も文字のまま戻す
    report = tool.run(fake.api(settings), settings, COPY_ID, "工番マスタ", apply=True)
    assert report["applied"] == ["工番マスタ"]
    assert tab.get(1, 1).value == original and tab.get(extra_row, 0).value in (None, "")
    assert tab.get(2, 1).value == "=1+1"
    bodies = [b for m, u, b in fake.writes() if "values:batchUpdate" in u]
    assert bodies and all(b["valueInputOption"] == "RAW" for b in bodies)
    assert all(STAGING_SID in u for m, u, b in fake.writes())  # 書くのはこの環境の DB だけ
    assert not fake.batch_requests()  # 行・タブを足す・消す要求を出さない


@pytest.mark.parametrize("tabs", ["all", "Requests,all"])
def test_all_excludes_protected_tabs(restore_env, settings, tabs):
    tool, fake, *_ = restore_env
    report = tool.run(fake.api(settings), settings, COPY_ID, tabs, apply=False)
    names = [t["tab"] for t in report["tabs"]]
    assert "DeptApprovers" not in names and "Settings" not in names and "Requests" in names


def test_protected_tabs_only_when_named(restore_env, settings):
    tool, fake, *_ = restore_env
    report = tool.run(fake.api(settings), settings, COPY_ID, "DeptApprovers,Settings", apply=False)
    assert [t["tab"] for t in report["tabs"]] == ["DeptApprovers", "Settings"]


@pytest.mark.parametrize("tabs", ["StampMap", "休憩マスタ", "FormMap", ""])
def test_other_tabs_are_refused(restore_env, settings, tabs):
    tool, fake, *_ = restore_env
    with pytest.raises(tool.RestoreStop):
        tool.run(fake.api(settings), settings, COPY_ID, tabs, apply=False)


def test_refuses_non_backup_copy_and_target_as_source(restore_env, settings):
    tool, fake, db_book, backup = restore_env
    with pytest.raises(tool.RestoreStop):
        tool.run(fake.api(settings), settings, STAGING_SID, "工番マスタ", apply=False)
    with pytest.raises(tool.RestoreStop):
        tool.run(fake.api(settings), settings, config.MASTER_SOURCE_SPREADSHEET_ID_FIXED, "工番マスタ", apply=False)
    backup.title = "別のファイル"
    with pytest.raises(tool.RestoreStop):
        tool.run(fake.api(settings), settings, COPY_ID, "工番マスタ", apply=False)


def test_restore_target_guard_is_this_environment_only(settings):
    from app import safety

    with pytest.raises(safety.UnsafeWriteError):
        safety.assert_restore_target(config.PROD_SPREADSHEET_ID_FIXED, "Requests", settings)
    with pytest.raises(safety.UnsafeWriteError):
        safety.assert_restore_target(STAGING_SID, "DeptApprovers", settings)
    assert safety.assert_restore_target(STAGING_SID, "DeptApprovers", settings, protected_listed=True) == STAGING_SID


def test_main_exit_codes(restore_env, settings, capsys):
    tool, fake, *_ = restore_env
    assert tool.main(["--from-copy", COPY_ID, "--tabs", "工番マスタ"], api=fake.api(settings), settings=settings) == 0
    assert "書いていません" in capsys.readouterr().err
    assert tool.main(["--from-copy", COPY_ID, "--tabs", "StampMap"], api=fake.api(settings), settings=settings) == 2
