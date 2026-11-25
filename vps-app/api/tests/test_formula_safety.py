"""セキュリティ 高-1: VPS がシートへ書く値は、どの経路でも式にならない（'=1+1'・'+81…'・'-部品'・'@x' は文字のまま）。

- 行の反映・照合（row_sync・reconcile）: updateCells の stringValue で書く（式として読まれない＝RAW と同じ）。
- マスタ同期（master_sync）: GAS と同じ USER_ENTERED だが、= + - @ で始まる文字は ' を付けて文字にする。
- 移行の道具（tools/migrate_from_sheets.py）はシートを読むだけ。
"""
from __future__ import annotations

from pathlib import Path

import pytest
from sqlalchemy import select

from app import sheet_queue
from app.models import Order, Request
from app.services import master_sync, row_sync
from app.services import sheet_codec as codec
from fake_sheets import FormulaCell
from sheet_fixtures import R4, REQUEST_HEADERS, STAGING_SID, make_fake, migrate

DANGEROUS = ["=1+1", "+81-90-1234-5678", "-部品", "@x", "=HYPERLINK(\"https://example.com\",\"PDF\")", "=IMPORTDATA(A2)"]


@pytest.mark.parametrize("text", DANGEROUS)
def test_codec_writes_text_columns_as_string_value(text):
    rule = codec.REQUESTS.rule("workContent")
    enc = codec.encode_cell(rule, text, codec.STRING, "Asia/Tokyo")
    assert enc["cell"] == {"userEnteredValue": {"stringValue": text}}


def test_no_formula_value_is_ever_built():
    app_dir = Path(__file__).resolve().parents[1] / "app"
    tools_dir = Path(__file__).resolve().parents[2] / "tools"
    for path in list(app_dir.rglob("*.py")) + list(tools_dir.rglob("*.py")):
        assert "formulaValue" not in path.read_text(encoding="utf-8"), path


@pytest.mark.parametrize("text", DANGEROUS)
def test_row_sync_lands_user_text_as_text(text, engine, session_factory, settings, db, monkeypatch):
    from conftest import freeze_now
    from datetime import datetime, timezone

    freeze_now(monkeypatch, datetime(2026, 11, 26, 3, 0, tzinfo=timezone.utc))
    row_sync.reset_lane()
    fake = make_fake()
    migrate(fake, engine, settings)
    req = db.get(Request, R4)
    req.work_content = text
    req.reason_detail = text
    req.dept = text
    sheet_queue.mark_dirty(db, "requests", R4)
    db.commit()
    result = row_sync.run_once(session_factory, api=fake.api(settings), app_settings=settings)
    assert result["written"] == 1, result
    tab = fake.books[STAGING_SID].tabs["Requests"]
    for header in ("workContent", "reasonDetail", "dept"):
        cell = tab.get(4, REQUEST_HEADERS.index(header))
        assert cell.value == text and not isinstance(cell.value, FormulaCell)
    row_sync.reset_lane()


@pytest.mark.parametrize("text", DANGEROUS)
def test_master_sync_quotes_formula_leads(text, db, settings, monkeypatch):
    monkeypatch.setattr(master_sync, "sleep", lambda s: None)
    fake = make_fake()
    api = fake.api(settings)
    meta = api.get_spreadsheet(STAGING_SID)
    master_sync.write_tab(api, STAGING_SID, "工番マスタ", [["工番", "品名"], ["W1", text], ["-5", "+3"]], meta)
    values = fake.books[STAGING_SID].tabs["工番マスタ"].values()
    assert values[1][1] == text and not isinstance(values[1][1], FormulaCell)
    assert values[2] == [-5, 3]  # ただの数は GAS の setValues と同じく数になる


def test_formula_safe_rules():
    assert master_sync.formula_safe("=1+1") == "'=1+1"
    assert master_sync.formula_safe("-部品") == "'-部品"
    assert master_sync.formula_safe("@x") == "'@x"
    assert master_sync.formula_safe("-5") == "-5" and master_sync.formula_safe("+1.5") == "+1.5"
    assert master_sync.formula_safe("部品-1") == "部品-1" and master_sync.formula_safe(None) == ""
    assert master_sync.formula_safe(12) == 12
