"""services/master_sync: 80 人・1 列目が空の行・書式だけの列・Date の規則（時間帯に依らない）・USER_ENTERED・余りの行は値だけ消す・
空なら上書きしない・タブごとの失敗・DB の表・行やタブを消さない。"""
from __future__ import annotations

import pytest
from sqlalchemy import select

from app.config import MASTER_SOURCE_SPREADSHEET_ID_FIXED
from app.models import CompanyCalendar, Order, SyncRun, Worker
from app.services import master_sync
from fake_sheets import serial
from sheet_fixtures import STAGING_SID, make_fake


@pytest.fixture(autouse=True)
def _no_sleep(monkeypatch):
    monkeypatch.setattr(master_sync, "sleep", lambda s: None)


def _sync(db, settings, fake):
    result = master_sync.sync_masters(db, api=fake.api(settings), app_settings=settings)
    db.commit()
    return result


def test_workers_capped_at_80_after_dropping_blank_first_column(db, settings):
    fake = make_fake()
    result = _sync(db, settings, fake)
    assert result["ok"] is True
    assert result["tabs"]["作業員マスタ"] == {"ok": True, "rows": 80, "skipped_reason": "", "cols": 6}
    tab = fake.books[STAGING_SID].tabs["作業員マスタ"]
    values = tab.values()
    assert values[0] == ["作業員コード", "氏名", "部署", "担当業務", "拠点", "スタッフ種類"]
    assert len(values) == 81 and values[1][0] == 1001 and values[80][0] == 1080  # 元の並びの先頭 80 人
    assert all(v[1] not in ("1 列目が空の行", "1 列目が 0 の行") for v in values)
    assert all(len(v) <= 6 for v in values)  # 書式だけの列（見出しの無い列）は写さない
    workers = db.scalars(select(Worker).order_by(Worker.sort_order)).all()
    assert len(workers) == 80 and workers[0].worker_code == "1001" and workers[0].dept == "組立・塗装"
    assert workers[0].raw["拠点"] == "本社工場" and workers[0].google_account == ""


def test_surplus_rows_cleared_values_only_and_nothing_deleted(db, settings):
    fake = make_fake()
    tab = fake.books[STAGING_SID].tabs["作業員マスタ"]
    rows_before = tab.row_count
    _sync(db, settings, fake)
    assert tab.row_count == rows_before  # 行は消さない
    assert tab.last_row() == 81  # 古い 102 行の残りは値だけ消えた
    for method, url, body in fake.writes():
        assert "values:batchUpdate" in url or "values:batchClear" in url
        assert STAGING_SID in url
    clears = [b for m, u, b in fake.writes() if "batchClear" in u]
    assert {"ranges": ["'作業員マスタ'!A82:AE"]} in clears or any("'作業員マスタ'!A82" in r for c in clears
                                                                    for r in c["ranges"])
    assert all(m == "POST" for m, u, b in fake.writes())
    assert not fake.batch_requests()  # deleteDimension などの spreadsheets.batchUpdate を出さない


def test_user_entered_and_dates_are_timezone_independent(db, settings):
    fake = make_fake(source_tz="Etc/GMT")  # マスタ元のファイルの時間帯が東京でなくても日付がずれない
    _sync(db, settings, fake)
    body = [b for m, u, b in fake.writes() if "values:batchUpdate" in u]
    assert all(b["valueInputOption"] == "USER_ENTERED" for b in body)
    sent = {d["range"].split("!")[0]: d["values"] for b in body for d in b["data"]}
    # GMT の 2026-04-04 00:00 は JST の 09:00 → 日時の文字（GAS の normalizeDatesForWrite_ と同じ）
    assert sent["'社内カレンダーマスタ'"][1] == ["2026-04-04 09:00:00", "休日", "Sat", ""]
    assert sent["'工番マスタ'"][2][6] == "2025-09-01 22:45:00"


def test_dates_in_tokyo_source(db, settings):
    fake = make_fake()
    _sync(db, settings, fake)
    sent = {d["range"].split("!")[0]: d["values"] for m, u, b in fake.writes() if "values:batchUpdate" in u
            for d in b["data"]}
    assert sent["'社内カレンダーマスタ'"][1:3] == [["2026-04-04", "休日", "Sat", ""], ["2026-04-05", "休日", "Sun", ""]]
    assert [r[6] for r in sent["'工番マスタ'"][1:]] == ["2025-09-01", "2025-09-01 13:45:00", "9/1/2025"]
    assert len(sent["'社内カレンダーマスタ'"][0]) == 4  # 5 列目の空の見出しは写さない
    # DB のタブは USER_ENTERED で日付のセルになる（GAS の setValues と同じ）
    cal = fake.books[STAGING_SID].tabs["社内カレンダーマスタ"]
    assert cal.get(1, 0).value == serial("2026-04-04") and cal.get(1, 0).formatted() == "2026-04-04"
    rows = db.scalars(select(CompanyCalendar).order_by(CompanyCalendar.sort_order)).all()
    assert [(r.date_text, r.kind, r.weekday) for r in rows][:2] == [("2026-04-04", "休日", "Sat"),
                                                                    ("2026-04-05", "休日", "Sun")]
    orders = db.scalars(select(Order).order_by(Order.sort_order)).all()
    assert [(o.code, o.customer, o.dest, o.quantity) for o in orders][:2] == [("LW00001", "架空商事", "架空工業", "1"),
                                                                             ("LW00002", "架空商事", "架空工業", "2")]
    # 業務NOマスタは表を持たず、タブだけ（全角の見出しもそのまま）
    jobs = fake.books[STAGING_SID].tabs["業務NOマスタ"]
    assert jobs.row_values(1) == ["業務ＩＤ", "業務NO", "業務名", "部署", "説明"]
    assert jobs.get(2, 1).value == 42  # '0042' は USER_ENTERED で数になる（GAS の setValues と同じ）


def test_header_only_or_missing_source_does_not_overwrite(db, settings):
    fake = make_fake()
    _sync(db, settings, fake)
    fake2 = make_fake()
    src = fake2.books[MASTER_SOURCE_SPREADSHEET_ID_FIXED]
    src.tabs["工番マスタ"].rows = src.tabs["工番マスタ"].rows[:1]  # 見出しだけ
    del src.tabs["社内カレンダーマスタ"]
    src.tabs["作業員マスタ"].rows = [src.tabs["作業員マスタ"].rows[0]] + [r for r in src.tabs["作業員マスタ"].rows[1:]
                                                                    if r[0].value in ("", 0)]  # 有効な行 0
    before = {t: fake2.books[STAGING_SID].tabs[t].values() for t in ("工番マスタ", "社内カレンダーマスタ", "作業員マスタ")}
    result = _sync(db, settings, fake2)
    assert result["ok"] is False
    assert result["tabs"]["工番マスタ"]["skipped_reason"].startswith("見出しだけ")
    assert result["tabs"]["作業員マスタ"]["skipped_reason"].startswith("見出しだけ")
    assert result["tabs"]["社内カレンダーマスタ"]["skipped_reason"] == "元シートなし"
    assert result["tabs"]["業務NO.マスタ"]["ok"] is True  # ほかのタブは続ける
    for t, values in before.items():
        assert fake2.books[STAGING_SID].tabs[t].values() == values
    assert len(db.scalars(select(Order)).all()) == 3  # DB の表も前のまま
    assert len(db.scalars(select(Worker)).all()) == 80


def test_read_retries_like_gas(db, settings, monkeypatch):
    fake = make_fake()
    calls = {"n": 0}
    real = fake.api(settings).read_both

    class Flaky:
        def __getattr__(self, name):
            return getattr(api, name)

        def read_both(self, sid, ranges):
            calls["n"] += 1
            if calls["n"] <= 2:
                raise RuntimeError("Service error")
            return real(sid, ranges)

    api = fake.api(settings)
    result = master_sync.sync_masters(db, api=Flaky(), app_settings=settings)
    assert result["tabs"]["作業員マスタ"]["ok"] is True and calls["n"] == 2 + 4


def test_read_gives_up_after_three_attempts(db, settings):
    fake = make_fake()
    api = fake.api(settings)

    class Broken:
        def __getattr__(self, name):
            return getattr(api, name)

        def read_both(self, sid, ranges):
            if "作業員" in ranges[0]:
                raise RuntimeError("Service error")
            return api.read_both(sid, ranges)

    result = master_sync.sync_masters(db, api=Broken(), app_settings=settings)
    assert result["tabs"]["作業員マスタ"]["skipped_reason"].startswith("ERROR Read failed after all retries")
    assert result["tabs"]["工番マスタ"]["ok"] is True


def test_missing_destination_tab_is_skipped(db, settings):
    fake = make_fake()
    del fake.books[STAGING_SID].tabs["業務NOマスタ"]
    result = _sync(db, settings, fake)
    assert result["tabs"]["業務NO.マスタ"]["skipped_reason"] == "転記先シートなし"


def test_never_writes_to_master_source(db, settings):
    fake = make_fake()
    _sync(db, settings, fake)
    assert all(MASTER_SOURCE_SPREADSHEET_ID_FIXED not in u for m, u, b in fake.writes())


def test_records_sync_run_and_registers_schedule(db, settings):
    from app import worker

    _sync(db, settings, make_fake())
    run = db.scalars(select(SyncRun).where(SyncRun.kind == "sync_masters")).one()
    assert run.ok and run.summary_json["tabs"]["工番マスタ"]["rows"] == 3
    master_sync.register_worker()
    assert worker.schedules()["sync_masters"].at == (0, 30)


def test_compact_and_detect_last_col_like_gas():
    assert master_sync.detect_last_col(["a", "b", " ", ""]) == 2
    assert master_sync.detect_last_col(["", ""]) == 1
    data = master_sync.prepare([["h1", "h2"], ["", "x"], [0, "y"], [False, "z"], ["  ", "w"], ["k", "v"]],
                               [], master_sync.codec.JST, None)
    assert data == [["h1", "h2"], ["k", "v"]]
    assert master_sync.as_gas_reads("0042") == "42" and master_sync.as_gas_reads("2026-04-04") == "2026-04-04"
    assert master_sync.as_gas_reads(1.0) == "1"
