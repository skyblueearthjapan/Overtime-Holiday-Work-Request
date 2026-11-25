"""services/sheet_codec: シリアル値と時間帯の往復・cell_kinds・auto・表示形式・WorkLogs の '@'・見出し。"""
from __future__ import annotations

from datetime import date, datetime, timezone
from zoneinfo import ZoneInfo

import pytest

from app.services import sheet_codec as c

TOKYO = ZoneInfo("Asia/Tokyo")
NY = ZoneInfo("America/New_York")
R = c.REQUESTS.rule
W = c.WORK_LOGS.rule
B = c.BATCH_LOGS.rule


def test_serial_epoch_and_millis():
    assert c.wall_to_serial(datetime(1899, 12, 30)) == 0
    assert c.wall_to_serial(datetime(2026, 3, 2)) == 46083
    s = c.wall_to_serial(datetime(2026, 3, 2, 8, 15, 30, 250000))
    assert c.serial_to_wall(s) == datetime(2026, 3, 2, 8, 15, 30, 250000)


@pytest.mark.parametrize("tz", [TOKYO, NY, ZoneInfo("Etc/UTC")])
def test_datetime_round_trip_in_sheet_timezone(tz):
    moment = datetime(2026, 11, 26, 21, 30, 5, 123000, tzinfo=timezone.utc)  # JST では 27 日 06:30
    serial = c.datetime_to_serial(moment, tz)
    assert c.serial_to_datetime(serial, tz) == moment
    wall = c.serial_to_wall(serial)
    assert wall == moment.astimezone(tz).replace(tzinfo=None)


def test_tokyo_wall_clock_serial():
    moment = datetime(2026, 11, 26, 21, 30, tzinfo=timezone.utc)
    assert c.serial_to_wall(c.datetime_to_serial(moment, TOKYO)) == datetime(2026, 11, 27, 6, 30)


def test_encode_datetime_serial_with_format():
    fmt = {"type": "DATE_TIME", "pattern": "m/d/yyyy h:mm:ss"}
    out = c.encode_cell(R("submittedAt"), datetime(2026, 3, 1, 23, 0, tzinfo=timezone.utc), "serial", TOKYO, fmt)
    assert out["cell"]["userEnteredValue"] == {"numberValue": 46083 + 8 / 24}
    assert out["cell"]["userEnteredFormat"]["numberFormat"] == fmt
    assert out["fields"] == "userEnteredValue,userEnteredFormat.numberFormat" and out["kind"] == "serial"


def test_encode_datetime_without_reference_uses_default_format():
    out = c.encode_cell(R("approvedAt"), datetime(2026, 3, 1, 23, 0, tzinfo=timezone.utc), "serial", TOKYO)
    assert out["cell"]["userEnteredFormat"]["numberFormat"] == c.DEFAULT_DATETIME_FORMAT
    out = c.encode_cell(R("targetDate"), date(2026, 3, 2), "auto", TOKYO)
    assert out["cell"] == {"userEnteredValue": {"numberValue": 46083.0},
                           "userEnteredFormat": {"numberFormat": c.DEFAULT_DATE_FORMAT}}


def test_encode_string_kinds():
    out = c.encode_cell(R("modifiedAt"), datetime(2026, 3, 1, 23, 0, tzinfo=timezone.utc), "string", TOKYO)
    assert out["cell"] == {"userEnteredValue": {"stringValue": "2026-03-02 08:00:00"}}
    assert out["fields"] == "userEnteredValue"
    out = c.encode_cell(R("targetDate"), date(2026, 3, 7), "string", TOKYO)
    assert out["cell"] == {"userEnteredValue": {"stringValue": "2026-03-07"}}
    out = c.encode_cell(R("targetDate"), "来週", "serial", TOKYO)  # 日付にならない元の文字はそのまま
    assert out["cell"] == {"userEnteredValue": {"stringValue": "来週"}}


def test_free_text_is_never_parsed():
    out = c.encode_cell(R("workContent"), "1/2", "string", TOKYO)
    assert out["cell"] == {"userEnteredValue": {"stringValue": "1/2"}} and out["fields"] == "userEnteredValue"
    assert c.encode_cell(R("workContent"), "'=1+1", "string", TOKYO)["cell"]["userEnteredValue"] == {
        "stringValue": "'=1+1"}


def test_blank_clears_value():
    for rule in (R("approvedAt"), R("approvedMinutes"), R("pdfFileId"), R("targetDate")):
        out = c.encode_cell(rule, None if rule.kind in ("datetime", "int") else "", "serial", TOKYO)
        assert "userEnteredValue" not in out["cell"]


def test_worklogs_jst_text_is_text_format():
    out = c.encode_cell(W("actualEndAt"), "2026-03-02 18:30:00", "serial", TOKYO)
    assert out["cell"] == {"userEnteredValue": {"stringValue": "2026-03-02 18:30:00"},
                           "userEnteredFormat": {"numberFormat": c.TEXT_FORMAT}}
    cleared = c.encode_cell(W("actualStartAt"), "", "string", TOKYO)
    assert cleared["cell"] == {"userEnteredFormat": {"numberFormat": c.TEXT_FORMAT}}


def test_worklogs_updated_at_kinds():
    serial = c.encode_cell(W("updatedAt"), "2026-03-02 18:30:05", "serial", TOKYO)
    assert serial["cell"]["userEnteredValue"]["numberValue"] == pytest.approx(
        c.wall_to_serial(datetime(2026, 3, 2, 18, 30, 5)))
    text = c.encode_cell(W("updatedAt"), "2026-03-02 18:30:05", "string", TOKYO)
    assert text["cell"]["userEnteredValue"] == {"stringValue": "2026-03-02 18:30:05"}
    assert text["cell"]["userEnteredFormat"]["numberFormat"] == c.TEXT_FORMAT
    # シートが東京でなければ、JST の時刻をシートの壁時計に直す
    ny = c.encode_cell(W("updatedAt"), "2026-03-02 18:30:05", "serial", NY)
    assert c.serial_to_wall(ny["cell"]["userEnteredValue"]["numberValue"]) == datetime(2026, 3, 2, 4, 30, 5)


def test_decode_kinds():
    assert c.decode_cell(R("targetDate"), 46083, "2026-03-02", TOKYO) == (date(2026, 3, 2), "serial")
    assert c.decode_cell(R("targetDate"), "2026-03-02", "2026-03-02", TOKYO) == (date(2026, 3, 2), "string")
    assert c.decode_cell(R("targetDate"), "未定", "未定", TOKYO) == ("未定", "string")
    assert c.decode_cell(R("modifiedAt"), "2026-03-10 08:00:00", "", TOKYO) == (
        datetime(2026, 3, 9, 23, 0, tzinfo=timezone.utc), "string")
    assert c.decode_cell(R("approvedMinutes"), 60, "60", TOKYO) == (60, "serial")
    assert c.decode_cell(R("approvedMinutes"), "", "", TOKYO) == (None, None)
    assert c.decode_cell(R("workerCode"), 1001, "1001", TOKYO) == ("1001", "serial")
    # 文字の列に日付のセルがあれば、表示の文字で持つ（GAS が日付に変えた自由入力）
    assert c.decode_cell(R("workContent"), 46023, "1/2/2026", TOKYO) == ("1/2/2026", "serial")
    with pytest.raises(c.CodecError):
        c.decode_cell(R("submittedAt"), "そのうち", "", TOKYO)


def test_decode_worklogs_old_date_cells_like_gas():
    """GAS の toJstString_: シートの時間帯で書式化した文字（壁時計のまま）を JST とみなす。"""
    serial = c.wall_to_serial(datetime(2026, 2, 10, 8, 0))
    assert c.decode_cell(W("actualStartAt"), serial, "", NY) == ("2026-02-10 08:00:00", "serial")
    assert c.decode_cell(W("actualStartAt"), "2026-02-10 08:00:00", "", NY) == ("2026-02-10 08:00:00", "string")
    assert c.decode_cell(W("updatedAt"), serial, "", TOKYO) == ("2026-02-10 08:00:00", "serial")


def test_expected_kind_chain():
    rule = R("targetDate")
    assert c.expected_kind(rule, {}, {}) == "auto"
    assert c.expected_kind(rule, {}, {"targetDate": "string"}) == "string"
    assert c.expected_kind(rule, {"targetDate": "serial"}, {"targetDate": "string"}) == "serial"
    assert c.expected_kind(rule, {"targetDate": "auto"}, {"targetDate": "string"}) == "string"
    assert c.expected_kind(R("submittedAt"), {}, {}) == "serial"


def test_same_value_seconds():
    a = datetime(2026, 3, 2, 0, 0, 0, 100000, tzinfo=timezone.utc)
    b = datetime(2026, 3, 2, 0, 0, 0, 900000, tzinfo=timezone.utc)
    assert c.same_value(R("submittedAt"), a, b)
    assert not c.same_value(R("submittedAt"), a, datetime(2026, 3, 2, 0, 0, 2, tzinfo=timezone.utc))
    assert c.same_value(R("approvedMinutes"), 60, 60) and not c.same_value(R("approvedMinutes"), 0, None)
    assert c.same_value(R("workerCode"), "1001", "1001")


def test_canonical_and_snapshot():
    row = {"target_date": date(2026, 3, 2), "target_date_text": "", "executed_at": None, "batch_name": "x",
           "success": 2, "skip": None, "fail": 0, "error_detail": ""}
    snap = c.snapshot(c.BATCH_LOGS, row)
    assert snap == {"実行日時": "", "バッチ名": "x", "対象日": "2026-03-02", "成功": 2, "スキップ": "", "失敗": 0,
                    "エラー詳細": ""}
    assert c.canonical(R("submittedAt"), datetime(2026, 3, 2, 1, 2, 3, 456789, tzinfo=timezone.utc)) == \
        "2026-03-02T01:02:03.456+00:00"


def test_header_index_like_gas():
    idx = c.header_index([" requestId ", "ｄｅｐｔ", "", "dept"])
    assert idx == {"requestId": 0, "dept": 3}
    assert c.missing_headers(c.BATCH_LOGS, ["実行日時", "バッチ名"]) == ["対象日", "成功", "スキップ", "失敗", "エラー詳細"]


def test_rules_cover_every_db_column():
    from app.models import BatchLog, Request, WorkLog

    for spec, model, extra in ((c.REQUESTS, Request, {"old_pdf_file_id_pending", "target_date_text",
                                                     "source_pdf_file_id", "source_pdf_folder_id"}),
                               (c.WORK_LOGS, WorkLog, {"id"}), (c.BATCH_LOGS, BatchLog, {"id", "target_date_text"})):
        attrs = {r.attr for r in spec.columns}
        columns = {col.name for col in model.__table__.columns} - {"row_version", "sheet_row", "cell_kinds"}
        assert columns - attrs == extra
    assert len(c.REQUESTS.columns) == 39


def test_looks_like_date_display():
    assert c.looks_like_date_display("3/2/2026 8:15:30")
    assert c.looks_like_date_display("2026-03-02")
    assert not c.looks_like_date_display("46083.34")
