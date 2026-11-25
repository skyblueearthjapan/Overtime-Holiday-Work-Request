"""services/sheets_api.SheetsApi: 書く関数は必ず安全装置を通る・読む関数は通さない・再試行。偽物のセッションで動かす。"""
from __future__ import annotations

import json

import pytest

from app.config import MASTER_SOURCE_SPREADSHEET_ID_FIXED, PROD_SPREADSHEET_ID_FIXED
from app.safety import UnsafeWriteError
from app.services.sheets_api import SheetsApi, SheetsError, a1, column_letter, spreadsheet_timezone
from conftest import TEST_ENV

STAGING = TEST_ENV["STAGING_SPREADSHEET_ID"]
META = {"properties": {"title": "DB", "timeZone": "Asia/Tokyo"},
        "sheets": [{"properties": {"sheetId": 11, "title": "Requests"}},
                   {"properties": {"sheetId": 22, "title": "工番マスタ"}},
                   {"properties": {"sheetId": 33, "title": "Settings"}}]}


class FakeResp:
    def __init__(self, status, body):
        self.status_code = status
        self._body = body
        self.text = json.dumps(body)
        self.content = self.text.encode()

    def json(self):
        return self._body


class FakeSession:
    def __init__(self, responder):
        self.calls = []
        self.responder = responder

    def request(self, method, url, params=None, json=None, timeout=None):
        self.calls.append((method, url, params, json))
        return self.responder(method, url, params, json)


def make_api(settings, responder=None, sleeps=None):
    def default(method, url, params, body):
        if method == "GET" and url.endswith(STAGING) or url.endswith(MASTER_SOURCE_SPREADSHEET_ID_FIXED):
            return FakeResp(200, META)
        if "values:batchGet" in url:
            return FakeResp(200, {"valueRanges": [{"values": [["a", 1]]}]})
        return FakeResp(200, {"ok": True})

    session = FakeSession(responder or default)
    api = SheetsApi(lambda: object(), session_factory=lambda creds: session, app_settings=settings,
                    sleep=(sleeps.append if sleeps is not None else (lambda s: None)))
    return api, session


def test_reads_are_not_guarded(settings):
    api, session = make_api(settings)
    assert api.get_spreadsheet(MASTER_SOURCE_SPREADSHEET_ID_FIXED)["properties"]["timeZone"] == "Asia/Tokyo"
    assert api.batch_get_values(MASTER_SOURCE_SPREADSHEET_ID_FIXED, ["'作業員マスタ'!A1:Z", "x"]) == [[["a", 1]], []]
    params = dict(session.calls[-1][2])
    assert params["valueRenderOption"] == "UNFORMATTED_VALUE" and params["dateTimeRenderOption"] == "SERIAL_NUMBER"
    with pytest.raises(ValueError):
        api.batch_get_values(STAGING, ["A1"], "RAW")


def test_update_cells_goes_to_allowed_tab(settings):
    api, session = make_api(settings)
    req = {"updateCells": {"range": {"sheetId": 11, "startRowIndex": 4, "endRowIndex": 5},
                           "rows": [{"values": [{"userEnteredValue": {"stringValue": "x"}}]}],
                           "fields": "userEnteredValue,userEnteredFormat.numberFormat"}}
    api.batch_update(STAGING, [req], tab="Requests", purpose="row_sync")
    method, url, _, body = session.calls[-1]
    assert method == "POST" and url.endswith(f"{STAGING}:batchUpdate") and body == {"requests": [req]}


@pytest.mark.parametrize("sid", [PROD_SPREADSHEET_ID_FIXED, MASTER_SOURCE_SPREADSHEET_ID_FIXED, "other"])
def test_staging_never_writes_elsewhere(settings, sid):
    api, session = make_api(settings)
    with pytest.raises(UnsafeWriteError):
        api.batch_update(sid, [{"updateCells": {"range": {"sheetId": 11}}}], tab="Requests", purpose="row_sync")
    with pytest.raises(UnsafeWriteError):
        api.values_batch_update(sid, [{"range": "'工番マスタ'!A1", "values": [["x"]]}], tab="工番マスタ",
                                purpose="master_sync")
    with pytest.raises(UnsafeWriteError):
        api.values_batch_clear(sid, ["'工番マスタ'!A5:H"], tab="工番マスタ", purpose="master_sync")
    with pytest.raises(UnsafeWriteError):
        api.append_rows(sid, "Requests", 5)
    assert all(c[0] == "GET" for c in session.calls)  # 書く要求は 1 つも出ていない


def test_forbidden_requests_and_tabs(settings):
    api, session = make_api(settings)
    with pytest.raises(UnsafeWriteError):
        api.batch_update(STAGING, [{"deleteDimension": {"range": {"sheetId": 11}}}], tab="Requests",
                         purpose="row_sync")
    with pytest.raises(UnsafeWriteError):  # sheetId が tab と違う（Settings のシート）
        api.batch_update(STAGING, [{"updateCells": {"range": {"sheetId": 33}}}], tab="Requests", purpose="row_sync")
    with pytest.raises(UnsafeWriteError):
        api.batch_update(STAGING, [{"updateCells": {"range": {"sheetId": 33}}}], tab="Settings", purpose="row_sync")
    with pytest.raises(UnsafeWriteError):  # 範囲がほかのタブ
        api.values_batch_update(STAGING, [{"range": "'Settings'!A1", "values": [["x"]]}], tab="工番マスタ",
                                purpose="master_sync")
    with pytest.raises(UnsafeWriteError):  # A 区分の値の消去は使わない
        api.values_batch_clear(STAGING, ["'Requests'!A2:A"], tab="Requests", purpose="row_sync")
    with pytest.raises(UnsafeWriteError):  # C 区分に行を足さない
        api.append_rows(STAGING, "工番マスタ", 1, purpose="master_sync")
    assert all(c[0] == "GET" for c in session.calls)


def test_master_sync_writes_and_clears(settings):
    api, session = make_api(settings)
    api.values_batch_update(STAGING, [{"range": "'工番マスタ'!A1:B1", "values": [["工番", "受注先"]]}],
                            tab="工番マスタ", purpose="master_sync", input_option="USER_ENTERED")
    assert session.calls[-1][3]["valueInputOption"] == "USER_ENTERED"
    api.values_batch_clear(STAGING, ["'工番マスタ'!A3:H"], tab="工番マスタ", purpose="master_sync")
    assert session.calls[-1][1].endswith("values:batchClear")


def test_append_rows_uses_append_dimension(settings):
    api, session = make_api(settings)
    api.append_rows(STAGING, "Requests", 100)
    body = session.calls[-1][3]
    assert body == {"requests": [{"appendDimension": {"sheetId": 11, "dimension": "ROWS", "length": 100}}]}
    with pytest.raises(ValueError):
        api.append_rows(STAGING, "Requests", 0)


def test_missing_tab_is_not_created(settings):
    api, _ = make_api(settings)
    with pytest.raises(SheetsError):
        api.append_rows(STAGING, "BatchLogs", 1)


def test_retries_on_429_and_5xx(settings):
    sleeps = []
    responses = [FakeResp(429, {}), FakeResp(503, {}), FakeResp(200, {"valueRanges": []})]
    api, session = make_api(settings, lambda *a: responses.pop(0), sleeps)
    assert api.batch_get_values(STAGING, ["A1"]) == [[]]
    assert sleeps == [1, 2]


def test_gives_up_after_retries(settings):
    api, _ = make_api(settings, lambda *a: FakeResp(500, {"error": "x"}), [])
    with pytest.raises(SheetsError) as exc:
        api.get_spreadsheet(STAGING)
    assert exc.value.status == 500


def test_client_error_is_not_retried(settings):
    calls = []
    api, _ = make_api(settings, lambda *a: calls.append(1) or FakeResp(403, {"error": "denied"}), [])
    with pytest.raises(SheetsError):
        api.get_spreadsheet(STAGING)
    assert len(calls) == 1


def test_helpers():
    assert a1("It's", "A1:B2") == "'It''s'!A1:B2"
    assert column_letter(0) == "A" and column_letter(25) == "Z" and column_letter(26) == "AA"
    assert spreadsheet_timezone(META) == "Asia/Tokyo"
