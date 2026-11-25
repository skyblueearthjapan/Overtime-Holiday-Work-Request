"""Sebastian の入口 POST /api/sebastian（実装仕様 4-4。src/api.gs の doPost と同じ取り決め、01 §1.2 の表）。"""
from __future__ import annotations

import json
import types
from datetime import date

import pytest

from app.domain import gasjobs, sebastian
from conftest import freeze_now
from test_domain_helpers import add_request, add_wl, jobs_of, jst, wls

TOKEN = "test-sebastian-token-0123456789abcdef0123"
NOW = jst(2026, 11, 26, 19, 30, 15, 123)


@pytest.fixture(autouse=True)
def _now(monkeypatch):
    freeze_now(monkeypatch, NOW)


@pytest.fixture
def pdf_result(monkeypatch):
    box = {"result": None, "calls": []}

    def fake_wait(factory, job_id, timeout):
        box["calls"].append((job_id, timeout))
        return box["result"]

    monkeypatch.setattr(gasjobs, "wait_for_result", fake_wait)
    return box


def post(client, body, content_type="application/json"):
    raw = body if isinstance(body, (bytes, str)) else json.dumps(body)
    resp = client.post("/api/sebastian", content=raw, headers={"content-type": content_type})
    assert resp.status_code == 200
    return resp.json()


def call(client, action, **kw):
    return post(client, {"token": TOKEN, "action": action, **kw})


# ===== 入口・認証 =====


def test_ping_text_plain(client):
    assert post(client, {"token": TOKEN, "action": "ping"}, "text/plain") == {
        "ok": True, "pong": "2026-11-26T10:30:15.123Z"}


def test_invalid_json(client):
    res = post(client, b"{not json")
    assert set(res) == {"error"} and res["error"].startswith("invalid JSON: ")


@pytest.mark.parametrize("body", [b"", {"action": "ping"}, {"token": "wrong", "action": "ping"},
                                  {"token": 123, "action": "ping"}, b"null", b"[1,2]"])
def test_unauthorized(client, body):
    assert post(client, body) == {"error": "unauthorized"}


def test_token_not_set(db):
    s = types.SimpleNamespace(sebastian_api_token="")
    assert sebastian.handle(db, b'{"token":"","action":"ping"}', s) == {
        "error": "server misconfigured: SEBASTIAN_API_TOKEN not set"}


def test_unknown_action_and_missing_request_id(client):
    assert call(client, "explode") == {"error": "unknown action: explode", "action": "explode"}
    assert post(client, {"token": TOKEN}) == {"error": "unknown action: ", "action": ""}
    for action in ("completeOvertime", "markHolidayStart", "markHolidayDone"):
        assert call(client, action) == {"error": "requestId is required", "action": action}
        assert call(client, action, requestId="") == {"error": "requestId is required", "action": action}


@pytest.mark.parametrize("action,kind,error", [
    ("completeOvertime", "missing", "申請が見つかりません: R1"),
    ("completeOvertime", "holiday", "残業申請ではありません: R1"),
    ("completeOvertime", "canceled", "キャンセル済みの申請です: R1"),
    ("markHolidayStart", "missing", "申請が見つかりません: R1"),
    ("markHolidayStart", "overtime", "休日出勤申請ではありません: R1"),
    ("markHolidayStart", "canceled_h", "キャンセル済みの申請です: R1"),
    ("markHolidayDone", "overtime", "休日出勤申請ではありません: R1"),
    ("markHolidayDone", "holiday", "休日開始が記録されていません（先にstartを実行してください）: R1"),
    ("markHolidayDone", "holiday_empty_start", "休日開始が記録されていません（先にstartを実行してください）: R1"),
])
def test_action_errors(client, db, action, kind, error):
    if kind == "holiday":
        add_request(db, "R1", request_type="holiday")
    elif kind == "holiday_empty_start":
        add_request(db, "R1", request_type="holiday")
        add_wl(db, "R1")
    elif kind == "overtime":
        add_request(db, "R1")
    elif kind == "canceled":
        add_request(db, "R1", status="canceled")
    elif kind == "canceled_h":
        add_request(db, "R1", request_type="holiday", status="canceled")
    assert call(client, action, requestId="R1") == {"error": error, "action": action}
    db.expire_all()
    assert jobs_of(db) == []


# ===== completeOvertime =====


def test_complete_overtime_submitted(client, db, pdf_result):
    add_request(db, "R1", worker_name=" 山田　太郎 ", target=date(2026, 11, 26), worker_email="other@x.example")
    res = call(client, "completeOvertime", requestId="R1", manualEndTime="18:00")  # 手入力は無視
    assert res == {"ok": True, "requestId": "R1", "workerName": " 山田　太郎 ",
                   "targetDate": "2026-11-25T15:00:00.000Z", "actualMinutes": 130, "breakMinutes": 10,
                   "netMinutes": 120, "status": "submitted", "pdf": None}
    assert pdf_result["calls"] == []
    db.expire_all()
    [wl] = wls(db, "R1")
    assert wl.updated_by == "sebastian-agent" and wl.actual_end_at == "2026-11-26 19:30:15"
    assert wl.cell_kinds == {"updatedAt": "string"}
    [job] = jobs_of(db)
    assert job["kind"] == "internal.afterComplete" and job["priority"] == 10
    assert job["args"] == {"source": "sebastian", "statusAtEnqueue": "submitted"}  # wlOverride なし


def test_complete_overtime_uses_clock_windows_not_break_master_23(client, db, pdf_result, monkeypatch):
    add_request(db, "R1", target=date(2026, 11, 26))
    freeze_now(monkeypatch, jst(2026, 11, 26, 21, 40))
    res = call(client, "completeOvertime", requestId="R1")
    assert (res["actualMinutes"], res["breakMinutes"], res["netMinutes"]) == (260, 20, 240)


@pytest.mark.parametrize("result,pdf", [
    ({"ok": True, "pdf": {"status": "created", "pdfFileId": "F1", "pdfName": "n.pdf", "folderId": "D1",
                          "pdfGeneratedAt": "2026-11-26 19:31:00"}},
     {"ok": True, "pdfFileId": "F1", "pdfName": "n.pdf", "folderId": "D1"}),
    ({"ok": True, "pdf": {"status": "already", "pdfFileId": "F0"}}, {"already": True, "pdfFileId": "F0"}),
    ({"ok": False, "error": "internal", "message": "テンプレートが開けません", "pdf": {"status": "error",
                                                                              "error": "テンプレートが開けません"}},
     {"error": "テンプレートが開けません"}),
    ({"ok": False, "error": "busy", "message": "ロック待ち切れ"}, {"error": "ロック待ち切れ"}),
    (None, {"error": "PDF作成の結果待ちが時間切れになりました（作成は続いています）"}),
])
def test_complete_overtime_approved_waits_for_pdf(client, db, pdf_result, result, pdf):
    add_request(db, "R1", status="approved", target=date(2026, 11, 26), cell_kinds={"targetDate": "string"})
    pdf_result["result"] = result
    res = call(client, "completeOvertime", requestId="R1")
    assert res["status"] == "approved" and res["targetDate"] == "2026-11-26"
    assert res["pdf"] == pdf
    [(job_id, timeout)] = pdf_result["calls"]
    assert timeout == 40
    db.expire_all()
    assert jobs_of(db)[0]["args"]["statusAtEnqueue"] == "approved"  # 待つ前に確定している


def test_target_date_text_cell_is_returned_as_is(client, db, pdf_result):
    add_request(db, "R1", target=None, target_text="2026/11/26", cell_kinds={"targetDate": "string"})
    res = call(client, "completeOvertime", requestId="R1")
    assert res["targetDate"] == "2026/11/26" and res["actualMinutes"] == 130


def test_numeric_request_id_is_stringified(client, db):
    add_request(db, "12345", request_type="holiday")
    res = call(client, "markHolidayStart", requestId=12345)
    assert res["requestId"] == "12345" and res["ok"] is True


# ===== 休日 =====


def test_holiday_start_and_done(client, db, pdf_result):
    add_request(db, "H1", request_type="holiday", status="approved", target=date(2026, 11, 28))
    res = call(client, "markHolidayStart", requestId="H1", manualStartTime="08:00")
    assert res == {"ok": True, "requestId": "H1", "workerName": "山田太郎", "targetDate": "2026-11-27T15:00:00.000Z",
                   "actualStartAt": "2026-11-28 08:00:00"}
    db.expire_all()
    assert jobs_of(db) == []
    res = call(client, "markHolidayStart", requestId="H1", manualStartTime="8:00")  # 形が違えば今
    assert res["actualStartAt"] == "2026-11-26 19:30:15"
    call(client, "markHolidayStart", requestId="H1", manualStartTime="08:00")
    pdf_result["result"] = {"ok": True, "pdf": {"status": "created", "pdfFileId": "F", "pdfName": "p",
                                                "folderId": "D"}}
    res = call(client, "markHolidayDone", requestId="H1", manualEndTime="17:00")
    assert res == {"ok": True, "requestId": "H1", "workerName": "山田太郎", "targetDate": "2026-11-27T15:00:00.000Z",
                   "actualStartAt": "2026-11-28 08:00:00", "actualEndAt": "2026-11-28 17:00:00",
                   "actualMinutes": 540, "breakMinutes": 65, "netMinutes": 475, "status": "approved",
                   "pdf": {"ok": True, "pdfFileId": "F", "pdfName": "p", "folderId": "D"}}
    db.expire_all()
    [job] = jobs_of(db)
    assert job["priority"] == 10 and job["args"]["source"] == "sebastian" and "wlOverride" not in job["args"]
    assert wls(db, "H1")[0].updated_by == "sebastian-agent"


def test_holiday_start_bad_time_is_error_with_action(client, db):
    add_request(db, "H1", request_type="holiday")
    assert call(client, "markHolidayStart", requestId="H1", manualStartTime="25:00") == {
        "error": "時刻が不正です: 25:00", "action": "markHolidayStart"}


def test_pdf_from_result_shapes():
    assert sebastian.pdf_from_result({"ok": True, "pdf": {"status": "skipped"}}) is None
    assert sebastian.pdf_from_result({"already": True, "pdfFileId": "X"}) == {"already": True, "pdfFileId": "X"}
    assert sebastian.pdf_from_result({"ok": True, "pdfFileId": "X", "pdfName": "n", "folderId": "f"}) == {
        "ok": True, "pdfFileId": "X", "pdfName": "n", "folderId": "f"}


# ===== セキュリティ確認 中-1・正しさ 低-19 =====


def test_unexpected_exception_does_not_leak_internal_text(client, db, monkeypatch):
    add_request(db, "R1", request_type="holiday")

    def boom(*a, **k):
        raise RuntimeError('relation "work_logs" does not exist SELECT secret')

    monkeypatch.setattr(sebastian.worklog, "core_holiday_start", boom)
    res = call(client, "markHolidayStart", requestId="R1")
    assert res == {"error": "internal error", "action": "markHolidayStart"}


def test_exception_while_waiting_for_pdf_keeps_ok_true(client, db, monkeypatch):
    add_request(db, "R1", status="approved", target=date(2026, 11, 26))

    def boom(*a, **k):
        raise RuntimeError("db connection lost")

    monkeypatch.setattr(gasjobs, "wait_for_result", boom)
    res = call(client, "completeOvertime", requestId="R1")
    assert res["ok"] is True and res["status"] == "approved"
    assert res["pdf"] == {"error": sebastian.PDF_WAIT_ERROR}
    assert "db connection" not in json.dumps(res)
    db.expire_all()
    assert wls(db, "R1")  # WorkLogs は確定している
