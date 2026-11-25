"""作業実績（4-3 の 8・9・10・26。worklog.gs:143-342・adminApi.gs:432-518）。"""
from __future__ import annotations

from datetime import date

import pytest

from app.domain import worklog as wlmod
from app.rpc import RpcError
from test_domain_helpers import ADMIN, USER, add_request, add_wl, ctx, dirty, jobs_of, jst, wls


def done(db, rid, manual=None, now=jst(2026, 11, 26, 19, 30), email=USER):
    c = ctx(db, email=email, now=now)
    return wlmod.api_markOvertimeDone(c, rid) if manual is None else wlmod.api_markOvertimeDone(c, rid, manual)


def test_overtime_done_now_creates_worklog_and_enqueues(db):
    req = add_request(db, status="approved")
    res = done(db, req.request_id)
    assert res == {"ok": True, "requestId": req.request_id, "actualMinutes": 130, "breakMinutes": 10,
                   "netMinutes": 120, "pdf": None}
    [wl] = wls(db, req.request_id)
    assert (wl.actual_start_at, wl.actual_end_at) == ("2026-11-26 17:20:00", "2026-11-26 19:30:00")
    assert (wl.actual_minutes, wl.break_minutes, wl.net_minutes) == (130, 10, 120)
    assert wl.updated_by == USER and wl.updated_at == "2026-11-26 19:30:00"
    assert wl.cell_kinds == {"updatedAt": "string"}
    assert ("work_logs", str(wl.id)) in dirty(db)
    [job] = jobs_of(db)
    assert job["kind"] == "internal.afterComplete" and job["priority"] == 0
    assert job["args"] == {"source": "screen", "statusAtEnqueue": "approved", "wlOverride": {
        "actualStartAt": "2026-11-26 17:20:00", "actualEndAt": "2026-11-26 19:30:00", "actualMinutes": 130,
        "breakMinutes": 10, "netMinutes": 120}}
    # gas_jobs.requires_for_request: 一度も変えていない行（申請）はシートと同じなので待たない。変えた実績の行だけ待つ
    assert {r["table"] for r in job["requires"]} == {"work_logs"}


def test_overtime_done_submitted_still_enqueues_for_ledger(db):
    req = add_request(db, status="submitted")
    done(db, req.request_id)
    assert jobs_of(db)[0]["args"]["statusAtEnqueue"] == "submitted"


def test_overtime_done_manual_time_and_rerun_overwrites_q9(db):
    req = add_request(db, target=date(2026, 11, 25))
    res = done(db, req.request_id, "21:40")
    assert (res["actualMinutes"], res["breakMinutes"], res["netMinutes"]) == (260, 20, 240)
    res = done(db, req.request_id, "18:00")
    assert (res["actualMinutes"], res["breakMinutes"], res["netMinutes"]) == (40, 0, 40)
    [wl] = wls(db, req.request_id)  # 上書き（行は増えない）
    assert wl.actual_end_at == "2026-11-25 18:00:00"
    assert len(jobs_of(db)) == 2


def test_overtime_done_bad_manual_format_uses_now_g13(db):
    req = add_request(db)
    res = done(db, req.request_id, "9:30", now=jst(2026, 11, 26, 18, 20))
    assert res["actualMinutes"] == 60
    assert wls(db, req.request_id)[0].actual_end_at == "2026-11-26 18:20:00"


def test_overtime_done_before_1720_is_zero(db):
    req = add_request(db)
    res = done(db, req.request_id, now=jst(2026, 11, 26, 16, 0))
    assert (res["actualMinutes"], res["breakMinutes"], res["netMinutes"]) == (0, 0, 0)


def test_overtime_done_out_of_range_time_is_400(db):
    req = add_request(db)
    with pytest.raises(RpcError) as e:
        done(db, req.request_id, "25:00")
    assert e.value.message == "時刻が不正です: 25:00"


def test_overtime_done_24_00_is_next_midnight_like_v8(db):
    req = add_request(db)
    res = done(db, req.request_id, "24:00")
    assert res["actualMinutes"] == 400 and res["breakMinutes"] == 0  # 日をまたぐと休憩 0


@pytest.mark.parametrize("setup,error", [
    ("missing", "申請が見つかりません。"),
    ("holiday", "残業申請ではありません。"),
    ("canceled", "キャンセル済みです。"),
    ("other_email", "本人のみ操作可能です。"),
    ("upper_email", "本人のみ操作可能です。"),
])
def test_overtime_done_errors(db, setup, error):
    rid = "missing"
    if setup == "holiday":
        rid = add_request(db, request_type="holiday").request_id
    elif setup == "canceled":
        rid = add_request(db, status="canceled").request_id
    elif setup == "other_email":
        rid = add_request(db, worker_email="someone@lineworks-local.info").request_id
    elif setup == "upper_email":
        rid = add_request(db, worker_email="Tanaka@lineworks-local.info").request_id  # 大文字小文字を区別（Q24）
    with pytest.raises(RpcError) as e:
        done(db, rid)
    assert e.value.message == error
    assert jobs_of(db) == []


def test_assert_self_passes_for_same_email_and_empty(db):
    a = add_request(db, worker_email=USER)
    b = add_request(db, worker_email="")
    assert done(db, a.request_id)["ok"] and done(db, b.request_id)["ok"]


# ===== 休日 =====


def test_holiday_start_and_done(db):
    req = add_request(db, request_type="holiday", status="approved")
    c = ctx(db, now=jst(2026, 11, 26, 8, 5, 30))
    res = wlmod.api_markHolidayStart(c, req.request_id)
    assert res == {"ok": True, "requestId": req.request_id, "actualStartAt": "2026-11-26 08:05:30"}
    assert jobs_of(db) == []  # 開始は何も積まない
    res = wlmod.api_markHolidayStart(c, req.request_id, "08:00")  # 何度でも上書き
    assert res["actualStartAt"] == "2026-11-26 08:00:00"
    res = wlmod.api_markHolidayDone(ctx(db, now=jst(2026, 11, 26, 17, 0)), req.request_id)
    # 8:00-17:00 = 540 分、窓 10:00(10)+12:15(45)+15:00(10)=65
    assert (res["actualMinutes"], res["breakMinutes"], res["netMinutes"]) == (540, 65, 475)
    [job] = jobs_of(db)
    assert job["kind"] == "internal.afterComplete" and job["args"]["source"] == "screen"
    # 完了の後も開始をやり直せる（gap-1 G14。分は計算し直さない）
    wlmod.api_markHolidayStart(c, req.request_id, "09:00")
    [wl] = wls(db, req.request_id)
    assert wl.actual_start_at == "2026-11-26 09:00:00" and wl.actual_minutes == 540


def test_holiday_done_end_before_start_is_zero(db):
    req = add_request(db, request_type="holiday")
    add_wl(db, req.request_id, start="2026-11-26 15:00:00")
    res = wlmod.api_markHolidayDone(ctx(db), req.request_id, "14:00")
    assert (res["actualMinutes"], res["breakMinutes"], res["netMinutes"]) == (0, 0, 0)


def test_holiday_done_uses_first_worklog_row(db):
    req = add_request(db, request_type="holiday")
    add_wl(db, req.request_id, start="2026-11-26 09:00:00")
    add_wl(db, req.request_id, start="2026-11-26 13:00:00")
    res = wlmod.api_markHolidayDone(ctx(db), req.request_id, "12:00")
    assert res["actualMinutes"] == 180
    first, second = wls(db, req.request_id)
    assert first.actual_end_at == "2026-11-26 12:00:00" and second.actual_end_at == ""


@pytest.mark.parametrize("setup", ["no_row", "empty_start", "garbage_start"])
def test_holiday_done_without_start(db, setup):
    req = add_request(db, request_type="holiday")
    if setup == "empty_start":
        add_wl(db, req.request_id)
    elif setup == "garbage_start":
        add_wl(db, req.request_id, start="きのう")
    with pytest.raises(RpcError) as e:
        wlmod.api_markHolidayDone(ctx(db), req.request_id)
    assert e.value.message == "休日開始が記録されていません（開始ボタンを押してください）。"


@pytest.mark.parametrize("fn", ["api_markHolidayStart", "api_markHolidayDone"])
def test_holiday_errors(db, fn):
    ot = add_request(db)
    with pytest.raises(RpcError) as e:
        getattr(wlmod, fn)(ctx(db), ot.request_id)
    assert e.value.message == "休日申請ではありません。"
    cn = add_request(db, request_type="holiday", status="canceled")
    with pytest.raises(RpcError) as e:
        getattr(wlmod, fn)(ctx(db), cn.request_id)
    assert e.value.message == "キャンセル済みです。"
    other = add_request(db, request_type="holiday", worker_email="x@lineworks-local.info")
    with pytest.raises(RpcError) as e:
        getattr(wlmod, fn)(ctx(db), other.request_id)
    assert e.value.message == "本人のみ操作可能です。"


# ===== 総務の手動入力 =====


def manual(db, rid, end=None, start=None):
    c = ctx(db, email=ADMIN, now=jst(2026, 11, 27, 9))
    args = [rid] + ([end] if end is not None or start is not None else []) + ([start] if start is not None else [])
    return wlmod.api_adminManualComplete(c, *args)


def test_manual_complete_overtime_default_start(db):
    req = add_request(db, status="approved", worker_email="x@lineworks-local.info")  # 本人確認はしない
    res = manual(db, req.request_id, "19:30")
    assert res == {"ok": True, "actualMinutes": 130, "breakMinutes": 10, "netMinutes": 120}
    [wl] = wls(db, req.request_id)
    assert wl.updated_by == ADMIN and wl.actual_start_at == "2026-11-26 17:20:00"
    [job] = jobs_of(db)
    assert job["args"]["source"] == "admin" and job["args"]["wlOverride"]["actualEndAt"] == "2026-11-26 19:30:00"


def test_manual_complete_holiday_start_from_worklog(db):
    req = add_request(db, request_type="holiday")
    add_wl(db, req.request_id, start="2026-11-26 08:30:00")
    res = manual(db, req.request_id, "12:00")
    assert res["actualMinutes"] == 210 and res["breakMinutes"] == 10


@pytest.mark.parametrize("setup,end,start,error", [
    ("missing", "19:00", None, "申請が見つかりません。"),
    ("canceled", "19:00", None, "キャンセル済みの申請には入力できません。"),
    ("holiday_nostart", "19:00", None, "休日出勤の開始時刻が不明です。開始時刻を指定してください。"),
    ("ot", "", None, "終了時刻を指定してください。"),
    ("ot", "19:00", "9:00", "開始時刻が不正です: 9:00"),
    ("ot", "19:0", None, "終了時刻が不正です: 19:0"),
    ("ot", "17:20", None, "終了時刻は開始時刻より後にしてください。"),
    ("ot", "10:00", "12:00", "終了時刻は開始時刻より後にしてください。"),
])
def test_manual_complete_errors(db, setup, end, start, error):
    rid = "missing"
    if setup == "canceled":
        rid = add_request(db, status="canceled").request_id
    elif setup == "holiday_nostart":
        rid = add_request(db, request_type="holiday").request_id
    elif setup == "ot":
        rid = add_request(db).request_id
    with pytest.raises(RpcError) as e:
        manual(db, rid, end, start)
    assert e.value.message == error
