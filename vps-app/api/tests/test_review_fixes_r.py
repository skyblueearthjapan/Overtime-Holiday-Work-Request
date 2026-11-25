"""レビュー（docs/reviews/*.md）の指摘の直しの試験（team R）。"""
from __future__ import annotations

import pytest

from app.domain import approval, secondary, store
from app.jsutil import js_array_like_items, js_object_key_order, js_property_key, js_string
from app.rpc import RpcError
from test_domain_helpers import ADMIN, APPROVER, add_request, add_wl, ctx, jst, set_dept_approvers

NOW = jst(2026, 11, 26, 12, 34, 56)


@pytest.fixture
def approvers(db):
    set_dept_approvers(db, [("製造1課", "boss@lineworks-local.info", "上司")])


# ===== 正しさ 低-5: 配列でない一括承認の引数を包まない =====


def test_batch_string_argument_is_split_into_characters_and_approves_nothing(db, approvers):
    a = add_request(db)
    res = approval.api_approveRequestsBatch(ctx(db, email=APPROVER, now=NOW), a.request_id[:50])
    assert res["ok"] is True and all(r["error"] == "見つかりません" for r in res["results"])
    db.refresh(a)
    assert a.status == "submitted"
    with pytest.raises(RpcError) as e:  # 36 字の UUID でも 51 字以上の文字列は件数の上限
        approval.api_approveRequestsBatch(ctx(db, email=APPROVER, now=NOW), "x" * 51)
    assert e.value.message == "一括承認は50件までです。"


@pytest.mark.parametrize("value", [5, {"0": "x"}, True])
def test_batch_non_array_without_length_returns_empty(db, approvers, value):
    a = add_request(db)
    assert approval.api_approveRequestsBatch(ctx(db, email=APPROVER, now=NOW), value) == {"ok": True, "results": []}
    db.refresh(a)
    assert a.status == "submitted"


def test_secondary_batch_string_argument_does_not_approve(db):
    a = add_request(db, status="approved")
    res = secondary.api_secondaryApproveBatch(ctx(db, email=ADMIN, now=NOW), a.request_id)
    assert all(r["error"] == "見つかりません" for r in res["results"])
    db.refresh(a)
    assert not a.approved_by2
    assert secondary.api_secondaryApproveBatch(ctx(db, email=ADMIN, now=NOW), 7) == {"ok": True, "results": []}


# ===== 正しさ 低-6: 見つからない ID の書き方と並び =====


def test_batch_missing_ids_use_js_keys_and_order(db, approvers):
    add_request(db)
    res = approval.api_approveRequestsBatch(ctx(db, email=APPROVER, now=NOW), ["b", "5", None, 1.0, "01"])
    assert [r["requestId"] for r in res["results"]] == ["1", "5", "b", "null", "01"]


def test_js_helpers():
    assert js_string(["R1"]) == "R1" and js_string([1, None, "a"]) == "1,,a" and js_string({}) == "[object Object]"
    assert js_property_key(None) == "null" and js_property_key(1.0) == "1"
    assert js_object_key_order(["b", "5", "01", "2"]) == ["2", "5", "b", "01"]
    assert js_array_like_items("a😀") == ["a", "\ud83d", "\ude00"] and js_array_like_items(3) == []


# ===== 推測 1（ID の前後の空白）: GAS の normalize_ と同じく行の ID の空白を無視して探す =====


def test_request_with_trailing_space_is_found_like_gas(db, approvers):
    add_request(db, "R-1 ")
    res = approval.api_approveRequest(ctx(db, email=APPROVER, now=NOW), "R-1")
    assert res["ok"] is True
    req = store.lock_request(db, "R-1")
    assert req is not None and req.request_id == "R-1 " and req.status == "approved"
    # GAS は引数を normalize_ しないので、空白付きの引数は当たらない
    assert store.lock_request(db, " R-1") is None


def test_exact_id_wins_and_like_wildcards_are_escaped(db):
    add_request(db, "A_1")
    add_request(db, "AX1 ")
    assert store.find_request_id(db, "A_1") == "A_1"
    assert store.find_request_id(db, "A%1") is None
    assert store.find_request_id(db, "AX1") == "AX1 "


def test_worklog_with_spaces_is_found_like_gas(db):
    add_request(db, "R-2")
    add_wl(db, " R-2", end="2026-11-26 19:00:00")
    assert store.first_worklog(db, "R-2").request_id == " R-2"
    assert store.worklog_view(db, "R-2").actualEndAt == "2026-11-26 19:00:00"


# ===== 正しさ 低-1・低-8・低-9・低-10・低-16・低-17、セキュリティ 低-2 =====

from app.domain import ops, summaries, worklog  # noqa: E402
from app.domain import requests as rq  # noqa: E402
from app.domain.jsdate import HM_RE, YMD_RE, js_number, valid_ymd_text  # noqa: E402


def test_regex_end_anchor_rejects_trailing_newline():
    assert not HM_RE.match("19:30\n") and not YMD_RE.match("2026-11-26\n")
    assert HM_RE.match("19:30") and YMD_RE.match("2026-11-26")


def test_manual_time_with_newline_is_ignored_like_gas(db):
    add_request(db, "H9", request_type="holiday")
    now = jst(2026, 11, 26, 9, 5)
    req, start = worklog.core_holiday_start(db, "H9", "08:00\n", now,
                                            msgs=worklog.SEBASTIAN, email=None, updated_by="t")
    assert start == now  # 形に合わないので今の時刻（G13）


def test_js_number_rejects_loose_radix_literals():
    import math

    assert math.isnan(js_number("0x1_0")) and math.isnan(js_number("0x 1")) and math.isnan(js_number("0x"))
    assert js_number("0x10") == 16 and js_number("0b11") == 3 and js_number("0o17") == 15


def _base(**kw):
    data = {"requestType": "holiday", "dept": "製造1課", "workerLabel": "A001 山田 太郎", "targetDate": "2026-11-28",
            "reason": "急ぎ作業を要するため", "hours": "半日"}
    data.update(kw)
    return data


def test_year_zero_is_valid_but_before_fiscal_year(db):
    assert valid_ymd_text("0000-02-29") and not valid_ymd_text("0000-13-01")
    res = rq.api_submitRequest(ctx(db, now=NOW), _base(targetDate="0000-01-01"))
    assert res == {"ok": False, "error": "年度開始日（2026-03-16）より前の日付は申請できません。"}


@pytest.mark.parametrize("hours", ["1e308", "1e20"])
def test_huge_planned_hours_are_rejected_not_500(db, hours):
    with pytest.raises(RpcError) as e:
        rq.api_submitRequest(ctx(db, now=NOW), _base(requestType="overtime", hours=hours, targetDate=""))
    assert e.value.message == f"残業の予定時間が不正です: {hours}"


def test_text_longer_than_a_sheet_cell_is_refused_like_gas(db):
    with pytest.raises(RpcError) as e:
        rq.api_submitRequest(ctx(db, now=NOW), _base(workContent="あ" * 50001))
    assert e.value.message == "業務内容が長すぎます（50000字まで）。"
    assert rq.api_submitRequest(ctx(db, now=NOW), _base(workContent="あ" * 50000))["ok"] is True


@pytest.mark.parametrize("ym", ["9999-12", "10000-01", "0.5-01"])
def test_out_of_range_year_in_summaries_is_400_not_500(db, ym):
    with pytest.raises(RpcError):
        summaries.api_adminMonthlySummary(ctx(db, email=ADMIN, now=NOW), ym, "ALL")
    with pytest.raises(RpcError):
        summaries.api_getClosingSummary(ctx(db, email=ADMIN, now=NOW), ym)


@pytest.mark.parametrize("value", [1.9, True, " 3 ", "3.0", None])
def test_retry_job_accepts_only_integer_ids(db, value):
    with pytest.raises(RpcError) as e:
        ops.ops_retryJob(ctx(db, email=ADMIN), value)
    assert e.value.message == "仕事の番号が正しくありません。"


def test_worker_label_split_uses_js_whitespace():
    """正しさ 低-11。"""
    assert rq.parse_worker_choice("001﻿山田") == ("001", "山田")
    assert rq.parse_worker_choice("001\x1c山田") == ("001\x1c山田", "")
    assert rq.parse_worker_choice("001 山田 太郎") == ("001", "山田 太郎")
