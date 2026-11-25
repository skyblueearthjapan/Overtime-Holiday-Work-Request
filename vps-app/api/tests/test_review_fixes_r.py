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
