"""登録（38 関数）と、担当 M の gas_jobs の呼び方（10-2 の名前と引数）。JS の小道具と日本語の照合。"""
from __future__ import annotations

import sys
import types
from datetime import date

import pytest
from sqlalchemy.orm import sessionmaker

from app import jobs, rpc
from app.domain import approval, collation, gasjobs, jsdate, register, requests, secondary, worklog
from app.models import Job
from test_domain_helpers import ADMIN, APPROVER, add_request, add_wl, ctx, set_dept_approvers


def test_register_all_covers_the_closed_list():
    names = register.register_all(rpc.register)
    assert sorted(names) == sorted(rpc.RPC_FUNCTIONS)
    assert rpc.implemented_functions() == list(rpc.RPC_FUNCTIONS)
    assert rpc.load_domain_handlers() == list(rpc.RPC_FUNCTIONS)


@pytest.fixture
def m_stub(monkeypatch):
    calls: list = []
    mod = types.ModuleType("app.services.gas_jobs")

    def rec(name):
        def fn(db, *args, **kw):
            calls.append((name, args, kw))
            return 99
        return fn

    for name in ("enqueue_after_approve", "enqueue_after_complete", "enqueue_regenerate", "enqueue_ledger_upsert",
                 "enqueue_after_delete", "enqueue_ledger_rebuild"):
        setattr(mod, name, rec(name))
    mod.wait_for_result = lambda factory, job_id, timeout: {"jobId": job_id}
    monkeypatch.setitem(sys.modules, "app.services.gas_jobs", mod)
    return calls


def test_calls_m_entry_points_with_10_2_signatures(db, m_stub):
    set_dept_approvers(db, [("製造1課", APPROVER, "")])
    a = add_request(db)
    add_wl(db, a.request_id, end="2026-11-26 19:00:00")
    approval.api_approveRequest(ctx(db, email=APPROVER), a.request_id)
    b = add_request(db, status="approved", pdf_file_id="P")
    worklog.api_markOvertimeDone(ctx(db), b.request_id, "19:30")
    secondary.api_secondaryApprove(ctx(db, email=ADMIN), b.request_id)
    c = add_request(db, status="approved")
    secondary.api_secondaryApprove(ctx(db, email=ADMIN), c.request_id)
    requests.api_deleteRequest(ctx(db), c.request_id)
    assert m_stub == [
        ("enqueue_after_approve", (a.request_id,), {}),
        ("enqueue_after_complete", (b.request_id,), {
            "source": "screen", "status_at_enqueue": "approved", "priority": 0, "wl_override": {
                "actualStartAt": "2026-11-26 17:20:00", "actualEndAt": "2026-11-26 19:30:00", "actualMinutes": 130,
                "breakMinutes": 10, "netMinutes": 120}}),
        ("enqueue_regenerate", (b.request_id,), {"reason": "secondary", "old_pdf_file_id": "P", "batch_id": ""}),
        ("enqueue_ledger_upsert", (c.request_id,), {"batch_id": ""}),
        ("enqueue_after_delete", (c.request_id,), {"request_type": "overtime", "target_date": "2026-11-26",
                                                   "old_pdf_file_id": ""}),
    ]
    assert gasjobs.wait_for_result(None, 5, 1) == {"jobId": 5}


def test_fallback_wait_for_result_reads_jobs_table(db, session_factory, monkeypatch):
    monkeypatch.setitem(sys.modules, "app.services.gas_jobs", None)
    jid = jobs.enqueue(db, "internal.afterComplete", {}, request_id="r")
    db.commit()
    assert gasjobs.wait_for_result(session_factory, jid, 0) is None
    job = db.get(Job, jid)
    job.state, job.result_json = "done", '{"ok": true, "pdf": {"status": "already", "pdfFileId": "X"}}'
    db.commit()
    assert gasjobs.wait_for_result(session_factory, jid, 0)["pdf"]["status"] == "already"
    job.state, job.result_json, job.last_error = "failed", "", "bad_request"
    db.commit()
    assert gasjobs.wait_for_result(sessionmaker(bind=db.get_bind()), jid, 0) == {"ok": False, "error": "bad_request"}


# ===== JS の小道具 =====


@pytest.mark.parametrize("text,iso", [
    ("2026-11-24", "2026-11-24T00:00:00.000Z"),
    ("2026/11/24", "2026-11-23T15:00:00.000Z"),
    ("2026-11-24 10:00:00", "2026-11-24T01:00:00.000Z"),
    ("2026-11-24T09:00", "2026-11-24T00:00:00.000Z"),
    ("2026-02-30T00:00:00+09:00", "2026-03-01T15:00:00.000Z"),
    ("2026-11-24T24:00:00+09:00", "2026-11-24T15:00:00.000Z"),
    (None, "1970-01-01T00:00:00.000Z"),
])
def test_js_new_date_matches_v8(text, iso):
    assert jsdate.iso_utc_ms(jsdate.js_new_date(text)) == iso


@pytest.mark.parametrize("text", ["", "2026-13-01T00:00:00+09:00", "2026-11-24T25:00:00+09:00",
                                  "2026-11-24T24:01:00+09:00", "2026-00-10", "2026-11-00", "abc", jsdate.MISSING])
def test_js_new_date_invalid(text):
    assert jsdate.js_new_date(text) is None


def test_js_number_and_truthy():
    import math

    assert jsdate.js_number("") == 0 and jsdate.js_number(" 1.5 ") == 1.5 and jsdate.js_number("0x10") == 16
    assert math.isnan(jsdate.js_number("1,5")) and jsdate.js_number(None) == 0
    assert not jsdate.js_truthy(0) and not jsdate.js_truthy("") and jsdate.js_truthy([]) and jsdate.js_truthy("0")


def test_labels():
    assert jsdate.label_ja("2026-11-07", "2026-01-01") == "11/7(土)"
    assert jsdate.label_ja("2025-11-07", "2026-01-01") == "2025/11/7(金)"
    assert jsdate.label_ja("x", "2026-01-01") == "x"
    assert jsdate.label_md_dow("2026-11-26") == "11/26(木)"
    assert jsdate.days_ago("2026-11-20", "2026-11-26") == 6 and jsdate.days_ago("2026-11-27", "2026-11-26") == 0
    assert jsdate.is_retroactive(None, "2026-11-20") is False


# node（ICU 入り）で `a.sort((x, y) => x.localeCompare(y, 'ja'))` とした結果（2026-11-26 に手元で取った表）
NODE_JA_ORDER = [
    " ", "　", "_", "-", "(", "（", "1", "2", "a", "ａ", "A", "Ａ", "abc", "ABC", "b", "x1", "x10", "x2", "zeta",
    "Zeta", "ぁ", "ァ", "あ", "ア", "ｱ", "あいう", "アイウ", "か", "が", "ガ", "かき", "カキ", "ｶｷ", "がき", "は",
    "ハ", "ば", "ぱ", "パ", "ﾊﾟ", "一", "営業", "課長", "漢", "機械設計山田", "係長", "高橋", "佐藤", "製造1課",
    "製造1課佐藤", "斉藤", "総務鈴木", "第1", "第１", "第10", "第2", "丁", "品質管理田中", "齋藤", "髙",
]


def test_ja_collation_matches_node_icu_table():
    import random

    shuffled = NODE_JA_ORDER[:]
    random.Random(7).shuffle(shuffled)
    assert sorted(shuffled, key=collation.ja_key) == NODE_JA_ORDER


def test_target_date_text_fallback(db):
    r = add_request(db, target=None, target_text="2026/11/24")
    from app.domain import store

    assert store.target_ymd(r) == "2026-11-24" and store.target_date_of(r) == date(2026, 11, 24)
