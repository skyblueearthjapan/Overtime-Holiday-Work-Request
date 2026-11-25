"""RPC の入口と呼び出し（実装仕様 4-1）: Origin・415・404・400・501・401・403・冪等キーの再生と 409・監査。"""
from __future__ import annotations

import json

import pytest
from sqlalchemy import func, select

from app import rpc
from app.models import AuditLog, IdempotencyKey, Setting
from conftest import make_admin

ME = "tanaka@lineworks-local.info"
ORIGIN = "https://overtime.lineworks-local.net"


@pytest.fixture
def logged_in(client, login_as):
    login_as(ME)
    return client


def count(db, model):
    return db.execute(select(func.count()).select_from(model)).scalar()


# ===== 閉じたリスト =====


def test_exactly_38_functions_are_public():
    assert len(rpc.SCREEN_FUNCTIONS) == 29 == len(set(rpc.SCREEN_FUNCTIONS))
    assert len(rpc.OPS_FUNCTIONS) == 9
    assert len(rpc.RPC_FUNCTIONS) == 38 == len(set(rpc.RPC_FUNCTIONS))
    # 画面から呼ばれない公開関数・旧フォーム・保守の関数は出さない（9-1 #11・#13）
    for name in ("api_getFormUrl", "api_getTodayRequestsForDept", "api_getTodayRequestsForWorker",
                 "recalcWorkLogsBreakMinutes", "rebuildAccumulationSS", "syncAllMasters", "doGet", "doPost",
                 "getSettings_", "api_getHolidayCandidateDates_legacy_"):
        assert name not in rpc.RPC_FUNCTIONS
        with pytest.raises(ValueError):
            rpc.register(name, lambda ctx: None)


def test_write_and_admin_classification():
    assert rpc.WRITE_FUNCTIONS <= set(rpc.RPC_FUNCTIONS)
    assert rpc.ADMIN_ONLY_FUNCTIONS <= set(rpc.RPC_FUNCTIONS)
    assert set(rpc.OPS_FUNCTIONS) <= rpc.ADMIN_ONLY_FUNCTIONS
    assert "api_secondaryApproveBatch" not in rpc.ADMIN_ONLY_FUNCTIONS  # 件数の確かめが先（approval.gs:776-779）
    for name in ("api_submitRequest", "api_approveRequest", "api_deleteRequest", "api_markHolidayStart",
                 "ops_revokeSessions"):
        assert name in rpc.WRITE_FUNCTIONS
    for name in ("api_getDashboard", "api_getDeptList", "ops_getStatus", "api_adminDailyDetail"):
        assert name not in rpc.WRITE_FUNCTIONS
    for name in ("api_getDashboard", "api_submitRequest", "api_approveRequest", "api_getClosingSummary",
                 "api_getApproverDepts"):
        assert name not in rpc.ADMIN_ONLY_FUNCTIONS


def test_register_decorator():
    @rpc.register("api_getDeptList")
    def handler(ctx):
        return ["A"]

    assert rpc.get_handler("api_getDeptList") is handler
    assert rpc.implemented_functions() == ["api_getDeptList"]


def test_load_domain_handlers_without_domain_package():
    assert rpc.load_domain_handlers() == [] or isinstance(rpc.load_domain_handlers(), list)


# ===== 入口 =====


def test_unknown_function_is_404(logged_in):
    assert logged_in.post("/api/rpc/api_getFormUrl", json=[]).status_code == 404
    assert logged_in.post("/api/rpc/__init__", json=[]).status_code == 404


def test_unregistered_function_is_501(logged_in):
    resp = logged_in.post("/api/rpc/api_getDashboard", json=[])
    assert resp.status_code == 501
    assert resp.json()["ok"] is False


def test_login_required_is_401_json(client):
    resp = client.post("/api/rpc/api_getDashboard", json=[])
    assert resp.status_code == 401
    assert resp.json() == {"ok": False, "error": "ログインが必要です。", "login": "/auth/login"}


def test_success_shape_and_context(logged_in):
    seen = {}

    def get_dashboard(ctx):
        seen["user"], seen["fn"], seen["now"] = ctx.user, ctx.fn, ctx.now
        return {"today": "2026-11-26", "overtime": [], "holiday": []}

    rpc.register("api_getDashboard", get_dashboard)
    resp = logged_in.post("/api/rpc/api_getDashboard", json=[])
    assert resp.status_code == 200
    assert resp.json() == {"ok": True, "result": {"today": "2026-11-26", "overtime": [], "holiday": []}}
    assert resp.headers["cache-control"] == "no-store"
    assert seen["user"].email == ME and seen["fn"] == "api_getDashboard"
    assert seen["now"].tzinfo is not None


@pytest.mark.parametrize("result", [None, {"error": "x"}, {"ok": False, "error": "作業員が指定されていません。"},
                                    {"ok": True, "results": [{"requestId": "a", "ok": False, "error": "e"}]}])
def test_failure_shapes_inside_success_are_passed_through(logged_in, result):
    rpc.register("api_getWorkerInfo", lambda ctx: result)
    resp = logged_in.post("/api/rpc/api_getWorkerInfo", json=[])
    assert resp.status_code == 200 and resp.json() == {"ok": True, "result": result}


def test_business_error_is_400_and_rolls_back(logged_in, db):
    def delete_request(ctx, request_id):
        ctx.db.add(Setting(key="X", value="x"))
        ctx.db.flush()
        raise rpc.RpcError("既にキャンセル済みです。")

    rpc.register("api_deleteRequest", delete_request)
    resp = logged_in.post("/api/rpc/api_deleteRequest", json=["R1"])
    assert resp.status_code == 400
    assert resp.json() == {"ok": False, "error": "既にキャンセル済みです。"}
    db.expire_all()
    assert db.get(Setting, "X") is None
    # 失敗も監査に残る（業務の書き込みは戻した後で）
    row = db.scalars(select(AuditLog)).one()
    assert row.fn == "api_deleteRequest" and row.ok is False and json.loads(row.request_ids) == ["R1"]
    assert row.error == "既にキャンセル済みです。"


def test_unexpected_error_is_500_without_details(logged_in):
    def boom(ctx):
        raise KeyError("secret detail")

    rpc.register("api_getDeptList", boom)
    resp = logged_in.post("/api/rpc/api_getDeptList", json=[])
    assert resp.status_code == 500
    assert "secret" not in resp.text
    assert resp.json()["error"] == "サーバーでエラーが起きました。時間をおいてもう一度お試しください。"


def test_optional_args_and_wrong_arity(logged_in):
    rpc.register("api_markOvertimeDone", lambda ctx, request_id, manual_end_time=None: [request_id, manual_end_time])
    assert logged_in.post("/api/rpc/api_markOvertimeDone", json=["R"]).json()["result"] == ["R", None]
    assert logged_in.post("/api/rpc/api_markOvertimeDone", json=["R", "19:30"]).json()["result"] == ["R", "19:30"]
    assert logged_in.post("/api/rpc/api_markOvertimeDone", json=[]).status_code == 400
    # 余った引数は GAS と同じく捨てる（正しさ 低-7）
    assert logged_in.post("/api/rpc/api_markOvertimeDone", json=["R", None, "extra"]).json()["result"] == ["R", None]


@pytest.mark.parametrize("body", [b"{}", b"\"x\"", b"not json", b"123", b"null"])
def test_body_must_be_json_array(logged_in, body):
    rpc.register("api_getWorkersByDept", lambda ctx, *a: None)
    resp = logged_in.post("/api/rpc/api_getWorkersByDept", content=body, headers={"content-type": "application/json"})
    assert resp.status_code == 400


def test_empty_body_means_no_args(logged_in):
    rpc.register("api_getDeptList", lambda ctx: "ok")
    resp = logged_in.post("/api/rpc/api_getDeptList", content=b"", headers={"content-type": "application/json"})
    assert resp.json() == {"ok": True, "result": "ok"}


def test_content_type_must_be_json_415(logged_in):
    rpc.register("api_getDeptList", lambda ctx: "ok")
    assert logged_in.post("/api/rpc/api_getDeptList").status_code == 415
    for ct in ("text/plain", "application/x-www-form-urlencoded", "multipart/form-data; boundary=x"):
        assert logged_in.post("/api/rpc/api_getDeptList", content=b"[]", headers={"content-type": ct}).status_code == 415
    ok = logged_in.post("/api/rpc/api_getDeptList", content=b"[]",
                        headers={"content-type": "application/json; charset=utf-8"})
    assert ok.status_code == 200


def test_missing_origin_requires_same_origin_fetch_403(logged_in):
    rpc.register("api_getDeptList", lambda ctx: "ok")
    for site in ("cross-site", "same-site", "none"):
        assert logged_in.post("/api/rpc/api_getDeptList", json=[], headers={"sec-fetch-site": site}).status_code == 403
    del logged_in.headers["sec-fetch-site"]
    assert logged_in.post("/api/rpc/api_getDeptList", json=[]).status_code == 403
    assert logged_in.post("/api/rpc/api_getDeptList", json=[], headers={"origin": ORIGIN}).status_code == 200


def test_cross_origin_is_403(logged_in):
    rpc.register("api_getDeptList", lambda ctx: "ok")
    for origin in ("https://evil.example.com", "http://overtime.lineworks-local.net", "null",
                   "https://overtime.lineworks-local.net.evil.example.com"):
        resp = logged_in.post("/api/rpc/api_getDeptList", json=[], headers={"origin": origin})
        assert resp.status_code == 403
        assert resp.json() == {"ok": False, "error": "この画面からは操作できません。"}
    assert logged_in.post("/api/rpc/api_getDeptList", json=[], headers={"origin": ORIGIN + "/"}).status_code == 200


# ===== 管理者だけの関数（403） =====


@pytest.mark.parametrize("fn", sorted(rpc.ADMIN_ONLY_FUNCTIONS))
def test_admin_only_functions_are_403_for_non_admin(logged_in, db, fn):
    called = []
    rpc.register(fn, lambda ctx, *a: called.append(1) or {"ok": True})
    resp = logged_in.post(f"/api/rpc/{fn}", json=[])
    assert resp.status_code == 403
    assert resp.json() == {"ok": False, "error": "総務部（管理者）権限がありません。"}
    assert called == []
    make_admin(db, ME)
    assert logged_in.post(f"/api/rpc/{fn}", json=[]).status_code == 200
    assert called == [1]


def test_bootstrap_admin_is_admin(client, login_as):
    rpc.register("ops_getStatus", lambda ctx: {"ok": True})
    login_as("owner@lineworks-local.info")
    assert client.post("/api/rpc/ops_getStatus", json=[]).status_code == 200


def test_handler_side_admin_check_keeps_gas_order(logged_in, db):
    """api_secondaryApproveBatch: 件数の確かめが先（50 件超は 400）、権限が後（403）。"""

    def batch(ctx, ids):
        if not ids:
            return {"ok": True, "results": []}
        if len(ids) > 50:
            raise rpc.RpcError("一括承認は50件までです。")
        ctx.require_admin()
        return {"ok": True, "results": [{"requestId": i, "ok": True} for i in ids]}

    rpc.register("api_secondaryApproveBatch", batch)
    url = "/api/rpc/api_secondaryApproveBatch"
    assert logged_in.post(url, json=[[]]).json() == {"ok": True, "result": {"ok": True, "results": []}}
    assert logged_in.post(url, json=[[str(i) for i in range(51)]]).status_code == 400
    resp = logged_in.post(url, json=[["a"]])
    assert resp.status_code == 403 and resp.json()["error"] == "総務部（管理者）権限がありません。"
    make_admin(db, ME)
    assert logged_in.post(url, json=[["a"]]).status_code == 200


# ===== 冪等キー（9-1 #5・罠 17） =====


def test_idempotency_key_replays_first_result(logged_in, db):
    calls = []

    def submit(ctx, data):
        calls.append(data)
        ctx.db.add(Setting(key=f"k{len(calls)}", value="v"))
        return {"ok": True, "success": True, "requestId": f"REQ-{len(calls)}"}

    rpc.register("api_submitRequest", submit)
    h = {"X-Idempotency-Key": "abcdefgh-1234"}
    first = logged_in.post("/api/rpc/api_submitRequest", json=[{"a": 1}], headers=h)
    second = logged_in.post("/api/rpc/api_submitRequest", json=[{"a": 1}], headers=h)
    assert first.status_code == second.status_code == 200
    assert first.json() == second.json() == {"ok": True, "result": {"ok": True, "success": True, "requestId": "REQ-1"}}
    assert len(calls) == 1
    assert count(db, Setting) == 1  # 二重押しで 2 件できない
    row = db.get(IdempotencyKey, (ME, "abcdefgh-1234"))
    assert row.state == "done" and row.fn == "api_submitRequest"
    third = logged_in.post("/api/rpc/api_submitRequest", json=[{"a": 1}], headers={"X-Idempotency-Key": "abcdefgh-5678"})
    assert third.json()["result"]["requestId"] == "REQ-2"
    # 同じキーを別の関数に使うと 400
    rpc.register("api_deleteRequest", lambda ctx, rid: {"ok": True})
    assert logged_in.post("/api/rpc/api_deleteRequest", json=["R"], headers=h).status_code == 400
    # 監査: 成功は 2 回分（再生は数えない）。申請 ID は返事の requestId
    rows = db.scalars(select(AuditLog).where(AuditLog.ok.is_(True)).order_by(AuditLog.id)).all()
    assert [json.loads(r.request_ids) for r in rows] == [["REQ-1"], ["REQ-2"]]


def test_in_progress_key_is_409(logged_in, db):
    from app.timeutil import now_utc

    rpc.register("api_approveRequest", lambda ctx, rid: {"ok": True})
    db.add(IdempotencyKey(user_email=ME, key="busy-key-0001", fn="api_approveRequest", state="running",
                          created_at=now_utc()))
    db.commit()
    resp = logged_in.post("/api/rpc/api_approveRequest", json=["R"], headers={"X-Idempotency-Key": "busy-key-0001"})
    assert resp.status_code == 409
    assert resp.json()["ok"] is False


def test_read_only_functions_ignore_key_and_are_not_audited(logged_in, db):
    rpc.register("api_getDashboard", lambda ctx: {"n": 1})
    logged_in.post("/api/rpc/api_getDashboard", json=[], headers={"X-Idempotency-Key": "read-key-0001"})
    assert count(db, IdempotencyKey) == 0
    assert count(db, AuditLog) == 0


def test_business_error_does_not_keep_key(logged_in, db):
    state = {"n": 0}

    def approve(ctx, rid):
        state["n"] += 1
        if state["n"] == 1:
            raise rpc.RpcError("申請が見つかりません: R")
        return {"ok": True, "requestId": rid}

    rpc.register("api_approveRequest", approve)
    h = {"X-Idempotency-Key": "retry-key-0001"}
    assert logged_in.post("/api/rpc/api_approveRequest", json=["R"], headers=h).status_code == 400
    assert db.get(IdempotencyKey, (ME, "retry-key-0001")) is None  # 鍵が残らない（もう一度押せる）
    assert logged_in.post("/api/rpc/api_approveRequest", json=["R"], headers=h).json() == {
        "ok": True, "result": {"ok": True, "requestId": "R"}}
    assert count(db, IdempotencyKey) == 1


def test_crash_does_not_keep_key(logged_in, db):
    def crash(ctx, rid, end, start=None):
        raise RuntimeError("db down")

    rpc.register("api_adminManualComplete", crash)
    make_admin(db, ME)
    resp = logged_in.post("/api/rpc/api_adminManualComplete", json=["R", "19:30"],
                          headers={"X-Idempotency-Key": "crash-key-0001"})
    assert resp.status_code == 500
    assert count(db, IdempotencyKey) == 0


def test_bad_idempotency_key_is_400(logged_in):
    rpc.register("api_submitRequest", lambda ctx, p: "x")
    for key in ("short", "has space 123", "x" * 129, "semi;colon-key1"):
        resp = logged_in.post("/api/rpc/api_submitRequest", json=[{}], headers={"X-Idempotency-Key": key})
        assert resp.status_code == 400


def test_idempotency_is_per_user(client, login_as):
    calls = []
    rpc.register("api_submitRequest", lambda ctx, p: calls.append(ctx.user.email) or len(calls))
    h = {"X-Idempotency-Key": "shared-key-0001"}
    login_as("a@lineworks-local.info")
    assert client.post("/api/rpc/api_submitRequest", json=[{}], headers=h).json()["result"] == 1
    client.cookies.clear()
    login_as("b@lineworks-local.info")
    assert client.post("/api/rpc/api_submitRequest", json=[{}], headers=h).json()["result"] == 2


def test_purge_idempotency_keys(db):
    from datetime import timedelta

    from app.timeutil import now_utc

    now = now_utc()
    db.add_all([IdempotencyKey(user_email=ME, key="old-key-0001", fn="f", state="done",
                               created_at=now - timedelta(hours=25)),
                IdempotencyKey(user_email=ME, key="new-key-0001", fn="f", state="done",
                               created_at=now - timedelta(hours=23))])
    db.commit()
    assert rpc.purge_idempotency_keys(db, now) == 1
    db.commit()
    assert [k.key for k in db.scalars(select(IdempotencyKey))] == ["new-key-0001"]


# ===== 監査 =====


def test_batch_request_ids_are_audited(logged_in, db):
    rpc.register("api_approveRequestsBatch", lambda ctx, ids: {"ok": True, "results": []})
    logged_in.post("/api/rpc/api_approveRequestsBatch", json=[["A", "B"]])
    row = db.scalars(select(AuditLog)).one()
    assert row.ok and row.user_email == ME and json.loads(row.request_ids) == ["A", "B"]


def test_audit_is_in_same_transaction_as_business_write(logged_in, db):
    def edit(ctx, rid, patch):
        ctx.db.add(Setting(key="E", value="1"))
        return {"ok": True, "newStatus": "submitted", "newMinutes": 60}

    rpc.register("api_editRequest", edit)
    assert logged_in.post("/api/rpc/api_editRequest", json=["R9", {"minutes": 60}]).status_code == 200
    db.expire_all()
    assert db.get(Setting, "E") is not None
    assert json.loads(db.scalars(select(AuditLog)).one().request_ids) == ["R9"]


def test_healthz(client):
    resp = client.get("/healthz")
    assert resp.status_code == 200 and resp.json() == {"ok": True, "env": "staging"}


def test_healthz_detail_only_for_admin(client, db, login_as):
    assert client.get("/healthz?detail=1").json() == {"ok": True, "env": "staging"}
    login_as("someone@lineworks-local.info")
    assert "detail" not in client.get("/healthz?detail=1").json()
    make_admin(db, "someone@lineworks-local.info")
    body = client.get("/healthz?detail=1").json()
    assert body["detail"]["rowSync"] == {"pending": 0, "blocked": 0, "oldestSeconds": 0, "delayed": False}
    assert body["detail"]["gasPing"] is None


def test_sebastian_route_is_always_200_json(client, monkeypatch):
    import types

    import app.main as main_module

    # app.domain.sebastian がまだ無いとき
    monkeypatch.setattr(main_module, "_optional_router", lambda name: None)
    resp = client.post("/api/sebastian", content=b"not json", headers={"content-type": "text/plain"})
    assert resp.status_code == 200 and resp.json() == {"error": "service unavailable"}

    seen = {}

    def handle(db, raw):
        seen["raw"] = raw
        return {"ok": True, "pong": "2026-11-26T00:00:00.000Z"}

    monkeypatch.setattr(main_module, "_optional_router", lambda name: types.SimpleNamespace(handle=handle))
    resp = client.post("/api/sebastian", content=b'{"action":"ping","token":"t"}',
                       headers={"content-type": "text/plain", "origin": "https://evil.example.com"})
    assert resp.status_code == 200 and resp.json()["ok"] is True
    assert seen["raw"] == b'{"action":"ping","token":"t"}'  # 生の本文（Content-Type・Origin・ログインを見ない）

    def crash(db, raw):
        raise RuntimeError("x")

    monkeypatch.setattr(main_module, "_optional_router", lambda name: types.SimpleNamespace(handle=crash))
    resp = client.post("/api/sebastian", content=b"{}")
    assert resp.status_code == 200 and resp.json() == {"error": "internal error"}
