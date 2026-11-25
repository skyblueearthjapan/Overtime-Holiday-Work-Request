"""tools/qa/acceptance.py（受け入れテストの道具）と tools/qa/staging_session.py の試験（手元だけで動く）。

- 読み方・表・判定の小さな関数を 1 つずつ。
- 本物の app（SQLite・TestClient）に向けて道具をまるごと動かし、読むだけの確かめと書き込みの試験が通ること
  （nginx・GAS の worker は無いので、その項目は FAIL か SKIP になることも確かめる）。
tools/ は本番・試験のイメージに入らない（罠 5）ので、ファイルが無ければこの試験は飛ばす。
"""
from __future__ import annotations

import importlib.util
import json
import sys
from datetime import timedelta
from pathlib import Path
from types import SimpleNamespace

import pytest
from sqlalchemy import select
from conftest import TEST_ENV

QA_DIR = Path(__file__).resolve().parents[2] / "tools" / "qa"
pytestmark = pytest.mark.skipif(not (QA_DIR / "acceptance.py").is_file(),
                                reason="tools/qa はこのイメージに入っていない（手元で動かす試験）")


def _load(name: str):
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, QA_DIR / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def acc():
    return _load("acceptance")


@pytest.fixture(scope="module")
def ss():
    return _load("staging_session")


# ===== 取り決めの写しが app と同じ =====


def test_function_lists_match_rpc(acc):
    from app import rpc

    assert acc.SCREEN_FUNCTIONS == rpc.SCREEN_FUNCTIONS
    assert acc.OPS_FUNCTIONS == rpc.OPS_FUNCTIONS
    assert len(acc.RPC_FUNCTIONS) == 38
    assert acc.WRITE_FUNCTIONS == rpc.WRITE_FUNCTIONS


def test_constants_match_app(acc):
    from app import auth
    from app.services import gas_jobs

    assert acc.SESSION_COOKIE == auth.SESSION_COOKIE
    assert (acc.K_AFTER_APPROVE, acc.K_AFTER_COMPLETE, acc.K_REGENERATE, acc.K_LEDGER_UPSERT, acc.K_AFTER_DELETE) == (
        gas_jobs.KIND_AFTER_APPROVE, gas_jobs.KIND_AFTER_COMPLETE, gas_jobs.KIND_REGENERATE,
        gas_jobs.KIND_LEDGER_UPSERT, gas_jobs.KIND_AFTER_DELETE)


def test_csp_matches_nginx_conf(acc):
    conf = (Path(__file__).resolve().parents[2] / "web" / "nginx.conf").read_text("utf-8")
    assert f'add_header Content-Security-Policy "{acc.CSP_EXPECTED}" always;' in conf
    for k, v in acc.HEADERS_EXPECTED.items():
        assert v in conf


# ===== 読み方 =====


def _resp(acc, status, body, headers=None):
    text = body if isinstance(body, str) else json.dumps(body, ensure_ascii=False)
    return acc.HttpResponse(status, headers or {}, text)


def test_parse_rpc_success_error_and_garbage(acc):
    r = acc.parse_rpc(_resp(acc, 200, {"ok": True, "result": {"ok": False, "error": "x"}}))
    assert (r.status, r.ok, r.result, r.error) == (200, True, {"ok": False, "error": "x"}, None)
    r = acc.parse_rpc(_resp(acc, 200, {"ok": True, "result": None}))
    assert r.ok is True and r.result is None
    r = acc.parse_rpc(_resp(acc, 400, {"ok": False, "error": "申請が見つかりません。"}))
    assert (r.ok, r.error) == (False, "申請が見つかりません。") and r.brief() == "400 申請が見つかりません。"
    r = acc.parse_rpc(_resp(acc, 502, "<html>Bad gateway</html>"))
    assert r.ok is None and r.error.startswith("JSON ではありません") and r.brief().startswith("502 ")
    r = acc.parse_rpc(_resp(acc, 200, [1, 2]))
    assert r.ok is None and r.error.startswith("形が違います")


def test_cookie_header_accepts_value_pair_and_header(acc):
    assert acc.cookie_header("abc") == "__Host-overtime_session=abc"
    assert acc.cookie_header(" __Host-overtime_session=abc ") == "__Host-overtime_session=abc"
    assert acc.cookie_header("Cookie: __Host-overtime_session=abc") == "__Host-overtime_session=abc"


def test_write_refusal_only_staging(acc):
    assert acc.write_refusal({"ok": True, "env": "staging"}) is None
    assert "production" in acc.write_refusal({"ok": True, "env": "production"})
    assert acc.write_refusal(None) is not None


# ===== 表と終了コード =====


def test_report_table_and_exit_code(acc):
    rep = acc.Report()
    rep.add("A-1", "通る", True, "見せない詳細")
    rep.add("A-2", "飛ばす", None, "理由")
    assert rep.exit_code() == 0
    rep.problems("A-3", "落ちる", ["一つ目", "二つ|目"])
    rep.problems("A-4", "問題なし", [])
    rep.problems("A-5", "理由つきで飛ばす", ["無視される"], skip_reason="なし")
    assert [i.status for i in rep.items] == ["PASS", "SKIP", "FAIL", "PASS", "SKIP"]
    assert rep.exit_code() == 1
    table = rep.table()
    lines = table.splitlines()
    assert lines[0] == "| # | 結果 | 項目 | 詳細 |"
    assert "| A-1 | PASS | 通る |  |" in lines  # PASS の詳細は出さない
    assert "| A-2 | SKIP | 飛ばす | 理由 |" in lines
    assert "| A-3 | FAIL | 落ちる | 一つ目; 二つ\\|目 |" in lines
    assert "PASS 2 / FAIL 1 / SKIP 2（全 5 項目）" in table and lines[-1] == "FAIL: A-3"
    data = rep.to_json()
    assert data["counts"] == {"PASS": 2, "FAIL": 1, "SKIP": 2} and data["exitCode"] == 1


# ===== 形の確かめ =====


CTX = {"today": "2026-11-26", "ym": "2026-11", "dept": "製造1課", "me": "owner@lineworks-local.info"}


@pytest.mark.parametrize("fn,result,bad", [
    ("api_getDeptList", ["製造1課"], False),
    ("api_getDeptList", [], True),
    ("api_getWorkersByDept", ["A001 山田"], False),
    ("api_getWorkersByDept", ["A001"], True),
    ("api_getWorkersByDept", [], True),  # 部署を渡して 0 人
    ("api_getWorkerInfo", None, False),
    ("api_getWorkerInfo", {"workerCode": "A"}, True),
    ("api_getOrderList", [{"code": "1", "product": "", "dest": ""}], False),
    ("api_getHolidayCandidateDates", [{"date": "2026-11-28", "label": "11/28(土)", "type": "weekend"}], False),
    ("api_getHolidayCandidateDates", [{"date": "2026/11/28", "label": "x", "type": "weekend"}], True),
    ("api_getDashboard", {"today": "2026-11-26", "overtime": [], "holiday": []}, False),
    ("api_getDashboard", {"today": "", "overtime": [], "holiday": [], "error": "boom"}, True),
    ("api_getDashboard", {"today": "2026-11-26", "overtime": [{"actualEndAt": "2026-11-26T18:00"}], "holiday": []},
     True),
    ("api_getApproverDashboard", {"today": "2026-11-26", "overtime": [], "holiday": [], "isAdmin": True,
                                  "selectedDepts": ["製造1課"]}, False),
    ("api_getApproverDepts", {"isAdmin": True, "approvers": [], "currentEmail": "owner@lineworks-local.info"}, False),
    ("api_getApproverDepts", {"isAdmin": False, "approvers": [], "currentEmail": "owner@lineworks-local.info"}, True),
    ("api_adminMonthlySummary", {"yearMonth": "2026-11", "kpi": {}, "people": [], "charts": {}}, False),
    ("api_adminMonthlySummary", {"yearMonth": "2026-10", "kpi": {}, "people": [], "charts": {}}, True),
    ("api_getSecondaryApprovalList", {"date": "2026-11-26", "items": []}, False),
    ("api_getSecondaryApprovalList", {"date": "2026-11-25", "items": []}, True),
    ("api_adminDeptOptions", ["ALL", "製造1課"], False),
    ("api_adminDeptOptions", ["製造1課"], True),
    ("api_adminFiscalTrend", {"fiscalYearStart": "", "fiscalYearEnd": "", "deptFilter": "ALL",
                              "labels": ["x"] * 12, "values": [0] * 12}, False),
    ("api_adminFiscalTrend", {"fiscalYearStart": "", "fiscalYearEnd": "", "deptFilter": "ALL",
                              "labels": [], "values": []}, True),
    ("api_adminDailyDetail", {"date": "2026-11-26", "deptFilter": "ALL", "total": 1, "pdfMissing": 0, "items": [
        {"requestId": "r", "targetDate": "2026-11-26", "pdfGeneratedAt": "", "actualStartAt": "",
         "actualEndAt": "", "isRetroactive": False, "daysAgo": 0}]}, False),
    ("api_adminDailyDetail", {"date": "2026-11-26", "deptFilter": "ALL", "total": 1, "pdfMissing": 0, "items": [
        {"requestId": "r", "targetDate": None, "pdfGeneratedAt": "", "actualStartAt": "",
         "actualEndAt": "", "isRetroactive": False, "daysAgo": 0}]}, True),  # GAS の null（#22 の直し）
    ("api_adminSpecialClauseCount", {"fiscalYear": "2026-04-01 ~ 2027-03-31", "list": []}, False),
    ("api_adminSpecialClauseCount", {"fiscalYear": "2026", "list": []}, True),
    ("api_adminPendingWatch", [], False),
    ("api_getClosingSummary", None, False),
    ("api_getClosingSummary", {"mode": "list", "meta": {}}, False),
    ("ops_getStatus", {"rowSync": {}, "gasJobs": {}, "gasPing": {}, "lastRuns": {}}, False),
    ("ops_getStatus", {"rowSync": {}}, True),
    ("api_unknown", {}, True),
])
def test_shape_problems(acc, fn, result, bad):
    assert bool(acc.shape_problems(fn, result, CTX)) is bad


def test_check_headers(acc):
    good = {"content-security-policy": acc.CSP_EXPECTED, "x-content-type-options": "nosniff",
            "x-frame-options": "SAMEORIGIN", "referrer-policy": "strict-origin-when-cross-origin",
            "strict-transport-security": "max-age=31536000", "permissions-policy": "camera=()"}
    assert acc.check_headers(good) == []
    bad = dict(good, **{"content-security-policy": "default-src *"})
    del bad["permissions-policy"]
    problems = acc.check_headers(bad)
    assert len(problems) == 2 and problems[0].startswith("CSP が違います")


def test_check_page_html(acc):
    ok = ('<title>残業・休日出勤申請</title><script type="application/json" id="gas-template-data">{}</script>'
          '<script src="/static/js/top.js"></script>')
    assert acc.check_page_html(ok) == []
    assert acc.check_page_html(ok + "<script>alert(1)</script>") == ["インラインの script があります"]
    assert acc.check_page_html(ok + '<button onclick="x()">') == ["on 属性があります"]
    assert acc.check_page_html("<title>x</title>") == ["題名が「残業・休日出勤申請」ではありません"]


def test_check_job_sequence(acc):
    exp = [{acc.K_AFTER_COMPLETE}, {acc.K_REGENERATE, acc.K_LEDGER_UPSERT}, {acc.K_AFTER_DELETE}]
    jobs = [{"id": 9, "kind": acc.K_AFTER_DELETE}, {"id": 3, "kind": acc.K_AFTER_COMPLETE},
            {"id": 5, "kind": acc.K_LEDGER_UPSERT}]
    assert acc.check_job_sequence(jobs, exp) == []  # id の順に並べ直して比べる
    jobs.append({"id": 6, "kind": acc.K_AFTER_APPROVE})
    assert acc.check_job_sequence(jobs, exp)[0].startswith("仕事の並びが違います")
    assert acc.check_job_sequence([], exp)[0].startswith("仕事の並びが違います: （なし）")


# ===== 本物の app に向けて動かす =====


def _transport(acc, client):
    def send(method, path, headers, body):
        r = client.request(method, path, headers=headers, content=body)
        return acc.HttpResponse(r.status_code, {k.lower(): v for k, v in r.headers.items()}, r.text)

    return send


def _fake_web(acc, headers):
    def send(method, path, h, body):
        if path == "/static/js/gas-shim.js":
            return acc.HttpResponse(200, dict(headers), "window.google = {};")
        return acc.HttpResponse(200, dict(headers), '{"ok":true,"env":"staging"}')

    return send


@pytest.fixture
def live(acc, ss, client, db, settings, session_factory):
    from app import rpc
    from test_domain_helpers import add_worker

    rpc.load_domain_handlers()
    add_worker(db, 1, "A001", "山田 太郎", "製造1課")
    add_worker(db, 2, "B001", "佐藤 花子", "製造2課")
    minted = ss.mint(db, settings)
    headers = {"content-security-policy": acc.CSP_EXPECTED, "x-content-type-options": "nosniff",
               "x-frame-options": "SAMEORIGIN", "referrer-policy": "strict-origin-when-cross-origin",
               "strict-transport-security": "max-age=31536000", "permissions-policy": "camera=()"}
    return SimpleNamespace(api=_transport(acc, client), web=_fake_web(acc, headers), cookie=minted.cookie_value,
                           minted=minted, probe=acc.DbJobProbe(session_factory))


def _by_id(report):
    return {i.id: i for i in report.items}


def test_runner_read_only_against_app(acc, live, session_factory):
    runner = acc.Runner(live.api, web=live.web, cookie=live.cookie, sebastian_token=TEST_ENV["SEBASTIAN_API_TOKEN"],
                        expect_env="staging", job_wait=0, log=lambda *_: None)
    rep = runner.run()
    items = _by_id(rep)
    # GAS の照合（gas_ping）はこの試験では一度も走っていないので FAIL になる。それ以外は通るか、理由つきで飛ばす。
    assert items["O-1"].status == "FAIL"
    failed = [(i.id, i.detail) for i in rep.items if i.status == "FAIL" and i.id != "O-1"]
    assert failed == []
    assert rep.exit_code() == 1
    for must in ("H-1", "H-2", "H-3", "E-1", "E-2", "E-3", "E-4", "E-5", "E-6", "E-7", "E-8", "E-9", "E-10", "E-11",
                 "S-0", "P-1", "P-5", "F-all", "F-api_adminDailyDetail", "F-api_submitRequest",
                 "F-api_approveRequest", "SB-1", "SB-2", "SB-3", "SB-4", "SB-5"):
        assert items[must].status == "PASS", (must, items[must].detail)
    skipped = {i.id for i in rep.items if i.status == "SKIP"}
    assert skipped == {"F-ops_syncSheetConfigNow", "F-ops_runMasterSync", "F-ops_recalcBreaks", "F-ops_rebuildLedger"}
    assert runner.me == "owner@lineworks-local.info"
    # 書かなかった（申請も仕事も無い）
    from app.models import Job, Request

    with session_factory() as s:
        assert s.scalars(select(Request)).first() is None and s.scalars(select(Job)).first() is None
    assert runner.created == []


def test_runner_without_session_or_web_skips_with_reasons(acc, live):
    rep = acc.Runner(live.api, cookie=None, sebastian_token=None, job_wait=0, log=lambda *_: None).run()
    items = _by_id(rep)
    assert items["R"].status == "SKIP" and items["E-10"].status == "SKIP" and items["SB-4"].status == "SKIP"
    assert items["E-3"].status == "PASS" and items["SB-2"].status == "PASS"
    assert rep.exit_code() == 0


def test_runner_reports_missing_security_headers(acc, live):
    rep = acc.Runner(live.api, web=live.api, cookie=live.cookie, job_wait=0, log=lambda *_: None).run()
    # api に直接つなぐと nginx のヘッダーが無い
    assert _by_id(rep)["E-10"].status == "FAIL"


def test_runner_write_scenario_against_app(acc, live, session_factory):
    from app.models import Job, Request

    runner = acc.Runner(live.api, web=live.web, cookie=live.cookie, sebastian_token=TEST_ENV["SEBASTIAN_API_TOKEN"],
                        probe=live.probe, write=True, job_wait=0, run_id="T1", log=lambda *_: None)
    rep = runner.run()
    items = _by_id(rep)
    w = {k: v for k, v in items.items() if k.startswith("W-")}
    assert {k: v.status for k, v in w.items() if v.status != "PASS"} == {"W-8": "SKIP", "W-11": "SKIP",
                                                                         "W-18": "SKIP"}, \
        [(k, v.detail) for k, v in w.items() if v.status == "FAIL"]
    assert items["F-ops_recalcBreaks"].status == "PASS"
    # 作った 2 件は【受け入れテスト】の印つきで、最後は canceled
    with session_factory() as s:
        reqs = list(s.scalars(select(Request).order_by(Request.submitted_at)))
        assert len(reqs) == 2 and len(runner.created) == 2
        assert all(r.work_content.startswith("【受け入れテスト】T1") and r.status == "canceled" for r in reqs)
        kinds = [j.kind for j in s.scalars(select(Job).order_by(Job.id))]
    assert kinds.count(acc.K_AFTER_APPROVE) == 1  # 2 つのタブで承認しても 1 つ
    assert runner.deleted == set(runner.created)


def test_runner_cleans_up_when_scenario_breaks(acc, live, session_factory, monkeypatch):
    from app.models import Request

    runner = acc.Runner(live.api, cookie=live.cookie, write=True, job_wait=0, run_id="T2", log=lambda *_: None)
    original = runner.complete

    def boom(rid):
        original(rid)
        raise RuntimeError("途中で止まった")

    monkeypatch.setattr(runner, "complete", boom)
    with pytest.raises(RuntimeError):
        runner.run()
    with session_factory() as s:
        reqs = list(s.scalars(select(Request)))
    assert len(reqs) == 1 and reqs[0].status == "canceled"
    assert any(n.startswith("後片付け:") for n in runner.report.notes)


def test_main_refuses_write_when_not_staging(acc, capsys):
    calls = []

    def api(method, path, headers, body):
        calls.append((method, path))
        return acc.HttpResponse(200, {}, '{"ok":true,"env":"production"}')

    assert acc.main(["--write", "--cookie", "x"], api=api) == 2
    assert calls == [("GET", "/healthz")]  # healthz の後は何も送らない
    assert "REFUSED" in capsys.readouterr().err


def test_main_prints_table_and_exit_code(acc, live, capsys, tmp_path):
    out = tmp_path / "r.json"
    code = acc.main(["--cookie", live.cookie, "--sebastian-token", TEST_ENV["SEBASTIAN_API_TOKEN"], "--job-wait", "0",
                     "--json-out", str(out)], api=live.api, web=live.web)
    text = capsys.readouterr().out
    assert code == 1  # gas_ping（O-1）
    assert text.splitlines()[0] == "| # | 結果 | 項目 | 詳細 |" and "FAIL: O-1" in text and "速さ（中央値 ms）" in text
    data = json.loads(out.read_text("utf-8"))
    assert data["exitCode"] == 1 and data["counts"]["FAIL"] == 1


# ===== staging_session =====


def test_mint_refuses_production_and_non_admin(ss, db, settings):
    with pytest.raises(ss.Refused):
        ss.mint(db, SimpleNamespace(app_env="production", bootstrap_admin_list=["owner@lineworks-local.info"]))
    with pytest.raises(ss.Refused):
        ss.mint(db, settings, "someone@lineworks-local.info")
    with pytest.raises(ss.Refused):
        ss.mint(db, settings, hours=48)


def test_mint_session_logs_in_and_does_not_outlive_hours(ss, db, settings, client):
    from app.models import UserSession

    m = ss.mint(db, settings, "OWNER@lineworks-local.info", hours=2)
    assert m.email == "owner@lineworks-local.info"
    r = client.get("/auth/me", headers={"Cookie": m.cookie_header})
    assert r.status_code == 200 and r.json()["email"] == "owner@lineworks-local.info"
    db.expire_all()
    row = db.get(UserSession, m.token_hash)
    assert row.name == ss.SESSION_NAME
    # 使っても期限は作ったときの期限より先に延びない（auth の 90 日の上限に合わせてある）
    assert row.expires_at <= m.expires_at + timedelta(seconds=1)


def test_revoke_and_purge_only_touch_acceptance_sessions(ss, db, settings, login_as):
    from app.models import UserSession

    login_as("owner@lineworks-local.info")  # 本物のログイン
    a = ss.mint(db, settings)
    ss.mint(db, settings)
    assert ss.revoke(db, a.token_hash) == 1
    assert ss.purge(db) == 1
    rows = list(db.scalars(select(UserSession)))
    assert len(rows) == 1 and rows[0].name != ss.SESSION_NAME
