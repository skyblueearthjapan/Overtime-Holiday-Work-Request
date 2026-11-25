"""受け入れテスト（HTTP で動いている環境を確かめる。実装仕様 12 章・docs/acceptance-checklist.md）。

項目ごとに PASS／FAIL／SKIP を表に出し、FAIL が 1 つでもあれば終了コード 1。書き込みの拒否は 2。

動かし方は 2 通り:

1) api コンテナと同じネットワークの中（おすすめ。DB を見て仕事の並びも確かめる）

   sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app \
       -v /opt/overtime/tools:/app/tools:ro api \
       python tools/qa/acceptance.py --base-url http://api:8000 --web-url http://web --mint-session [--write]"

   --mint-session: staging_session.py で BOOTSTRAP_ADMINS の 1 人目のセッションを作り、最後に必ず消す
   （staging でだけ作れる）。Sebastian のトークンは環境変数 SEBASTIAN_API_TOKEN を使う。

2) 外から（クッキーを渡す。DB は見ないので仕事の並びの項目は SKIP）

   python tools/qa/acceptance.py --base-url https://overtime.lineworks-local.net --cookie "<値>" [--write]

   クッキーの値は staging_session.py mint で作る（終わったら revoke）。

書き込みの試験（--write）:
- /healthz が env=staging を返すときだけ動く。それ以外は何も書かずに終了コード 2。
- 作る申請の業務内容（workContent）は「【受け入れテスト】<実行 ID>」で始まる。最後に 2 件とも削除（canceled）する。
- 流れ: 申請（同じ冪等キーで 2 回）→ 承認（2 つのタブ）→ 完了 → PDF を待つ → 二次承認 → 削除。
  もう 1 件で 編集 → 完了 → 完了後の編集の拒否 → 承認（2 つのタブ）→ 削除 → 2 回目の削除の拒否。
- DB を見られるとき（--mint-session か --db-probe）は、申請ごとの GAS の仕事が決まった順に積まれ、失敗が無いことも見る。

書き込みをしないとき（既定）:
- 書く関数は、staging でだけ「無い申請 ID」などで呼び、GAS と同じ文の 400 で止まることを見る（何も書かない）。
  production では書く関数を呼ばない（監査の記録も残さないため）。
- 重い管理の関数（台帳の作り直し）はどちらでも呼ばない（SKIP）。共有タブの取り込み・マスタ同期・休憩の再計算は --write のときだけ。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import statistics
import sys
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable, Iterable

PASS, FAIL, SKIP = "PASS", "FAIL", "SKIP"
TEST_PREFIX = "【受け入れテスト】"
JST = timezone(timedelta(hours=9))
DEFAULT_ORIGIN = "https://overtime.lineworks-local.net"
SESSION_COOKIE = "__Host-overtime_session"
USER_AGENT = "overtime-acceptance/1"
BOGUS_ID = "ACCEPTANCE-NO-SUCH-REQUEST"

# nginx のヘッダー（実装仕様 3-7）
CSP_EXPECTED = ("default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; "
                "connect-src 'self'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'")
HEADERS_EXPECTED = {
    "x-content-type-options": "nosniff",
    "x-frame-options": "SAMEORIGIN",
    "referrer-policy": "strict-origin-when-cross-origin",
}
HEADERS_PRESENT = ("strict-transport-security", "permissions-policy")

# 閉じたリスト（app/rpc.py と同じ。test_acceptance_runner.py が食い違いを止める）
SCREEN_FUNCTIONS: tuple[str, ...] = (
    "api_getDeptList", "api_getWorkersByDept", "api_getWorkerInfo", "api_getOrderList",
    "api_getHolidayCandidateDates", "api_getDashboard", "api_submitRequest", "api_markOvertimeDone",
    "api_markHolidayStart", "api_markHolidayDone", "api_editRequest", "api_deleteRequest", "api_getApproverDepts",
    "api_getApproverDashboard", "api_approveRequest", "api_rejectRequest", "api_approveRequestsBatch",
    "api_approverMonthlySummary", "api_getSecondaryApprovalList", "api_secondaryApprove",
    "api_secondaryApproveBatch", "api_adminDeptOptions", "api_adminMonthlySummary", "api_adminFiscalTrend",
    "api_adminDailyDetail", "api_adminManualComplete", "api_adminSpecialClauseCount", "api_adminPendingWatch",
    "api_getClosingSummary",
)
OPS_FUNCTIONS: tuple[str, ...] = (
    "ops_getStatus", "ops_retryJob", "ops_syncSheetConfigNow", "ops_runMasterSync", "ops_regeneratePdf",
    "ops_batchPdfForDate", "ops_rebuildLedger", "ops_recalcBreaks", "ops_revokeSessions",
)
RPC_FUNCTIONS: tuple[str, ...] = SCREEN_FUNCTIONS + OPS_FUNCTIONS
WRITE_FUNCTIONS = frozenset({
    "api_submitRequest", "api_markOvertimeDone", "api_markHolidayStart", "api_markHolidayDone", "api_editRequest",
    "api_deleteRequest", "api_approveRequest", "api_rejectRequest", "api_approveRequestsBatch",
    "api_secondaryApprove", "api_secondaryApproveBatch", "api_adminManualComplete", "ops_retryJob",
    "ops_syncSheetConfigNow", "ops_runMasterSync", "ops_regeneratePdf", "ops_batchPdfForDate", "ops_rebuildLedger",
    "ops_recalcBreaks", "ops_revokeSessions",
})

# GAS の仕事の種類（services/gas_jobs.py の jobs.kind）
K_AFTER_APPROVE = "internal.afterApprove"
K_AFTER_COMPLETE = "internal.afterComplete"
K_REGENERATE = "internal.regenerate"
K_LEDGER_UPSERT = "internal.ledgerUpsert"
K_AFTER_DELETE = "internal.afterDelete"
JOB_FINISHED_OK = ("done", "skipped")

YMD_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
YMD_HMS_RE = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
ISO_MS_Z_RE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$")
INLINE_SCRIPT_RE = re.compile(r"<script(?![^>]*\bsrc=)(?![^>]*application/json)[^>]*>", re.I)
ON_ATTR_RE = re.compile(r"<[^>]+\son[a-z]+\s*=", re.I)


# ===== HTTP =====


@dataclass
class HttpResponse:
    status: int
    headers: dict[str, str]  # 小文字のキー
    text: str

    def json(self) -> Any:
        return json.loads(self.text)


Transport = Callable[[str, str, dict, "bytes | None"], HttpResponse]


class UrllibTransport:
    """標準ライブラリだけの送り手（リダイレクトは追わない）。base_url + path に送る。"""

    def __init__(self, base_url: str, timeout: float = 120.0):
        import urllib.request

        self.base_url = base_url.rstrip("/")
        self.timeout = timeout

        class _NoRedirect(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, *a, **k):  # noqa: ANN002
                return None

        self._opener = urllib.request.build_opener(_NoRedirect)

    def __call__(self, method: str, path: str, headers: dict, body: bytes | None) -> HttpResponse:
        import urllib.error
        import urllib.request

        req = urllib.request.Request(self.base_url + path, data=body, method=method,
                                     headers={"User-Agent": USER_AGENT, **headers})
        try:
            with self._opener.open(req, timeout=self.timeout) as r:
                return HttpResponse(r.status, {k.lower(): v for k, v in r.headers.items()},
                                    r.read().decode("utf-8", "replace"))
        except urllib.error.HTTPError as e:
            return HttpResponse(e.code, {k.lower(): v for k, v in (e.headers or {}).items()},
                                e.read().decode("utf-8", "replace") if e.fp else "")


# ===== RPC の返事 =====


@dataclass
class RpcReply:
    status: int
    ok: bool | None
    result: Any = None
    error: str | None = None
    raw: str = ""
    ms: float = 0.0

    def brief(self) -> str:
        if self.status == 200 and self.ok:
            return "200 " + _short(self.result)
        return f"{self.status} {self.error if self.error is not None else _short(self.raw)}"


def parse_rpc(resp: HttpResponse, ms: float = 0.0) -> RpcReply:
    """{"ok":true,"result":R} / {"ok":false,"error":msg} を読む。JSON でなければ ok=None。"""
    try:
        body = resp.json()
    except ValueError:
        return RpcReply(resp.status, None, error="JSON ではありません: " + _short(resp.text), raw=resp.text, ms=ms)
    if not isinstance(body, dict) or "ok" not in body:
        return RpcReply(resp.status, None, error="形が違います: " + _short(resp.text), raw=resp.text, ms=ms)
    if body["ok"] is True:
        return RpcReply(resp.status, True, result=body.get("result"), raw=resp.text, ms=ms)
    return RpcReply(resp.status, False, error=str(body.get("error", "")), raw=resp.text, ms=ms)


def _short(value: Any, n: int = 300) -> str:
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False, default=str)
    return text if len(text) <= n else text[:n] + "…"


# ===== 表 =====


@dataclass
class Item:
    id: str
    title: str
    status: str
    detail: str = ""


@dataclass
class Report:
    items: list[Item] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def add(self, item_id: str, title: str, ok: bool | None, detail: Any = "") -> Item:
        """ok=True→PASS、False→FAIL、None→SKIP（detail に理由）。"""
        status = SKIP if ok is None else (PASS if ok else FAIL)
        text = detail if isinstance(detail, str) else _short(detail)
        item = Item(item_id, title, status, text)
        self.items.append(item)
        return item

    def problems(self, item_id: str, title: str, problems: list[str], skip_reason: str | None = None) -> Item:
        if skip_reason is not None:
            return self.add(item_id, title, None, skip_reason)
        return self.add(item_id, title, not problems, "; ".join(problems))

    def count(self, status: str) -> int:
        return sum(1 for i in self.items if i.status == status)

    @property
    def failed(self) -> list[Item]:
        return [i for i in self.items if i.status == FAIL]

    def exit_code(self) -> int:
        return 1 if self.failed else 0

    def table(self) -> str:
        lines = ["| # | 結果 | 項目 | 詳細 |", "|---|---|---|---|"]
        for i in self.items:
            detail = i.detail.replace("|", "\\|").replace("\n", " ")
            lines.append(f"| {i.id} | {i.status} | {i.title} | {detail if i.status != PASS else ''} |")
        lines.append("")
        lines.append(f"PASS {self.count(PASS)} / FAIL {self.count(FAIL)} / SKIP {self.count(SKIP)}"
                     f"（全 {len(self.items)} 項目）")
        if self.failed:
            lines.append("FAIL: " + ", ".join(i.id for i in self.failed))
        return "\n".join(lines)

    def to_json(self) -> dict:
        return {"items": [i.__dict__ for i in self.items], "notes": self.notes,
                "counts": {s: self.count(s) for s in (PASS, FAIL, SKIP)}, "exitCode": self.exit_code()}


# ===== 返す形の確かめ（GAS の形。実装仕様 4-3） =====


def missing_keys(obj: Any, keys: Iterable[str]) -> list[str]:
    if not isinstance(obj, dict):
        return [f"dict ではありません: {_short(obj, 80)}"]
    return [f"キー {k} がありません" for k in keys if k not in obj]


def _is_list_of(value: Any, typ: type) -> bool:
    return isinstance(value, list) and all(isinstance(v, typ) for v in value)


def shape_problems(fn: str, result: Any, ctx: dict) -> list[str]:
    """読む関数の返事が GAS の形か。ctx: today・ym・dept・me。"""
    p: list[str] = []
    if fn == "api_getDeptList":
        if not _is_list_of(result, str) or not result:
            p.append("部署名の配列（1 件以上）ではありません")
    elif fn == "api_getWorkersByDept":
        if not _is_list_of(result, str):
            p.append("文字の配列ではありません")
        elif ctx.get("dept") and not result:
            p.append("部署の作業員が 0 人です")
        elif any(" " not in s for s in result):
            p.append("「<コード> <氏名>」の形でない行があります")
    elif fn == "api_getWorkerInfo":
        if result is not None:
            p += missing_keys(result, ("workerCode", "workerName", "dept", "email"))
    elif fn == "api_getOrderList":
        if not isinstance(result, list):
            p.append("配列ではありません")
        elif result:
            p += missing_keys(result[0], ("code", "product", "dest"))
    elif fn == "api_getHolidayCandidateDates":
        if not isinstance(result, list):
            p.append("配列ではありません")
        for d in result if isinstance(result, list) else []:
            p += missing_keys(d, ("date", "label", "type"))
            if isinstance(d, dict) and (not YMD_RE.match(str(d.get("date", ""))) or
                                        d.get("type") not in ("weekend", "holiday")):
                p.append("date か type の形が違います: " + _short(d, 80))
                break
    elif fn in ("api_getDashboard", "api_getApproverDashboard"):
        keys = ("today", "overtime", "holiday") + (("isAdmin", "selectedDepts") if fn != "api_getDashboard" else ())
        p += missing_keys(result, keys)
        if isinstance(result, dict):
            if result.get("error"):
                p.append("error: " + str(result["error"]))
            if not YMD_RE.match(str(result.get("today", ""))):
                p.append("today が yyyy-MM-dd ではありません")
            for it in (result.get("overtime") or []) + (result.get("holiday") or []):
                for k in ("actualStartAt", "actualEndAt"):
                    v = it.get(k, "") if isinstance(it, dict) else ""
                    if v and not YMD_HMS_RE.match(str(v)):
                        p.append(f"{k} の形が違います: {v}")
                        break
    elif fn == "api_getApproverDepts":
        p += missing_keys(result, ("isAdmin", "approvers", "currentEmail"))
        if isinstance(result, dict):
            if result.get("isAdmin") is not True:
                p.append("管理者のセッションなのに isAdmin が true ではありません")
            if ctx.get("me") and result.get("currentEmail") != ctx["me"]:
                p.append("currentEmail がログインのメールと違います")
    elif fn in ("api_approverMonthlySummary", "api_adminMonthlySummary"):
        p += missing_keys(result, ("yearMonth", "kpi", "people", "charts"))
        if isinstance(result, dict) and result.get("yearMonth") != ctx.get("ym"):
            p.append("yearMonth が違います: " + str(result.get("yearMonth")))
    elif fn == "api_getSecondaryApprovalList":
        p += missing_keys(result, ("date", "items"))
        if isinstance(result, dict):
            if result.get("date") != ctx.get("today"):
                p.append("date が今日ではありません")
            for it in result.get("items") or []:
                p += missing_keys(it, ("requestId", "status", "targetDate", "isRetroactive", "daysAgo"))
                break
    elif fn == "api_adminDeptOptions":
        if not _is_list_of(result, str) or not result or result[0] != "ALL":
            p.append("先頭が ALL の配列ではありません")
    elif fn == "api_adminFiscalTrend":
        p += missing_keys(result, ("fiscalYearStart", "fiscalYearEnd", "deptFilter", "labels", "values"))
        if isinstance(result, dict) and (len(result.get("labels") or []) != 12 or len(result.get("values") or []) != 12):
            p.append("labels・values が 12 か月ではありません")
    elif fn == "api_adminDailyDetail":
        p += missing_keys(result, ("date", "deptFilter", "total", "pdfMissing", "items"))
        for it in (result.get("items") or []) if isinstance(result, dict) else []:
            # 9-1 #22: 日時を文字で返す（GAS は null だった）
            for k in ("targetDate", "pdfGeneratedAt", "actualStartAt", "actualEndAt"):
                if not isinstance(it.get(k), str):
                    p.append(f"{k} が文字ではありません（{it.get('requestId')}）")
            if not isinstance(it.get("isRetroactive"), bool) or not isinstance(it.get("daysAgo"), int):
                p.append(f"isRetroactive・daysAgo がありません（{it.get('requestId')}）")
            if p:
                break
    elif fn == "api_adminSpecialClauseCount":
        p += missing_keys(result, ("fiscalYear", "list"))
        if isinstance(result, dict) and not re.match(r"^\d{4}-04-01 ~ \d{4}-03-31$", str(result.get("fiscalYear"))):
            p.append("fiscalYear の形が違います: " + str(result.get("fiscalYear")))
    elif fn == "api_adminPendingWatch":
        if not isinstance(result, list):
            p.append("配列ではありません")
    elif fn == "api_getClosingSummary":
        if result is not None:
            p += missing_keys(result, ("mode", "meta"))
    elif fn == "ops_getStatus":
        p += missing_keys(result, ("rowSync", "gasJobs", "gasPing", "lastRuns"))
    else:
        p.append("形の確かめがありません: " + fn)
    return p


def check_headers(headers: dict[str, str]) -> list[str]:
    p = []
    csp = headers.get("content-security-policy")
    if csp != CSP_EXPECTED:
        p.append("CSP が違います: " + str(csp))
    for k, v in HEADERS_EXPECTED.items():
        if headers.get(k) != v:
            p.append(f"{k} が {v} ではありません: {headers.get(k)}")
    for k in HEADERS_PRESENT:
        if not headers.get(k):
            p.append(f"{k} がありません")
    return p


def check_page_html(text: str) -> list[str]:
    p = []
    if "<title>残業・休日出勤申請</title>" not in text:
        p.append("題名が「残業・休日出勤申請」ではありません")
    if INLINE_SCRIPT_RE.search(text):
        p.append("インラインの script があります")
    if ON_ATTR_RE.search(text):
        p.append("on 属性があります")
    return p


def check_job_sequence(jobs: list[dict], expected: list[set[str]]) -> list[str]:
    """jobs（id の順）の kind が expected（各位置で受け入れる kind の集合）と同じ並びか。"""
    ordered = sorted(jobs, key=lambda j: j["id"])
    kinds = [j["kind"] for j in ordered]
    if len(kinds) != len(expected) or any(k not in allowed for k, allowed in zip(kinds, expected)):
        want = " → ".join("|".join(sorted(a)) for a in expected)
        return [f"仕事の並びが違います: {' → '.join(kinds) or '（なし）'}（期待: {want}）"]
    return []


def write_refusal(health: dict | None) -> str | None:
    """--write を断る理由（staging のときだけ None）。"""
    env = (health or {}).get("env")
    if env != "staging":
        return f"/healthz の env が staging ではありません（{env}）。書き込みの試験はステージングでだけ動かします。"
    return None


def cookie_header(value: str) -> str:
    value = value.strip()
    if value.startswith(SESSION_COOKIE + "="):
        return value
    if value.lower().startswith("cookie:"):
        return value.split(":", 1)[1].strip()
    return f"{SESSION_COOKIE}={value}"


# ===== DB を見る（コンテナの中だけ） =====


class DbJobProbe:
    """jobs 表を申請 ID で読む。"""

    def __init__(self, session_factory):
        self.session_factory = session_factory

    @classmethod
    def from_app(cls) -> "DbJobProbe":
        from app.db import get_sessionmaker

        return cls(get_sessionmaker())

    def jobs_for(self, request_id: str) -> list[dict]:
        from sqlalchemy import select

        from app.models import Job

        with self.session_factory() as s:
            rows = s.scalars(select(Job).where(Job.request_id == request_id).order_by(Job.id))
            return [{"id": r.id, "kind": r.kind, "state": r.state, "error": r.last_error or ""} for r in rows]


# ===== 本体 =====


class Runner:
    def __init__(self, api: Transport, *, web: Transport | None = None, origin: str = DEFAULT_ORIGIN,
                 cookie: str | None = None, sebastian_token: str | None = None, probe: DbJobProbe | None = None,
                 write: bool = False, expect_env: str | None = None, job_wait: float = 240.0,
                 poll_interval: float = 3.0, dept: str | None = None, worker_label: str | None = None,
                 run_id: str | None = None, sleep: Callable[[float], None] = time.sleep,
                 clock: Callable[[], float] = time.monotonic, log: Callable[[str], None] = print):
        self.api = api
        self.web = web
        self.origin = origin.rstrip("/")
        self.cookie = cookie_header(cookie) if cookie else None
        self.sebastian_token = sebastian_token
        self.probe = probe
        self.write = write
        self.expect_env = expect_env
        self.job_wait = job_wait
        self.poll_interval = poll_interval
        self.dept = dept
        self.worker_label = worker_label
        self.run_id = run_id or datetime.now(JST).strftime("%Y%m%d-%H%M%S")
        self.sleep = sleep
        self.clock = clock
        self.log = log
        self.report = Report()
        self.times: dict[str, list[float]] = {}
        self.health: dict | None = None
        self.env: str | None = None
        self.me: str = ""
        self.today: str = datetime.now(JST).strftime("%Y-%m-%d")
        self.created: list[str] = []
        self.deleted: set[str] = set()

    # ---- 送る ----

    def http(self, method: str, path: str, headers: dict | None = None, body: bytes | None = None,
             with_cookie: bool = False, transport: Transport | None = None) -> HttpResponse:
        h = dict(headers or {})
        if with_cookie and self.cookie:
            h["Cookie"] = self.cookie
        return (transport or self.api)(method, path, h, body)

    def rpc(self, fn: str, *args: Any, key: str | None = "auto", origin: str | None = None,
            content_type: str = "application/json", with_cookie: bool = True, raw_body: bytes | None = None
            ) -> RpcReply:
        headers = {"Content-Type": content_type, "Origin": origin or self.origin}
        if fn in WRITE_FUNCTIONS and key:
            headers["X-Idempotency-Key"] = str(uuid.uuid4()) if key == "auto" else key
        body = raw_body if raw_body is not None else json.dumps(list(args), ensure_ascii=False).encode("utf-8")
        t = time.perf_counter()
        resp = self.http("POST", "/api/rpc/" + fn, headers, body, with_cookie=with_cookie)
        ms = (time.perf_counter() - t) * 1000
        self.times.setdefault(fn, []).append(ms)
        return parse_rpc(resp, ms)

    def sebastian(self, body: bytes) -> HttpResponse:
        return self.http("POST", "/api/sebastian", {"Content-Type": "application/json"}, body)

    # ---- 項目 ----

    def run(self) -> Report:
        self.check_health()
        if self.write:
            reason = write_refusal(self.health)
            if reason:
                raise WriteRefused(reason)
        self.check_entry()
        self.check_web_headers()
        if self.cookie:
            self.check_session()
            self.check_pages()
            self.check_rpcs()
        else:
            self.report.add("R", "ログインしての確かめ（画面・38 関数・書き込み）", None,
                            "セッションがありません（--mint-session か --cookie）")
        self.check_sebastian()
        if self.write and self.cookie:
            self.write_scenario()
        elif self.write:
            self.report.add("W", "書き込みの試験", False, "セッションがありません")
        self.report.notes.append("速さ（中央値 ms）: " + json.dumps(
            {k: round(statistics.median(v), 1) for k, v in sorted(self.times.items())}, ensure_ascii=False))
        return self.report

    def check_health(self) -> None:
        r = self.http("GET", "/healthz")
        try:
            body = r.json()
        except ValueError:
            body = None
        ok = r.status == 200 and isinstance(body, dict) and body.get("ok") is True and body.get("env") in (
            "staging", "production")
        self.health = body if isinstance(body, dict) else None
        self.env = (self.health or {}).get("env")
        self.report.add("H-1", "/healthz が 200 {ok:true, env}", ok, f"{r.status} {_short(r.text)}")
        if self.expect_env:
            self.report.add("H-2", f"/healthz の env が {self.expect_env}", self.env == self.expect_env,
                            f"env={self.env}")
        if isinstance(body, dict):
            extra = sorted(set(body) - {"ok", "env"})
            self.report.add("H-3", "ログインなしの /healthz は ok と env だけ（詳細を出さない）", not extra,
                            f"余分なキー: {extra}")

    def check_entry(self) -> None:
        r = self.http("GET", "/")
        loc = r.headers.get("location", "")
        self.report.add("E-1", "/ はログインなしで 302 /auth/login?next=…", r.status in (302, 303, 307) and
                        loc.startswith("/auth/login?next="), f"{r.status} location={loc}")
        r = self.http("GET", "/?page=approve2&date=2026-11-25")
        loc = r.headers.get("location", "")
        self.report.add("E-2", "ログインへの next が page と date を保つ", r.status in (302, 303, 307) and
                        "approve2" in loc and "2026-11-25" in loc, f"{r.status} location={loc}")
        reply = self.rpc("api_getDeptList", with_cookie=False)
        want = {"ok": False, "error": "ログインが必要です。", "login": "/auth/login"}
        try:
            body = json.loads(reply.raw)
        except ValueError:
            body = None
        self.report.add("E-3", "ログインなしの RPC は 401 {ok:false, ログインが必要です。, login}",
                        reply.status == 401 and body == want, reply.brief())
        r = self.http("GET", "/auth/me")
        self.report.add("E-4", "ログインなしの /auth/me は 401", r.status == 401, f"{r.status}")
        if not self.cookie:
            for i, t in (("E-5", "JSON でない本文は 415"), ("E-6", "別サイトの Origin は 403"),
                         ("E-7", "知らない関数は 404"), ("E-8", "出さない GAS の関数（api_getFormUrl）は 404"),
                         ("E-9", "本文が配列でなければ 400")):
                self.report.add(i, t, None, "セッションがありません（401 が先に返る）")
            return
        r = self.rpc("api_getDeptList", content_type="text/plain")
        self.report.add("E-5", "JSON でない本文は 415", r.status == 415 and r.error == "呼び出しの形が正しくありません。",
                        r.brief())
        r = self.rpc("api_getDeptList", origin="https://evil.example")
        self.report.add("E-6", "別サイトの Origin は 403", r.status == 403 and r.error == "この画面からは操作できません。",
                        r.brief())
        r = self.rpc("api_noSuchFunction")
        self.report.add("E-7", "知らない関数は 404", r.status == 404 and r.error == "その操作はありません。", r.brief())
        r = self.rpc("api_getFormUrl")
        self.report.add("E-8", "出さない GAS の関数（api_getFormUrl）は 404", r.status == 404, r.brief())
        r = self.rpc("api_getDeptList", raw_body=b'{"a":1}')
        self.report.add("E-9", "本文が配列でなければ 400", r.status == 400, r.brief())

    def check_web_headers(self) -> None:
        if self.web is None:
            self.report.add("E-10", "CSP などのセキュリティヘッダー（nginx）", None, "--web-url がありません")
            return
        r = self.http("GET", "/healthz", transport=self.web)
        self.report.problems("E-10", "CSP などのセキュリティヘッダー（nginx）", check_headers(r.headers))
        r = self.http("GET", "/static/js/gas-shim.js", transport=self.web)
        self.report.add("E-11", "/static/js/gas-shim.js を nginx が返す", r.status == 200 and "google" in r.text,
                        f"{r.status}")

    def check_session(self) -> None:
        r = self.http("GET", "/auth/me", with_cookie=True)
        try:
            body = r.json()
        except ValueError:
            body = {}
        self.me = str(body.get("email", "")) if isinstance(body, dict) else ""
        self.report.add("S-0", "セッションで /auth/me が 200 {ok, email}", r.status == 200 and bool(self.me),
                        f"{r.status} {_short(r.text)}")

    def check_pages(self) -> None:
        for i, page in enumerate(("top", "approver", "admin", "approve2", "staffSummary"), start=1):
            r = self.http("GET", "/?page=" + page, with_cookie=True)
            problems = [] if r.status == 200 else [f"HTTP {r.status}"]
            if r.status == 200:
                problems += check_page_html(r.text)
                if r.headers.get("cache-control") != "no-store":
                    problems.append("Cache-Control が no-store ではありません")
            self.report.problems(f"P-{i}", f"画面 ?page={page} が 200 で開く（題名・インライン script なし）", problems)

    # ---- 38 関数 ----

    def check_rpcs(self) -> None:
        depts = self.rpc("api_getDeptList")
        dept = self.dept or (depts.result[0] if depts.ok and isinstance(depts.result, list) and depts.result else "")
        dash = self.rpc("api_getDashboard")
        if dash.ok and isinstance(dash.result, dict) and YMD_RE.match(str(dash.result.get("today", ""))):
            self.today = dash.result["today"]
        ym = self.today[:7]
        ctx = {"today": self.today, "ym": ym, "dept": dept, "me": self.me}
        self._ctx = ctx
        read_calls: list[tuple[str, tuple]] = [
            ("api_getDeptList", ()), ("api_getWorkersByDept", (dept,)), ("api_getWorkerInfo", ()),
            ("api_getOrderList", ()), ("api_getHolidayCandidateDates", ()), ("api_getDashboard", ()),
            ("api_getApproverDepts", ()), ("api_getApproverDashboard", ([dept],)),
            ("api_approverMonthlySummary", (ym, [dept])), ("api_getSecondaryApprovalList", (self.today,)),
            ("api_adminDeptOptions", ()), ("api_adminMonthlySummary", (ym, "ALL")),
            ("api_adminFiscalTrend", (self.today, "ALL")), ("api_adminDailyDetail", (self.today, "ALL")),
            ("api_adminSpecialClauseCount", (self.today, "ALL")), ("api_adminPendingWatch", ()),
            ("api_getClosingSummary", (ym, dept, "")), ("ops_getStatus", ()),
        ]
        seen: set[str] = set()
        for fn, args in read_calls:
            r = self.rpc(fn, *args)
            problems = [] if (r.status == 200 and r.ok) else [r.brief()]
            if not problems:
                problems = shape_problems(fn, r.result, ctx)
            self.report.problems(f"F-{fn}", f"{fn}{_args_text(args)} が GAS の形を返す", problems)
            seen.add(fn)
            if fn == "ops_getStatus" and r.ok and isinstance(r.result, dict):
                ping = r.result.get("gasPing") or {}
                self.report.add("O-1", "GAS の内部用の入口の照合（gas_ping）が通っている", ping.get("ok") is True and
                                not ping.get("problems"), _short(ping))
                failed = ((r.result.get("gasJobs") or {}).get("counts") or {}).get("failed", 0)
                self.report.add("O-2", "失敗した GAS の仕事が 0 件", not failed, f"failed={failed}")
                blocked = (r.result.get("rowSync") or {}).get("blocked") or []
                self.report.add("O-3", "止まっている行の反映が 0 件", not blocked, _short(blocked))
        r = self.rpc("api_getWorkersByDept", "")
        self.report.add("F-api_getWorkersByDept-empty", "api_getWorkersByDept('') は []", r.ok and r.result == [],
                        r.brief())

        # 書く関数: 何も書かない引数で、GAS と同じ文で止まること（staging だけ）
        not_found = "申請が見つかりません。"
        approval_missing = ("requestIdが見つかりません。", "Requestsが空です。")
        probes: list[tuple[str, tuple, Callable[[RpcReply], bool]]] = [
            ("api_submitRequest", ({},), lambda r: r.status == 200 and r.result == {
                "ok": False, "error": "申請種別が不正です: "}),
            ("api_markOvertimeDone", (BOGUS_ID,), lambda r: r.status == 400 and r.error == not_found),
            ("api_markHolidayStart", (BOGUS_ID,), lambda r: r.status == 400 and r.error == not_found),
            ("api_markHolidayDone", (BOGUS_ID,), lambda r: r.status == 400 and r.error == not_found),
            ("api_editRequest", (BOGUS_ID, {"hours": "1.0"}), lambda r: r.status == 400 and r.error == not_found),
            ("api_deleteRequest", (BOGUS_ID,), lambda r: r.status == 400 and r.error == not_found),
            ("api_approveRequest", (BOGUS_ID,), lambda r: r.status == 400 and r.error in approval_missing),
            ("api_rejectRequest", (BOGUS_ID,), lambda r: r.status == 400 and r.error in approval_missing),
            ("api_approveRequestsBatch", ([],), lambda r: r.ok is True and r.result == {"ok": True, "results": []}),
            ("api_secondaryApprove", (BOGUS_ID,), lambda r: r.status == 400 and r.error in approval_missing),
            ("api_secondaryApproveBatch", ([],), lambda r: r.ok is True and r.result == {"ok": True, "results": []}),
            ("api_adminManualComplete", (BOGUS_ID, "18:00"), lambda r: r.status == 400 and r.error == not_found),
            ("ops_retryJob", ("x",), lambda r: r.status == 400 and r.error == "仕事の番号が正しくありません。"),
            ("ops_regeneratePdf", (BOGUS_ID,), lambda r: r.status == 400 and r.error == not_found),
            ("ops_batchPdfForDate", ("bad",), lambda r: r.status == 400 and r.error == "日付が不正です（yyyy-MM-dd）"),
            ("ops_revokeSessions", ("",), lambda r: r.status == 400 and r.error == "メールアドレスを指定してください。"),
        ]
        for fn, args, good in probes:
            title = f"{fn}{_args_text(args)} は何も書かずに GAS と同じ文で止まる"
            seen.add(fn)
            if self.env != "staging":
                self.report.add(f"F-{fn}", title, None, "staging ではないので書く関数は呼ばない")
                continue
            r = self.rpc(fn, *args)
            self.report.add(f"F-{fn}", title, good(r), r.brief())
        r = self.rpc("api_submitRequest", {"requestType": "overtime"}) if self.env == "staging" else None
        if r is not None:
            self.report.add("F-api_submitRequest-dept", "申請の入力の誤りは 200 {ok:false, error}（部署なし）",
                            r.status == 200 and r.result == {"ok": False, "error": "部署が指定されていません。"},
                            r.brief())
        for fn, why in (("ops_syncSheetConfigNow", "--write のときに呼ぶ"), ("ops_runMasterSync", "--write のときに呼ぶ"),
                        ("ops_recalcBreaks", "--write のときに呼ぶ"),
                        ("ops_rebuildLedger", "台帳を作り直す重い処理なので手で確かめる（JEV）")):
            seen.add(fn)
            if not (self.write and fn != "ops_rebuildLedger"):
                self.report.add(f"F-{fn}", f"{fn}()", None, why)
        missing = [fn for fn in RPC_FUNCTIONS if fn not in seen]
        self.report.add("F-all", "38 関数をすべて確かめの対象にした", not missing, f"漏れ: {missing}")

    # ---- Sebastian ----

    def check_sebastian(self) -> None:
        r = self.sebastian(b"not json")
        body = _json_or_none(r.text)
        self.report.add("SB-1", "Sebastian: JSON でない本文は 200 {error: invalid JSON: …}", r.status == 200 and
                        isinstance(body, dict) and str(body.get("error", "")).startswith("invalid JSON: ") and
                        set(body) == {"error"}, f"{r.status} {_short(r.text)}")
        r = self.sebastian(json.dumps({"action": "ping", "token": "wrong-token-" + uuid.uuid4().hex}).encode())
        self.report.add("SB-2", "Sebastian: トークン違いは 200 {error: unauthorized}",
                        r.status == 200 and _json_or_none(r.text) == {"error": "unauthorized"},
                        f"{r.status} {_short(r.text)}")
        if not self.sebastian_token:
            for i, t in (("SB-3", "Sebastian: 知らない action"), ("SB-4", "Sebastian: ping"),
                         ("SB-5", "Sebastian: requestId なしの completeOvertime")):
                self.report.add(i, t, None, "トークンがありません（SEBASTIAN_API_TOKEN）")
            return
        tok = self.sebastian_token
        r = self.sebastian(json.dumps({"action": "acceptanceNoSuchAction", "token": tok}).encode())
        self.report.add("SB-3", "Sebastian: 知らない action は {error: unknown action: X, action: X}",
                        r.status == 200 and _json_or_none(r.text) == {
                            "error": "unknown action: acceptanceNoSuchAction", "action": "acceptanceNoSuchAction"},
                        f"{r.status} {_short(r.text)}")
        r = self.sebastian(json.dumps({"action": "ping", "token": tok}).encode())
        body = _json_or_none(r.text)
        self.report.add("SB-4", "Sebastian: ping は {ok:true, pong:<ISO ミリ秒 Z>}", r.status == 200 and
                        isinstance(body, dict) and body.get("ok") is True and
                        bool(ISO_MS_Z_RE.match(str(body.get("pong", "")))), f"{r.status} {_short(r.text)}")
        r = self.sebastian(json.dumps({"action": "completeOvertime", "token": tok}).encode())
        self.report.add("SB-5", "Sebastian: requestId なしは {error: requestId is required, action}（書かない）",
                        r.status == 200 and _json_or_none(r.text) == {"error": "requestId is required",
                                                                      "action": "completeOvertime"},
                        f"{r.status} {_short(r.text)}")

    # ---- 書き込みの試験（staging だけ） ----

    def daily_item(self, request_id: str) -> dict | None:
        r = self.rpc("api_adminDailyDetail", self.today, "ALL")
        if not (r.ok and isinstance(r.result, dict)):
            return None
        for it in r.result.get("items") or []:
            if it.get("requestId") == request_id:
                return it
        return None

    def secondary_item(self, request_id: str) -> dict | None:
        r = self.rpc("api_getSecondaryApprovalList", self.today)
        for it in (r.result or {}).get("items") or [] if r.ok else []:
            if it.get("requestId") == request_id:
                return it
        return None

    def wait_for(self, predicate: Callable[[], Any]) -> Any:
        deadline = self.clock() + self.job_wait
        while True:
            value = predicate()
            if value or self.clock() >= deadline:
                return value
            self.sleep(self.poll_interval)

    def complete(self, rid: str) -> RpcReply:
        """本人の完了（18:20）。作業員マスタのメールが自分と違えば、総務の手動入力で同じ時刻を入れる。"""
        r = self.rpc("api_markOvertimeDone", rid, "18:20")
        if r.status == 400 and r.error == "本人のみ操作可能です。":
            self.report.notes.append(f"{rid}: 作業員マスタのメールが自分と違うので、総務の手動入力（18:20）で完了にしました")
            r = self.rpc("api_adminManualComplete", rid, "18:20")
        return r

    def write_scenario(self) -> None:
        rep = self.report
        ctx = getattr(self, "_ctx", {"today": self.today})
        dept = ctx.get("dept") or ""
        workers = self.rpc("api_getWorkersByDept", dept)
        label = self.worker_label or (workers.result[0] if workers.ok and workers.result else "")
        if not dept or not label:
            rep.add("W-0", "書き込みの試験の準備（部署と作業員）", False, f"dept={dept} worker={label}")
            return
        code = label.split(" ", 1)[0]
        marker = f"{TEST_PREFIX}{self.run_id}"
        base = {"requestType": "overtime", "dept": dept, "workerLabel": label, "targetDate": "", "workNo1": "",
                "workNo2": "", "workNo3": "", "reason": "急ぎ作業を要するため", "reasonDetail": "",
                "hours": "1.0", "isRetroactive": False}
        try:
            r = self.rpc("api_submitRequest", dict(base, reason="", workContent=marker))
            rep.add("W-1", "申請: 理由なしは 200 {ok:false, 理由を選択してください。}（書かない）",
                    r.status == 200 and r.result == {"ok": False, "error": "理由を選択してください。"}, r.brief())

            before = self._count_worker_today(code)
            key = str(uuid.uuid4())
            a1 = self.rpc("api_submitRequest", dict(base, workContent=marker + " A"), key=key)
            a2 = self.rpc("api_submitRequest", dict(base, workContent=marker + " A"), key=key)
            rid_a = (a1.result or {}).get("requestId") if a1.ok else None
            if rid_a:
                self.created.append(rid_a)
            after = self._count_worker_today(code)
            rep.add("W-2", "申請: 同じ冪等キーで 2 回送っても 1 件（同じ requestId）",
                    bool(rid_a) and a1.result == a2.result and a1.result.get("ok") is True and
                    before is not None and after == before + 1,
                    f"a1={a1.brief()} a2={a2.brief()} 件数 {before}→{after}")
            if not rid_a:
                return
            it = self.secondary_item(rid_a)
            rep.add("W-3", "申請 A が二次承認の一覧に submitted で出る", bool(it) and it.get("status") == "submitted",
                    _short(it))

            t1 = self.rpc("api_approveRequest", rid_a)
            t2 = self.rpc("api_approveRequest", rid_a)
            rep.add("W-4", "2 つのタブで承認: 1 回目は承認、2 回目は「既に承認済みです。」",
                    t1.ok is True and (t1.result or {}).get("approvedBy") == self.me and
                    YMD_HMS_RE.match(str((t1.result or {}).get("approvedAt", ""))) is not None and
                    t2.ok is True and t2.result == {"ok": True, "message": "既に承認済みです。"},
                    f"1={t1.brief()} 2={t2.brief()}")
            it = self.secondary_item(rid_a)
            rep.add("W-5", "申請 A が approved になる", bool(it) and it.get("status") == "approved", _short(it))

            c = self.complete(rid_a)
            rep.add("W-6", "完了（17:20〜18:20）: 実績 60・休憩 0・正味 60",
                    c.ok is True and (c.result or {}).get("actualMinutes") == 60 and
                    (c.result or {}).get("breakMinutes") == 0 and (c.result or {}).get("netMinutes") == 60, c.brief())
            d = self.daily_item(rid_a)
            rep.add("W-7", "総務の日次一覧に完了が文字の日時で出る（#22）",
                    bool(d) and str(d.get("actualEndAt", "")).startswith(self.today + " 18:20") and
                    d.get("netMinutes") == 60, _short(d))

            old_pdf = ""
            if self.job_wait > 0:
                d = self.wait_for(lambda: (self.daily_item(rid_a) or {}).get("pdfFileId"))
                old_pdf = d or ""
                rep.add("W-8", "完了の後、GAS が PDF を作り pdfFileId が入る", bool(old_pdf), f"pdfFileId={old_pdf}")
            else:
                rep.add("W-8", "完了の後、GAS が PDF を作り pdfFileId が入る", None, "--job-wait 0")

            s1 = self.rpc("api_secondaryApprove", rid_a)
            s2 = self.rpc("api_secondaryApprove", rid_a)
            rep.add("W-9", "二次承認: 1 回目は approvedBy2、2 回目は「既に二次承認済みです。」",
                    s1.ok is True and (s1.result or {}).get("approvedBy2") == self.me and
                    (s1.result or {}).get("pdfRegenerated") is False and
                    s2.ok is True and s2.result == {"ok": True, "message": "既に二次承認済みです。"},
                    f"1={s1.brief()} 2={s2.brief()}")
            rep.add("W-10", "二次承認した申請は二次承認の一覧から消える", self.secondary_item(rid_a) is None, rid_a)
            if self.job_wait > 0 and old_pdf:
                new_pdf = self.wait_for(lambda: (lambda v: v if v and v != old_pdf else None)(
                    (self.daily_item(rid_a) or {}).get("pdfFileId")))
                rep.add("W-11", "二次承認の後、PDF が作り直される（別の pdfFileId）", bool(new_pdf),
                        f"old={old_pdf} new={new_pdf}")
            else:
                rep.add("W-11", "二次承認の後、PDF が作り直される（別の pdfFileId）", None,
                        "--job-wait 0 か、先の PDF が無い")

            b = self.rpc("api_submitRequest", dict(base, workContent=marker + " B"))
            rid_b = (b.result or {}).get("requestId") if b.ok else None
            if not rid_b:
                rep.add("W-12", "申請 B を作る", False, b.brief())
                return
            self.created.append(rid_b)
            e1 = self.rpc("api_editRequest", rid_b, {"hours": "2.0"})
            e2 = self.rpc("api_editRequest", rid_b, {"hours": "9.9"})
            rep.add("W-12", "編集: 2.0 時間で {ok, newStatus:submitted, newMinutes:120}、9.9 は止まる",
                    e1.ok is True and e1.result == {"ok": True, "newStatus": "submitted", "newMinutes": 120} and
                    e2.status == 400 and e2.error == "予定時間が不正です: 9.9", f"1={e1.brief()} 2={e2.brief()}")
            c = self.complete(rid_b)
            e3 = self.rpc("api_editRequest", rid_b, {"hours": "1.0"})
            rep.add("W-13", "完了した申請の編集は「作業完了済みの申請は編集できません。」",
                    c.ok is True and e3.status == 400 and e3.error == "作業完了済みの申請は編集できません。",
                    f"完了={c.brief()} 編集={e3.brief()}")
            t1 = self.rpc("api_approveRequest", rid_b)
            t2 = self.rpc("api_approveRequest", rid_b)
            rep.add("W-14", "完了後の承認を 2 つのタブで: 2 回目は「既に承認済みです。」",
                    t1.ok is True and (t1.result or {}).get("requestId") == rid_b and
                    t2.ok is True and t2.result == {"ok": True, "message": "既に承認済みです。"},
                    f"1={t1.brief()} 2={t2.brief()}")

            d_a = self.rpc("api_deleteRequest", rid_a)
            d_b = self.rpc("api_deleteRequest", rid_b)
            d_b2 = self.rpc("api_deleteRequest", rid_b)
            for rid, rr in ((rid_a, d_a), (rid_b, d_b)):
                if rr.ok:
                    self.deleted.add(rid)
            rep.add("W-15", "削除: {ok:true}、2 回目は「既にキャンセル済みです。」",
                    d_a.result == {"ok": True} and d_b.result == {"ok": True} and d_b2.status == 400 and
                    d_b2.error == "既にキャンセル済みです。", f"A={d_a.brief()} B={d_b.brief()} B2={d_b2.brief()}")
            ap = self.rpc("api_approveRequest", rid_a)
            rep.add("W-16", "削除した申請は一覧から消え、承認は「キャンセル済みです。」",
                    self.secondary_item(rid_a) is None and self.secondary_item(rid_b) is None and
                    ap.status == 400 and ap.error == "キャンセル済みです。", ap.brief())

            self.check_jobs(rid_a, rid_b)
            self.write_ops()
        finally:
            self.cleanup()

    def _count_worker_today(self, code: str) -> int | None:
        r = self.rpc("api_adminDailyDetail", self.today, "ALL")
        if not (r.ok and isinstance(r.result, dict)):
            return None
        return sum(1 for it in r.result.get("items") or [] if it.get("workerCode") == code)

    def check_jobs(self, rid_a: str, rid_b: str) -> None:
        rep = self.report
        expected_a = [{K_AFTER_COMPLETE}, {K_REGENERATE, K_LEDGER_UPSERT}, {K_AFTER_DELETE}]
        expected_b = [{K_AFTER_COMPLETE}, {K_AFTER_APPROVE}, {K_AFTER_DELETE}]
        title_order = "GAS の仕事が決まった順に積まれた（A: 完了→作り直し→削除、B: 完了→承認→削除。承認は 1 回）"
        title_done = "GAS の仕事がすべて終わり、失敗が無い"
        if self.probe is None:
            rep.add("W-17", title_order, None, "DB を見られません（--mint-session か --db-probe で動かす）")
            rep.add("W-18", title_done, None, "DB を見られません")
            return
        jobs_a, jobs_b = self.probe.jobs_for(rid_a), self.probe.jobs_for(rid_b)
        rep.problems("W-17", title_order, check_job_sequence(jobs_a, expected_a) + check_job_sequence(jobs_b, expected_b))
        if self.job_wait <= 0:
            rep.add("W-18", title_done, None, "--job-wait 0")
            return

        def finished() -> list[dict] | None:
            js = self.probe.jobs_for(rid_a) + self.probe.jobs_for(rid_b)
            return js if js and all(j["state"] in JOB_FINISHED_OK + ("failed",) for j in js) else None

        js = self.wait_for(finished) or (self.probe.jobs_for(rid_a) + self.probe.jobs_for(rid_b))
        bad = [j for j in js if j["state"] not in JOB_FINISHED_OK]
        rep.add("W-18", title_done, not bad, _short(bad))

    def write_ops(self) -> None:
        rep = self.report
        r = self.rpc("ops_syncSheetConfigNow")
        rep.add("F-ops_syncSheetConfigNow", "ops_syncSheetConfigNow() が {ok, …}",
                r.ok is True and isinstance(r.result, dict) and r.result.get("ok") is True, r.brief())
        r = self.rpc("ops_runMasterSync")
        rep.add("F-ops_runMasterSync", "ops_runMasterSync() が {ok, tabs}",
                r.ok is True and isinstance(r.result, dict) and r.result.get("ok") is True and "tabs" in r.result,
                r.brief())
        r = self.rpc("ops_recalcBreaks")
        rep.add("F-ops_recalcBreaks", "ops_recalcBreaks() が {ok:true, changed:数}",
                r.ok is True and isinstance(r.result, dict) and r.result.get("ok") is True and
                isinstance(r.result.get("changed"), int), r.brief())

    def cleanup(self) -> None:
        """途中で止まっても、作った試験の申請を削除（canceled）しておく。"""
        for rid in self.created:
            if rid in self.deleted:
                continue
            r = self.rpc("api_deleteRequest", rid)
            self.report.notes.append(f"後片付け: {rid} を削除 → {r.brief()}")
            if r.ok or r.error == "既にキャンセル済みです。":
                self.deleted.add(rid)
        if self.created:
            self.report.notes.append("試験の申請 ID: " + ", ".join(self.created))


class WriteRefused(RuntimeError):
    pass


def _json_or_none(text: str) -> Any:
    try:
        return json.loads(text)
    except ValueError:
        return None


def _args_text(args: tuple) -> str:
    return "(" + ", ".join(json.dumps(a, ensure_ascii=False) for a in args) + ")"


# ===== 入口 =====


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="残業・休日出勤申請 VPS 版の受け入れテスト（HTTP）")
    p.add_argument("--base-url", default="http://localhost:8000", help="api の URL（コンテナの中は http://api:8000）")
    p.add_argument("--web-url", default=None, help="nginx の URL（ヘッダーを見る。外からなら --base-url と同じ）")
    p.add_argument("--origin", default=os.environ.get("PUBLIC_ORIGIN", DEFAULT_ORIGIN))
    p.add_argument("--cookie", default=os.environ.get("ACCEPTANCE_COOKIE"), help="セッションのクッキーの値")
    p.add_argument("--mint-session", action="store_true", help="staging でセッションを作り、最後に消す（コンテナの中）")
    p.add_argument("--email", default=None, help="--mint-session のメール（BOOTSTRAP_ADMINS の中）")
    p.add_argument("--db-probe", action="store_true", help="DB の jobs を見る（コンテナの中）")
    p.add_argument("--sebastian-token", default=os.environ.get("SEBASTIAN_API_TOKEN"))
    p.add_argument("--write", action="store_true", help="書き込みの試験（staging だけ）")
    p.add_argument("--expect-env", choices=["staging", "production"], default=None)
    p.add_argument("--job-wait", type=float, default=240.0, help="GAS の仕事・PDF を待つ秒（0 で待たない）")
    p.add_argument("--dept", default=None)
    p.add_argument("--worker-label", default=None, help="申請に使う作業員（「<コード> <氏名>」）")
    p.add_argument("--json-out", default=None, help="結果の JSON の書き出し先")
    p.add_argument("--timeout", type=float, default=120.0)
    return p


def main(argv: list[str] | None = None, *, api: Transport | None = None, web: Transport | None = None) -> int:
    args = build_parser().parse_args(argv)
    api = api or UrllibTransport(args.base_url, args.timeout)
    if web is None and args.web_url:
        web = UrllibTransport(args.web_url, args.timeout)
    minted = None
    probe = None
    cookie = args.cookie
    db = None
    if args.mint_session or args.db_probe:  # コンテナの中（app と DB を使える）
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        probe = DbJobProbe.from_app()
    try:
        if args.mint_session:
            import staging_session
            from app.config import get_settings
            from app.db import get_sessionmaker

            db = get_sessionmaker()()
            try:
                minted = staging_session.mint(db, get_settings(), args.email, hours=2)
            except staging_session.Refused as exc:
                print("REFUSED: " + str(exc), file=sys.stderr)
                return 2
            cookie = minted.cookie_value
        runner = Runner(api, web=web, origin=args.origin, cookie=cookie, sebastian_token=args.sebastian_token,
                        probe=probe, write=args.write, expect_env=args.expect_env, job_wait=args.job_wait,
                        dept=args.dept, worker_label=args.worker_label)
        try:
            report = runner.run()
        except WriteRefused as exc:
            print("REFUSED: " + str(exc), file=sys.stderr)
            return 2
        print(report.table())
        for note in report.notes:
            print(note)
        if args.json_out:
            Path(args.json_out).write_text(json.dumps(report.to_json(), ensure_ascii=False, indent=1), "utf-8")
        return report.exit_code()
    finally:
        if minted is not None and db is not None:
            import staging_session

            n = staging_session.revoke(db, minted.token_hash)
            print(f"session removed: {n == 1}")
        if db is not None:
            db.close()


if __name__ == "__main__":
    sys.exit(main())
