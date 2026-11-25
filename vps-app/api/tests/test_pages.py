"""画面の入口 `GET /`（app/pages.py）の試験（仕様 3-2・3-3・12-2 の U）。

ログイン・権限・設定は偽物に差し替え、pages の振り分けとテンプレートの差し込みだけを確かめる。
"""
from __future__ import annotations

import json
import re
from html import unescape
from types import SimpleNamespace
from urllib.parse import parse_qs, urlsplit

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from app import auth, pages
from app.config import get_settings
from app.db import get_db

ADMIN = "imaizumi@lineworks-local.info"
APPROVER = "shonin@lineworks-local.info"
STAFF = "genba@lineworks-local.info"
PORTAL = "https://script.google.com/a/macros/lineworks-local.info/s/PORTAL_TEST/exec"
ORIGIN = "https://overtime.lineworks-local.net"


@pytest.fixture
def make_client(monkeypatch):
    def _make(email: str | None):
        app = FastAPI()
        app.include_router(pages.router)
        user = None if email is None else SimpleNamespace(email=email, name=email.split("@")[0])
        app.dependency_overrides[auth.get_optional_user] = lambda: user
        app.dependency_overrides[get_db] = lambda: None
        app.dependency_overrides[get_settings] = lambda: SimpleNamespace(public_origin=ORIGIN, portal_url=PORTAL)
        monkeypatch.setattr(pages.permissions, "is_admin", lambda db, e: e == ADMIN)
        monkeypatch.setattr(pages.permissions, "is_any_dept_approver", lambda db, e: e == APPROVER)
        return TestClient(app, base_url="https://testserver")

    return _make


def title_of(html: str) -> str:
    m = re.search(r"<title>(.*?)</title>", html)
    return m.group(1) if m else ""


def data_of(html: str) -> dict:
    m = re.search(r'<script type="application/json" id="gas-template-data">(.*?)</script>', html, re.S)
    assert m, "テンプレートの値の塊が無い"
    return json.loads(m.group(1))


def page_js(html: str) -> str:
    srcs = re.findall(r'<script src="/static/js/([^"]+)\.js"></script>', html)
    assert srcs[0] == "gas-shim"
    return srcs[1]


# ---- ログイン ----

@pytest.mark.parametrize("path", ["/", "/?page=admin&date=2026-11-25", "/?page=approve2&date=2026-11-25&x=1"])
def test_not_logged_in_redirects_to_login_with_next(make_client, path):
    r = make_client(None).get(path, follow_redirects=False)
    assert r.status_code == 302
    loc = urlsplit(r.headers["location"])
    assert loc.path == "/auth/login"
    assert parse_qs(loc.query)["next"] == [path]
    assert r.headers["cache-control"] == "no-store"


# ---- 振り分け ----

@pytest.mark.parametrize("query,expected", [
    ("", "top"), ("?page=", "top"), ("?page=top", "top"), ("?page=unknown", "top"),
    ("?page=staffSummary", "staffSummary"), ("?page=approver", "approver"),
    ("?page=admin", "admin"), ("?page=approve2", "approve2"),
    ("?page=no_auth", "top"), ("?page=../admin", "top"),
    ("?page=admin&page=top", "admin"),  # GAS の e.parameter と同じく最初の値
])
def test_page_routing_for_admin(make_client, query, expected):
    r = make_client(ADMIN).get("/" + query)
    assert r.status_code == 200
    assert page_js(r.text) == expected
    assert data_of(r.text)["PAGE"] == expected
    assert title_of(r.text) == "残業・休日出勤申請"
    assert r.headers["cache-control"] == "no-store"


def test_redirect_page_goes_to_top(make_client):
    r = make_client(STAFF).get("/?page=redirect&type=overtime&dept=x", follow_redirects=False)
    assert r.status_code == 302
    assert r.headers["location"] == "/?page=top"


@pytest.mark.parametrize("page", ["top", "staffSummary"])
def test_open_pages_need_no_role(make_client, page):
    r = make_client(STAFF).get("/?page=" + page)
    assert r.status_code == 200
    assert page_js(r.text) == page


@pytest.mark.parametrize("page", ["admin", "approve2"])
def test_admin_pages_reject_non_admin(make_client, page):
    for email in (STAFF, APPROVER):
        r = make_client(email).get("/?page=" + page)
        assert r.status_code == 200
        assert title_of(r.text) == "権限エラー"
        msg = unescape(re.search(r'<div class="msg">(.*?)</div>', r.text, re.S).group(1))
        assert msg == (
            "総務部（管理者）権限がありません。\n"
            "ApproverMap シートに role=admin で登録されているか、\n"
            "Settings シートの ADMIN_EMAILS にメールアドレスが含まれているか確認してください。\n"
            "（現在のアカウント: " + email + "）"
        )
        assert "デバッグ情報" not in r.text
        assert "gas-shim.js" not in r.text


def test_approver_page(make_client):
    assert page_js(make_client(APPROVER).get("/?page=approver").text) == "approver"
    assert page_js(make_client(ADMIN).get("/?page=approver").text) == "approver"
    r = make_client(STAFF).get("/?page=approver")
    assert title_of(r.text) == "権限エラー"
    msg = unescape(re.search(r'<div class="msg">(.*?)</div>', r.text, re.S).group(1))
    assert msg == (
        "承認者権限がありません。\n"
        "DeptApprovers または ApproverMap に登録されているか確認してください。\n"
        "（現在のアカウント: " + STAFF + "）"
    )


def test_no_auth_links_and_escaping(make_client):
    r = make_client('a"<b>@lineworks-local.info').get("/?page=admin")
    assert "<b>" not in r.text  # メールは自動エスケープ
    assert f'href="{ORIGIN}/?page=top"' in r.text
    assert f'href="{PORTAL}"' in r.text


# ---- テンプレートの値 ----

@pytest.mark.parametrize("value,expected", [
    ("2026-11-25", "2026-11-25"),
    ("2024-02-29", "2024-02-29"),
    ("2026-02-30", ""),
    ("2026-13-01", ""),
    ("2026-1-05", ""),
    ("2026-11-25 ", ""),
    ("2026-11-25\n", ""),
    ("２０２６-11-25", ""),
    ("", ""),
    ("</script><script>alert(1)</script>", ""),
    ("2026-11-25';alert(1)//", ""),
])
def test_date_param_is_validated(make_client, value, expected):
    r = make_client(ADMIN).get("/", params={"page": "admin", "date": value})
    assert data_of(r.text)["DATE_PARAM"] == expected
    assert "alert(1)" not in r.text


def test_json_block_cannot_close_script():
    text = pages.json_for_script({"x": "</script><script>alert(1)</script>&"})
    assert "<" not in text and ">" not in text and "&" not in text
    assert json.loads(text) == {"x": "</script><script>alert(1)</script>&"}


def test_template_values(make_client, monkeypatch):
    monkeypatch.setattr(pages.constants, "REASONS", ["急ぎ作業を要するため", "その他（→補足理由必須）"])
    monkeypatch.setattr(pages.constants, "OT_HOURS", ["0.5", "1.0"])
    monkeypatch.setattr(pages.constants, "HD_HOURS", ["半日", "1日"])
    r = make_client(STAFF).get("/?page=top")
    d = data_of(r.text)
    assert d["REASONS"] == ["急ぎ作業を要するため", "その他（→補足理由必須）"]
    assert d["OT_HOURS"] == ["0.5", "1.0"]
    assert d["HD_HOURS"] == ["半日", "1日"]
    assert d["DATE_PARAM"] == ""
    # ヘッダーのリンクは GAS と同じ `<APP_URL>?page=...`
    for p in ("staffSummary", "approver", "admin"):
        assert f'href="{ORIGIN}/?page={p}"' in r.text
    assert f'href="{PORTAL}"' in r.text
    assert "{{" not in r.text and "<?" not in r.text
