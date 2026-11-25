import time
from urllib.parse import parse_qs, urlparse

import jwt
import pytest
from cryptography.hazmat.primitives.asymmetric import rsa

from app import auth, rpc
from app.models import UserSession
from conftest import make_admin

CLIENT_ID = "test-client-id.apps.googleusercontent.com"
KEY = rsa.generate_private_key(public_exponent=65537, key_size=2048)
OTHER_KEY = rsa.generate_private_key(public_exponent=65537, key_size=2048)


def make_id_token(nonce, key=KEY, **overrides):
    now = int(time.time())
    claims = {
        "iss": "https://accounts.google.com",
        "aud": CLIENT_ID,
        "azp": CLIENT_ID,
        "sub": "1234567890",
        "iat": now,
        "exp": now + 600,
        "nonce": nonce,
        "email": "Tanaka@lineworks-local.info",
        "email_verified": True,
        "hd": "lineworks-local.info",
    }
    claims.update(overrides)
    claims = {k: v for k, v in claims.items() if v is not None}
    return jwt.encode(claims, key, algorithm="RS256", headers={"kid": "k1"})


@pytest.fixture(autouse=True)
def fake_google(monkeypatch):
    monkeypatch.setattr(auth, "resolve_signing_key", lambda token: KEY.public_key())
    holder = {"claims": {}, "key": KEY, "calls": []}

    def _exchange(code, redirect_uri, verifier, settings):
        holder["calls"].append((code, redirect_uri, verifier))
        claims = {"nonce": holder["nonce"], **holder["claims"]}
        return {"id_token": make_id_token(claims.pop("nonce"), key=holder["key"], **claims)}

    monkeypatch.setattr(auth, "exchange_code", _exchange)
    return holder


def start_login(client, next_path=None):
    resp = client.get("/auth/login", params={"next": next_path} if next_path else None)
    assert resp.status_code == 302
    q = parse_qs(urlparse(resp.headers["location"]).query)
    return resp, {k: v[0] for k, v in q.items()}


def finish_login(client, fake_google, params, state=None):
    fake_google["nonce"] = params["nonce"]
    return client.get("/auth/callback", params={"code": "abc", "state": state or params["state"]})


def test_login_redirect_has_state_nonce_pkce_and_hd(client):
    resp, q = start_login(client)
    assert resp.headers["location"].startswith(auth.GOOGLE_AUTH_URL)
    assert q["client_id"] == CLIENT_ID
    assert q["redirect_uri"] == "https://overtime.lineworks-local.net/auth/callback"
    assert q["hd"] == "lineworks-local.info"
    assert q["code_challenge_method"] == "S256"
    assert len(q["state"]) > 30 and len(q["nonce"]) > 30
    assert len(q["code_challenge"]) == 43
    assert "prompt" not in q  # select_account・consent を付けない（gap-2 S2）
    assert q["scope"] == "openid email profile"
    cookie = resp.headers["set-cookie"]
    assert auth.LOGIN_FLOW_COOKIE in cookie and "HttpOnly" in cookie and "Secure" in cookie
    assert "samesite=lax" in cookie.lower()


def test_successful_login_returns_to_request_detail(client, db, fake_google):
    _, q = start_login(client, "/?page=approve2&date=2026-11-25")
    resp = finish_login(client, fake_google, q)
    assert resp.status_code == 302
    assert resp.headers["location"] == "/?page=approve2&date=2026-11-25"
    set_cookie = ";".join(resp.headers.get_list("set-cookie"))
    assert auth.SESSION_COOKIE in set_cookie
    row = db.query(UserSession).one()
    assert row.email == "tanaka@lineworks-local.info" and row.name == "tanaka"
    # PKCE の verifier がトークン交換に渡っている
    assert fake_google["calls"][0][2]


def test_state_mismatch_is_rejected(client, fake_google, db):
    _, q = start_login(client)
    resp = finish_login(client, fake_google, q, state="wrong-state")
    assert resp.status_code == 403
    assert db.query(UserSession).count() == 0


def test_callback_without_flow_cookie_is_rejected(client, fake_google, db):
    fake_google["nonce"] = "n"
    resp = client.get("/auth/callback", params={"code": "abc", "state": "s"})
    assert resp.status_code == 403
    assert db.query(UserSession).count() == 0


@pytest.mark.parametrize(
    "claims",
    [
        {"nonce": "other-nonce"},
        {"aud": "someone-else"},
        {"azp": "someone-else"},
        {"iss": "https://evil.example.com"},
        {"exp": int(time.time()) - 3600, "iat": int(time.time()) - 7200},
        {"email_verified": False},
        {"email_verified": None},
        {"hd": "gmail.com"},
        {"hd": None},
        {"email": "tanaka@gmail.com"},
        {"sub": None},
    ],
)
def test_bad_id_token_claims_are_rejected(client, fake_google, db, claims):
    fake_google["claims"] = claims
    _, q = start_login(client)
    resp = finish_login(client, fake_google, q)
    assert resp.status_code == 403
    assert db.query(UserSession).count() == 0


def test_bad_signature_is_rejected(client, fake_google, db):
    fake_google["key"] = OTHER_KEY
    _, q = start_login(client)
    assert finish_login(client, fake_google, q).status_code == 403
    assert db.query(UserSession).count() == 0


def test_google_error_param_is_rejected(client, db):
    _, q = start_login(client)
    resp = client.get("/auth/callback", params={"error": "access_denied", "state": q["state"]})
    assert resp.status_code == 403


@pytest.mark.parametrize(
    "value,expected",
    [
        ("/?page=approve2&date=2026-11-25", "/?page=approve2&date=2026-11-25"),
        ("//evil.example.com/", "/"),
        ("https://evil.example.com/", "/"),
        ("/\\evil.example.com", "/"),
        ("/auth/login", "/"),
        ("", "/"),
        (None, "/"),
        ("/a\nb", "/"),
    ],
)
def test_safe_next(value, expected):
    assert auth.safe_next(value) == expected


def test_rpc_requires_login_and_logout_revokes(client, db, login_as):
    resp = client.post("/api/rpc/api_getDeptList", json=[])
    assert resp.status_code == 401
    assert resp.json() == {
        "ok": False,
        "error": "ログインが必要です。",
        "login": "/auth/login",
    }
    login_as("Someone@lineworks-local.info")
    assert client.post("/api/rpc/api_getDeptList", json=[]).status_code == 501  # ログイン済み・未実装
    saved = dict(client.cookies)
    assert client.post("/auth/logout").status_code == 200
    assert db.query(UserSession).count() == 0
    for k, v in saved.items():  # 古いクッキーを持ち続けても使えない
        client.cookies.set(k, v)
    assert client.post("/api/rpc/api_getDeptList", json=[]).status_code == 401


def test_logout_is_post_only(client, db, login_as):
    login_as("someone@lineworks-local.info")
    assert client.get("/auth/logout").status_code == 405
    assert db.query(UserSession).count() == 1


def test_session_lasts_90_days_even_without_use(client, db, login_as, settings, monkeypatch):
    from datetime import timedelta

    from app import timeutil

    rpc.register("api_getDeptList", lambda ctx: "ok")
    login_as("someone@lineworks-local.info")
    row = db.query(UserSession).one()
    # ログイン直後の期限は 90 日後（9-1 #24。無操作でも 90 日は切らない）
    assert abs((row.expires_at - row.created_at) - timedelta(days=90)) < timedelta(seconds=5)
    assert settings.session_idle_hours == 2160 and settings.session_max_age_days == 90
    base = timeutil.now_utc()

    # 89 日使わなくても、まだ使える
    monkeypatch.setattr(timeutil, "now_utc", lambda: base + timedelta(days=89))
    assert client.post("/api/rpc/api_getDeptList", json=[]).status_code == 200

    # ログインから 90 日を過ぎたら切れる（使い続けても延びない）
    monkeypatch.setattr(timeutil, "now_utc", lambda: base + timedelta(days=90, minutes=1))
    assert client.post("/api/rpc/api_getDeptList", json=[]).status_code == 401


def test_sliding_renewal_with_shorter_idle_setting(client, db, settings, monkeypatch):
    """SESSION_IDLE_HOURS を短くした場合: 使うたびに延び、使わなければ切れる（ログインから最大 90 日）。"""
    from datetime import timedelta

    from app import timeutil
    from conftest import login, make_settings

    short = make_settings(session_idle_hours=12)
    client.app.dependency_overrides[auth.get_settings] = lambda: short
    rpc.register("api_getDeptList", lambda ctx: "ok")
    login(client, db, "someone@lineworks-local.info", short)
    base = timeutil.now_utc()
    later = base + timedelta(hours=11)
    monkeypatch.setattr(timeutil, "now_utc", lambda: later)
    assert client.post("/api/rpc/api_getDeptList", json=[]).status_code == 200
    db.expire_all()
    assert db.query(UserSession).one().expires_at >= later + timedelta(hours=12) - timedelta(seconds=5)
    monkeypatch.setattr(timeutil, "now_utc", lambda: later + timedelta(hours=13))
    assert client.post("/api/rpc/api_getDeptList", json=[]).status_code == 401


def test_session_never_outlives_max_age(settings):
    from datetime import datetime, timedelta, timezone

    created = datetime(2026, 11, 1, tzinfo=timezone.utc)
    near_end = created + timedelta(days=settings.session_max_age_days) - timedelta(hours=1)
    assert auth._session_expiry(created, near_end, settings) == created + timedelta(days=settings.session_max_age_days)


def test_revoke_sessions_removes_all_sessions_of_one_email(db, settings):
    from fastapi.responses import Response

    for email in ("a@lineworks-local.info", "a@lineworks-local.info", "b@lineworks-local.info"):
        auth.create_session(db, Response(), email, settings)
    assert auth.revoke_sessions(db, " A@lineworks-local.info ") == 2
    db.commit()
    assert [s.email for s in db.query(UserSession).all()] == ["b@lineworks-local.info"]
    assert auth.revoke_sessions(db, "") == 0


def test_tampered_session_cookie_is_rejected(client, login_as):
    login_as("someone@lineworks-local.info")
    value = client.cookies.get(auth.SESSION_COOKIE)
    client.cookies.set(auth.SESSION_COOKIE, value[:-2] + ("AA" if not value.endswith("AA") else "BB"))
    assert client.post("/api/rpc/api_getDeptList", json=[]).status_code == 401


def test_bootstrap_admin_is_the_owner_only(db, settings):
    assert auth.is_bootstrap_admin("owner@lineworks-local.info", settings)
    assert auth.is_bootstrap_admin(" OWNER@lineworks-local.info ", settings)
    assert not auth.is_bootstrap_admin("someone@lineworks-local.info", settings)
    assert not auth.is_bootstrap_admin("", settings)
    # Settings の ADMIN_EMAILS の管理者は業務の管理者（permissions.is_admin）で、委任の開始はできない
    make_admin(db, "boss@lineworks-local.info")
    assert not auth.is_bootstrap_admin("boss@lineworks-local.info", settings)


def test_auth_me(client, login_as):
    resp = client.get("/auth/me")
    assert resp.status_code == 401
    assert resp.json() == {"ok": False, "error": "ログインが必要です。", "login": "/auth/login"}
    login_as("Someone@lineworks-local.info")
    resp = client.get("/auth/me")
    assert resp.status_code == 200
    assert resp.json() == {"ok": True, "email": "someone@lineworks-local.info", "name": "someone"}
    assert resp.headers["cache-control"] == "no-store"


def test_hd_is_sent_and_checked(client, fake_google, db):
    _, q = start_login(client)
    assert q["hd"] == "lineworks-local.info"
    fake_google["claims"] = {"hd": "lineworks.co.jp", "email": "tanaka@lineworks.co.jp"}
    assert finish_login(client, fake_google, q).status_code == 403
    assert db.query(UserSession).count() == 0


def test_cookie_names_are_overtime(client, db, fake_google):
    assert auth.SESSION_COOKIE == "__Host-overtime_session"
    assert auth.LOGIN_FLOW_COOKIE == "__Host-overtime_login"
    _, q = start_login(client)
    resp = finish_login(client, fake_google, q)
    cookie = [c for c in resp.headers.get_list("set-cookie") if c.startswith(auth.SESSION_COOKIE)][0]
    assert "Max-Age=7776000" in cookie  # 90 日
    assert "HttpOnly" in cookie and "Secure" in cookie and "Path=/" in cookie


def test_user_name_is_local_part():
    u = auth.user_from_email(" Tanaka.Taro@lineworks-local.info ")
    assert u.email == "tanaka.taro@lineworks-local.info" and u.name == "tanaka.taro"
