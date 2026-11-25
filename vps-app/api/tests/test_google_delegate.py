from urllib.parse import parse_qs, urlparse

import pytest

from app import auth, google_delegate
from app.google_delegate import DelegateNotAuthorized, get_delegate_credentials, get_delegate_status
from app.models import GoogleCredential

from test_auth import KEY, make_id_token

ADMIN = "owner@lineworks-local.info"
ALL_SCOPES = " ".join(google_delegate.DELEGATE_SCOPES)


@pytest.fixture
def fake_google(monkeypatch):
    monkeypatch.setattr(auth, "resolve_signing_key", lambda token: KEY.public_key())
    holder = {"email": ADMIN, "refresh_token": "1//refresh-token", "scope": ALL_SCOPES}

    def _exchange(code, redirect_uri, verifier, settings):
        holder["redirect_uri"] = redirect_uri
        tokens = {"id_token": make_id_token(holder["nonce"], email=holder["email"]), "scope": holder["scope"]}
        if holder["refresh_token"]:
            tokens["refresh_token"] = holder["refresh_token"]
        return tokens

    monkeypatch.setattr(auth, "exchange_code", _exchange)
    return holder


def start(client):
    resp = client.get("/auth/google-delegate/start")
    assert resp.status_code == 302, resp.text
    return {k: v[0] for k, v in parse_qs(urlparse(resp.headers["location"]).query).items()}


def test_not_authorized_status(db, settings):
    assert get_delegate_status(db) == {"authorized": False, "accountEmail": "", "grantedBy": "", "grantedAt": "", "scopes": []}
    with pytest.raises(DelegateNotAuthorized):
        get_delegate_credentials(db, settings)


def test_start_requires_admin(client, login_as):
    assert client.get("/auth/google-delegate/start").status_code == 302  # 未ログインはログインへ
    login_as("someone@lineworks-local.info")
    assert client.get("/auth/google-delegate/start").status_code == 403


def test_start_requests_offline_consent_for_delegate(client, login_as):
    login_as(ADMIN)
    q = start(client)
    assert q["access_type"] == "offline" and q["prompt"] == "consent"
    assert q["login_hint"] == ADMIN
    assert q["redirect_uri"] == "https://overtime.lineworks-local.net/auth/google-delegate/callback"
    # 実装仕様 1-1 のスコープだけ（gmail は要らない。後から足さない＝罠 7）
    assert sorted(q["scope"].split()) == sorted(["openid", "email", "https://www.googleapis.com/auth/spreadsheets",
                                                 "https://www.googleapis.com/auth/drive"])


def test_callback_stores_encrypted_token(client, db, login_as, fake_google, settings):
    login_as(ADMIN)
    q = start(client)
    fake_google["nonce"] = q["nonce"]
    resp = client.get("/auth/google-delegate/callback", params={"code": "c", "state": q["state"]})
    assert resp.status_code == 200, resp.text
    row = db.get(GoogleCredential, "delegate")
    assert row.account_email == ADMIN and row.granted_by_email == ADMIN
    assert "1//refresh-token" not in row.refresh_token_encrypted
    status = get_delegate_status(db)
    assert status["authorized"] is True and status["accountEmail"] == ADMIN and len(status["grantedAt"]) == 19
    creds = get_delegate_credentials(db, settings)
    assert creds.refresh_token == "1//refresh-token"
    assert creds.client_id == settings.google_oauth_client_id


@pytest.mark.parametrize(
    "change",
    [
        {"email": "someone@lineworks-local.info"},  # 委任は DELEGATE_EMAIL だけ
        {"refresh_token": ""},
        {"scope": "openid email https://www.googleapis.com/auth/drive"},  # スプレッドシートを外された
    ],
)
def test_callback_rejects_bad_grants(client, db, login_as, fake_google, change):
    login_as(ADMIN)
    q = start(client)
    fake_google.update(change)
    fake_google["nonce"] = q["nonce"]
    resp = client.get("/auth/google-delegate/callback", params={"code": "c", "state": q["state"]})
    assert resp.status_code == 403
    assert db.get(GoogleCredential, "delegate") is None


def test_wrong_fernet_key_is_reported(db, settings):
    from conftest import make_settings
    from cryptography.fernet import Fernet

    google_delegate.save_delegate_token(
        db, account_email=ADMIN, refresh_token="t", scopes=ALL_SCOPES, granted_by=ADMIN, settings=settings
    )
    db.commit()
    other = make_settings(fernet_key=Fernet.generate_key().decode())
    with pytest.raises(DelegateNotAuthorized):
        get_delegate_credentials(db, other)


def test_delegate_cookie_name_and_scopes():
    assert google_delegate.DELEGATE_FLOW_COOKIE == "__Host-overtime_delegate"
    assert google_delegate.REQUIRED_API_SCOPES == ("https://www.googleapis.com/auth/spreadsheets",
                                                   "https://www.googleapis.com/auth/drive")


def test_credentials_use_granted_scopes(db, settings):
    google_delegate.save_delegate_token(db, account_email=ADMIN, refresh_token="t",
                                        scopes="https://www.googleapis.com/auth/spreadsheets openid",
                                        granted_by=ADMIN, settings=settings)
    db.commit()
    creds = get_delegate_credentials(db, settings)
    # 許可されたスコープで更新する（罠 7: 未許可のスコープを要求すると invalid_scope で全部止まる）
    assert list(creds.scopes) == ["https://www.googleapis.com/auth/spreadsheets", "openid"]


def test_get_access_token_refreshes_with_stored_token(monkeypatch, session_factory, settings):
    from app import db as app_db

    with session_factory() as s:
        google_delegate.save_delegate_token(s, account_email=ADMIN, refresh_token="1//r", scopes=ALL_SCOPES,
                                            granted_by=ADMIN, settings=settings)
        s.commit()
    monkeypatch.setattr(app_db, "get_sessionmaker", lambda: session_factory)
    seen = {}

    def fake_refresh(self, request):
        seen["refresh_token"] = self.refresh_token
        self.token = "access-token-1"

    from google.oauth2.credentials import Credentials

    monkeypatch.setattr(Credentials, "refresh", fake_refresh)
    assert google_delegate.get_access_token(settings=settings) == "access-token-1"
    assert seen["refresh_token"] == "1//r"
