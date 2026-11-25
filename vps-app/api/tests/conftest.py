"""試験の土台。SQLite（メモリー）に models の表を作り、FastAPI を https://testserver で動かす。

使い方（担当 U・W・D・M）:
- `db` フィクスチャ: 空の表がそろった Session。`session_factory` も使える（worker・jobs の試験）。
- `client` フィクスチャ: TestClient（業務関数は登録しない。rpc.register で差し込む）。
- `login_as(email)`: そのメールでログインした状態のクッキーを client に付ける。
- `settings` フィクスチャ: 試験用の Settings（APP_ENV=staging）。`make_settings(**上書き)` で別の値。
- `make_admin(db, email)`: Settings の ADMIN_EMAILS にメールを入れる（管理者にする）。
- `freeze_now(monkeypatch, dt)`: timeutil.now_utc を固定する。
外の API（Google・GAS）は呼ばない（偽物を入れる）。PostgreSQL だけの機能（行ロック・skip_locked・advisory lock）は
SQLite では何もしないので、同じコードがそのまま動く。TEST_DATABASE_URL を渡すと PostgreSQL で動く
（DB 名に test を含まないものは拒む。罠 51: `python -m pytest` で呼ぶ）。
"""
from __future__ import annotations

import os

from cryptography.fernet import Fernet

# app を import する前に環境変数をそろえる（config は起動時に必須値を確かめる）。
# 本番の ID は config の定数と同じでなければ起動しないので、ここでは config から読んだ値を入れる。
os.environ.setdefault("APP_ENV", "staging")
from app import config as _config  # noqa: E402

TEST_ENV = {
    "APP_ENV": "staging",
    "DATABASE_URL": "sqlite+pysqlite:///:memory:",
    "PUBLIC_ORIGIN": "https://overtime.lineworks-local.net",
    "GOOGLE_OAUTH_CLIENT_ID": "test-client-id.apps.googleusercontent.com",
    "GOOGLE_OAUTH_CLIENT_SECRET": "test-client-secret",
    "SESSION_SECRET": "test-session-secret-0123456789abcdef0123456789",
    "SESSION_MAX_AGE_DAYS": "90",
    "SESSION_IDLE_HOURS": "2160",
    "ALLOWED_DOMAIN": "lineworks-local.info",
    "BOOTSTRAP_ADMINS": "owner@lineworks-local.info",
    "FERNET_KEY": Fernet.generate_key().decode(),
    "DELEGATE_EMAIL": "owner@lineworks-local.info",
    "PROD_SPREADSHEET_ID": _config.PROD_SPREADSHEET_ID_FIXED,
    "STAGING_SPREADSHEET_ID": "staging-db-sheet-id",
    "MASTER_SOURCE_SPREADSHEET_ID": _config.MASTER_SOURCE_SPREADSHEET_ID_FIXED,
    "PROD_PDF_FOLDER_ID": _config.PROD_PDF_FOLDER_ID_FIXED,
    "STAGING_PDF_FOLDER_ID": "staging-pdf-folder-id",
    "PROD_LEDGER_SPREADSHEET_ID": _config.PROD_LEDGER_SPREADSHEET_ID_FIXED,
    "STAGING_LEDGER_SPREADSHEET_ID": "staging-ledger-id",
    "GAS_INTERNAL_URL": "https://script.google.com/macros/s/AKfycbTESTstagingInternal/exec",
    "GAS_INTERNAL_TOKEN": "test-gas-internal-token-0123456789abcdef",
    "GAS_CALL_TIMEOUT_SECONDS": "330",
    "SEBASTIAN_API_TOKEN": "test-sebastian-token-0123456789abcdef0123",
    "SEBASTIAN_PDF_WAIT_SECONDS": "40",
    "DB_SHEET_TIMEZONE": "Asia/Tokyo",
    "PORTAL_URL": "https://portal.example.lineworks-local.info/",
    "SHEET_CONFIG_SYNC_MINUTES": "5",
}
for _k, _v in TEST_ENV.items():
    os.environ[_k] = _v

from datetime import datetime  # noqa: E402

import pytest  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402
from sqlalchemy.orm import Session, sessionmaker  # noqa: E402

from app import jobs as _jobs  # noqa: E402
from app import models as _models  # noqa: E402,F401 - すべての表を読み込む
from app import rpc, timeutil, worker  # noqa: E402
from app.auth import create_session  # noqa: E402
from app.config import Settings, get_settings  # noqa: E402
from app.db import Base, get_db, make_engine  # noqa: E402
from app.main import create_app  # noqa: E402

get_settings.cache_clear()


def make_settings(**overrides) -> Settings:
    values = {k.lower(): v for k, v in TEST_ENV.items()}
    values.update(overrides)
    return Settings(_env_file=None, **values)  # type: ignore[call-arg]


@pytest.fixture
def settings() -> Settings:
    return make_settings()


def _test_database_url() -> str:
    """既定は SQLite（メモリー）。TEST_DATABASE_URL で PostgreSQL の結合試験にする（DB 名に test が必要）。"""
    url = os.environ.get("TEST_DATABASE_URL", "").strip()
    if not url:
        return "sqlite+pysqlite:///:memory:"
    dbname = url.rsplit("/", 1)[-1].split("?", 1)[0]
    if "test" not in dbname.lower():
        raise RuntimeError("TEST_DATABASE_URL の DB 名には test を含めてください（表を消すため）")
    return url


@pytest.fixture
def engine():
    eng = make_engine(_test_database_url())
    Base.metadata.drop_all(eng)
    Base.metadata.create_all(eng)
    yield eng
    Base.metadata.drop_all(eng)
    eng.dispose()


@pytest.fixture
def session_factory(engine):
    return sessionmaker(bind=engine, expire_on_commit=False)


@pytest.fixture
def db(session_factory) -> Session:
    s = session_factory()
    yield s
    s.close()


@pytest.fixture(autouse=True)
def _open_write_gate(request, monkeypatch):
    """シートへ書く前の関所（データ安全 高-3）は、ほかの試験では開けておく。関所そのものの試験は
    @pytest.mark.real_write_gate を付ける（tests/test_write_gate.py）。"""
    if request.node.get_closest_marker("real_write_gate"):
        return
    from app.services import write_gate

    monkeypatch.setattr(write_gate, "check", lambda *a, **k: (True, ""))


@pytest.fixture(autouse=True)
def _reset_registries():
    """登録（RPC・仕事の種類・予定）を試験ごとに空へ戻す。"""
    rpc.reset_registry()
    _jobs.unregister_all()
    worker.unregister_schedules()
    yield
    rpc.reset_registry()
    _jobs.unregister_all()
    worker.unregister_schedules()


@pytest.fixture
def app(session_factory, settings):
    application = create_app(load_handlers=False)

    def _db():
        s = session_factory()
        try:
            yield s
        finally:
            s.close()

    application.dependency_overrides[get_db] = _db
    application.dependency_overrides[get_settings] = lambda: settings
    return application


@pytest.fixture
def client(app) -> TestClient:
    # ブラウザが同じサイトの fetch に付ける Sec-Fetch-Site（RPC の入口で確かめる）を既定で付ける
    return TestClient(app, base_url="https://testserver", follow_redirects=False,
                      headers={"Sec-Fetch-Site": "same-origin"})


def login(client: TestClient, db: Session, email: str, settings: Settings | None = None) -> None:
    """email でログインしたセッションを作り、クッキーを client に付ける。"""
    from fastapi.responses import Response

    response = Response()
    create_session(db, response, email, settings or make_settings())
    for header in response.headers.getlist("set-cookie"):
        name, _, rest = header.partition("=")
        client.cookies.set(name, rest.split(";", 1)[0])


@pytest.fixture
def login_as(client, db, settings):
    """login_as('someone@lineworks-local.info') でその人としてログインする。"""
    return lambda email: login(client, db, email, settings)


def make_admin(db: Session, *emails: str) -> None:
    """Settings の ADMIN_EMAILS にメールを入れる（config_sync が取り込んだのと同じ形）。"""
    from app.models import Setting

    row = db.get(Setting, "ADMIN_EMAILS")
    value = ",".join(emails)
    if row is None:
        db.add(Setting(key="ADMIN_EMAILS", value=value, sheet_row=3))
    else:
        row.value = value
    db.commit()


def freeze_now(monkeypatch, moment: datetime) -> None:
    """timeutil.now_utc を固定する（時差付きの datetime を渡す）。"""
    monkeypatch.setattr(timeutil, "now_utc", lambda: moment.astimezone(timeutil.UTC))
