"""設定（fail-closed）と安全装置（実装仕様 1-1・9 章）。本番とステージングの ID の取り違えを全部の組み合わせで拒む。"""
from __future__ import annotations

import re
from pathlib import Path

import pytest
from pydantic import ValidationError

from app import config, safety
from app.config import (
    KNOWN_STAGING_IDS,
    MASTER_SOURCE_SPREADSHEET_ID_FIXED,
    PROD_GAS_DEPLOYMENT_IDS_FIXED,
    PROD_GAS_SCRIPT_ID_FIXED,
    PROD_LEDGER_SPREADSHEET_ID_FIXED,
    PROD_PDF_FOLDER_ID_FIXED,
    PROD_SPREADSHEET_ID_FIXED,
    PRODUCTION_IDS,
    STAGING_GAS_DEPLOYMENT_IDS_KNOWN,
    STAGING_GAS_SCRIPT_ID_KNOWN,
    STAGING_LEDGER_SPREADSHEET_ID_KNOWN,
    STAGING_PDF_FOLDER_ID_KNOWN,
    STAGING_SPREADSHEET_ID_KNOWN,
)
from app.safety import UnsafeWriteError
from conftest import TEST_ENV, make_settings

APP_DIR = Path(__file__).resolve().parents[1] / "app"
STAGING_SHEET = TEST_ENV["STAGING_SPREADSHEET_ID"]


def gas_url(deployment: str) -> str:
    return f"https://script.google.com/macros/s/{deployment}/exec"


def prod_settings(**overrides):
    values = {"app_env": "production", "staging_spreadsheet_id": "", "staging_pdf_folder_id": "",
              "staging_ledger_spreadsheet_id": "", "gas_internal_url": gas_url("AKfycbTESTprodInternal")}
    values.update(overrides)
    return make_settings(**values)


# ===== 必須の値・形 =====


@pytest.mark.parametrize("key", [
    "app_env", "database_url", "public_origin", "google_oauth_client_id", "google_oauth_client_secret",
    "session_secret", "allowed_domain", "bootstrap_admins", "fernet_key", "delegate_email", "prod_spreadsheet_id",
    "master_source_spreadsheet_id", "prod_pdf_folder_id", "prod_ledger_spreadsheet_id", "gas_internal_url",
    "gas_internal_token", "sebastian_api_token", "db_sheet_timezone", "portal_url",
])
def test_missing_required_value_refuses_to_start(key, monkeypatch):
    monkeypatch.delenv(key.upper(), raising=False)
    values = {k.lower(): v for k, v in TEST_ENV.items() if k.lower() != key}
    with pytest.raises(ValidationError):
        config.Settings(_env_file=None, **values)  # type: ignore[call-arg]


def test_defaults_match_spec(settings):
    assert settings.session_max_age_days == 90
    assert settings.session_idle_hours == 2160  # 90 日（9-1 #24）
    assert settings.gas_call_timeout_seconds == 330
    assert settings.sebastian_pdf_wait_seconds == 40
    assert settings.sheet_config_sync_minutes == 5
    assert settings.app_url == "https://overtime.lineworks-local.net/"


@pytest.mark.parametrize("overrides", [
    {"app_env": "dev"},
    {"public_origin": "http://overtime.lineworks-local.net"},
    {"public_origin": "https://overtime.lineworks-local.net/"},
    {"session_secret": "short"},
    {"session_max_age_days": 91},
    {"fernet_key": "not-a-key"},
    {"bootstrap_admins": ""},
    {"bootstrap_admins": " , "},
    {"bootstrap_admins": "not-an-email"},
    {"delegate_email": "owner@gmail.com"},
    {"gas_internal_url": "https://script.google.com/a/macros/lineworks-local.info/s/AKfycbX/exec"},
    {"gas_internal_url": "https://evil.example.com/macros/s/AKfycbX/exec"},
    {"gas_internal_url": "https://script.google.com/macros/s/AKfycbX/dev"},
    {"gas_internal_token": "too-short"},
    {"db_sheet_timezone": "Mars/Base"},
    {"portal_url": "http://portal"},
])
def test_bad_values_refuse_to_start(overrides):
    with pytest.raises(ValidationError):
        make_settings(**overrides)


@pytest.mark.parametrize("key,wrong", [
    ("prod_spreadsheet_id", STAGING_SPREADSHEET_ID_KNOWN),
    ("prod_spreadsheet_id", "1xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"),
    ("master_source_spreadsheet_id", PROD_SPREADSHEET_ID_FIXED),
    ("prod_pdf_folder_id", STAGING_PDF_FOLDER_ID_KNOWN),
    ("prod_ledger_spreadsheet_id", STAGING_LEDGER_SPREADSHEET_ID_KNOWN),
])
@pytest.mark.parametrize("env", ["staging", "production"])
def test_production_ids_must_equal_constants(env, key, wrong):
    base = prod_settings if env == "production" else make_settings
    with pytest.raises(ValidationError):
        base(**{key: wrong})


# ===== staging は本番の ID をどこにも書けない =====


@pytest.mark.parametrize("key", ["staging_spreadsheet_id", "staging_pdf_folder_id", "staging_ledger_spreadsheet_id"])
@pytest.mark.parametrize("prod_id", sorted(PRODUCTION_IDS))
def test_staging_refuses_every_production_id(key, prod_id):
    with pytest.raises(ValidationError):
        make_settings(**{key: prod_id})


@pytest.mark.parametrize("key", ["staging_spreadsheet_id", "staging_pdf_folder_id", "staging_ledger_spreadsheet_id"])
def test_staging_requires_its_ids(key):
    with pytest.raises(ValidationError):
        make_settings(**{key: ""})


def test_staging_ids_must_not_overlap():
    with pytest.raises(ValidationError):
        make_settings(staging_ledger_spreadsheet_id=STAGING_SHEET)


@pytest.mark.parametrize("deployment", PROD_GAS_DEPLOYMENT_IDS_FIXED)
def test_staging_refuses_production_gas_deployment(deployment):
    with pytest.raises(ValidationError):
        make_settings(gas_internal_url=gas_url(deployment))


def test_staging_accepts_the_real_staging_values():
    s = make_settings(staging_spreadsheet_id=STAGING_SPREADSHEET_ID_KNOWN,
                      staging_ledger_spreadsheet_id=STAGING_LEDGER_SPREADSHEET_ID_KNOWN,
                      staging_pdf_folder_id=STAGING_PDF_FOLDER_ID_KNOWN,
                      gas_internal_url=gas_url(STAGING_GAS_DEPLOYMENT_IDS_KNOWN[0]))
    assert s.db_spreadsheet_id == STAGING_SPREADSHEET_ID_KNOWN
    assert s.ledger_spreadsheet_id == STAGING_LEDGER_SPREADSHEET_ID_KNOWN
    assert s.pdf_folder_id == STAGING_PDF_FOLDER_ID_KNOWN
    safety.verify_startup(s)


# ===== production はステージングの ID を指せない =====


@pytest.mark.parametrize("deployment", STAGING_GAS_DEPLOYMENT_IDS_KNOWN)
def test_production_refuses_staging_gas_deployment(deployment):
    with pytest.raises(ValidationError):
        prod_settings(gas_internal_url=gas_url(deployment))


@pytest.mark.parametrize("key", ["staging_spreadsheet_id", "staging_pdf_folder_id", "staging_ledger_spreadsheet_id"])
@pytest.mark.parametrize("prod_id", sorted(PRODUCTION_IDS))
def test_production_refuses_staging_fields_pointing_to_production(key, prod_id):
    with pytest.raises(ValidationError):
        prod_settings(**{key: prod_id})


def test_production_does_not_need_staging_values():
    s = prod_settings()
    assert s.db_spreadsheet_id == PROD_SPREADSHEET_ID_FIXED
    assert s.pdf_folder_id == PROD_PDF_FOLDER_ID_FIXED
    assert s.ledger_spreadsheet_id == PROD_LEDGER_SPREADSHEET_ID_FIXED
    safety.verify_startup(s)


# ===== 書き込みの直前の許可リスト =====

A_TABS = ["Requests", "WorkLogs", "BatchLogs"]
C_TABS = ["作業員マスタ", "工番マスタ", "社内カレンダーマスタ", "業務NOマスタ"]
FORBIDDEN_TABS = ["Settings", "DeptApprovers", "ApproverMap", "StampMap", "休憩マスタ", "申請書フォーム", "操作",
                  "休憩計算_フォーム", "FormMap", "FormTemplates", "ExportHistory", "部署マスタ", "Sheet1"]


@pytest.mark.parametrize("tab", A_TABS)
@pytest.mark.parametrize("purpose", ["row_sync", "reconcile"])
def test_a_tabs_writable_by_row_sync_and_reconcile(settings, tab, purpose):
    assert safety.assert_writable(STAGING_SHEET, tab, purpose, settings) == STAGING_SHEET
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(STAGING_SHEET, tab, "master_sync", settings)


@pytest.mark.parametrize("tab", C_TABS)
def test_c_tabs_writable_only_by_master_sync(settings, tab):
    assert safety.assert_writable(STAGING_SHEET, tab, "master_sync", settings)
    for purpose in ("row_sync", "reconcile"):
        with pytest.raises(UnsafeWriteError):
            safety.assert_writable(STAGING_SHEET, tab, purpose, settings)


@pytest.mark.parametrize("tab", FORBIDDEN_TABS)
@pytest.mark.parametrize("purpose", ["row_sync", "reconcile", "master_sync"])
def test_b_d_e_tabs_are_never_writable(settings, tab, purpose):
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(STAGING_SHEET, tab, purpose, settings)


def test_unknown_purpose_is_refused(settings):
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(STAGING_SHEET, "Requests", "export", settings)


ALL_KNOWN_IDS = sorted(PRODUCTION_IDS | KNOWN_STAGING_IDS | {
    TEST_ENV["STAGING_LEDGER_SPREADSHEET_ID"], TEST_ENV["STAGING_PDF_FOLDER_ID"], "some-other-sheet"})


@pytest.mark.parametrize("sid", ALL_KNOWN_IDS)
def test_staging_writes_only_to_staging_db(settings, sid):
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(sid, "Requests", "row_sync", settings)


@pytest.mark.parametrize("sid", sorted((PRODUCTION_IDS | KNOWN_STAGING_IDS | {"some-other-sheet", STAGING_SHEET})
                                       - {PROD_SPREADSHEET_ID_FIXED}))
def test_production_writes_only_to_production_db(sid):
    s = prod_settings()
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(sid, "Requests", "row_sync", s)
    assert safety.assert_writable(PROD_SPREADSHEET_ID_FIXED, "Requests", "row_sync", s)


def test_production_refuses_known_staging_db_even_with_staging_values_set():
    s = prod_settings(staging_spreadsheet_id=STAGING_SPREADSHEET_ID_KNOWN)
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(STAGING_SPREADSHEET_ID_KNOWN, "Requests", "row_sync", s)


@pytest.mark.parametrize("sid", ["", "   "])
def test_empty_id_is_refused(settings, sid):
    with pytest.raises(UnsafeWriteError):
        safety.assert_writable(sid, "Requests", "row_sync", settings)


@pytest.mark.parametrize("kind", ["deleteDimension", "insertDimension", "deleteSheet", "addSheet", "duplicateSheet",
                                  "sortRange", "moveDimension", "updateSheetProperties", "deleteRange",
                                  "repeatCell", "pasteData", "clearBasicFilter"])
def test_forbidden_batch_requests_are_refused(kind):
    with pytest.raises(UnsafeWriteError):
        safety.assert_write_requests([{kind: {"range": {"sheetId": 7}}}], "Requests", 7)


def test_allowed_batch_requests():
    safety.assert_write_requests([{"updateCells": {"range": {"sheetId": 7}, "rows": [], "fields": "*"}},
                                  {"appendDimension": {"sheetId": 7, "dimension": "ROWS", "length": 3}}],
                                 "Requests", 7)
    with pytest.raises(UnsafeWriteError):  # 列は足さない
        safety.assert_write_requests([{"appendDimension": {"sheetId": 7, "dimension": "COLUMNS", "length": 1}}],
                                     "Requests", 7)
    with pytest.raises(UnsafeWriteError):  # ほかのシート
        safety.assert_write_requests([{"updateCells": {"range": {"sheetId": 8}}}], "Requests", 7)
    with pytest.raises(UnsafeWriteError):  # C 区分に行を足さない
        safety.assert_write_requests([{"appendDimension": {"sheetId": 7, "dimension": "ROWS", "length": 1}}],
                                     "工番マスタ", 7)
    with pytest.raises(UnsafeWriteError):
        safety.assert_write_requests([], "Requests", 7)


def test_write_methods():
    safety.assert_write_method("values.batchClear", "工番マスタ")
    for tab in A_TABS:
        with pytest.raises(UnsafeWriteError):  # A 区分の clear は使わない
            safety.assert_write_method("values.batchClear", tab)
    with pytest.raises(UnsafeWriteError):
        safety.assert_write_method("values.clear", "工番マスタ")


@pytest.mark.parametrize("a1,tab", [("'Requests'!A1:B2", "Requests"), ("Requests!A1", "Requests"),
                                    ("'It''s'!A1", "It's"), ("'工番マスタ'", "工番マスタ")])
def test_a1_sheet_name(a1, tab):
    assert safety.a1_sheet_name(a1) == tab


def test_ranges_must_stay_in_tab():
    safety.assert_ranges_in_tab(["'工番マスタ'!A1:B2"], "工番マスタ")
    with pytest.raises(UnsafeWriteError):
        safety.assert_ranges_in_tab(["'工番マスタ'!A1", "'Settings'!A1"], "工番マスタ")


# ===== GAS の照合（verify_gas_ping） =====


def staging_ping(**overrides):
    ping = {"ok": True, "scriptId": "staging-script-id", "spreadsheetId": STAGING_SHEET,
            "spreadsheetTimeZone": "Asia/Tokyo", "scriptTimeZone": "Asia/Tokyo", "migrated": False,
            "settings": {"TEMPLATE_SSID": STAGING_SHEET, "PDF_ROOT_FOLDER_ID": TEST_ENV["STAGING_PDF_FOLDER_ID"],
                         "ACCUMULATION_SS_ID": TEST_ENV["STAGING_LEDGER_SPREADSHEET_ID"], "PDF_MODE": "direct",
                         "SOURCE_SSID": MASTER_SOURCE_SPREADSHEET_ID_FIXED}, "apiVersion": 1}
    ping.update(overrides)
    return ping


def prod_ping(**overrides):
    ping = {"ok": True, "scriptId": PROD_GAS_SCRIPT_ID_FIXED, "spreadsheetId": PROD_SPREADSHEET_ID_FIXED,
            "spreadsheetTimeZone": "Asia/Tokyo",
            "settings": {"TEMPLATE_SSID": PROD_SPREADSHEET_ID_FIXED, "PDF_ROOT_FOLDER_ID": PROD_PDF_FOLDER_ID_FIXED,
                         "ACCUMULATION_SS_ID": PROD_LEDGER_SPREADSHEET_ID_FIXED}, "apiVersion": 1}
    ping.update(overrides)
    return ping


def with_setting(ping, key, value):
    ping = dict(ping)
    ping["settings"] = {**ping["settings"], key: value}
    return ping


def test_gas_ping_passes_when_everything_matches(settings):
    assert safety.verify_gas_ping(staging_ping(), settings) == []
    assert safety.verify_gas_ping(staging_ping(scriptId=STAGING_GAS_SCRIPT_ID_KNOWN), settings) == []
    assert safety.verify_gas_ping(prod_ping(), prod_settings()) == []


@pytest.mark.parametrize("ping", [
    None, "x", {"ok": False, "error": "unauthorized"},
    staging_ping(spreadsheetId=PROD_SPREADSHEET_ID_FIXED),
    staging_ping(scriptId=PROD_GAS_SCRIPT_ID_FIXED),
    staging_ping(scriptId=""),
    staging_ping(spreadsheetTimeZone="Etc/GMT"),
    staging_ping(spreadsheetId="another-sheet"),
    with_setting(staging_ping(), "TEMPLATE_SSID", PROD_SPREADSHEET_ID_FIXED),
    with_setting(staging_ping(), "PDF_ROOT_FOLDER_ID", PROD_PDF_FOLDER_ID_FIXED),
    with_setting(staging_ping(), "ACCUMULATION_SS_ID", PROD_LEDGER_SPREADSHEET_ID_FIXED),
    with_setting(staging_ping(), "ACCUMULATION_SS_ID", ""),
])
def test_gas_ping_staging_mismatches(settings, ping):
    assert safety.verify_gas_ping(ping, settings)


@pytest.mark.parametrize("ping", [
    prod_ping(spreadsheetId=STAGING_SPREADSHEET_ID_KNOWN),
    prod_ping(scriptId=STAGING_GAS_SCRIPT_ID_KNOWN),
    prod_ping(spreadsheetTimeZone="Etc/GMT"),
    with_setting(prod_ping(), "TEMPLATE_SSID", STAGING_SPREADSHEET_ID_KNOWN),
    with_setting(prod_ping(), "PDF_ROOT_FOLDER_ID", STAGING_PDF_FOLDER_ID_KNOWN),
    with_setting(prod_ping(), "ACCUMULATION_SS_ID", STAGING_LEDGER_SPREADSHEET_ID_KNOWN),
    with_setting(prod_ping(), "ACCUMULATION_SS_ID", "some-other-ledger"),
])
def test_gas_ping_production_mismatches(ping):
    assert safety.verify_gas_ping(ping, prod_settings())


def test_gas_ping_reports_production_id_by_name(settings):
    problems = safety.verify_gas_ping(with_setting(staging_ping(), "PDF_ROOT_FOLDER_ID", PROD_PDF_FOLDER_ID_FIXED),
                                      settings)
    assert any("PDF_ROOT_FOLDER_ID が本番の ID" in p for p in problems)


# ===== ほかのモジュールに ID を直書きしない（実装仕様 1-1） =====

# ID の頭 8 文字（定数から作る。この試験のファイルにも ID を直書きしない）と、オーナーのメールの頭
ID_PATTERNS = sorted({i[:8] for i in PRODUCTION_IDS | KNOWN_STAGING_IDS} | {"AKfycb", "imaizumi" + "@"})


def test_ids_are_written_only_in_config_py():
    offenders = []
    for path in APP_DIR.rglob("*.py"):
        if path.name == "config.py" and path.parent == APP_DIR:
            continue
        text = path.read_text(encoding="utf-8")
        for pattern in ID_PATTERNS:
            if pattern in text:
                offenders.append(f"{path.relative_to(APP_DIR)}: {pattern}")
    assert offenders == []


def test_config_defines_the_fixed_constants():
    for name in ("PROD_SPREADSHEET_ID_FIXED", "MASTER_SOURCE_SPREADSHEET_ID_FIXED", "PROD_GAS_SCRIPT_ID_FIXED",
                 "PROD_PDF_FOLDER_ID_FIXED", "PROD_LEDGER_SPREADSHEET_ID_FIXED"):
        assert re.fullmatch(r"[A-Za-z0-9_-]{25,}", getattr(config, name))
    assert not (PRODUCTION_IDS & KNOWN_STAGING_IDS)


# ===== Sebastian の token の強さ（セキュリティ確認 中-1） =====


def _env_values(app_env, token):
    values = {k.lower(): v for k, v in TEST_ENV.items()}
    values.update(app_env=app_env, sebastian_api_token=token)
    if app_env == "production":
        values.update(gas_internal_url="https://script.google.com/macros/s/AKfycbPRODinternal/exec")
    return values


def test_staging_requires_32_char_sebastian_token(monkeypatch):
    monkeypatch.delenv("SEBASTIAN_API_TOKEN", raising=False)
    with pytest.raises(ValidationError, match="32 字以上"):
        config.Settings(_env_file=None, **_env_values("staging", "x" * 31))  # type: ignore[call-arg]
    s = config.Settings(_env_file=None, **_env_values("staging", "x" * 32))  # type: ignore[call-arg]
    assert s.sebastian_token_is_weak is False


def test_production_accepts_existing_short_sebastian_token_with_warning(monkeypatch, caplog):
    """verify-fixes R-2: 本番は Sebastian が今使う token をそのまま使う（8 字以上なら起動、32 字未満は警告）。"""
    import logging

    monkeypatch.delenv("SEBASTIAN_API_TOKEN", raising=False)
    monkeypatch.delenv("APP_ENV", raising=False)
    monkeypatch.delenv("GAS_INTERNAL_URL", raising=False)
    secret = "abcDEF1234"  # 10 字
    with caplog.at_level(logging.WARNING, logger="app.config"):
        s = config.Settings(_env_file=None, **_env_values("production", secret))  # type: ignore[call-arg]
    assert s.sebastian_token_is_weak is True
    assert any("SEBASTIAN_API_TOKEN" in r.getMessage() for r in caplog.records)
    assert all(secret not in r.getMessage() for r in caplog.records)  # 値はログに出さない
    with pytest.raises(ValidationError):
        config.Settings(_env_file=None, **_env_values("production", "x" * 7))  # type: ignore[call-arg]
    with pytest.raises(ValidationError):
        config.Settings(_env_file=None, **_env_values("production", ""))  # type: ignore[call-arg]
    assert config.Settings(_env_file=None, **_env_values("production", "y" * 32)).sebastian_token_is_weak is False


def test_weak_sebastian_token_is_shown_in_ops_status(db, monkeypatch):
    from app.services import ops_status

    weak = type("S", (), {"sebastian_token_is_weak": True, "is_staging": False, "sheet_write_ping_max_minutes": 10})()
    monkeypatch.setattr(ops_status, "get_settings", lambda: weak)
    warnings = ops_status.get_ops_status(db)["warnings"]
    assert warnings and "32 字未満" in warnings[0]


def test_sebastian_token_must_differ_from_gas_internal_token():
    with pytest.raises(ValidationError, match="同じ値"):
        make_settings(sebastian_api_token=TEST_ENV["GAS_INTERNAL_TOKEN"])


# ===== 本番の内部用の GAS のデプロイ（データ安全 低-6） =====

PROD_INTERNAL = "AKfycbPRODinternalDeployment0123456789abcdefXYZ"


def _prod_values(**kw):
    values = {k.lower(): v for k, v in TEST_ENV.items()}
    values.update(app_env="production", gas_internal_url=f"https://script.google.com/macros/s/{PROD_INTERNAL}/exec")
    values.update(kw)
    return values


def test_prod_internal_deployment_id_must_match_in_production(monkeypatch):
    for k in ("APP_ENV", "GAS_INTERNAL_URL"):
        monkeypatch.delenv(k, raising=False)
    s = config.Settings(_env_file=None, **_prod_values(prod_gas_internal_deployment_id=PROD_INTERNAL))
    assert PROD_INTERNAL in s.prod_gas_deployment_ids
    with pytest.raises(ValidationError, match="PROD_GAS_INTERNAL_DEPLOYMENT_ID"):
        config.Settings(_env_file=None, **_prod_values(
            prod_gas_internal_deployment_id=PROD_INTERNAL,
            gas_internal_url="https://script.google.com/macros/s/AKfycbOTHERdeployment0123456789abcdefgh/exec"))


def test_staging_refuses_the_configured_prod_internal_deployment():
    with pytest.raises(ValidationError, match="本番の GAS のデプロイ"):
        make_settings(prod_gas_internal_deployment_id=PROD_INTERNAL,
                      gas_internal_url=f"https://script.google.com/macros/s/{PROD_INTERNAL}/exec")
    assert make_settings(prod_gas_internal_deployment_id=PROD_INTERNAL).prod_gas_internal_deployment_id == PROD_INTERNAL


@pytest.mark.parametrize("bad", ["not-an-id", "AKfycshort", config.STAGING_GAS_DEPLOYMENT_IDS_KNOWN[0]])
def test_prod_internal_deployment_id_is_validated(bad):
    with pytest.raises(ValidationError):
        make_settings(prod_gas_internal_deployment_id=bad)


def test_gas_client_refuses_other_prod_deployment(monkeypatch):
    from app.services.gas_client import GasClientConfigError, assert_gas_url_allowed

    for k in ("APP_ENV", "GAS_INTERNAL_URL"):
        monkeypatch.delenv(k, raising=False)
    s = config.Settings(_env_file=None, **_prod_values(prod_gas_internal_deployment_id=PROD_INTERNAL))
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed("https://script.google.com/macros/s/AKfycbOTHERdeployment0123456789abcdefgh/exec", s)
    staging = make_settings(prod_gas_internal_deployment_id=PROD_INTERNAL)
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed(f"https://script.google.com/macros/s/{PROD_INTERNAL}/exec", staging)



def test_gas_ping_empty_timezone_is_accepted_only_for_gmt_family(settings):
    """Apps Script は GMT のシートで getSpreadsheetTimeZone() に空を返す（2026-11-26 確認）。
    GMT の系統を DB_SHEET_TIMEZONE にしたときだけ、空を同じ時間帯として合格にする。"""
    gmt = settings.model_copy(update={"db_sheet_timezone": "Etc/GMT"})
    assert safety.verify_gas_ping(staging_ping(spreadsheetTimeZone=""), gmt) == []
    assert safety.verify_gas_ping(staging_ping(spreadsheetTimeZone="Etc/GMT"), gmt) == []
    assert safety.verify_gas_ping(staging_ping(spreadsheetTimeZone="Asia/Tokyo"), gmt)
    # 東京の設定では、空はこれまでどおり不合格
    assert safety.verify_gas_ping(staging_ping(spreadsheetTimeZone=""), settings)
