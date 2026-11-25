"""環境変数の読み込み（実装仕様 1-1）。必須の値が欠けている・危ない組み合わせのときは起動しない（fail-closed）。

業務用の ID を読むのはこのファイルだけ。本番の ID と、取り違えを防ぐために知っておくステージングの ID は、
コードにも定数として持つ（ほかのモジュールに ID を直書きしない。tests/test_config_safety.py が拒む）。

- APP_ENV=production: 本番の ID が定数と同じでなければ起動しない。ステージングの ID（既知の値）を
  どこかで指していれば起動しない。
- APP_ENV=staging: STAGING_* が必須。本番の ID（定数）をどこかで指していれば起動しない。
"""
from __future__ import annotations

import hmac
import logging
import re
from functools import lru_cache
from typing import Literal
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from cryptography.fernet import Fernet
from pydantic import Field, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

# ===== 本番の ID（計画書 4-0。.env の値と違えば起動しない） =====
# 本番 DB スプレッドシート（PDF テンプレも同じファイル）
PROD_SPREADSHEET_ID_FIXED = "1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ"
# マスタ元（SOURCE_SSID。読むだけ。どの環境からも書かない）
MASTER_SOURCE_SPREADSHEET_ID_FIXED = "1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ"
# 本番 GAS の scriptId（internal.ping の照合）
PROD_GAS_SCRIPT_ID_FIXED = "1GPNBR9U9mw7TZPnl1pxtdVjFagUPxt9S6qSYlbNIEuHNvRLVspNIQNc7"
# 本番 PDF フォルダ（VPS は書かない。internal.ping の照合だけ）
PROD_PDF_FOLDER_ID_FIXED = "1Bs1FvgRvlCAkpARZ0XN3YkXJS9eoIN1B"
# 蓄積台帳（VPS は書かない。GAS が書く）
PROD_LEDGER_SPREADSHEET_ID_FIXED = "1tKlT57SJxJijgfDVGafZ434sdVjscJJqN2uQ7P3sUyU"
# 本番 GAS の公開版（画面用 @131）。ステージングの GAS_INTERNAL_URL がこれを指していれば起動しない。
# 本番の内部用のデプロイの deploymentId は、作った後に .env の PROD_GAS_INTERNAL_DEPLOYMENT_ID で渡す
# （両方の環境の .env に書く。Settings.prod_gas_deployment_ids がこの定数と合わせて使う。データ安全 低-6）。
PROD_GAS_DEPLOYMENT_IDS_FIXED: tuple[str, ...] = (
    "AKfycbwHnaECeni3FmmgtVzBZIM_xDMH3bw9Qx5sQMrW9J3XoRCOs_RmdyBensZX5Xta6sfIxA",
)

# ===== ステージングの ID（計画書 4-0。本番がこれを指していれば起動しない・書かない） =====
STAGING_SPREADSHEET_ID_KNOWN = "1I0dPlH6w14UaiJVbO9i6wFCXlvGyPEJvDus8ekGptWY"
STAGING_LEDGER_SPREADSHEET_ID_KNOWN = "16M5PxyozS_tTEpQmR6jizICgZov4SoK2M1IO4TG1oHI"
STAGING_PDF_FOLDER_ID_KNOWN = "1KsWanVEQr-gMsb_N9mxzplEjKFibsnrc"
STAGING_GAS_SCRIPT_ID_KNOWN = "1y027cB_59N1P2o3toawpCyPZ2k_W6mPETtgiDpPFtpY7TKgBWslnyLNM"
STAGING_GAS_DEPLOYMENT_IDS_KNOWN: tuple[str, ...] = (
    "AKfycbw9uPC4C5zlY5S3YqR9Yo0TECZETvUUgGY2ih447eL8E8XS0_VhVpwtftinTr4Nj5A",
)

# 本番のファイル・フォルダ・スクリプトの ID（ステージングはどれも指してはいけない）
PRODUCTION_IDS: frozenset[str] = frozenset({
    PROD_SPREADSHEET_ID_FIXED,
    MASTER_SOURCE_SPREADSHEET_ID_FIXED,
    PROD_GAS_SCRIPT_ID_FIXED,
    PROD_PDF_FOLDER_ID_FIXED,
    PROD_LEDGER_SPREADSHEET_ID_FIXED,
    *PROD_GAS_DEPLOYMENT_IDS_FIXED,
})
# 既知のステージングの ID（本番はどれも指してはいけない）
KNOWN_STAGING_IDS: frozenset[str] = frozenset({
    STAGING_SPREADSHEET_ID_KNOWN,
    STAGING_LEDGER_SPREADSHEET_ID_KNOWN,
    STAGING_PDF_FOLDER_ID_KNOWN,
    STAGING_GAS_SCRIPT_ID_KNOWN,
    *STAGING_GAS_DEPLOYMENT_IDS_KNOWN,
})

log = logging.getLogger(__name__)

SEBASTIAN_TOKEN_MIN_PRODUCTION = 8
SEBASTIAN_TOKEN_RECOMMENDED = 32
GAS_DEPLOYMENT_ID_RE = re.compile(r"^AKfyc[A-Za-z0-9_-]{30,}\Z")
GAS_INTERNAL_URL_RE = re.compile(r"^https://script\.google\.com/macros/s/([A-Za-z0-9_-]+)/exec\Z")
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\Z")


def split_emails(value: str | None) -> list[str]:
    """カンマ・改行・セミコロン区切りのメール。小文字にして空を除く。"""
    return [e.strip().lower() for e in re.split(r"[,\n;]", value or "") if e.strip()]


def gas_deployment_id(url: str) -> str:
    """GAS_INTERNAL_URL の deploymentId（形が違えば ''）。"""
    m = GAS_INTERNAL_URL_RE.match(url or "")
    return m.group(1) if m else ""


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=None, extra="ignore", case_sensitive=False)

    app_env: Literal["staging", "production"]
    database_url: str = Field(min_length=1)
    public_origin: str = Field(min_length=1)

    # Google ログイン（このアプリ専用の OAuth クライアント）
    google_oauth_client_id: str = Field(min_length=1)
    google_oauth_client_secret: str = Field(min_length=1)
    session_secret: str = Field(min_length=32)
    # 9-1 #24: 共有アカウントのログインの持ち時間は 90 日（無操作でも 90 日は切らない）
    session_max_age_days: int = Field(default=90, ge=1, le=90)
    session_idle_hours: int = Field(default=2160, ge=1, le=24 * 90)
    allowed_domain: str = Field(min_length=1)
    # 管理者（Settings の ADMIN_EMAILS・GENERAL_AFFAIRS_EMAILS に足す）。委任の開始もこの人だけ。空なら起動しない
    bootstrap_admins: str = Field(min_length=1)

    # 委任（DELEGATE_EMAIL の権限で Sheets を使う）
    fernet_key: str = Field(min_length=1)
    delegate_email: str = Field(min_length=1)

    # 業務用の ID（読むのはこのファイルだけ）
    prod_spreadsheet_id: str = Field(min_length=1)
    staging_spreadsheet_id: str = ""
    master_source_spreadsheet_id: str = Field(min_length=1)
    prod_pdf_folder_id: str = Field(min_length=1)
    staging_pdf_folder_id: str = ""
    prod_ledger_spreadsheet_id: str = Field(min_length=1)
    staging_ledger_spreadsheet_id: str = ""

    # GAS の内部用の入口（実装仕様 7-1）
    gas_internal_url: str = Field(min_length=1)
    gas_internal_token: str = Field(min_length=32)
    gas_call_timeout_seconds: float = Field(default=330, gt=0, le=600)
    # 本番の内部用の GAS のデプロイの deploymentId（作ったら両方の環境に書く。空なら未設定）。
    # production: GAS_INTERNAL_URL がこれと違えば起動しない。staging: GAS_INTERNAL_URL がこれなら起動しない。
    prod_gas_internal_deployment_id: str = ""

    # Sebastian の入口（実装仕様 4-4）。社外から認証なしで届く入口（セキュリティ 中-1・verify-fixes R-2）:
    # - staging: 32 字以上を必須（ステージング用に作った乱数）。
    # - production: Sebastian が今使っている token（GAS のプロパティの値）をそのまま使うので変えられない。
    #   8 字以上なら起動し、32 字未満は WARNING をログ（値は出さない）と管理者の欄（ops_getStatus）に出す。
    # GAS の内部用の token と同じ値はどちらでも拒む。
    sebastian_api_token: str = Field(min_length=SEBASTIAN_TOKEN_MIN_PRODUCTION)
    sebastian_pdf_wait_seconds: float = Field(default=40, gt=0, le=300)

    # シートへ書く前の関所（services/write_gate。データ安全 高-3）: 合格した gas_ping がこの分数より古ければ書かない
    sheet_write_ping_max_minutes: int = Field(default=10, ge=1, le=60)

    # DB スプレッドシートの時間帯（IANA 名。シートの値と違えば行の反映を止める）
    db_sheet_timezone: str = Field(min_length=1)
    # 社内ポータルの固定 URL（画面の「社内ポータルへ移動」）
    portal_url: str = Field(min_length=1)
    # 共有タブ（Settings・DeptApprovers・ApproverMap）の取り込みの間隔（分）
    sheet_config_sync_minutes: int = Field(default=5, ge=1, le=60)

    @field_validator("public_origin")
    @classmethod
    def _origin(cls, v: str) -> str:
        v = v.strip()
        if not v.startswith("https://") or v.endswith("/"):
            raise ValueError("PUBLIC_ORIGIN は https:// で始め、末尾に / を付けないでください")
        return v

    @field_validator("portal_url")
    @classmethod
    def _portal(cls, v: str) -> str:
        v = v.strip()
        if not v.startswith("https://"):
            raise ValueError("PORTAL_URL は https:// で始めてください")
        return v

    @field_validator("allowed_domain", "delegate_email")
    @classmethod
    def _lower(cls, v: str) -> str:
        return v.strip().lower()

    @field_validator(
        "prod_spreadsheet_id", "staging_spreadsheet_id", "master_source_spreadsheet_id", "prod_pdf_folder_id",
        "staging_pdf_folder_id", "prod_ledger_spreadsheet_id", "staging_ledger_spreadsheet_id", "gas_internal_url",
    )
    @classmethod
    def _strip(cls, v: str) -> str:
        return v.strip()

    @field_validator("prod_gas_internal_deployment_id")
    @classmethod
    def _prod_deployment(cls, v: str) -> str:
        v = v.strip()
        if v and not GAS_DEPLOYMENT_ID_RE.match(v):
            raise ValueError("PROD_GAS_INTERNAL_DEPLOYMENT_ID が GAS の deploymentId の形（AKfyc…）ではありません")
        if v in KNOWN_STAGING_IDS:
            raise ValueError("PROD_GAS_INTERNAL_DEPLOYMENT_ID がステージングの GAS のデプロイです")
        return v

    @field_validator("gas_internal_url")
    @classmethod
    def _gas_url(cls, v: str) -> str:
        if not GAS_INTERNAL_URL_RE.match(v):
            raise ValueError("GAS_INTERNAL_URL は https://script.google.com/macros/s/<deploymentId>/exec の形にしてください")
        return v

    @field_validator("db_sheet_timezone")
    @classmethod
    def _tz(cls, v: str) -> str:
        v = v.strip()
        try:
            ZoneInfo(v)
        except (ZoneInfoNotFoundError, ValueError) as exc:
            raise ValueError(f"DB_SHEET_TIMEZONE が時間帯の名前（IANA）ではありません: {v}") from exc
        return v

    @field_validator("fernet_key")
    @classmethod
    def _fernet(cls, v: str) -> str:
        try:
            Fernet(v.encode())
        except Exception as exc:  # noqa: BLE001 - 形式違いはすべて起動失敗にする
            raise ValueError("FERNET_KEY が Fernet の鍵の形式ではありません") from exc
        return v

    @model_validator(mode="after")
    def _safety(self) -> "Settings":
        if self.prod_spreadsheet_id != PROD_SPREADSHEET_ID_FIXED:
            raise ValueError("PROD_SPREADSHEET_ID が本番 DB スプレッドシートの ID と違います")
        if self.master_source_spreadsheet_id != MASTER_SOURCE_SPREADSHEET_ID_FIXED:
            raise ValueError("MASTER_SOURCE_SPREADSHEET_ID がマスタ元の ID と違います")
        if self.prod_pdf_folder_id != PROD_PDF_FOLDER_ID_FIXED:
            raise ValueError("PROD_PDF_FOLDER_ID が本番 PDF フォルダの ID と違います")
        if self.prod_ledger_spreadsheet_id != PROD_LEDGER_SPREADSHEET_ID_FIXED:
            raise ValueError("PROD_LEDGER_SPREADSHEET_ID が蓄積台帳の ID と違います")
        if not self.bootstrap_admin_list:
            raise ValueError("BOOTSTRAP_ADMINS が空です")
        for email in self.bootstrap_admin_list:
            if not _EMAIL_RE.match(email):
                raise ValueError("BOOTSTRAP_ADMINS にメールアドレスでない値があります")
        if self.is_staging and len(self.sebastian_api_token) < SEBASTIAN_TOKEN_RECOMMENDED:
            raise ValueError(f"ステージングの SEBASTIAN_API_TOKEN は {SEBASTIAN_TOKEN_RECOMMENDED} 字以上にしてください")
        if self.sebastian_token_is_weak:
            log.warning("SEBASTIAN_API_TOKEN が %d 字未満です（Sebastian と一緒に入れ替えることをすすめます）",
                        SEBASTIAN_TOKEN_RECOMMENDED)
        if hmac.compare_digest(self.sebastian_api_token.encode(), self.gas_internal_token.encode()):
            raise ValueError("SEBASTIAN_API_TOKEN と GAS_INTERNAL_TOKEN に同じ値を使わないでください")
        if not self.delegate_email.endswith("@" + self.allowed_domain):
            raise ValueError("DELEGATE_EMAIL は社内ドメインのアドレスにしてください")

        deployment = gas_deployment_id(self.gas_internal_url)
        staging_values = {
            "STAGING_SPREADSHEET_ID": self.staging_spreadsheet_id,
            "STAGING_PDF_FOLDER_ID": self.staging_pdf_folder_id,
            "STAGING_LEDGER_SPREADSHEET_ID": self.staging_ledger_spreadsheet_id,
        }
        if self.app_env == "staging":
            missing = [name for name, value in staging_values.items() if not value]
            if missing:
                raise ValueError("ステージングに必要な値がありません: " + ", ".join(missing))
            for name, value in staging_values.items():
                if value in PRODUCTION_IDS:
                    raise ValueError(f"{name} が本番の ID を指しています")
            if len(set(staging_values.values())) != len(staging_values):
                raise ValueError("ステージングの DB・台帳・PDF フォルダの ID が重なっています")
            if deployment in PRODUCTION_IDS or deployment in self.prod_gas_deployment_ids:
                raise ValueError("ステージングの GAS_INTERNAL_URL が本番の GAS のデプロイを指しています")
        else:
            if deployment in KNOWN_STAGING_IDS:
                raise ValueError("本番の GAS_INTERNAL_URL がステージングの GAS のデプロイを指しています")
            if self.prod_gas_internal_deployment_id and deployment != self.prod_gas_internal_deployment_id:
                raise ValueError("本番の GAS_INTERNAL_URL が PROD_GAS_INTERNAL_DEPLOYMENT_ID のデプロイと違います")
            for name, value in staging_values.items():
                if value and value in PRODUCTION_IDS:
                    raise ValueError(f"{name} が本番の ID を指しています")
        return self

    @property
    def sebastian_token_is_weak(self) -> bool:
        """Sebastian の token が 32 字未満か（本番だけ起こりうる。値は出さない。管理者の欄の警告に使う）。"""
        return len(self.sebastian_api_token) < SEBASTIAN_TOKEN_RECOMMENDED

    @property
    def prod_gas_deployment_ids(self) -> tuple[str, ...]:
        """本番 GAS のデプロイ（定数の画面用＋設定した内部用）。"""
        extra = (self.prod_gas_internal_deployment_id,) if self.prod_gas_internal_deployment_id else ()
        return PROD_GAS_DEPLOYMENT_IDS_FIXED + extra

    @property
    def bootstrap_admin_list(self) -> list[str]:
        return split_emails(self.bootstrap_admins)

    @property
    def is_staging(self) -> bool:
        return self.app_env == "staging"

    @property
    def db_spreadsheet_id(self) -> str:
        """この環境の DB スプレッドシート（行の反映・照合・マスタ同期の書き込み先）。"""
        return self.staging_spreadsheet_id if self.is_staging else self.prod_spreadsheet_id

    @property
    def pdf_folder_id(self) -> str:
        """この環境の PDF フォルダ（GAS が書く。internal.ping の照合だけに使う）。"""
        return self.staging_pdf_folder_id if self.is_staging else self.prod_pdf_folder_id

    @property
    def ledger_spreadsheet_id(self) -> str:
        """この環境の蓄積台帳（GAS が書く。internal.ping の照合だけに使う）。"""
        return self.staging_ledger_spreadsheet_id if self.is_staging else self.prod_ledger_spreadsheet_id

    @property
    def app_url(self) -> str:
        """画面のテンプレート変数 APP_URL（PUBLIC_ORIGIN + '/'）。"""
        return self.public_origin + "/"

    @property
    def oauth_redirect_uri(self) -> str:
        return self.public_origin + "/auth/callback"

    @property
    def delegate_redirect_uri(self) -> str:
        return self.public_origin + "/auth/google-delegate/callback"

    @property
    def db_sheet_tz(self) -> ZoneInfo:
        return ZoneInfo(self.db_sheet_timezone)


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()  # type: ignore[call-arg]
