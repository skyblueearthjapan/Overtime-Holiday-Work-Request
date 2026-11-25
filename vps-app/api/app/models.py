"""DB の表（実装仕様 5 章）。担当 F が全表を作る（alembic 0001_initial_schema と同じ形）。

共通:
- 文字の列は NOT NULL DEFAULT ''（空とシートの空を同じに扱う）。
- 日時は UTCDateTime（timestamptz。時差のない値は拒む）。
- enum は CheckConstraint。
- 業務の表（A 区分: requests・work_logs・batch_logs）は row_version（書くたびに +1。sheet_queue.mark_dirty）、
  sheet_row（シートの行番号。まだ書いていない新しい行は NULL）、cell_kinds（型が混ざる列の行ごとの型。6-3）を持つ。
- シートの値は正規化せずに生のまま入れる（GAS が normalize_ するところは業務の側で jsutil.normalize_gas を通す）。
"""
from __future__ import annotations

from datetime import date, datetime
from typing import Any

from sqlalchemy import (
    JSON,
    BigInteger,
    Boolean,
    CheckConstraint,
    Date,
    Index,
    Integer,
    Sequence,
    String,
    Text,
)
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import Mapped, mapped_column

from app.constants import (
    IDEMPOTENCY_STATES,
    JOB_LANES,
    JOB_STATES,
    REQUEST_STATUSES,
    REQUEST_TYPES,
    SYNC_STATES,
)
from app.db import Base, UTCDateTime

# 自動で増える整数の主キー（PostgreSQL は BIGINT、SQLite は INTEGER でないと自動採番されない）
BigIntPK = BigInteger().with_variant(Integer, "sqlite")
# PostgreSQL は jsonb、SQLite は JSON（文字）
JsonB = JSON().with_variant(JSONB(), "postgresql")

# sheet_sync_state.seq（新しい行を作った順）の連番。PostgreSQL だけで使う（SQLite は sheet_queue が max+1）
SHEET_SYNC_SEQ = Sequence("sheet_sync_state_seq", metadata=Base.metadata)


def _in(column: str, values: tuple[str, ...], allow_empty: bool = False) -> str:
    allowed = list(values) + ([""] if allow_empty else [])
    quoted = ", ".join("'" + v.replace("'", "''") + "'" for v in allowed)
    return f"{column} IN ({quoted})"


def _text(**kw: Any):
    return mapped_column(Text, nullable=False, default="", server_default="", **kw)


def _int_null():
    return mapped_column(Integer, nullable=True)


def _row_version():
    return mapped_column(Integer, nullable=False, default=1, server_default="1")


def _sheet_row():
    return mapped_column(Integer, nullable=True)


def _cell_kinds():
    return mapped_column(JsonB, nullable=False, default=dict, server_default="{}")


# ===== 5-1. 業務の表（A 区分。行単位でシートへ反映する） =====


class Request(Base):
    """シート Requests（見出し 1 行目）。列は見出しの文字で対応させる（services/sheet_codec.COLUMN_RULES）。"""

    __tablename__ = "requests"

    request_id: Mapped[str] = mapped_column(String(64), primary_key=True)  # requestId（UUID 36 字）
    request_type: Mapped[str] = _text()  # requestType(overtime/holiday)
    status: Mapped[str] = _text()  # status(submitted/approved/canceled)
    dept: Mapped[str] = _text()
    worker_code: Mapped[str] = _text()
    worker_name: Mapped[str] = _text()
    worker_email: Mapped[str] = _text()
    target_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    target_date_text: Mapped[str] = _text()  # 日付にならない元の文字（gap-3 G3-L2）
    submitted_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    approved_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    approved_by: Mapped[str] = _text()
    approved_minutes: Mapped[int | None] = _int_null()  # 申請時の予定（承認で変わらない）
    reason: Mapped[str] = _text()
    reason_detail: Mapped[str] = _text()
    work_content: Mapped[str] = _text()
    job_id1: Mapped[str] = _text()
    job_id2: Mapped[str] = _text()
    job_id3: Mapped[str] = _text()
    hr_mail_sent_at: Mapped[str] = _text()
    export_error: Mapped[str] = _text()
    work_no1: Mapped[str] = _text()
    work_no2: Mapped[str] = _text()
    work_no3: Mapped[str] = _text()
    order_no1: Mapped[str] = _text()
    order_no2: Mapped[str] = _text()
    order_no3: Mapped[str] = _text()
    customer1: Mapped[str] = _text()
    customer2: Mapped[str] = _text()
    customer3: Mapped[str] = _text()
    product1: Mapped[str] = _text()
    product2: Mapped[str] = _text()
    product3: Mapped[str] = _text()
    pdf_file_id: Mapped[str] = _text()
    pdf_folder_id: Mapped[str] = _text()
    pdf_generated_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    approved_by2: Mapped[str] = _text()
    approved_at2: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    modified_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    modified_by: Mapped[str] = _text()
    # PDF の 3 列を空にした操作の、まだゴミ箱へ移していない前の ID（7 章の仕事が使う）
    old_pdf_file_id_pending: Mapped[str] = _text()
    # ステージングへの移行で空にした、本番の PDF の 3 列のうち 2 列の元の値（参照用。仕事は使わない。
    # ステージングの GAS が本番の PDF を触らないようにするため。データ安全 高-1）。シートには書かない
    source_pdf_file_id: Mapped[str] = _text()
    source_pdf_folder_id: Mapped[str] = _text()

    row_version: Mapped[int] = _row_version()
    sheet_row: Mapped[int | None] = _sheet_row()
    cell_kinds: Mapped[dict] = _cell_kinds()

    __table_args__ = (
        CheckConstraint(_in("request_type", REQUEST_TYPES), name="ck_requests_request_type"),
        CheckConstraint(_in("status", REQUEST_STATUSES), name="ck_requests_status"),
        Index("ix_requests_status", "status"),
        Index("ix_requests_target_date", "target_date"),
        Index("ix_requests_dept", "dept"),
    )


class WorkLog(Base):
    """シート WorkLogs（1 行目は説明、見出しは 2 行目、データは 3 行目から）。

    Requests と 1 対 1 ではない（孤児の行・行の無い申請がある）ので外部キーを張らない。
    同じ requestId が複数あれば sheet_row の小さい行を使う（GAS の findWorkLogRowNo_ と同じ）。
    """

    __tablename__ = "work_logs"

    id: Mapped[int] = mapped_column(BigIntPK, primary_key=True, autoincrement=True)
    request_id: Mapped[str] = mapped_column(String(64), nullable=False, default="", server_default="", index=True)
    # GAS が読んで返す JST の 'yyyy-MM-dd HH:mm:ss'（日時のセルだった行も文字にして持ち、cell_kinds に元の型を残す）
    actual_start_at: Mapped[str] = _text()
    actual_end_at: Mapped[str] = _text()
    actual_minutes: Mapped[int | None] = _int_null()  # 空は NULL（0 と区別する）
    break_minutes: Mapped[int | None] = _int_null()
    net_minutes: Mapped[int | None] = _int_null()
    updated_at: Mapped[str] = _text()
    updated_by: Mapped[str] = _text()

    row_version: Mapped[int] = _row_version()
    sheet_row: Mapped[int | None] = _sheet_row()
    cell_kinds: Mapped[dict] = _cell_kinds()


class BatchLog(Base):
    """シート BatchLogs（見出し 1 行目: 実行日時・バッチ名・対象日・成功・スキップ・失敗・エラー詳細）。"""

    __tablename__ = "batch_logs"

    id: Mapped[int] = mapped_column(BigIntPK, primary_key=True, autoincrement=True)
    executed_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)  # 実行日時
    batch_name: Mapped[str] = _text()  # バッチ名
    target_date: Mapped[date | None] = mapped_column(Date, nullable=True)  # 対象日
    target_date_text: Mapped[str] = _text()
    success: Mapped[int | None] = _int_null()  # 成功
    skip: Mapped[int | None] = _int_null()  # スキップ
    fail: Mapped[int | None] = _int_null()  # 失敗
    error_detail: Mapped[str] = _text()  # エラー詳細（複数は \n でつなぐ）

    row_version: Mapped[int] = _row_version()
    sheet_row: Mapped[int | None] = _sheet_row()
    cell_kinds: Mapped[dict] = _cell_kinds()


# ===== 5-2. シートを正とする表（B 区分。VPS は書かない。取り込むだけ） =====


class Setting(Base):
    """Settings（1 行目説明・2 行目見出し・3 行目から A=key、B=value）。同じキーは下の行が勝つ。"""

    __tablename__ = "settings"

    key: Mapped[str] = mapped_column(String(256), primary_key=True)  # 前後の空白を取る
    value: Mapped[str] = _text()
    sheet_row: Mapped[int | None] = _sheet_row()


class DeptApprover(Base):
    """DeptApprovers（見出し 1 行目。データ行だけを持つ＝sheet_row は 2 から）。"""

    __tablename__ = "dept_approvers"

    sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    dept: Mapped[str] = _text()  # A
    emails: Mapped[str] = _text()  # B（生の文字）
    name: Mapped[str] = _text()  # C
    col_d: Mapped[str] = _text()  # D 役職（持つだけ）
    col_e: Mapped[str] = _text()  # E 通知メール（持つだけ）


class ApproverMapRow(Base):
    """ApproverMap（1 行目から生のまま）。permissions は GAS と同じく 1 行目を見出しとして読む。"""

    __tablename__ = "approver_map_rows"

    sheet_row: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    cells: Mapped[list] = mapped_column(JsonB, nullable=False, default=list, server_default="[]")


# ===== 5-3. マスタの表（C 区分。マスタ元から作る） =====


class Worker(Base):
    """作業員マスタ。"""

    __tablename__ = "workers"

    sort_order: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    worker_code: Mapped[str] = _text()
    worker_name: Mapped[str] = _text()
    dept: Mapped[str] = _text()
    active_flag: Mapped[str] = _text()
    google_account: Mapped[str] = _text()  # 「Googleアカウント」で始まる列（無ければ空）
    raw: Mapped[dict] = mapped_column(JsonB, nullable=False, default=dict, server_default="{}")


class Order(Base):
    """工番マスタ。"""

    __tablename__ = "orders"

    sort_order: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    code: Mapped[str] = _text()
    customer: Mapped[str] = _text()  # 受注先
    dest: Mapped[str] = _text()  # 納入先
    dest_address: Mapped[str] = _text()
    product: Mapped[str] = _text()
    quantity: Mapped[str] = _text()
    imported_at: Mapped[str] = _text()


class CompanyCalendar(Base):
    """社内カレンダーマスタ（A〜D 列を位置で）。"""

    __tablename__ = "company_calendar"

    sort_order: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=False)
    date_text: Mapped[str] = _text()  # A
    kind: Mapped[str] = _text()  # B
    weekday: Mapped[str] = _text()  # C（文字のまま）
    note: Mapped[str] = _text()  # D


# ===== 5-4. 土台の表 =====


class UserSession(Base):
    """ログインのセッション。クッキーには署名付きの乱数、DB には その SHA-256 だけ。"""

    __tablename__ = "sessions"

    token_hash: Mapped[str] = mapped_column(String(64), primary_key=True)
    email: Mapped[str] = mapped_column(String(256), nullable=False, index=True)
    name: Mapped[str] = _text()
    created_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    # min(最後に使った時刻 + SESSION_IDLE_HOURS, created_at + SESSION_MAX_AGE_DAYS)。90 日（9-1 #24）
    expires_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False, index=True)


class GoogleCredential(Base):
    """委任トークン（DELEGATE_EMAIL の Sheets・Drive の権限）。リフレッシュトークンは Fernet で暗号化。"""

    __tablename__ = "google_credentials"

    name: Mapped[str] = mapped_column(String(32), primary_key=True)  # 'delegate'
    account_email: Mapped[str] = mapped_column(String(256), nullable=False, default="", server_default="")
    refresh_token_encrypted: Mapped[str] = mapped_column(Text, nullable=False)
    scopes: Mapped[str] = _text()  # 空白区切り（実際に許可されたスコープ）
    granted_by_email: Mapped[str] = mapped_column(String(256), nullable=False, default="", server_default="")
    granted_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)


class IdempotencyKey(Base):
    """二重押し対策（9-1 #5）。同じ人の同じキーの 2 回目は、1 回目の結果を返す。24 時間で消す。"""

    __tablename__ = "idempotency_keys"

    user_email: Mapped[str] = mapped_column(String(256), primary_key=True)
    key: Mapped[str] = mapped_column(String(128), primary_key=True)
    fn: Mapped[str] = mapped_column(String(64), nullable=False)
    state: Mapped[str] = mapped_column(String(16), nullable=False, default="running", server_default="running")
    response_json: Mapped[str] = _text()
    created_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False, index=True)

    __table_args__ = (CheckConstraint(_in("state", IDEMPOTENCY_STATES), name="ck_idempotency_keys_state"),)


class Job(Base):
    """後回しの仕事（outbox）。業務の書き込みと同じトランザクションで積む（app/jobs.py）。

    lane: 'gas'（GAS の内部用の入口を呼ぶ。全体で同時に 1 本）／'default'。
    requires_json: [{table, pk, version}]。この版まで行の反映が済むまで動かさない。
    guard_json: 積んだ時点の status など（担当 M が実行前に照合する）。
    """

    __tablename__ = "jobs"

    id: Mapped[int] = mapped_column(BigIntPK, primary_key=True, autoincrement=True)
    kind: Mapped[str] = mapped_column(String(64), nullable=False)
    lane: Mapped[str] = mapped_column(String(16), nullable=False, default="gas", server_default="gas")
    priority: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    request_id: Mapped[str] = mapped_column(String(64), nullable=False, default="", server_default="")
    args_json: Mapped[str] = mapped_column(Text, nullable=False, default="{}", server_default="{}")
    requires_json: Mapped[str] = mapped_column(Text, nullable=False, default="[]", server_default="[]")
    guard_json: Mapped[str] = mapped_column(Text, nullable=False, default="{}", server_default="{}")
    state: Mapped[str] = mapped_column(String(16), nullable=False, default="pending", server_default="pending")
    attempts: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    run_after: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    leased_until: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    last_error: Mapped[str] = _text()
    result_json: Mapped[str] = _text()
    batch_id: Mapped[str] = mapped_column(String(64), nullable=False, default="", server_default="")
    created_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    finished_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)

    __table_args__ = (
        CheckConstraint(_in("state", JOB_STATES), name="ck_jobs_state"),
        CheckConstraint(_in("lane", JOB_LANES), name="ck_jobs_lane"),
        Index("ix_jobs_lane_state_run_after", "lane", "state", "run_after"),
        Index("ix_jobs_request_id_state", "request_id", "state"),
        Index("ix_jobs_batch_id", "batch_id"),
    )


class SheetSyncState(Base):
    """行の反映の状態（決定 A。実装仕様 6-2）。desired_version > synced_version の行を worker が書く。"""

    __tablename__ = "sheet_sync_state"

    table_name: Mapped[str] = mapped_column(String(32), primary_key=True)
    pk: Mapped[str] = mapped_column(String(64), primary_key=True)
    desired_version: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    synced_version: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    # 最後にシートへ書いた、またはシートにあると分かっている値（列ごと）
    synced_values: Mapped[dict] = mapped_column(JsonB, nullable=False, default=dict, server_default="{}")
    state: Mapped[str] = mapped_column(String(16), nullable=False, default="ok", server_default="ok")
    attempts: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    next_try_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    last_error: Mapped[str] = _text()
    first_dirty_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    last_synced_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    # 新しい行を作った順（PostgreSQL は sheet_sync_state_seq、SQLite は max+1。sheet_queue が振る）
    seq: Mapped[int] = mapped_column(BigInteger, nullable=False, default=0, server_default="0")

    __table_args__ = (
        CheckConstraint(_in("state", SYNC_STATES), name="ck_sheet_sync_state_state"),
        Index("ix_sheet_sync_state_pending", "state", "next_try_at"),
        Index("ix_sheet_sync_state_seq", "seq"),
    )


class SyncRun(Base):
    """定期処理・同期・照合・gas_ping の記録。worker は kind='schedule:<名前>' で予定の枠を判定する。"""

    __tablename__ = "sync_runs"

    id: Mapped[int] = mapped_column(BigIntPK, primary_key=True, autoincrement=True)
    kind: Mapped[str] = mapped_column(String(64), nullable=False)
    started_at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False)
    finished_at: Mapped[datetime | None] = mapped_column(UTCDateTime, nullable=True)
    ok: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="false")
    summary_json: Mapped[dict] = mapped_column(JsonB, nullable=False, default=dict, server_default="{}")

    __table_args__ = (Index("ix_sync_runs_kind_started", "kind", "started_at"),)


class AuditLog(Base):
    """書き込みの RPC の記録（rpc.py が業務の書き込みと同じトランザクションで残す。消さない）。"""

    __tablename__ = "audit_log"

    id: Mapped[int] = mapped_column(BigIntPK, primary_key=True, autoincrement=True)
    at: Mapped[datetime] = mapped_column(UTCDateTime, nullable=False, index=True)
    user_email: Mapped[str] = mapped_column(String(256), nullable=False, default="", server_default="")
    fn: Mapped[str] = mapped_column(String(64), nullable=False)
    request_ids: Mapped[str] = mapped_column(Text, nullable=False, default="[]", server_default="[]")  # JSON 配列
    ok: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="false")
    error: Mapped[str] = _text()


# sheet_queue が扱う A 区分の表（table_name → モデル）
SYNCED_TABLES: dict[str, type[Base]] = {
    "requests": Request,
    "work_logs": WorkLog,
    "batch_logs": BatchLog,
}

__all__ = [
    "ApproverMapRow",
    "AuditLog",
    "BatchLog",
    "CompanyCalendar",
    "DeptApprover",
    "GoogleCredential",
    "IdempotencyKey",
    "Job",
    "Order",
    "Request",
    "SHEET_SYNC_SEQ",
    "SYNCED_TABLES",
    "Setting",
    "SheetSyncState",
    "SyncRun",
    "UserSession",
    "WorkLog",
    "Worker",
]
