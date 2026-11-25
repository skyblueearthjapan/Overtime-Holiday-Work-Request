"""画面から呼ぶ関数（RPC）の登録と呼び出し（実装仕様 4-1）。`POST /api/rpc/{fn}`、本文は引数の JSON 配列。

入口で確かめること: Origin が PUBLIC_ORIGIN（無ければ Sec-Fetch-Site: same-origin）、Content-Type: application/json、
fn が閉じたリスト RPC_FUNCTIONS（38 関数）にある、本文が配列。

応答（gas-shim.js の取り決め。実装仕様 3-4）:
- 成功        200 {"ok": true, "result": R}（R が null・{error}・{ok:false,…} でもそのまま）
- 業務エラー   400 {"ok": false, "error": "GAS と同じ文"}（RpcError を raise する）
- 未ログイン   401 {"ok": false, "error": "ログインが必要です。", "login": "/auth/login"}（main.py）
- 管理者だけ   403 {"ok": false, "error": "総務部（管理者）権限がありません。"}
- 別サイト     403、JSON でない 415、知らない関数 404、まだ登録していない関数 501
- 冪等キーが処理中 409、想定外のエラー 500（詳細はログだけ）

業務関数の形（担当 W）:

    from app.rpc import RpcContext, RpcError, register

    @register("api_markOvertimeDone")
    def api_markOvertimeDone(ctx: RpcContext, request_id, manual_end_time=None):
        ...                       # ctx.db・ctx.user.email（小文字）・ctx.now（UTC。1 回の呼び出しで同じ）・ctx.settings
        raise RpcError("本人のみ操作可能です。")

- 引数は GAS と同じ並びの位置引数。省略できる引数には既定値を持たせる。
- 業務関数は commit しない（rpc が冪等キー・監査と一緒に確定する）。
- 管理者だけの関数（ADMIN_ONLY_FUNCTIONS）は、rpc が呼ぶ前に permissions.is_admin を確かめる（403）。
  api_secondaryApproveBatch は GAS と同じく「件数の確かめが先、権限が後」なので、業務関数が
  ctx.require_admin() を呼ぶ（RpcForbidden → 403）。
- 書き込みの関数（WRITE_FUNCTIONS）は X-Idempotency-Key を受け、audit_log を同じトランザクションで残す。
"""
from __future__ import annotations

import inspect
import json
import logging
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

from fastapi import APIRouter, Depends, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import Response
from sqlalchemy import delete
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from app import audit, permissions, timeutil
from app.auth import CurrentUser, require_user
from app.config import Settings, get_settings
from app.db import get_db
from app.models import IdempotencyKey
from app.serialize import json_default

log = logging.getLogger(__name__)

# ===== 閉じたリスト（実装仕様 4-3 の 29 関数 ＋ 4-2 の 9 関数＝38） =====
SCREEN_FUNCTIONS: tuple[str, ...] = (
    "api_getDeptList",
    "api_getWorkersByDept",
    "api_getWorkerInfo",
    "api_getOrderList",
    "api_getHolidayCandidateDates",
    "api_getDashboard",
    "api_submitRequest",
    "api_markOvertimeDone",
    "api_markHolidayStart",
    "api_markHolidayDone",
    "api_editRequest",
    "api_deleteRequest",
    "api_getApproverDepts",
    "api_getApproverDashboard",
    "api_approveRequest",
    "api_rejectRequest",
    "api_approveRequestsBatch",
    "api_approverMonthlySummary",
    "api_getSecondaryApprovalList",
    "api_secondaryApprove",
    "api_secondaryApproveBatch",
    "api_adminDeptOptions",
    "api_adminMonthlySummary",
    "api_adminFiscalTrend",
    "api_adminDailyDetail",
    "api_adminManualComplete",
    "api_adminSpecialClauseCount",
    "api_adminPendingWatch",
    "api_getClosingSummary",
)
OPS_FUNCTIONS: tuple[str, ...] = (
    "ops_getStatus",
    "ops_retryJob",
    "ops_syncSheetConfigNow",
    "ops_runMasterSync",
    "ops_regeneratePdf",
    "ops_batchPdfForDate",
    "ops_rebuildLedger",
    "ops_recalcBreaks",
    "ops_revokeSessions",
)
RPC_FUNCTIONS: tuple[str, ...] = SCREEN_FUNCTIONS + OPS_FUNCTIONS

# 書く関数（冪等キーと監査の対象。実装仕様 4-3 の「書く」と 4-2 の書く関数）
WRITE_FUNCTIONS = frozenset({
    "api_submitRequest",
    "api_markOvertimeDone",
    "api_markHolidayStart",
    "api_markHolidayDone",
    "api_editRequest",
    "api_deleteRequest",
    "api_approveRequest",
    "api_rejectRequest",
    "api_approveRequestsBatch",
    "api_secondaryApprove",
    "api_secondaryApproveBatch",
    "api_adminManualComplete",
    "ops_retryJob",
    "ops_syncSheetConfigNow",
    "ops_runMasterSync",
    "ops_regeneratePdf",
    "ops_batchPdfForDate",
    "ops_rebuildLedger",
    "ops_recalcBreaks",
    "ops_revokeSessions",
})
READ_ONLY_FUNCTIONS = frozenset(RPC_FUNCTIONS) - WRITE_FUNCTIONS

# rpc が呼ぶ前に管理者を確かめる関数（GAS で先頭に assertAdmin_ がある関数と ops_*）
ADMIN_ONLY_FUNCTIONS = frozenset({
    "api_getSecondaryApprovalList",
    "api_secondaryApprove",
    "api_adminDeptOptions",
    "api_adminMonthlySummary",
    "api_adminFiscalTrend",
    "api_adminDailyDetail",
    "api_adminManualComplete",
    "api_adminSpecialClauseCount",
    "api_adminPendingWatch",
    *OPS_FUNCTIONS,
})
# 業務関数の中で ctx.require_admin() を呼ぶ関数（GAS の確かめの順を保つため）
ADMIN_CHECKED_IN_HANDLER = frozenset({"api_secondaryApproveBatch"})

# 第 1 引数が申請 ID の書く関数（監査の request_ids）
_FIRST_ARG_REQUEST_ID = frozenset({
    "api_markOvertimeDone", "api_markHolidayStart", "api_markHolidayDone", "api_editRequest", "api_deleteRequest",
    "api_approveRequest", "api_rejectRequest", "api_secondaryApprove", "api_adminManualComplete", "ops_regeneratePdf",
})
_FIRST_ARG_REQUEST_IDS = frozenset({"api_approveRequestsBatch", "api_secondaryApproveBatch"})

IDEMPOTENCY_HEADER = "X-Idempotency-Key"
IDEMPOTENCY_TTL = timedelta(hours=24)
_IDEMPOTENCY_KEY_RE = re.compile(r"^[A-Za-z0-9._:-]{8,128}\Z")

ERR_SHAPE = "呼び出しの形が正しくありません。"
ERR_FORBIDDEN_ORIGIN = "この画面からは操作できません。"
ERR_ARGS = "引数の数が正しくありません。"
ERR_INTERNAL = "サーバーでエラーが起きました。時間をおいてもう一度お試しください。"
ERR_IN_PROGRESS = "同じ操作を処理しています。少し待ってから画面を読み込み直してください。"
ERR_KEY_REUSED = "同じ操作キーが別の操作に使われています。画面を読み込み直してください。"
ERR_UNKNOWN_FN = "その操作はありません。"
ERR_NOT_ADMIN = "総務部（管理者）権限がありません。"


class RpcError(Exception):
    """業務エラー（400）。message は画面にそのまま出る（GAS の throw new Error('...') と同じ文にする）。"""

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


class RpcForbidden(RpcError):
    """管理者だけの操作を管理者以外が呼んだ（403）。"""

    def __init__(self, message: str = ERR_NOT_ADMIN):
        super().__init__(message)


class RpcNotImplemented(Exception):
    pass


@dataclass
class RpcContext:
    """業務関数に渡す。

    - db: この呼び出しのセッション（commit しない）
    - user: CurrentUser(email, name)。email は Google ログインのメール（小文字・空にならない）
    - now: この呼び出しの時刻（UTC。1 回の呼び出しの中で同じ値を使う）
    - settings: app.config.Settings
    - fn: 呼ばれた関数名
    """

    db: Session
    user: CurrentUser
    settings: Settings
    now: datetime = field(default_factory=lambda: timeutil.now_utc())
    request: Request | None = None
    fn: str = ""
    _is_admin: bool | None = field(default=None, repr=False)

    @property
    def is_admin(self) -> bool:
        if self._is_admin is None:
            self._is_admin = permissions.is_admin(self.db, self.user.email, self.settings)
        return self._is_admin

    def require_admin(self) -> str:
        """管理者でなければ RpcForbidden（403）。GAS の assertAdmin_ と同じくメールを返す。"""
        if not self.is_admin:
            raise RpcForbidden()
        return self.user.email


Handler = Callable[..., Any]


def _not_implemented(ctx: RpcContext, *args: Any) -> Any:
    raise RpcNotImplemented()


_registry: dict[str, Handler] = {}


def reset_registry() -> None:
    """すべてを「まだ作っていない（501）」に戻す。"""
    _registry.clear()
    _registry.update({name: _not_implemented for name in RPC_FUNCTIONS})


reset_registry()


def register(name: str, handler: Handler | None = None):
    """関数を登録する。38 関数以外の名前は ValueError。デコレーター `@register("api_x")` としても使える。"""
    if name not in RPC_FUNCTIONS:
        raise ValueError(f"公開できない関数名です: {name}")

    def _do(h: Handler) -> Handler:
        _registry[name] = h
        return h

    return _do(handler) if handler is not None else _do


def get_handler(name: str) -> Handler | None:
    return _registry.get(name)


def is_implemented(name: str) -> bool:
    return _registry.get(name, _not_implemented) is not _not_implemented


def implemented_functions() -> list[str]:
    return [name for name in RPC_FUNCTIONS if is_implemented(name)]


def load_domain_handlers() -> list[str]:
    """担当 W の app.domain.register.register_all(register) を呼ぶ（無ければ何もしない＝501 のまま）。"""
    try:
        from app.domain import register as domain_register  # type: ignore[attr-defined]
    except ModuleNotFoundError as exc:
        if exc.name in ("app.domain", "app.domain.register"):
            log.warning("rpc: app.domain.register がまだありません（全関数 501）")
            return []
        raise
    domain_register.register_all(register)
    return implemented_functions()


# ===== JSON =====


def dumps(body: Any) -> str:
    return json.dumps(body, ensure_ascii=False, default=json_default, separators=(",", ":"))


def json_response(body: dict[str, Any] | str, status_code: int = 200) -> Response:
    content = body if isinstance(body, str) else dumps(body)
    return Response(
        content=content,
        status_code=status_code,
        media_type="application/json; charset=utf-8",
        headers={"Cache-Control": "no-store"},
    )


def error_response(message: str, status_code: int) -> Response:
    return json_response({"ok": False, "error": message}, status_code)


# ===== 呼び出し =====


def _call(ctx: RpcContext, handler: Handler, args: list[Any]) -> Any:
    try:
        bound = inspect.signature(handler).bind(ctx, *args)
    except TypeError as exc:
        raise RpcError(ERR_ARGS) from exc
    return handler(*bound.args, **bound.kwargs)


def _request_ids(fn: str, args: list[Any], result: Any = None) -> list[str]:
    ids: list[str] = []
    if args:
        first = args[0]
        if fn in _FIRST_ARG_REQUEST_ID and isinstance(first, (str, int)):
            ids.append(str(first))
        elif fn in _FIRST_ARG_REQUEST_IDS and isinstance(first, list):
            ids.extend(str(x) for x in first if isinstance(x, (str, int)))
    if isinstance(result, dict) and isinstance(result.get("requestId"), str) and result["requestId"] not in ids:
        ids.append(result["requestId"])
    return ids


def _replay(db: Session, user_email: str, key: str, fn: str) -> Response | None:
    row = db.get(IdempotencyKey, (user_email, key))
    if row is None:
        return None
    if row.fn != fn:
        return error_response(ERR_KEY_REUSED, 400)
    if row.state != "done":
        return error_response(ERR_IN_PROGRESS, 409)
    return json_response(row.response_json)


def purge_idempotency_keys(db: Session, now: datetime | None = None) -> int:
    """24 時間より古い冪等キーを消す（worker の purge）。消した件数を返す。commit は呼び出し側。"""
    cutoff = (now or timeutil.now_utc()) - IDEMPOTENCY_TTL
    return int(db.execute(delete(IdempotencyKey).where(IdempotencyKey.created_at < cutoff)).rowcount or 0)


def _audit_failure(db: Session, user: CurrentUser, fn: str, args: list[Any], message: str) -> None:
    """失敗した書き込みの記録（業務の書き込みを戻した後、別のトランザクションで）。"""
    try:
        audit.record(db, user.email, fn, _request_ids(fn, args), ok=False, error=message)
        db.commit()
    except Exception:  # noqa: BLE001 - 記録の失敗で応答を変えない
        db.rollback()
        log.exception("rpc %s: could not write audit_log", fn)


def dispatch(fn: str, args: list[Any], *, db: Session, user: CurrentUser, settings: Settings,
             request: Request | None = None, idempotency_key: str | None = None) -> Response:
    handler = _registry.get(fn)
    if handler is None:
        return error_response(ERR_UNKNOWN_FN, 404)
    ctx = RpcContext(db=db, user=user, settings=settings, now=timeutil.now_utc(), request=request, fn=fn)
    writes = fn in WRITE_FUNCTIONS
    use_key = bool(idempotency_key) and writes
    key_row: IdempotencyKey | None = None

    try:
        if fn in ADMIN_ONLY_FUNCTIONS and not ctx.is_admin:
            db.rollback()
            return error_response(ERR_NOT_ADMIN, 403)
        if use_key:
            replay = _replay(db, user.email, idempotency_key, fn)  # type: ignore[arg-type]
            if replay is not None:
                db.rollback()
                return replay
            key_row = IdempotencyKey(user_email=user.email, key=idempotency_key, fn=fn, state="running",
                                     created_at=ctx.now)
            db.add(key_row)
            try:
                # PostgreSQL では、同じキーを処理中の別の呼び出しがあれば、その確定まで待ってから重複になる
                db.flush()
            except IntegrityError:
                db.rollback()
                replay = _replay(db, user.email, idempotency_key, fn)  # type: ignore[arg-type]
                db.rollback()
                return replay or error_response(ERR_IN_PROGRESS, 409)

        result = _call(ctx, handler, args)
        body = dumps({"ok": True, "result": result})
        if key_row is not None:
            key_row.response_json = body
            key_row.state = "done"
        if writes:
            audit.record(db, user.email, fn, _request_ids(fn, args, result), ok=True)
        db.commit()
        return json_response(body)
    except RpcForbidden as exc:
        db.rollback()
        return error_response(exc.message, 403)
    except RpcError as exc:
        db.rollback()
        if writes:
            _audit_failure(db, user, fn, args, exc.message)
        return error_response(exc.message, 400)
    except RpcNotImplemented:
        db.rollback()
        return error_response(f"この機能はまだ用意されていません。（{fn}）", 501)
    except Exception as exc:  # noqa: BLE001
        log.exception("rpc %s failed", fn)
        db.rollback()
        if writes:
            _audit_failure(db, user, fn, args, f"{type(exc).__name__}")
        return error_response(ERR_INTERNAL, 500)


router = APIRouter()


def check_same_origin(request: Request, settings: Settings) -> bool:
    """Origin があれば PUBLIC_ORIGIN と照合し、無ければブラウザの Sec-Fetch-Site: same-origin を求める。"""
    origin = request.headers.get("origin")
    if origin is not None:
        return origin.rstrip("/") == settings.public_origin
    return request.headers.get("sec-fetch-site") == "same-origin"


@router.post("/api/rpc/{fn}")
async def call_rpc(
    fn: str,
    request: Request,
    db: Session = Depends(get_db),
    settings: Settings = Depends(get_settings),
    user: CurrentUser = Depends(require_user),
) -> Response:
    # 入口の確認: 同じサイトの画面からの JSON の呼び出しだけを受ける（ほかのサイトからの書き込みを止める）
    if not check_same_origin(request, settings):
        return error_response(ERR_FORBIDDEN_ORIGIN, 403)
    content_type = request.headers.get("content-type", "").split(";", 1)[0].strip().lower()
    if content_type != "application/json":
        return error_response(ERR_SHAPE, 415)
    if fn not in _registry:
        return error_response(ERR_UNKNOWN_FN, 404)

    raw = await request.body()
    if raw.strip():
        try:
            args = json.loads(raw)
        except ValueError:
            return error_response(ERR_SHAPE, 400)
    else:
        args = []
    if not isinstance(args, list):
        return error_response(ERR_SHAPE, 400)

    key = request.headers.get(IDEMPOTENCY_HEADER)
    if key is not None and not _IDEMPOTENCY_KEY_RE.match(key):
        return error_response(ERR_SHAPE, 400)

    return await run_in_threadpool(
        dispatch, fn, args, db=db, user=user, settings=settings, request=request, idempotency_key=key
    )
