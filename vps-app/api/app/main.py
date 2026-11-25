"""FastAPI の入口。/api/rpc/{fn}、/api/sebastian、/auth/*、/healthz、画面（app.pages。担当 U）。

起動時に設定を読み（欠けていれば起動しない）、安全装置を確かめ、業務関数（app.domain.register。担当 W）を登録する。
起動: `uvicorn --factory app.main:create_app`
"""
from __future__ import annotations

import importlib
import logging
from datetime import timedelta

from fastapi import Depends, FastAPI, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import Response
from sqlalchemy import func, select, text
from sqlalchemy.orm import Session

from app import auth, google_delegate, permissions, rpc, safety, timeutil
from app.config import Settings, get_settings
from app.db import get_db

log = logging.getLogger(__name__)

SYNC_DELAY_ALERT = timedelta(minutes=5)


def _optional_router(module_name: str):
    """担当 U・W のモジュール（まだ無ければ None）。"""
    try:
        module = importlib.import_module(module_name)
    except ModuleNotFoundError as exc:
        if exc.name == module_name or (exc.name and module_name.startswith(exc.name + ".")):
            return None
        raise
    return module


def health_detail(db: Session) -> dict:
    """管理者だけが見る /healthz の詳細（行の反映の遅れ・GAS の照合）。"""
    from app.models import SheetSyncState, SyncRun

    now = timeutil.now_utc()
    pending = select(SheetSyncState).where(SheetSyncState.desired_version > SheetSyncState.synced_version)
    count = db.scalar(select(func.count()).select_from(pending.subquery())) or 0
    oldest = db.scalar(select(func.min(SheetSyncState.first_dirty_at)).where(
        SheetSyncState.desired_version > SheetSyncState.synced_version))
    blocked = db.scalar(select(func.count()).select_from(SheetSyncState).where(SheetSyncState.state == "blocked")) or 0
    ping = db.scalars(select(SyncRun).where(SyncRun.kind == "gas_ping").order_by(SyncRun.started_at.desc())
                      .limit(1)).first()
    oldest_age = int((now - oldest).total_seconds()) if oldest else 0
    return {
        "rowSync": {"pending": int(count), "blocked": int(blocked), "oldestSeconds": oldest_age,
                    "delayed": bool(oldest and now - oldest > SYNC_DELAY_ALERT)},
        "gasPing": None if ping is None else {"at": ping.started_at, "ok": ping.ok,
                                              "problems": (ping.summary_json or {}).get("problems", [])},
    }


def create_app(load_handlers: bool = True) -> FastAPI:
    settings = get_settings()  # fail-closed: 必須の値が欠けていればここで止まる
    safety.verify_startup(settings)  # 書き込み先が許可リストを通らなければ止まる
    logging.basicConfig(level=logging.INFO)
    log.info("starting overtime api (APP_ENV=%s)", settings.app_env)

    app = FastAPI(title="overtime", docs_url=None, redoc_url=None, openapi_url=None)

    @app.exception_handler(auth.AuthRequired)
    async def _auth_required(request: Request, exc: auth.AuthRequired) -> Response:
        return rpc.json_response({"ok": False, "error": auth.ERR_LOGIN_REQUIRED, "login": auth.LOGIN_URL}, 401)

    @app.get("/healthz")
    def healthz(request: Request, detail: int = 0, db: Session = Depends(get_db),
                s: Settings = Depends(get_settings)) -> Response:
        # 外には {"ok", "env"} だけ。詳細はログインした管理者が ?detail=1 で見る（9-1 の表の下）。
        try:
            db.execute(text("SELECT 1"))
        except Exception:  # noqa: BLE001
            log.exception("healthz: db check failed")
            return rpc.json_response({"ok": False}, 503)
        body: dict = {"ok": True, "env": s.app_env}
        if detail:
            user = auth.get_optional_user(request, db, s)
            if user is not None and permissions.is_admin(db, user.email, s):
                body["detail"] = health_detail(db)
        return rpc.json_response(body)

    @app.post("/api/sebastian")
    async def sebastian(request: Request, db: Session = Depends(get_db)) -> Response:
        """Sebastian の入口（実装仕様 4-4）。ログイン・Origin・冪等キー・Content-Type は確かめない。いつも 200。"""
        raw = await request.body()
        return await run_in_threadpool(_sebastian_call, db, raw)

    app.include_router(auth.router)
    app.include_router(google_delegate.router)
    app.include_router(rpc.router)

    pages = _optional_router("app.pages")
    if pages is not None and hasattr(pages, "router"):
        app.include_router(pages.router)
    else:
        log.warning("app.pages がまだありません（画面は出ません）")

    if load_handlers:
        log.info("rpc handlers: %s", rpc.load_domain_handlers())
    return app


def _sebastian_call(db: Session, raw: bytes) -> Response:
    module = _optional_router("app.domain.sebastian")
    if module is None:
        return rpc.json_response({"error": "service unavailable"})
    try:
        # handle は自分で commit する（仕事を積んで確定してから、その結果を待つため）
        result = module.handle(db, raw)
        db.commit()
    except Exception:  # noqa: BLE001 - Sebastian にはいつも 200 の JSON
        db.rollback()
        log.exception("sebastian: handle failed")
        result = {"error": "internal error"}
    return rpc.json_response(result if isinstance(result, dict) else {"error": "internal error"})
