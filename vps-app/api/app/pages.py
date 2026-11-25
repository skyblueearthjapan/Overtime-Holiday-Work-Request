"""画面の入口 `GET /`（仕様 3-2・3-3。GAS の doGet の代わり）。

- `?page=` で 5 ページを振り分ける（`adminApi.gs:535-545` と同じ。無指定・不明は top）。
- 全ページでログインが要る。未ログインは `/auth/login?next=<今のパスとクエリ>` へ。
- admin・approve2 は管理者だけ、approver は管理者か部署の承認者だけ。だめなら権限エラーの画面（デバッグ情報は出さない）。
- テンプレートは GAS の HTML をそのまま写したもの（`vps-app/web/dev/port-pages.js` が作る）。
  差し込む値は `{{ APP_URL }}`・`{{ PORTAL_URL }}`・`{{ message }}`（自動エスケープ）と、
  JS へ渡す値の JSON の塊 `{{ data_json|safe }}`（`<`・`>`・`&` を \\u 形式にしてあるので `</script>` を閉じられない）。
"""
from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path
from typing import Any
from urllib.parse import quote

import jinja2
from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from sqlalchemy.orm import Session

from app import auth, constants, permissions
from app.config import Settings, get_settings
from app.db import get_db

router = APIRouter()

TEMPLATES_DIR = Path(__file__).resolve().parent / "templates" / "pages"
_env = jinja2.Environment(
    loader=jinja2.FileSystemLoader(str(TEMPLATES_DIR)),
    autoescape=True,
    keep_trailing_newline=True,
)

PAGES = ("top", "approver", "admin", "approve2", "staffSummary")
PAGE_TITLE = "残業・休日出勤申請"
NO_AUTH_TITLE = "権限エラー"
_DATE_RE = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}")

ADMIN_DENIED = (
    "総務部（管理者）権限がありません。\n"
    "ApproverMap シートに role=admin で登録されているか、\n"
    "Settings シートの ADMIN_EMAILS にメールアドレスが含まれているか確認してください。\n"
    "（現在のアカウント: {email}）"
)
APPROVER_DENIED = (
    "承認者権限がありません。\n"
    "DeptApprovers または ApproverMap に登録されているか確認してください。\n"
    "（現在のアカウント: {email}）"
)

_NO_STORE = {"Cache-Control": "no-store"}


def validate_date_param(value: str | None) -> str:
    """`?date=` は `yyyy-MM-dd` の形で、暦の上で正しい日付のときだけ返す。ほかは ''（9-1 #13）。"""
    text = value or ""
    if not _DATE_RE.fullmatch(text):
        return ""
    try:
        date.fromisoformat(text)
    except ValueError:
        return ""
    return text


def json_for_script(data: Any) -> str:
    """`<script type="application/json">` に入れてよい JSON（仕様 3-6）。"""
    text = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return text.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")


def _first(request: Request, name: str) -> str | None:
    """GAS の e.parameter と同じく、同じ名前が複数あれば最初の値。"""
    values = request.query_params.getlist(name)
    return values[0] if values else None


def _render(template: str, status_code: int = 200, **context: Any) -> HTMLResponse:
    html = _env.get_template(template).render(**context)
    return HTMLResponse(html, status_code=status_code, headers=_NO_STORE)


def _common(settings: Settings) -> dict[str, str]:
    return {
        "APP_URL": settings.public_origin.rstrip("/") + "/",
        "PORTAL_URL": settings.portal_url,
    }


def _no_auth(settings: Settings, message: str) -> HTMLResponse:
    return _render("no_auth.html", message=message, **_common(settings))


@router.get("/", include_in_schema=False)
def page_entry(
    request: Request,
    user: Any = Depends(auth.get_optional_user),
    db: Session = Depends(get_db),
    settings: Settings = Depends(get_settings),
) -> Response:
    if user is None:
        here = request.url.path + (("?" + request.url.query) if request.url.query else "")
        target = "/auth/login?next=" + quote(auth.safe_next(here), safe="")
        return RedirectResponse(target, status_code=302, headers=_NO_STORE)

    page = _first(request, "page") or "top"
    if page == "redirect":
        # 旧フォームの入口。フォームは閉じた（9-1 #11）ので TOP へ（Q9 の既定）
        return RedirectResponse("/?page=top", status_code=302, headers=_NO_STORE)
    safe_page = page if page in PAGES else "top"

    email = user.email
    if safe_page in ("admin", "approve2"):
        if not permissions.is_admin(db, email):
            return _no_auth(settings, ADMIN_DENIED.format(email=email))
    elif safe_page == "approver":
        if not (permissions.is_admin(db, email) or permissions.is_any_dept_approver(db, email)):
            return _no_auth(settings, APPROVER_DENIED.format(email=email))

    data = {
        "PAGE": safe_page,
        "DATE_PARAM": validate_date_param(_first(request, "date")),
        "REASONS": list(constants.REASONS),
        "OT_HOURS": list(constants.OT_HOURS),
        "HD_HOURS": list(constants.HD_HOURS),
    }
    return _render(safe_page + ".html", data_json=json_for_script(data), **_common(settings))
