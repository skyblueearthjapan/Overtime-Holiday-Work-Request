"""受け入れテスト用のログインのセッションを作る・消す（ステージングだけ。実装仕様 9-2・12-1）。

Google のログインを通らずに、BOOTSTRAP_ADMINS の人のセッションを DB に 1 行作り、クッキーの値を出す。
tools/qa/acceptance.py が --mint-session で中から使う。コンテナの外から acceptance.py を動かすときは、
これで作ったクッキーの値を --cookie に渡す。

使い方（VPS。tools はイメージに入っていないので付けて動かす＝罠 5）:

  sg docker -c "docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro \
      api python tools/qa/staging_session.py mint [--email owner@…] [--hours 2]"
      → 1 行目 cookie=<クッキーの値>、2 行目 token_hash=<消すための値>、3 行目 expires_at=<期限 UTC>
  … staging_session.py revoke --token-hash <値>      このセッションだけ消す
  … staging_session.py purge                          受け入れテストで作ったセッションを全部消す

決まり:
- mint は APP_ENV=staging のときだけ動く。production では何もせずに終了コード 2。
- メールは BOOTSTRAP_ADMINS の人だけ（既定は 1 人目）。ほかのメールは拒む。
- 期限は作ってから --hours 時間（既定 2、最大 12）。使うたびに延びる仕組み（auth の 90 日）があっても、
  created_at を前にずらして「ログインから 90 日」の上限が同じ時刻になるようにするので、期限より先には延びない。
- 行の name は "acceptance-test"。purge はこの印の行だけを消す（本物のログインは消さない）。
"""
from __future__ import annotations

import argparse
import secrets
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

# 手元（リポジトリ）で動かすときは vps-app/api を読めるようにする。コンテナの中は PYTHONPATH=/app。
_API_DIR = Path(__file__).resolve().parents[2] / "api"
if _API_DIR.is_dir() and str(_API_DIR) not in sys.path:
    sys.path.insert(0, str(_API_DIR))

SESSION_NAME = "acceptance-test"
MAX_HOURS = 12


class Refused(RuntimeError):
    """作ってはいけない（production・管理者でないメールなど）。"""


@dataclass(frozen=True)
class MintedSession:
    email: str
    cookie_value: str
    token_hash: str
    expires_at: datetime

    @property
    def cookie_header(self) -> str:
        from app.auth import SESSION_COOKIE

        return f"{SESSION_COOKIE}={self.cookie_value}"


def ensure_staging(settings) -> None:
    if getattr(settings, "app_env", "") != "staging":
        raise Refused(f"APP_ENV が staging ではありません（{getattr(settings, 'app_env', '')}）。"
                      "受け入れテストのセッションはステージングでだけ作ります。")


def pick_email(settings, email: str | None) -> str:
    admins = list(settings.bootstrap_admin_list)
    if not admins:
        raise Refused("BOOTSTRAP_ADMINS が空です。")
    target = (email or admins[0]).strip().lower()
    if target not in admins:
        raise Refused(f"BOOTSTRAP_ADMINS にないメールです: {target}")
    return target


def mint(db, settings, email: str | None = None, hours: float = 2, now: datetime | None = None) -> MintedSession:
    """セッションを 1 行作って commit し、クッキーの値を返す。"""
    from app import auth, timeutil
    from app.models import UserSession

    ensure_staging(settings)
    if not (0 < float(hours) <= MAX_HOURS):
        raise Refused(f"--hours は 0 より大きく {MAX_HOURS} 以下にしてください: {hours}")
    target = pick_email(settings, email)
    moment = now or timeutil.now_utc()
    expires_at = moment + timedelta(hours=float(hours))
    # auth は使うたびに min(今＋無操作の期限, created_at＋SESSION_MAX_AGE_DAYS) まで延ばす。
    # created_at を前にずらし、後ろの上限を expires_at に合わせる（期限より先には延びない）。
    created_at = expires_at - timedelta(days=settings.session_max_age_days)
    token = secrets.token_urlsafe(32)
    token_hash = auth._token_hash(token)
    db.add(UserSession(token_hash=token_hash, email=target, name=SESSION_NAME, created_at=created_at,
                       expires_at=expires_at))
    db.commit()
    cookie_value = auth._serializer(settings, "session").dumps(token)
    return MintedSession(email=target, cookie_value=cookie_value, token_hash=token_hash, expires_at=expires_at)


def revoke(db, token_hash: str) -> int:
    from sqlalchemy import delete

    from app.models import UserSession

    n = int(db.execute(delete(UserSession).where(UserSession.token_hash == token_hash)).rowcount or 0)
    db.commit()
    return n


def purge(db) -> int:
    from sqlalchemy import delete

    from app.models import UserSession

    n = int(db.execute(delete(UserSession).where(UserSession.name == SESSION_NAME)).rowcount or 0)
    db.commit()
    return n


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="受け入れテスト用のセッション（staging だけ）")
    sub = parser.add_subparsers(dest="cmd", required=True)
    p_mint = sub.add_parser("mint", help="セッションを作り、クッキーの値を出す")
    p_mint.add_argument("--email", default=None, help="BOOTSTRAP_ADMINS の中のメール（既定は 1 人目）")
    p_mint.add_argument("--hours", type=float, default=2.0, help=f"期限（時間。最大 {MAX_HOURS}）")
    p_rev = sub.add_parser("revoke", help="mint で作ったセッションを消す")
    p_rev.add_argument("--token-hash", required=True)
    sub.add_parser("purge", help="受け入れテストのセッションを全部消す")
    args = parser.parse_args(argv)

    from app.config import get_settings
    from app.db import get_sessionmaker

    settings = get_settings()
    db = get_sessionmaker()()
    try:
        if args.cmd == "mint":
            try:
                m = mint(db, settings, args.email, args.hours)
            except Refused as exc:
                print("REFUSED: " + str(exc), file=sys.stderr)
                return 2
            print("cookie=" + m.cookie_value)
            print("token_hash=" + m.token_hash)
            print("expires_at=" + m.expires_at.isoformat())
            print(f"email={m.email}", file=sys.stderr)
            return 0
        if args.cmd == "revoke":
            print(f"revoked={revoke(db, args.token_hash)}")
            return 0
        print(f"purged={purge(db)}")
        return 0
    finally:
        db.close()


if __name__ == "__main__":
    sys.exit(main())
