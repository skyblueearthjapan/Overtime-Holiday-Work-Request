"""worker（実装仕様 10-3）。docker compose の worker サービスで `python -m app.worker`。

3 本のスレッドを動かす（スレッドごとに DB のセッションを持つ。SIGTERM で止まる）:
  ① 行の反映のレーン: services.row_sync.run_once(session_factory, limit=50) を 2 秒ごと（担当 D）
  ② GAS のレーン: lane='gas' の仕事を 1 本ずつ（全体で 1 本。PostgreSQL では pg_try_advisory_lock でも守る）
  ③ 予定（register_schedule で登録したもの）と lane='default' の仕事

予定は sync_runs（kind='schedule:<名前>'）で枠を判定する。枠を始めるときに行を書いて確定してから動かすので、
同じ枠では 1 回だけ（失敗しても次の枠まで待つ。worker を再起動しても同じ枠で動き直さない＝起動の直後に
走らせ直さない）。毎日の予定は「時刻 at から window_minutes 分の間」が枠で、その外で起動しても動かない。

登録の口（担当 D・M）: 下の REGISTRATION_MODULES のモジュールに `register_worker()` を書き、その中で
jobs.register_kind(...) と worker.register_schedule(...) を呼ぶ。モジュールがまだ無ければ飛ばす。

手で動かす: `python -m app.worker run <名前>`（例 `run sync_masters`、`run sheet_reconcile --check`）。
`--check` は予定の関数に check_only=True を渡す。
"""
from __future__ import annotations

import importlib
import inspect
import logging
import signal
import sys
import threading
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import select, text
from sqlalchemy.orm import Session

from app import timeutil
from app.models import SyncRun

log = logging.getLogger(__name__)

SCHEDULE_PREFIX = "schedule:"
ROW_SYNC_SECONDS = 2.0
GAS_LANE_SECONDS = 1.0
DEFAULT_LANE_SECONDS = 2.0
GAS_LANE_ADVISORY_LOCK = 739_310_001  # pg_try_advisory_lock の鍵（GAS のレーンを全体で 1 本に）

# 担当 D・M が register_worker() を置くモジュール
REGISTRATION_MODULES: tuple[str, ...] = (
    "app.services.row_sync",
    "app.services.reconcile",
    "app.services.config_sync",
    "app.services.master_sync",
    "app.services.gas_jobs",
    "app.services.morning_batch",
    "app.services.ops_status",
    "app.services.sheet_intake",
)

TaskFn = Callable[..., Any]


@dataclass(frozen=True)
class Schedule:
    name: str
    task: TaskFn  # task(session_factory, **options) -> dict
    at: tuple[int, int] | None = None  # 日本時間の毎日この時刻から
    every_minutes: int | None = None  # この分ごと
    window_minutes: int = 60  # 毎日の予定の枠の長さ
    run_at_startup: bool = False  # 起動したとき、今の枠がまだでも動かす（gas_ping）


_schedules: dict[str, Schedule] = {}


def register_schedule(name: str, *, task: TaskFn, at: tuple[int, int] | None = None,
                      every_minutes: int | None = None, window_minutes: int = 60,
                      run_at_startup: bool = False) -> None:
    """予定を登録する。at=(時, 分)（日本時間の毎日）か every_minutes=n のどちらか 1 つ。"""
    if (at is None) == (every_minutes is None):
        raise ValueError("at か every_minutes のどちらか 1 つを指定してください")
    if at is not None and not (0 <= at[0] < 24 and 0 <= at[1] < 60):
        raise ValueError(f"時刻が正しくありません: {at}")
    if every_minutes is not None and not (1 <= every_minutes <= 24 * 60):
        raise ValueError(f"間隔が正しくありません: {every_minutes}")
    _schedules[name] = Schedule(name=name, task=task, at=at, every_minutes=every_minutes,
                                window_minutes=window_minutes, run_at_startup=run_at_startup)


def unregister_schedules() -> None:
    """試験用。"""
    _schedules.clear()


def schedules() -> dict[str, Schedule]:
    return dict(_schedules)


def slot_start(sched: Schedule, now: datetime) -> datetime | None:
    """今の時刻が予定の枠の中なら、その枠の始まり（UTC）。枠の外なら None。"""
    jst = timeutil.to_jst(now)
    if sched.every_minutes is not None:
        midnight = jst.replace(hour=0, minute=0, second=0, microsecond=0)
        elapsed = int((jst - midnight).total_seconds() // 60)
        start = midnight + timedelta(minutes=elapsed - elapsed % sched.every_minutes)
        return start.astimezone(timeutil.UTC)
    hour, minute = sched.at  # type: ignore[misc]
    start = jst.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if start > jst:
        start -= timedelta(days=1)
    if jst - start >= timedelta(minutes=sched.window_minutes):
        return None
    return start.astimezone(timeutil.UTC)


def _slot_taken(db: Session, name: str, start: datetime) -> bool:
    return db.scalars(select(SyncRun.id).where(SyncRun.kind == SCHEDULE_PREFIX + name,
                                               SyncRun.started_at >= start).limit(1)).first() is not None


def _call_task(task: TaskFn, session_factory: Callable[[], Session], options: dict) -> Any:
    params = inspect.signature(task).parameters
    accepted = {k: v for k, v in options.items() if k in params}
    return task(session_factory, **accepted)


def run_task(session_factory: Callable[[], Session], name: str, *, marker: int | None = None,
             options: dict | None = None) -> SyncRun:
    """予定を 1 つ動かし、sync_runs（kind='schedule:<名前>'）に結果を書く。marker は先に書いた行の id。"""
    sched = _schedules[name]
    with session_factory() as db:
        if marker is None:
            row = SyncRun(kind=SCHEDULE_PREFIX + name, started_at=timeutil.now_utc(), ok=False,
                          summary_json={"state": "running"})
            db.add(row)
            db.commit()
            marker = row.id
    ok, summary = False, {}
    try:
        result = _call_task(sched.task, session_factory, options or {})
        ok = not (isinstance(result, dict) and result.get("ok") is False)
        summary = result if isinstance(result, dict) else {"result": result}
    except Exception as error:  # noqa: BLE001 - 記録して次の枠を待つ
        log.exception("scheduled task failed: %s", name)
        summary = {"error": f"{type(error).__name__}: {error}"[:2000]}
    with session_factory() as db:
        row = db.get(SyncRun, marker)
        row.finished_at = timeutil.now_utc()
        row.ok = ok
        try:
            import json

            json.dumps(summary, default=str)
            row.summary_json = summary
        except (TypeError, ValueError):
            row.summary_json = {"result": str(summary)[:2000]}
        db.commit()
        db.refresh(row)
        return row


def skip_daily_on_startup_env() -> bool:
    """SKIP_DAILY_ON_STARTUP=1: 起動した時に枠の中にある毎日の予定を、その回だけ飛ばす（データ安全 中-5）。"""
    import os

    return os.environ.get("SKIP_DAILY_ON_STARTUP", "").strip().lower() in ("1", "true", "yes")


@dataclass
class Scheduler:
    session_factory: Callable[[], Session]
    started: bool = False
    # 起動した時に枠の中にある毎日の予定（照合 2:10・マスタ同期 0:30・朝バッチ 6:50 など）を動かさず、
    # 枠を「飛ばした」として記録する。切り替えの夜に worker を起動するとき、確かめの前に書き始めないため。
    skip_daily_on_startup: bool = False

    def due(self, now: datetime | None = None) -> list[tuple[str, int]]:
        """今の枠で動かす予定（名前と、先に書いた sync_runs の id）。枠の行はここで書いて確定する。"""
        now = now or timeutil.now_utc()
        first = not self.started
        self.started = True
        out: list[tuple[str, int]] = []
        with self.session_factory() as db:
            for name, sched in _schedules.items():
                start = slot_start(sched, now)
                if start is None and not (first and sched.run_at_startup):
                    continue
                if start is not None and _slot_taken(db, name, start) and not (first and sched.run_at_startup):
                    continue
                if first and self.skip_daily_on_startup and sched.at is not None and start is not None:
                    db.add(SyncRun(kind=SCHEDULE_PREFIX + name, started_at=now, finished_at=now, ok=True,
                                   summary_json={"state": "skipped", "reason": "SKIP_DAILY_ON_STARTUP"}))
                    db.commit()
                    log.warning("scheduled task skipped on startup (SKIP_DAILY_ON_STARTUP): %s", name)
                    continue
                row = SyncRun(kind=SCHEDULE_PREFIX + name, started_at=now, ok=False,
                              summary_json={"state": "running"})
                db.add(row)
                db.commit()
                out.append((name, row.id))
        return out

    def tick(self, now: datetime | None = None) -> list[str]:
        ran = []
        for name, marker in self.due(now):
            log.info("scheduled task start: %s", name)
            run_task(self.session_factory, name, marker=marker)
            ran.append(name)
        return ran


# ===== F の予定: purge（毎日 3:30） =====


def purge(session_factory: Callable[[], Session]) -> dict:
    """冪等キー 24 時間・期限切れのセッション・done の仕事 90 日を消す（audit_log は消さない）。"""
    from sqlalchemy import delete

    from app import jobs, rpc
    from app.models import UserSession

    now = timeutil.now_utc()
    with session_factory() as db:
        keys = rpc.purge_idempotency_keys(db, now)
        sessions = int(db.execute(delete(UserSession).where(UserSession.expires_at < now)).rowcount or 0)
        done = jobs.purge_done(db, now=now)
        db.commit()
    return {"ok": True, "idempotencyKeys": keys, "sessions": sessions, "doneJobs": done}


def register_builtin_schedules() -> None:
    register_schedule("purge", task=purge, at=(3, 30))


def load_registrations() -> list[str]:
    """担当 D・M のモジュールの register_worker() を呼ぶ。呼べたモジュールの名前を返す。"""
    loaded = []
    for module_name in REGISTRATION_MODULES:
        try:
            module = importlib.import_module(module_name)
        except ModuleNotFoundError as exc:
            if exc.name == module_name:
                continue
            raise
        hook = getattr(module, "register_worker", None)
        if callable(hook):
            hook()
            loaded.append(module_name)
    return loaded


def _resolve_row_sync() -> Callable[..., Any] | None:
    try:
        module = importlib.import_module("app.services.row_sync")
    except ModuleNotFoundError as exc:
        if exc.name == "app.services.row_sync":
            return None
        raise
    return getattr(module, "run_once", None)


# ===== スレッド =====


@dataclass
class Worker:
    session_factory: Callable[[], Session]
    app_settings: Any = None
    stop_event: threading.Event = field(default_factory=threading.Event)

    def _loop(self, name: str, seconds: float, body: Callable[[], Any]) -> None:
        log.info("worker thread start: %s", name)
        while not self.stop_event.is_set():
            try:
                body()
            except Exception:  # noqa: BLE001 - 1 回の失敗でスレッドを止めない
                log.exception("worker thread %s failed", name)
            self.stop_event.wait(seconds)
        log.info("worker thread stop: %s", name)

    def row_sync_body(self) -> None:
        fn = _resolve_row_sync()
        if fn is not None:
            fn(self.session_factory, limit=50)

    def gas_lane_body(self) -> None:
        from app.jobs import JobRunner

        runner = JobRunner(self.session_factory, lane="gas", app_settings=self.app_settings)
        with self.session_factory() as lock_db:
            bind = lock_db.get_bind()
            if bind.dialect.name == "postgresql":
                got = lock_db.execute(text("SELECT pg_try_advisory_lock(:k)"), {"k": GAS_LANE_ADVISORY_LOCK}).scalar()
                if not got:
                    return
                try:
                    runner.run_once()
                finally:
                    lock_db.execute(text("SELECT pg_advisory_unlock(:k)"), {"k": GAS_LANE_ADVISORY_LOCK})
                    lock_db.commit()
                return
        runner.run_once()

    def default_lane_body(self, scheduler: Scheduler) -> None:
        from app.jobs import JobRunner

        scheduler.tick()
        JobRunner(self.session_factory, lane="default", app_settings=self.app_settings).run_once()

    def run_forever(self) -> None:
        scheduler = Scheduler(self.session_factory, skip_daily_on_startup=skip_daily_on_startup_env())
        threads = [
            threading.Thread(target=self._loop, args=("row_sync", ROW_SYNC_SECONDS, self.row_sync_body), daemon=True),
            threading.Thread(target=self._loop, args=("gas_lane", GAS_LANE_SECONDS, self.gas_lane_body), daemon=True),
            threading.Thread(target=self._loop,
                             args=("schedule_default", DEFAULT_LANE_SECONDS, lambda: self.default_lane_body(scheduler)),
                             daemon=True),
        ]
        signal.signal(signal.SIGTERM, lambda *_: self.stop_event.set())
        signal.signal(signal.SIGINT, lambda *_: self.stop_event.set())
        for t in threads:
            t.start()
        while not self.stop_event.is_set():
            self.stop_event.wait(1.0)
        for t in threads:
            t.join(timeout=30)


def main(argv: list[str] | None = None) -> None:  # pragma: no cover - VPS で動かす
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    from app import jobs, safety
    from app.config import get_settings
    from app.db import get_sessionmaker

    settings = get_settings()  # fail-closed
    safety.verify_startup(settings)
    sf = get_sessionmaker()
    register_builtin_schedules()
    loaded = load_registrations()
    args = list(sys.argv[1:] if argv is None else argv)
    if args[:1] == ["run"] and len(args) >= 2 and args[1] in _schedules:
        options = {"check_only": True} if "--check" in args[2:] else {}
        row = run_task(sf, args[1], options=options)
        print(f"{args[1]}: ok={row.ok} {row.summary_json}")
        sys.exit(0 if row.ok else 1)
    if args:
        sys.exit("usage: python -m app.worker [run <" + "|".join(sorted(_schedules)) + "> [--check]]")
    log.info("starting overtime worker (APP_ENV=%s, modules=%s, kinds=%s, schedules=%s)", settings.app_env, loaded,
             jobs.registered_kinds(), sorted(_schedules))
    Worker(session_factory=sf, app_settings=settings).run_forever()


if __name__ == "__main__":  # pragma: no cover
    main()
