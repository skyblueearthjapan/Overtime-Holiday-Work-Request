"""後回しの仕事（outbox）: 積む・取り出す・実行する・送り直す（担当 F の土台。実装仕様 5-4・7-3・10-3）。

kind ごとの実行は登録式（担当 M・D が自分のモジュールの register_worker() で呼ぶ）:

    from app.jobs import JobContext, PermanentJobError, RetryLater, SkipJob, register_kind

    def run_after_approve(ctx: JobContext) -> dict:
        ...                      # ctx.db で読み、GAS を呼ぶ。戻り値（dict）は jobs.result_json に入る
        raise RetryLater(60)     # 試行回数を増やさずに 60 秒後にもう一度
        raise SkipJob("canceled")            # skipped で終える
        raise PermanentJobError("bad_request")  # 再試行せずに failed

    register_kind("internal.afterApprove", run_after_approve, lane="gas", on_exhausted=None)

決まり:
- lane='gas' は全体で同時に 1 本（取り出しは 1 件ずつ。動いている gas の仕事があれば取り出さない。
  PostgreSQL では worker が pg_try_advisory_lock でも守る）。priority の高い順、同じなら id 順。
- 同じ申請（request_id）の仕事は積んだ順。前の仕事が done／skipped になるまで次は動かさない。
  前の仕事が failed のままなら、同じ申請の次の仕事も止まる（list_blocked で管理者の欄に出す）。
- requires（[{table, pk, version}]）の行の反映が済むまで動かさない（試行回数は増やさず 5 秒後に）。
- 失敗は 1・2・4・8・16・30 分、その後は 30 分ごとに、積んでから 24 時間まで送り直す（Q20 の既定）。
  PermanentJobError と 24 時間を過ぎたものは failed にして on_exhausted を別のトランザクションで呼ぶ。
- 実行中（running）のまま leased_until を過ぎた仕事は、worker が止まったとみなして拾い直す。
- ハンドラーが登録されていない種類は取り出さない（pending のまま残る）。
- 実行は 1 本ずつ 1 つのトランザクション。ハンドラーが戻ったら done にして commit する。
"""
from __future__ import annotations

import json
import logging
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import and_, func, or_, select
from sqlalchemy.orm import Session

from app import models, sheet_queue, timeutil
from app.constants import JOB_LANES

log = logging.getLogger(__name__)

DEFAULT_RETRY_DELAYS: tuple[timedelta, ...] = tuple(timedelta(minutes=m) for m in (1, 2, 4, 8, 16, 30))
RETRY_EVERY_AFTER = timedelta(minutes=30)
RETRY_DEADLINE = timedelta(hours=24)
REQUIRES_WAIT = timedelta(seconds=5)
LEASES: dict[str, timedelta] = {"gas": timedelta(minutes=10), "default": timedelta(minutes=15)}
DONE_STATES = ("done", "skipped")


class SkipJob(Exception):
    """状態が変わった・申請が無いなどで、出さずに捨てる（skipped）。"""


class RetryLater(Exception):
    """試行回数を増やさずに、seconds 秒後にもう一度動かす（requires を待つ・gas_ping を待つなど）。"""

    def __init__(self, seconds: float, reason: str = ""):
        super().__init__(reason or f"retry in {seconds}s")
        self.seconds = float(seconds)
        self.reason = reason


class PermanentJobError(Exception):
    """再試行しても直らない失敗（GAS の retryable:false など）。すぐ failed にする。"""


@dataclass
class JobContext:
    """ハンドラーに渡す。db はこの仕事のトランザクション（commit は仕組みが行う）。"""

    db: Session
    job_id: int
    kind: str
    lane: str
    args: dict
    request_id: str
    attempt: int  # 1 から
    now: datetime
    guard: dict = field(default_factory=dict)
    requires: list = field(default_factory=list)
    batch_id: str = ""
    priority: int = 0
    created_at: datetime | None = None
    app_settings: Any = None

    def commit(self) -> None:
        """途中の確定（例: GAS を呼ぶ前に何かを記録する）。"""
        self.db.commit()


JobHandler = Callable[[JobContext], Any]
ExhaustedHandler = Callable[[JobContext, str], None]


@dataclass
class JobKind:
    kind: str
    lane: str
    handler: JobHandler
    on_exhausted: ExhaustedHandler | None = None
    retry_delays: tuple[timedelta, ...] = DEFAULT_RETRY_DELAYS
    retry_deadline: timedelta = RETRY_DEADLINE


_kinds: dict[str, JobKind] = {}


def register_kind(kind: str, handler: JobHandler, *, lane: str, on_exhausted: ExhaustedHandler | None = None,
                  retry_delays: Sequence[timedelta] | None = None,
                  retry_deadline: timedelta | None = None) -> None:
    """kind の実行を差し込む（worker の起動時）。lane は 'gas'／'default'。"""
    if lane not in JOB_LANES:
        raise ValueError(f"知らないレーンです: {lane}")
    if not kind:
        raise ValueError("kind が空です")
    _kinds[kind] = JobKind(kind=kind, lane=lane, handler=handler, on_exhausted=on_exhausted,
                           retry_delays=tuple(retry_delays) if retry_delays else DEFAULT_RETRY_DELAYS,
                           retry_deadline=retry_deadline or RETRY_DEADLINE)


def unregister_all() -> None:
    """試験用: 差し込んだ実行を外す。"""
    _kinds.clear()


def registered_kinds(lane: str | None = None) -> list[str]:
    return [k for k, spec in _kinds.items() if lane is None or spec.lane == lane]


def get_kind(kind: str) -> JobKind | None:
    return _kinds.get(kind)


# ---------------------------------------------------------------------------
# 積む
# ---------------------------------------------------------------------------


def _dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, default=str, separators=(",", ":"))


def _loads(text: str | None, default: Any) -> Any:
    try:
        value = json.loads(text or "")
    except ValueError:
        return default
    return value if isinstance(value, type(default)) else default


def enqueue(db: Session, kind: str, args: Mapping[str, Any] | None = None, *, lane: str = "gas",
            request_id: str = "", priority: int = 0, requires: Iterable[Mapping[str, Any]] = (),
            guard: Mapping[str, Any] | None = None, batch_id: str = "", run_after: datetime | None = None) -> int:
    """仕事を 1 件積む（業務の書き込みと同じトランザクション。commit は呼び出し側）。仕事の id を返す。"""
    if lane not in JOB_LANES:
        raise ValueError(f"知らないレーンです: {lane}")
    if not kind:
        raise ValueError("kind が空です")
    req_list = []
    for r in requires:
        if not {"table", "pk", "version"} <= set(r):
            raise ValueError("requires の形は {table, pk, version} です")
        req_list.append({"table": str(r["table"]), "pk": str(r["pk"]), "version": int(r["version"])})
    now = timeutil.now_utc()
    job = models.Job(kind=kind, lane=lane, priority=int(priority), request_id=request_id or "",
                     args_json=_dumps(dict(args or {})), requires_json=_dumps(req_list),
                     guard_json=_dumps(dict(guard or {})), state="pending", attempts=0,
                     run_after=run_after or now, last_error="", result_json="", batch_id=batch_id or "",
                     created_at=now, updated_at=now)
    db.add(job)
    db.flush()
    return int(job.id)


# ---------------------------------------------------------------------------
# 管理者の欄・運用（ops_getStatus は担当 M がこれらを使う）
# ---------------------------------------------------------------------------


def _job_dict(r: models.Job) -> dict:
    return {"id": r.id, "kind": r.kind, "lane": r.lane, "requestId": r.request_id, "state": r.state,
            "attempts": r.attempts, "priority": r.priority, "error": r.last_error or "", "batchId": r.batch_id,
            "createdAt": r.created_at, "runAfter": r.run_after}


def counts_by_state(db: Session, lane: str | None = None) -> dict[str, int]:
    J = models.Job
    q = select(J.state, func.count()).group_by(J.state)
    if lane:
        q = q.where(J.lane == lane)
    counts = {s: 0 for s in ("pending", "running", "done", "failed", "skipped")}
    counts.update({state: int(n) for state, n in db.execute(q)})
    return counts


def list_jobs(db: Session, states: Sequence[str] = ("failed",), lane: str | None = None, limit: int = 200) -> list[dict]:
    J = models.Job
    q = select(J).where(J.state.in_(tuple(states))).order_by(J.id).limit(limit)
    if lane:
        q = q.where(J.lane == lane)
    return [_job_dict(r) for r in db.scalars(q)]


def list_failed(db: Session, lane: str | None = None) -> list[dict]:
    return list_jobs(db, ("failed",), lane)


def list_blocked(db: Session) -> list[dict]:
    """前の仕事が failed のため止まっている仕事（同じ申請の後の仕事）。"""
    J = models.Job
    failed = select(J.request_id).where(J.state == "failed", J.request_id != "")
    rows = db.scalars(select(J).where(J.state == "pending", J.request_id.in_(failed)).order_by(J.id)).all()
    out = []
    for r in rows:
        earlier = db.scalar(select(func.count()).select_from(J).where(
            J.request_id == r.request_id, J.state == "failed", J.id < r.id)) or 0
        if earlier:
            out.append(_job_dict(r))
    return out


def retry_failed(db: Session, job_id: int) -> bool:
    """失敗した仕事をもう一度 pending に戻す（ops_retryJob）。24 時間の締め切りは戻した時刻から数え直す。"""
    job = db.get(models.Job, int(job_id))
    if job is None or job.state != "failed":
        return False
    now = timeutil.now_utc()
    job.state, job.attempts, job.run_after, job.finished_at, job.leased_until = "pending", 0, now, None, None
    job.created_at = now
    job.updated_at = now
    return True


def purge_done(db: Session, older_than: timedelta = timedelta(days=90), now: datetime | None = None) -> int:
    """done の仕事を 90 日で消す（worker の purge）。commit は呼び出し側。"""
    from sqlalchemy import delete

    cutoff = (now or timeutil.now_utc()) - older_than
    J = models.Job
    return int(db.execute(delete(J).where(J.state == "done", J.updated_at < cutoff)).rowcount or 0)


# ---------------------------------------------------------------------------
# 取り出し・実行
# ---------------------------------------------------------------------------


@dataclass
class JobView:
    id: int
    kind: str
    lane: str
    request_id: str
    args: dict
    guard: dict
    requires: list
    attempts: int
    batch_id: str
    priority: int
    created_at: datetime


def _requires_met(db: Session, requires: list) -> bool:
    for r in requires:
        try:
            if not sheet_queue.is_synced(db, str(r["table"]), str(r["pk"]), int(r["version"])):
                return False
        except (KeyError, TypeError, ValueError):
            return False
    return True


@dataclass
class JobRunner:
    """レーン 1 本の取り出しと実行。worker の各スレッドが自分の JobRunner を持つ。"""

    session_factory: Callable[[], Session]
    lane: str = "gas"
    app_settings: Any = None
    now: Callable[[], datetime] = field(default=lambda: timeutil.now_utc())

    def _runnable_kinds(self) -> list[str]:
        return registered_kinds(self.lane)

    def claim(self, limit: int | None = None) -> list[JobView]:
        """動かせる仕事を running にして返す。gas レーンはいつも 1 件まで。"""
        if self.lane == "gas":
            limit = 1
        limit = limit or 10
        kinds = self._runnable_kinds()
        if not kinds:
            return []
        now = self.now()
        J = models.Job
        claimed: list[JobView] = []
        with self.session_factory() as s, s.begin():
            live_running = and_(J.state == "running", or_(J.leased_until.is_(None), J.leased_until >= now))
            if self.lane == "gas":
                busy = s.scalar(select(func.count()).select_from(J).where(J.lane == "gas", live_running)) or 0
                if busy:
                    return []
            due = or_(J.run_after.is_(None), J.run_after <= now)
            stale = and_(J.state == "running", J.leased_until < now)
            candidates = s.scalars(
                select(J).where(J.lane == self.lane, J.kind.in_(kinds), or_(and_(J.state == "pending", due), stale))
                .order_by(J.priority.desc(), J.id).limit(max(limit * 5, 20)).with_for_update(skip_locked=True)
            ).all()
            taken: set[str] = set()
            for job in candidates:
                if len(claimed) >= limit:
                    break
                if job.request_id:
                    if job.request_id in taken:
                        continue
                    # 同じ申請の、先に積まれてまだ done／skipped でない仕事があれば後回し（積んだ順・前が failed なら止まる）
                    earlier = s.scalar(select(func.count()).select_from(J).where(
                        J.request_id == job.request_id, J.id < job.id, J.state.notin_(DONE_STATES))) or 0
                    if earlier:
                        continue
                requires = _loads(job.requires_json, [])
                if requires and not _requires_met(s, requires):
                    # 行の反映を待つ（試行回数は増やさない）
                    job.run_after = now + REQUIRES_WAIT
                    job.updated_at = now
                    continue
                if job.request_id:
                    taken.add(job.request_id)
                job.state = "running"
                job.attempts = int(job.attempts or 0) + 1
                job.leased_until = now + LEASES.get(self.lane, LEASES["default"])
                job.updated_at = now
                claimed.append(JobView(id=job.id, kind=job.kind, lane=job.lane, request_id=job.request_id,
                                       args=_loads(job.args_json, {}), guard=_loads(job.guard_json, {}),
                                       requires=requires, attempts=job.attempts, batch_id=job.batch_id,
                                       priority=job.priority, created_at=job.created_at))
        return claimed

    def run_once(self, limit: int | None = None) -> dict[str, int]:
        result = {"done": 0, "skipped": 0, "retry": 0, "waiting": 0, "failed": 0}
        for job in self.claim(limit):
            result[self.run_job(job)] += 1
        return result

    def _ctx(self, db: Session, job: JobView) -> JobContext:
        return JobContext(db=db, job_id=job.id, kind=job.kind, lane=job.lane, args=dict(job.args),
                          request_id=job.request_id, attempt=job.attempts, now=self.now(), guard=dict(job.guard),
                          requires=list(job.requires), batch_id=job.batch_id, priority=job.priority,
                          created_at=job.created_at, app_settings=self.app_settings)

    def _finish(self, job_id: int, **fields: Any) -> None:
        with self.session_factory() as s, s.begin():
            row = s.get(models.Job, job_id)
            for k, v in fields.items():
                setattr(row, k, v)
            row.updated_at = self.now()

    def _next_delay(self, spec: JobKind, attempts: int) -> timedelta:
        if attempts - 1 < len(spec.retry_delays):
            return spec.retry_delays[attempts - 1]
        return RETRY_EVERY_AFTER

    def run_job(self, job: JobView) -> str:
        spec = _kinds.get(job.kind)
        if spec is None:  # 取り出した後に登録が外れた（試験だけ）
            self._finish(job.id, state="pending", attempts=max(0, job.attempts - 1), leased_until=None)
            return "waiting"
        message = ""
        permanent = False
        with self.session_factory() as db:
            try:
                result = spec.handler(self._ctx(db, job))
                row = db.get(models.Job, job.id)
                now = self.now()
                row.state, row.finished_at, row.last_error, row.updated_at = "done", now, "", now
                row.leased_until = None
                if result is not None:
                    row.result_json = _dumps(result)
                db.commit()
                return "done"
            except SkipJob as skip:
                db.rollback()
                log.info("job %s skipped: %s", job.id, skip)
                self._finish(job.id, state="skipped", finished_at=self.now(), leased_until=None,
                             last_error=str(skip)[:1000])
                return "skipped"
            except RetryLater as wait:
                db.rollback()
                self._finish(job.id, state="pending", attempts=max(0, job.attempts - 1), leased_until=None,
                             run_after=self.now() + timedelta(seconds=max(0.0, wait.seconds)),
                             last_error=(wait.reason or "")[:1000])
                return "waiting"
            except PermanentJobError as error:
                db.rollback()
                log.warning("job %s (%s %s) failed permanently: %s", job.id, job.kind, job.request_id, error)
                message = f"{type(error).__name__}: {error}"[:1000]
                permanent = True
            except Exception as error:  # noqa: BLE001 - すべて送り直しの対象
                db.rollback()
                log.exception("job %s (%s %s) failed", job.id, job.kind, job.request_id)
                message = f"{type(error).__name__}: {error}"[:1000]
        now = self.now()
        next_at = now + self._next_delay(spec, job.attempts)
        if not permanent and next_at <= job.created_at + spec.retry_deadline:
            self._finish(job.id, state="pending", run_after=next_at, leased_until=None, last_error=message)
            return "retry"
        self._finish(job.id, state="failed", finished_at=now, leased_until=None, last_error=message)
        if spec.on_exhausted is not None:
            with self.session_factory() as db:
                try:
                    spec.on_exhausted(self._ctx(db, job), message)
                    db.commit()
                except Exception:  # noqa: BLE001
                    db.rollback()
                    log.exception("job %s: on_exhausted failed", job.id)
        return "failed"
