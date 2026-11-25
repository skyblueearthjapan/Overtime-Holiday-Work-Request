"""担当 M の試験の小道具: 偽物の GasClient・時計・申請の行・レーンを流しきる関数。"""
from __future__ import annotations

import copy
from datetime import date, datetime, timedelta, timezone

from sqlalchemy import func, select

from app import jobs, models, sheet_queue
from app.jobs import JobRunner

T0 = datetime(2026, 11, 26, 0, 0, tzinfo=timezone.utc)  # 09:00 JST


class Clock:
    def __init__(self, at: datetime = T0):
        self.at = at

    def __call__(self) -> datetime:
        return self.at

    def advance(self, **kw) -> None:
        self.at = self.at + timedelta(**kw)


def created_pdf(request_id: str, n: int = 1) -> dict:
    return {"ok": True, "requestId": request_id, "pdf": {
        "status": "created", "pdfFileId": f"PDF-{request_id}-{n}", "pdfName": f"{request_id}.pdf", "folderId": "FOLDER-1",
        "pdfGeneratedAt": "2026-11-26 09:00:05"}, "ledger": {"status": "upserted"}, "trash": {"status": "skipped"},
        "pdfFileId": f"PDF-{request_id}-{n}", "pdfFolderId": "FOLDER-1", "pdfGeneratedAt": "2026-11-26 09:00:05"}


class FakeGas:
    """GasClient の代わり。calls に呼ばれた順で残す。同時に 2 本動いたら max_in_flight が 2 になる。"""

    def __init__(self):
        self.calls: list[dict] = []
        self.responses: dict[str, object] = {}
        self.in_flight = 0
        self.max_in_flight = 0
        self.during = None  # 呼び出しの最中に動かす関数（直列の確かめ）

    def call(self, action, *, job_id, request_id="", expect=None, args=None, timeout=None):
        self.in_flight += 1
        self.max_in_flight = max(self.max_in_flight, self.in_flight)
        try:
            self.calls.append({"action": action, "jobId": job_id, "requestId": request_id,
                               "expect": copy.deepcopy(expect), "args": copy.deepcopy(args)})
            if self.during is not None:
                self.during(action, request_id)
            r = self.responses.get(action)
            if isinstance(r, list):
                r = r.pop(0) if r else None
            if callable(r) and not isinstance(r, Exception):
                r = r(action=action, job_id=job_id, request_id=request_id, expect=expect, args=args)
            if isinstance(r, Exception):
                raise r
            if r is None:
                r = self.default(action, request_id)
            return copy.deepcopy(r)
        finally:
            self.in_flight -= 1

    def default(self, action, request_id):
        if action == "internal.ping":
            return good_ping()
        if action in ("internal.afterApprove", "internal.afterComplete", "internal.regenerate", "internal.batchPdf"):
            return created_pdf(request_id)
        if action == "internal.afterDelete":
            return {"ok": True, "requestId": request_id, "trash": {"status": "skipped"}, "ledger": {"status": "removed"}}
        if action == "vps.upsertLedger":
            return {"ok": True, "requestId": request_id, "ledger": {"status": "upserted"}}
        if action == "internal.ledgerRewriteDay":
            return {"ok": True, "targetDate": "", "ledger": {"status": "rewritten"}}
        return {"ok": True}

    def actions(self) -> list[str]:
        return [c["action"] for c in self.calls if c["action"] != "internal.ping"]


def good_ping() -> dict:
    """G の vpsPing_ の形（dbId・dbTimeZone）に、仕様の scriptId を足したもの（staging で合格する値）。"""
    return {"ok": True, "pong": "2026-11-26T00:00:00.000Z", "version": "2026-11-26.1", "migrationMode": False,
            "scriptId": "staging-script-copy", "dbId": "staging-db-sheet-id", "dbTimeZone": "Asia/Tokyo",
            "settings": {"TEMPLATE_SSID": "staging-db-sheet-id", "PDF_ROOT_FOLDER_ID": "staging-pdf-folder-id",
                         "ACCUMULATION_SS_ID": "staging-ledger-id", "PDF_MODE": "direct で直接書込方式",
                         "pdfModeIsDirect": True}}


def add_ping(db, at: datetime, ok: bool = True) -> None:
    db.add(models.SyncRun(kind="gas_ping", started_at=at, finished_at=at, ok=ok,
                          summary_json={"problems": [] if ok else ["scriptId が違います"]}))
    db.commit()


def add_request(db, request_id: str, *, status: str = "approved", request_type: str = "overtime",
                target: date = date(2026, 11, 25), dept: str = "製造部", worker_name: str = "山田 太郎",
                pdf_file_id: str = "", approved_by2: str = "", work_log_end: str | None = None,
                synced: bool = True) -> models.Request:
    req = models.Request(request_id=request_id, request_type=request_type, status=status, dept=dept,
                         worker_code="W001", worker_name=worker_name, worker_email="w@lineworks-local.info",
                         target_date=target, approved_minutes=60, pdf_file_id=pdf_file_id,
                         pdf_folder_id="FOLDER-0" if pdf_file_id else "", approved_by2=approved_by2)
    db.add(req)
    db.flush()
    sheet_queue.mark_dirty(db, "requests", request_id)
    if work_log_end is not None:
        wl = models.WorkLog(request_id=request_id, actual_start_at="2026-11-25 17:20:00", actual_end_at=work_log_end,
                            actual_minutes=60, break_minutes=0, net_minutes=60, updated_by="w@lineworks-local.info")
        db.add(wl)
        db.flush()
        sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
    db.commit()
    if synced:
        sync_all(db)
    return req


def sync_all(db) -> None:
    """行の反映のレーンが全部書き終えたことにする。"""
    for st in db.scalars(select(models.SheetSyncState)):
        st.synced_version = st.desired_version
    db.commit()


def job(db, job_id) -> models.Job:
    db.expire_all()
    return db.get(models.Job, job_id)


def drain(session_factory, clock, *, max_steps: int = 200, sync=None, on_step=None) -> list[str]:
    """GAS のレーンを流しきる。待ちの仕事しか無ければ、いちばん早い run_after まで時計を進める。"""
    runner = JobRunner(session_factory, lane="gas", now=clock)
    outcomes = []
    for _ in range(max_steps):
        if sync is not None:
            with session_factory() as s:
                sync(s)
        result = runner.run_once()
        moved = [k for k, v in result.items() if v]
        outcomes.extend(moved)
        if on_step is not None:
            on_step()
        if any(result.values()):
            continue
        with session_factory() as s:
            J = models.Job
            nxt = s.scalar(select(func.min(J.run_after)).where(J.lane == "gas", J.state == "pending",
                                                               J.kind.in_(jobs.registered_kinds("gas"))))
        if nxt is None:
            return outcomes
        if nxt <= clock():
            # 動かせない（前の仕事が failed で止まっている・requires 待ちで時刻が来ている）
            clock.advance(seconds=5)
            if len(outcomes) > max_steps:
                return outcomes
            continue
        clock.at = nxt
    return outcomes
