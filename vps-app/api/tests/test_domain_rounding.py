"""丸め（gap-1 G5・実装仕様 8-3）。Math.round は floor(x + 0.5)、滞留は floor。"""
from __future__ import annotations

from datetime import date

from app.domain import requests, summaries
from test_domain_helpers import ADMIN, add_request, add_wl, ctx, jst


def test_projection_15_5_rounds_up_to_16(db):
    r = add_request(db, target=date(2026, 12, 1), status="approved")
    add_wl(db, r.request_id, net=1)
    res = summaries.api_adminMonthlySummary(ctx(db, email=ADMIN, now=jst(2026, 12, 2, 12)), "2026-12", "ALL")
    p = res["people"][0]
    assert p["pacePerDay"] == 0.5 and p["projectedTotal"] == 16  # 1 分 / 2 日 × 31 日 = 15.5 → 16


def test_projection_integer_pace_is_int_in_json(db):
    r = add_request(db, target=date(2026, 12, 1), status="approved")
    add_wl(db, r.request_id, net=4)
    p = summaries.api_adminMonthlySummary(ctx(db, email=ADMIN, now=jst(2026, 12, 2)), "2026-12", "ALL")["people"][0]
    assert p["pacePerDay"] == 2 and isinstance(p["pacePerDay"], int)


def test_planned_minutes_round_half_up():
    assert requests.planned_minutes_overtime("0.5") == 30
    assert requests.planned_minutes_overtime("0.0125") == 1  # 0.75 → 1
    assert requests.planned_minutes_overtime("0.0083333") == 0  # 0.49999 → 0（0 以下の検査は予定時間の数で行う）
    assert requests.planned_minutes_overtime("0x10") == 960  # Number('0x10') = 16


def test_pending_hours_floor(db):
    add_request(db, submitted_at=jst(2026, 11, 24, 10, 0, 1))
    res = summaries.api_adminPendingWatch(ctx(db, email=ADMIN, now=jst(2026, 11, 26, 10)))
    assert res[0]["hoursPending"] == 47 and res[0]["isOver48h"] is False
