"""日本時間の「今日」・年度・締め月（決定 C・実装仕様 8-1、gap-1 G1）。"""
from __future__ import annotations

from datetime import date

import pytest

from app import timeutil
from app.domain import dashboards, fiscal, requests, summaries
from app.models import Request
from test_domain_helpers import ADMIN, add_request, add_wl, ctx, jst


def submit(db, now, **kw):
    data = {"requestType": "overtime", "dept": "製造1課", "workerLabel": "A001 山田", "reason": "急ぎ作業を要するため",
            "hours": "1.0", "targetDate": "2000-01-01"}
    data.update(kw)
    return requests.api_submitRequest(ctx(db, now=now), data)


def test_0630_jst_overtime_target_is_today(db):
    now = jst(2026, 11, 26, 6, 30)
    assert now.astimezone(timeutil.UTC).date() == date(2026, 11, 25)  # UTC ではまだ前日
    res = submit(db, now)
    assert db.get(Request, res["requestId"]).target_date == date(2026, 11, 26)


def test_0316_0800_rejects_0315_retroactive_overtime(db):
    res = submit(db, jst(2026, 3, 16, 8, 0), isRetroactive=True, targetDate="2026-03-15")
    assert res == {"ok": False, "error": "年度開始日（2026-03-16）より前の日付は申請できません。"}
    # 3/15 の 23:59 JST なら前の年度の中
    res = submit(db, jst(2026, 3, 15, 23, 59), isRetroactive=True, targetDate="2026-03-15")
    assert res["ok"] is True


def test_0316_holiday_boundary(db):
    res = submit(db, jst(2026, 3, 16, 0, 5), requestType="holiday", hours="1日", targetDate="2026-03-15")
    assert res["error"] == "年度開始日（2026-03-16）より前の日付は申請できません。"
    assert submit(db, jst(2026, 3, 16, 0, 5), requestType="holiday", hours="1日", targetDate="2026-03-16")["ok"]


@pytest.mark.parametrize("d,start", [
    (date(2026, 3, 15), date(2025, 3, 16)), (date(2026, 3, 16), date(2026, 3, 16)),
    (date(2026, 1, 1), date(2025, 3, 16)), (date(2026, 12, 31), date(2026, 3, 16)),
])
def test_fy_0316(d, start):
    assert fiscal.fy_start_0316(d) == start


@pytest.mark.parametrize("d,start", [
    (date(2026, 3, 31), date(2025, 4, 1)), (date(2026, 4, 1), date(2026, 4, 1)),
])
def test_fy_0401(d, start):
    assert fiscal.fy_start_0401(d) == start


def test_closing_month_15th_2330_vs_16th_0030(db):
    assert fiscal.closing_ym(timeutil.to_jst(jst(2026, 11, 15, 23, 30)).date()) == "2026-10"
    assert fiscal.closing_ym(timeutil.to_jst(jst(2026, 11, 16, 0, 30)).date()) == "2026-11"
    assert fiscal.closing_ym(date(2026, 1, 15)) == "2025-12"
    # 16 日 00:30 JST（UTC では 15 日 15:30）の既定の締め月
    res = summaries.api_getClosingSummary(ctx(db, now=jst(2026, 11, 16, 0, 30)))
    assert res["meta"]["yearMonth"] == "2026-11" and res["meta"]["periodTo"] == "12/15"
    res = summaries.api_getClosingSummary(ctx(db, now=jst(2026, 11, 15, 23, 30)))
    assert res["meta"]["yearMonth"] == "2026-10"


def test_closing_month_january_and_fiscal_year_label(db):
    res = summaries.api_getClosingSummary(ctx(db, now=jst(2027, 3, 10)), "2027-03")
    assert res["meta"]["periodFrom"] == "3/16" and res["meta"]["periodTo"] == "4/15"
    assert res["meta"]["fyLabel"] == "2027年度"
    res = summaries.api_getClosingSummary(ctx(db), "2027-02")
    assert res["meta"]["fyLabel"] == "2026年度" and res["meta"]["monthLabel"] == "2027年 2月分"
    res = summaries.api_getClosingSummary(ctx(db), "2026-12")
    assert res["meta"]["periodTo"] == "1/15"


def test_special_clause_and_trend_on_3_31_and_4_1(db):
    r1 = add_request(db, target=date(2026, 3, 31), status="approved")
    add_wl(db, r1.request_id, net=3600)
    r2 = add_request(db, target=date(2026, 4, 1), status="approved")
    add_wl(db, r2.request_id, net=3600)
    old = summaries.api_adminSpecialClauseCount(ctx(db, email=ADMIN), "2026-03-31", "ALL")
    new = summaries.api_adminSpecialClauseCount(ctx(db, email=ADMIN), "2026-04-01", "ALL")
    assert old["fiscalYear"] == "2025-04-01 ~ 2026-03-31" and old["list"][0]["count60"] == 1
    assert new["fiscalYear"] == "2026-04-01 ~ 2027-03-31" and new["list"][0]["count60"] == 1
    assert summaries.api_adminFiscalTrend(ctx(db, email=ADMIN), "2026-03-31", "ALL")["values"][-1] == 3600
    assert summaries.api_adminFiscalTrend(ctx(db, email=ADMIN), "2026-04-01", "ALL")["values"][0] == 3600


def test_dashboard_30_day_window_and_today_at_0630(db):
    now = jst(2026, 11, 26, 6, 30)
    inside = add_request(db, target=date(2026, 10, 27), status="approved")
    add_request(db, target=date(2026, 10, 26), status="approved")
    today = add_request(db, target=date(2026, 11, 26))
    res = dashboards.api_getDashboard(ctx(db, now=now))
    assert res["today"] == "2026-11-26"
    assert {i["requestId"] for i in res["overtime"]} == {inside.request_id, today.request_id}
