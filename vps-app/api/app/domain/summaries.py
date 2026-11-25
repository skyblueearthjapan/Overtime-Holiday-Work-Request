"""集計（4-3 の 18・22〜25・27〜29）。adminApi.gs:36-376・staffSummaryApi.gs の写し（実装仕様 8-4）。

- 土台は store.joined_records（status が空・canceled を除き、submitted も含む。WorkLogs は下の行が勝つ）。
- 月次: 40h=2400・60h=3600・注意 1800、判定はすべて >=。未来月は月末まで経過とみなす。単位は dept + '__' + 氏名。
- 年度推移・特別条項は 4/1 始まり、スタッフ集計は 3/16 年度と締め月（approved の net だけ）。
- 丸めは jsutil.js_round（予測が .5 なら上へ）。並びは JS の `<`（UTF-16 の符号順）と安定な並べ替え。
  localeCompare(…,'ja') の 2 か所（日次一覧・部署の円グラフ）は collation.ja_key。
"""
from __future__ import annotations

import calendar
import math
from datetime import date, timedelta
from typing import Any

from app import timeutil
from app.jsutil import js_floor, js_round, js_string, js_trim, normalize_gas
from app.rpc import RpcContext, RpcError

from . import store
from .collation import ja_key
from .fiscal import add_months, closing_ym, fy_start_0316, fy_start_0401
from .jsdate import MISSING, days_ago, is_retroactive, js_int_or_float, js_new_date, js_number, js_truthy, label_ja
from .masters import load_dept_list

LIMIT40 = 2400
LIMIT60 = 3600
WATCH = 1800


def _today(ctx: RpcContext) -> date:
    return timeutil.to_jst(ctx.now).date()


def _dept_filtered(recs: list[store.Joined], dept_filter: Any) -> list[store.Joined]:
    if js_truthy(dept_filter) and dept_filter != "ALL":
        return [r for r in recs if r.dept == dept_filter]
    return recs


def _js_int(value: float) -> int:
    """new Date(y, m, …) に渡す数（小数は切り捨て）。"""
    return int(math.trunc(value))


# ===== 月次サマリー（buildMonthlySummary_） =====


def build_monthly_summary(ctx: RpcContext, year_month: Any, dept_filter: Any) -> dict:
    pieces = js_string(year_month).split("-") if year_month is not None else [""]
    y = js_number(pieces[0]) if len(pieces) > 0 else math.nan
    m = js_number(pieces[1]) if len(pieces) > 1 else math.nan
    if not js_truthy(y) or not js_truthy(m) or m < 1 or m > 12:
        raise RpcError("yearMonthは yyyy-mm で指定してください")
    year, month = _js_int(y), _js_int(m)
    if not 1 <= year <= 9998:
        # Python の日付は 1〜9999 年だけ。範囲外は 500 にせず同じ文の 400（GAS は空の集計。仕様 8-2。正しさ 低-16）
        raise RpcError("yearMonthは yyyy-mm で指定してください")
    start = date(year, month, 1)
    end = add_months(start, 1)

    recs = _dept_filtered(store.joined_records(ctx.db, start, end), dept_filter)

    today = _today(ctx)
    days_in_month = calendar.monthrange(year, month)[1]
    in_same_month = start <= today < end
    day_of_month = today.day if in_same_month else days_in_month

    people_map: dict[str, dict] = {}
    dept_sum: dict[str, int] = {}
    for r in recs:
        key = f"{r.dept}__{r.workerName}"
        agg = people_map.get(key)
        if agg is None:
            agg = {
                "dept": r.dept, "workerName": r.workerName, "overtimeNet": 0, "holidayNet": 0, "totalNet": 0,
                "approvedTotal": 0, "pdfMissing": 0, "pacePerDay": 0, "projectedTotal": 0,
                "projectedOver40": False, "projectedOver60": False, "remain40": LIMIT40, "remain60": LIMIT60,
            }
            people_map[key] = agg
        agg["approvedTotal"] += int(r.approvedMinutes or 0)
        if r.requestType == "overtime":
            agg["overtimeNet"] += r.netMinutes
        if r.requestType == "holiday":
            agg["holidayNet"] += r.netMinutes
        agg["totalNet"] += r.netMinutes
        if r.status == "approved" and r.netMinutes > 0 and not r.pdfFileId:
            agg["pdfMissing"] += 1
        dept_sum[r.dept] = dept_sum.get(r.dept, 0) + r.netMinutes

    people = list(people_map.values())
    for p in people:
        pace = p["totalNet"] / max(1, day_of_month)
        p["pacePerDay"] = js_int_or_float(pace)
        p["projectedTotal"] = js_round(pace * days_in_month)
        p["projectedOver40"] = p["projectedTotal"] >= LIMIT40
        p["projectedOver60"] = p["projectedTotal"] >= LIMIT60
        p["remain40"] = max(0, LIMIT40 - p["totalNet"])
        p["remain60"] = max(0, LIMIT60 - p["totalNet"])
    people.sort(key=lambda p: -p["totalNet"])

    watch = [{
        "dept": p["dept"], "workerName": p["workerName"], "totalNet": p["totalNet"],
        "projectedTotal": p["projectedTotal"], "remain40": p["remain40"], "remain60": p["remain60"],
        "projectedOver40": p["projectedOver40"], "projectedOver60": p["projectedOver60"],
        "pdfMissing": p["pdfMissing"],
    } for p in people if p["totalNet"] >= WATCH or p["projectedOver40"]]

    dept_labels = sorted(dept_sum.keys(), key=ja_key)
    data = {
        "yearMonth": year_month,
        "deptFilter": dept_filter if js_truthy(dept_filter) else "ALL",
        "monthInfo": {"daysInMonth": days_in_month, "dayOfMonth": day_of_month, "isCurrentMonth": in_same_month},
        "kpi": {},
        "people": people,
        "watch": watch,
        "charts": {
            "people": {},
            "dept": {"labels": dept_labels, "values": [dept_sum.get(d, 0) for d in dept_labels]},
        },
    }
    recompute_people_summary(data)
    return data


def recompute_people_summary(data: dict) -> None:
    """kpi と charts.people を people から作る（承認者の複数部署版はこれだけ作り直す。charts.dept は残す）。"""
    people = data["people"]
    data["kpi"] = {
        "totalPeople": len(people),
        "over40": sum(1 for p in people if p["totalNet"] >= LIMIT40),
        "over60": sum(1 for p in people if p["totalNet"] >= LIMIT60),
        "projectedOver40": sum(1 for p in people if p["projectedOver40"]),
        "projectedOver60": sum(1 for p in people if p["projectedOver60"]),
        "pdfMissingTotal": sum(p["pdfMissing"] for p in people),
        "totalNetMinutes": sum(p["totalNet"] for p in people),
    }
    data["charts"]["people"] = {
        "labels": [p["workerName"] for p in people],
        "values": [p["totalNet"] for p in people],
        "projected": [p["projectedTotal"] for p in people],
        "limit40": LIMIT40,
        "limit60": LIMIT60,
    }


# ===== 総務の関数 =====


def api_adminDeptOptions(ctx: RpcContext):
    return ["ALL", *load_dept_list(ctx.db)]


def api_adminMonthlySummary(ctx: RpcContext, year_month=None, dept_filter=None):
    return build_monthly_summary(ctx, year_month, dept_filter)


def _jst_date_of(value: Any, error: str) -> date:
    moment = js_new_date(value)
    if moment is None:
        raise RpcError(error)
    return timeutil.to_jst(moment).date()


def api_adminDailyDetail(ctx: RpcContext, date_ymd=MISSING, dept_filter=None):
    """9-1 #22: targetDate・pdfGeneratedAt・actualStartAt・actualEndAt を文字で返し、isRetroactive・daysAgo を足す。"""
    day = _jst_date_of(date_ymd, "日付が不正です（yyyy-mm-dd）")
    recs = _dept_filtered(store.joined_records(ctx.db, day, day + timedelta(days=1)), dept_filter)
    pdf_missing = sum(1 for r in recs if r.status == "approved" and r.netMinutes > 0 and not r.pdfFileId)
    today = _today(ctx).isoformat()
    items = []
    for r in sorted(recs, key=lambda r: ja_key(r.dept + r.workerName)):
        ymd = r.targetDate.isoformat()
        items.append({
            "requestId": r.requestId, "status": r.status, "requestType": r.requestType, "dept": r.dept,
            "workerCode": r.workerCode, "workerName": r.workerName, "targetDate": ymd,
            "approvedMinutes": r.approvedMinutes, "actualStartAt": r.actualStartAt, "actualEndAt": r.actualEndAt,
            "actualMinutes": r.actualMinutes, "breakMinutes": r.breakMinutes, "netMinutes": r.netMinutes,
            "pdfFileId": r.pdfFileId, "pdfGeneratedAt": store.jst_text(r.pdfGeneratedAt),
            "isRetroactive": is_retroactive(r.submittedAt, ymd), "daysAgo": days_ago(ymd, today),
        })
    return {
        "date": day.isoformat(),
        "deptFilter": dept_filter if js_truthy(dept_filter) else "ALL",
        "total": len(recs),
        "pdfMissing": pdf_missing,
        "items": items,
    }


def api_adminFiscalTrend(ctx: RpcContext, base_date_ymd=MISSING, dept_filter=None):
    base = _jst_date_of(base_date_ymd, "baseDateYmdが不正です（yyyy-mm-dd）")
    fy_start = fy_start_0401(base)
    fy_end = date(fy_start.year + 1, 4, 1)
    recs = _dept_filtered(store.joined_records(ctx.db, fy_start, fy_end), dept_filter)
    sums: dict[str, int] = {}
    for r in recs:
        mk = r.targetDate.strftime("%Y-%m")
        sums[mk] = sums.get(mk, 0) + r.netMinutes
    labels, values = [], []
    for i in range(12):
        mk = add_months(fy_start, i).strftime("%Y-%m")
        labels.append(mk)
        values.append(sums.get(mk, 0))
    return {
        "fiscalYearStart": fy_start.isoformat(),
        "fiscalYearEnd": fy_end.isoformat(),
        "deptFilter": dept_filter if js_truthy(dept_filter) else "ALL",
        "labels": labels,
        "values": values,
    }


def api_adminSpecialClauseCount(ctx: RpcContext, base_date_str=MISSING, dept_filter=None):
    base = _jst_date_of(base_date_str, "基準日が不正です")
    year = base.year if base.month >= 4 else base.year - 1
    recs = [r for r in store.joined_records(ctx.db, date(year, 4, 1), date(year + 1, 4, 1)) if r.netMinutes > 0]

    monthly: dict[str, int] = {}
    for r in recs:
        key = f"{r.dept}|{r.workerName}|{r.targetDate.strftime('%Y-%m')}"
        monthly[key] = monthly.get(key, 0) + r.netMinutes
    counts: dict[str, int] = {}
    for key, total in monthly.items():
        if total >= LIMIT60:
            dept, name = (key.split("|") + [""])[:2]
            k = f"{dept}|{name}"
            counts[k] = counts.get(k, 0) + 1
    result = []
    for key, cnt in counts.items():
        dept, name = (key.split("|") + [""])[:2]
        if js_truthy(dept_filter) and dept_filter != "ALL" and dept != dept_filter:
            continue
        result.append({"dept": dept, "workerName": name, "count60": cnt, "isDanger": cnt >= 6, "isWarn": cnt >= 5})
    result.sort(key=lambda x: -x["count60"])
    return {"fiscalYear": f"{year}-04-01 ~ {year + 1}-03-31", "list": result}


def api_adminPendingWatch(ctx: RpcContext):
    now = ctx.now
    today = _today(ctx).isoformat()
    result = []
    for req in store.requests_in_sheet_order(ctx.db):
        if normalize_gas(req.status) != "submitted":
            continue
        submitted = req.submitted_at
        if submitted is None:
            continue  # GAS は関数全体が失敗する見込み。VPS はその行を飛ばす（受け入れる差。8-4）
        hours = (now - submitted).total_seconds() / 3600
        td = store.target_ymd(req) or ""
        result.append({
            "requestType": normalize_gas(req.request_type),
            "dept": normalize_gas(req.dept),
            "workerName": normalize_gas(req.worker_name),
            "targetDateLabel": label_ja(td, today),
            "isPastDate": bool(td) and td < today,
            "submittedAt": store.jst_text(submitted),
            "hoursPending": js_floor(hours),
            "isOver48h": hours >= 48,
            "targetDate": td,
            "isRetroactive": is_retroactive(submitted, td),
            "daysAgo": days_ago(td, today),
        })
    result.sort(key=lambda x: -x["hoursPending"])
    return result


# ===== スタッフ集計（staffSummaryApi.gs） =====


def _worker_key(code: str, dept: str, name: str) -> str:
    return ("C:" + code) if code else ("N:" + dept + "|" + name)


def api_getClosingSummary(ctx: RpcContext, year_month=None, dept=None, worker_token=None):
    if not js_truthy(year_month):
        year_month = closing_ym(_today(ctx))
    pieces = js_string(year_month).split("-")
    y = js_number(pieces[0]) if pieces else math.nan
    m = js_number(pieces[1]) if len(pieces) > 1 else math.nan
    if not js_truthy(y) or not js_truthy(m) or m < 1 or m > 12:
        raise RpcError("yearMonth は yyyy-MM で指定してください: " + js_string(year_month))
    year, month = _js_int(y), _js_int(m)
    if not 1 <= year <= 9998:  # 正しさ 低-16（上と同じ）
        raise RpcError("yearMonth は yyyy-MM で指定してください: " + js_string(year_month))
    period_from = date(year, month, 16)
    period_to = add_months(period_from, 1)
    ym_label = f"{year}-{month:02d}"

    fy_start = fy_start_0316(period_from)
    fy_end = date(fy_start.year + 1, 3, 16)
    dept_sel = dept if (js_truthy(dept) and dept != "ALL") else ""

    fy_recs = store.joined_records(ctx.db, fy_start, fy_end)
    if dept_sel:
        fy_recs = [r for r in fy_recs if r.dept == dept_sel]
    month_recs = [r for r in fy_recs if period_from <= r.targetDate < period_to]

    last = period_to - timedelta(days=1)
    meta = {
        "yearMonth": ym_label,
        "monthLabel": f"{period_from.year}年 {period_from.month}月分",
        "periodFrom": f"{period_from.month}/{period_from.day}",
        "periodTo": f"{last.month}/{last.day}",
        "fyLabel": f"{fy_start.year}年度",
        "limit40": LIMIT40,
        "limit60": LIMIT60,
    }

    if js_truthy(worker_token):
        token = js_trim(js_string(worker_token))
        sp = token.find(" ")
        code = token[:sp] if sp >= 0 else token
        name = js_trim(token[sp + 1:]) if sp >= 0 else ""

        def match(r: store.Joined) -> bool:
            if code and r.workerCode == code:
                return True
            return bool(name) and r.workerName == name and (not dept_sel or r.dept == dept_sel)

        mine = [r for r in month_recs if match(r)]
        ot_net = hol_net = pend_min = pend_count = 0
        for r in mine:
            actualized = r.status == "approved" and r.netMinutes > 0
            if r.status == "approved":
                if r.requestType == "overtime":
                    ot_net += r.netMinutes
                elif r.requestType == "holiday":
                    hol_net += r.netMinutes
            if not actualized:
                pend_min += int(r.approvedMinutes or 0)
                pend_count += 1
        detail = [{
            "date": f"{r.targetDate.month}/{r.targetDate.day}",
            "type": r.requestType,
            "approvedMinutes": int(r.approvedMinutes or 0),
            "netMinutes": int(r.netMinutes or 0),
            "breakMinutes": int(r.breakMinutes or 0),
            "status": r.status,
            "pdfFileId": r.pdfFileId or "",
        } for r in sorted(mine, key=lambda r: r.targetDate)]

        by_ym: dict[str, int] = {}
        for r in fy_recs:
            if not match(r) or r.status != "approved":
                continue
            ym = closing_ym(r.targetDate)
            by_ym[ym] = by_ym.get(ym, 0) + r.netMinutes
        annual_over = sum(1 for v in by_ym.values() if v >= LIMIT40)

        w_name, w_dept = name, dept_sel
        if mine:
            w_name = mine[0].workerName or name
            w_dept = mine[0].dept or dept_sel
        return {
            "mode": "personal",
            "meta": meta,
            "worker": {"workerCode": code, "workerName": w_name, "dept": w_dept},
            "kpi": {"otNet": ot_net, "holNet": hol_net, "totalNet": ot_net + hol_net, "pendMin": pend_min,
                    "pendCount": pend_count},
            "detail": detail,
            "annualOver": annual_over,
        }

    groups: dict[str, dict] = {}
    for r in month_recs:
        key = _worker_key(r.workerCode, r.dept, r.workerName)
        g = groups.get(key)
        if g is None:
            g = {"workerCode": r.workerCode or "", "workerName": r.workerName, "dept": r.dept, "otNet": 0,
                 "holNet": 0, "totalNet": 0, "pendMin": 0, "annualOver": 0}
            groups[key] = g
        if r.status == "approved":
            if r.requestType == "overtime":
                g["otNet"] += r.netMinutes
            elif r.requestType == "holiday":
                g["holNet"] += r.netMinutes
            g["totalNet"] = g["otNet"] + g["holNet"]
        if not (r.status == "approved" and r.netMinutes > 0):
            g["pendMin"] += int(r.approvedMinutes or 0)

    annual_by_key: dict[str, dict[str, int]] = {}
    for r in fy_recs:
        if r.status != "approved":
            continue
        key = _worker_key(r.workerCode, r.dept, r.workerName)
        ym = closing_ym(r.targetDate)
        bucket = annual_by_key.setdefault(key, {})
        bucket[ym] = bucket.get(ym, 0) + r.netMinutes

    people = list(groups.values())
    for person in people:
        key = _worker_key(person["workerCode"], person["dept"], person["workerName"])
        person["annualOver"] = sum(1 for v in annual_by_key.get(key, {}).values() if v >= LIMIT40)
    people.sort(key=lambda p: -p["totalNet"])
    return {"mode": "list", "meta": meta, "dept": dept_sel, "people": people}


HANDLERS = {
    "api_adminDeptOptions": api_adminDeptOptions,
    "api_adminMonthlySummary": api_adminMonthlySummary,
    "api_adminFiscalTrend": api_adminFiscalTrend,
    "api_adminDailyDetail": api_adminDailyDetail,
    "api_adminSpecialClauseCount": api_adminSpecialClauseCount,
    "api_adminPendingWatch": api_adminPendingWatch,
    "api_getClosingSummary": api_getClosingSummary,
}
