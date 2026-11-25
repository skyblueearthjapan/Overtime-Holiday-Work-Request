"""マスタを読む画面の関数（4-3 の 1〜5）。masters.gs・deptAuth.gs・formMap.gs の写し。

表はマスタ同期（担当 D）が作った workers・orders・company_calendar（並びは sort_order＝シートの行の順）。
"""
from __future__ import annotations

import functools
from datetime import date, timedelta

from sqlalchemy import select
from sqlalchemy.orm import Session

from app import timeutil
from app.jsutil import js_sort_key, normalize_gas
from app.models import CompanyCalendar, Order, Worker
from app.rpc import RpcContext

from .jsdate import DAY_NAMES, js_new_date

# ===== 作業員マスタ =====


def _workers(db: Session) -> list[Worker]:
    return list(db.scalars(select(Worker).order_by(Worker.sort_order)))


def _active(row: Worker) -> bool:
    """在籍フラグ: 空・'1'・'true'（大文字小文字を問わない）なら在籍。"""
    active = normalize_gas(row.active_flag)
    return not active or active == "1" or active.lower() == "true"


def load_workers_by_dept(db: Session) -> dict[str, list[str]]:
    """loadWorkersByDept_: 部署 → ['コード 氏名', …]（各部署の中は JS の既定の並び）。"""
    out: dict[str, list[str]] = {}
    for row in _workers(db):
        dept = normalize_gas(row.dept)
        code = normalize_gas(row.worker_code)
        name = normalize_gas(row.worker_name)
        if not dept or not code or not name:
            continue
        if not _active(row):
            continue
        out.setdefault(dept, []).append(f"{code} {name}")
    for labels in out.values():
        labels.sort(key=js_sort_key)
    return out


def load_dept_list(db: Session) -> list[str]:
    return sorted(load_workers_by_dept(db).keys(), key=js_sort_key)


def lookup_worker_email(db: Session, worker_code: str) -> str:
    """lookupWorkerEmail_: コードが一致する最初の行の「Googleアカウント」列（列が無ければ ''）。"""
    if not worker_code:
        return ""
    if not _has_google_account_column(db):
        return ""
    for row in _workers(db):
        if normalize_gas(row.worker_code) == worker_code:
            return normalize_gas(row.google_account) or ""
    return ""


def _has_google_account_column(db: Session) -> bool:
    first = db.scalars(select(Worker).order_by(Worker.sort_order).limit(1)).first()
    if first is None:
        return False
    return any(normalize_gas(k).startswith("Googleアカウント") for k in (first.raw or {}).keys())


def api_getDeptList(ctx: RpcContext):
    return load_dept_list(ctx.db)


def api_getWorkersByDept(ctx: RpcContext, dept=None):
    if not dept:
        return []
    if not isinstance(dept, str):
        return []
    return load_workers_by_dept(ctx.db).get(dept, [])


def api_getWorkerInfo(ctx: RpcContext):
    """作業員マスタの「Googleアカウント」で始まる列とログインのメールを大文字小文字を区別して比べる。列が無ければ null。"""
    email = ctx.user.email
    if not email:
        return None
    if not _has_google_account_column(ctx.db):
        return None
    for row in _workers(ctx.db):
        if not _active(row):
            continue
        if normalize_gas(row.google_account) == email:
            return {
                "workerCode": normalize_gas(row.worker_code),
                "workerName": normalize_gas(row.worker_name),
                "dept": normalize_gas(row.dept),
                "email": email,
            }
    return None


# ===== 工番マスタ =====


def _orders(db: Session) -> list[Order]:
    return list(db.scalars(select(Order).order_by(Order.sort_order)))


def lookup_order_info(db: Session, order_no: str) -> dict | None:
    for row in _orders(db):
        if normalize_gas(row.code) == order_no:
            return {
                "orderNo": order_no,
                "customer": normalize_gas(row.customer),
                "dest": normalize_gas(row.dest),
                "product": normalize_gas(row.product),
            }
    return None


def api_getOrderList(ctx: RpcContext):
    out: list[dict] = []
    seen: set[str] = set()
    for row in _orders(ctx.db):
        code = normalize_gas(row.code)
        if not code or code in seen:
            continue
        seen.add(code)
        out.append({"code": code, "product": normalize_gas(row.product), "dest": normalize_gas(row.dest)})

    def cmp(a: dict, b: dict) -> int:
        ka, kb = js_sort_key(a["code"]), js_sort_key(b["code"])
        return -1 if ka < kb else 1 if ka > kb else 0

    out.sort(key=functools.cmp_to_key(cmp))
    return out


# ===== 休日の候補日（formMap.gs:222-450） =====


class _Fallback(Exception):
    pass


def _ymd(d: date) -> str:
    return d.isoformat()


def _md(d: date) -> str:
    return f"{d.month}/{d.day}"


def _dow(d: date) -> str:
    return DAY_NAMES[(d.weekday() + 1) % 7]


def _week(today: date) -> tuple[date, date]:
    """今週の月曜と、来週の金曜（月曜 + 11 日）。"""
    js_dow = (today.weekday() + 1) % 7  # 0=日
    monday = today - timedelta(days=6 if js_dow == 0 else js_dow - 1)
    return monday, monday + timedelta(days=11)


def _calendar_date_text(raw: str) -> str:
    """マスタの A 列。GAS は日付のセルなら yyyy-MM-dd、文字ならその文字（normalize_）。"""
    text = normalize_gas(raw)
    if len(text) >= 8 and ("/" in text[:10]) and text[:4].isdigit():
        moment = js_new_date(text)
        if moment is not None:
            return timeutil.to_jst_text(moment, "yyyy-MM-dd")
    return text


def holiday_candidates(db: Session, today: date) -> list[dict]:
    monday, search_end = _week(today)
    monday_s, end_s, today_s = _ymd(monday), _ymd(search_end), _ymd(today)
    candidates: list[dict] = []
    try:
        rows = list(db.scalars(select(CompanyCalendar).order_by(CompanyCalendar.sort_order)))
        for row in rows:
            kubun = normalize_gas(row.kind)
            youbi = normalize_gas(row.weekday)
            biko = normalize_gas(row.note)
            if kubun == "出勤土曜":
                continue
            date_str = _calendar_date_text(row.date_text)
            if not date_str or len(date_str) < 10:
                continue
            if date_str < today_s or date_str < monday_s or date_str > end_s:
                continue
            moment = js_new_date(f"{date_str}T00:00:00+09:00")
            if moment is None:
                raise _Fallback()  # fmtDate_ が例外 → 従来方式（GAS の catch）
            d = moment.astimezone(timeutil.JST).date()
            label = f"{_md(d)}({youbi or _dow(d)})"
            kind = "weekend"
            holiday_name = ""
            if kubun == "祝日":
                kind = "holiday"
                holiday_name = biko or ""
                if biko:
                    label += " " + biko
            elif biko:
                holiday_name = biko
                label += " " + biko
            item = {"date": date_str, "label": label, "type": kind}
            if holiday_name:
                item["holidayName"] = holiday_name
            candidates.append(item)
    except _Fallback:
        return legacy_candidates(today)
    if not candidates:
        return legacy_candidates(today)
    return _sorted_by_date(candidates)


def _sorted_by_date(items: list[dict]) -> list[dict]:
    def cmp(a: dict, b: dict) -> int:
        return -1 if a["date"] < b["date"] else 1 if a["date"] > b["date"] else 0

    return sorted(items, key=functools.cmp_to_key(cmp))


def legacy_candidates(today: date) -> list[dict]:
    """api_getHolidayCandidateDates_legacy_（Calendar API は使わない＝計算の祝日だけ。gap-5 中-6）。"""
    monday, next_friday = _week(today)
    saturday, sunday = monday + timedelta(days=5), monday + timedelta(days=6)
    candidates = [
        {"date": _ymd(saturday), "label": _md(saturday) + "(土)", "type": "weekend"},
        {"date": _ymd(sunday), "label": _md(sunday) + "(日)", "type": "weekend"},
    ]
    for h in known_japanese_holidays(monday, next_friday):
        existing = next((c for c in candidates if c["date"] == h["date"]), None)
        if existing is not None:
            existing["holidayName"] = h["title"]
            continue
        hd = date.fromisoformat(h["date"])
        candidates.append({"date": h["date"], "label": f"{_md(hd)}({_dow(hd)}) {h['title']}", "type": "holiday"})
    today_s = _ymd(today)
    return [c for c in _sorted_by_date(candidates) if c["date"] >= today_s]


_FIXED = (
    (1, 1, "元日"), (1, -2, "成人の日"), (2, 11, "建国記念の日"), (2, 23, "天皇誕生日"), (3, 0, "春分の日"),
    (4, 29, "昭和の日"), (5, 3, "憲法記念日"), (5, 4, "みどりの日"), (5, 5, "こどもの日"), (7, -3, "海の日"),
    (8, 11, "山の日"), (9, -3, "敬老の日"), (9, 0, "秋分の日"), (10, -2, "スポーツの日"), (11, 3, "文化の日"),
    (11, 23, "勤労感謝の日"),
)


def _nth_monday(year: int, month: int, n: int) -> date:
    d = date(year, month, 1)
    count = 0
    while count < n:
        if d.weekday() == 0:
            count += 1
        if count < n:
            d += timedelta(days=1)
    return d


def _floor(x: float) -> int:
    import math

    return int(math.floor(x))


def _vernal(y: int) -> int:
    return _floor(20.8431 + 0.242194 * (y - 1980) - _floor((y - 1980) / 4))


def _autumnal(y: int) -> int:
    return _floor(23.2488 + 0.242194 * (y - 1980) - _floor((y - 1980) / 4))


def _js_date(y: int, month_index: int, day: int) -> date:
    """new Date(y, monthIndex, day)（日の繰り越しを含む）。"""
    return date(y + month_index // 12, month_index % 12 + 1, 1) + timedelta(days=day - 1)


def known_japanese_holidays(start: date, end: date) -> list[dict]:
    """getKnownJapaneseHolidays_（formMap.gs:381-450）を同じ順で写す。"""
    year = start.year
    range_start = _ymd(start)
    range_end = _ymd(end + timedelta(days=1))
    holidays: dict[str, str] = {}
    for y in (year - 1, year, year + 1):
        for month, day_or_type, title in _FIXED:
            if day_or_type < 0:
                d = _nth_monday(y, month, abs(day_or_type))
            elif day_or_type == 0 and month == 3:
                d = _js_date(y, 2, _vernal(y))
            elif day_or_type == 0 and month == 9:
                d = _js_date(y, 8, _autumnal(y))
            else:
                d = _js_date(y, month - 1, day_or_type)
            ds = _ymd(d)
            if range_start <= ds < range_end:
                holidays[ds] = title
            if d.weekday() == 6:  # 日曜 → 翌月曜（翌日も祝日ならさらに翌日）
                sub = d + timedelta(days=1)
                while holidays.get(_ymd(sub)):
                    sub += timedelta(days=1)
                sub_ds = _ymd(sub)
                if range_start <= sub_ds < range_end:
                    holidays[sub_ds] = title + "（振替休日）"
    return [{"date": ds, "title": t} for ds, t in holidays.items()]


def api_getHolidayCandidateDates(ctx: RpcContext):
    return holiday_candidates(ctx.db, timeutil.to_jst(ctx.now).date())


HANDLERS = {
    "api_getDeptList": api_getDeptList,
    "api_getWorkersByDept": api_getWorkersByDept,
    "api_getWorkerInfo": api_getWorkerInfo,
    "api_getOrderList": api_getOrderList,
    "api_getHolidayCandidateDates": api_getHolidayCandidateDates,
}
