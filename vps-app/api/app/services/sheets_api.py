"""Google Sheets API v4 の読み書き（DELEGATE_EMAIL の委任トークン）。担当 F の土台。担当 D が使う。

- 書く関数はすべて、直前に safety.assert_writable(spreadsheet_id, tab, purpose) を通す
  （staging は STAGING_SPREADSHEET_ID だけ、production は PROD_SPREADSHEET_ID だけ。タブと目的の許可リスト）。
  spreadsheets.batchUpdate の要求は updateCells と appendDimension（ROWS）だけで、すべて tab のシート
  （safety.assert_write_requests）。values.append は Requests の requestId だけ（RAW・OVERWRITE。行を挿入しない）。
  行・列・タブの挿入・削除・並べ替え・作り直しは出さない。
- 読む関数は照合しない（マスタ元・本番 DB を読むのは許す）。
- 429・5xx は指数的に待って 4 回まで送り直す。

使い方（担当 D）:
    api = SheetsApi.for_db(db)
    meta = api.get_spreadsheet(sid)                                   # properties.timeZone・シートの一覧
    unformatted, formatted = api.batch_get_values(sid, ["'Requests'!A1:AZ"], "UNFORMATTED_VALUE"), ...
    api.batch_update(sid, [{"updateCells": {...}}], tab="Requests", purpose="row_sync")
    api.append_rows(sid, "Requests", 10)                               # 末尾に行を足す（appendDimension）
    first_row = api.append_values(sid, "Requests", [[request_id]])     # 表の次の行を取る（RAW・OVERWRITE。決定 #26）
    api.values_batch_update(sid, [{"range": "'工番マスタ'!A1:H3", "values": [...]}],
                            tab="工番マスタ", purpose="master_sync", input_option="USER_ENTERED")
    api.values_batch_clear(sid, ["'工番マスタ'!A4:H"], tab="工番マスタ", purpose="master_sync")
"""
from __future__ import annotations

import logging
import time
import urllib.parse
from collections.abc import Callable, Mapping, Sequence
from typing import Any

from app import safety

log = logging.getLogger(__name__)

API_BASE = "https://sheets.googleapis.com/v4/spreadsheets"
RETRY_STATUS = (429, 500, 502, 503, 504)
RENDER_OPTIONS = ("UNFORMATTED_VALUE", "FORMATTED_VALUE", "FORMULA")
INPUT_OPTIONS = ("RAW", "USER_ENTERED")
META_FIELDS = "properties(title,timeZone,locale),sheets(properties(sheetId,title,index,gridProperties))"


class SheetsError(RuntimeError):
    """Sheets API の失敗（再試行が尽きた・4xx）。"""

    def __init__(self, message: str, status: int | None = None):
        super().__init__(message)
        self.status = status


class SheetsApi:
    """credentials_provider は委任の Credentials を返す関数。session_factory・sleep は試験用。"""

    def __init__(self, credentials_provider: Callable[[], Any], *, timeout: float = 60.0, retries: int = 4,
                 session_factory: Callable[[Any], Any] | None = None, app_settings: Any = None,
                 sleep: Callable[[float], None] = time.sleep):
        self._credentials_provider = credentials_provider
        self._timeout = timeout
        self._retries = retries
        self._session_factory = session_factory
        self._sleep = sleep
        self._meta_cache: dict[str, dict] = {}
        self.app_settings = app_settings

    @classmethod
    def for_db(cls, db, app_settings: Any = None) -> "SheetsApi":
        from app.google_delegate import get_delegate_credentials

        creds = get_delegate_credentials(db, app_settings)
        return cls(lambda: creds, app_settings=app_settings)

    # ----- 呼び出し -----

    def _session(self):
        creds = self._credentials_provider()
        if self._session_factory is not None:
            return self._session_factory(creds)
        from google.auth.transport.requests import AuthorizedSession

        return AuthorizedSession(creds)

    def _call(self, method: str, url: str, *, params: Any = None, body: Mapping | None = None) -> dict:
        session = self._session()
        for attempt in range(self._retries + 1):
            resp = session.request(method, url, params=params, json=body, timeout=self._timeout)
            if resp.status_code in RETRY_STATUS and attempt < self._retries:
                self._sleep(min(2 ** attempt, 30))
                continue
            if resp.status_code >= 400:
                raise SheetsError(f"Sheets API {method} が失敗しました（{resp.status_code}）: {resp.text[:500]}",
                                  resp.status_code)
            return resp.json() if resp.content else {}
        raise SheetsError("Sheets API の再試行が尽きました。")  # pragma: no cover

    # ----- 読む -----

    def get_spreadsheet(self, spreadsheet_id: str, fields: str = META_FIELDS) -> dict:
        """スプレッドシートの情報（properties.timeZone・シートの sheetId と名前）。"""
        meta = self._call("GET", f"{API_BASE}/{spreadsheet_id}", params={"fields": fields})
        if fields == META_FIELDS:
            self._meta_cache[spreadsheet_id] = meta
        return meta

    def get_spreadsheet_row_metadata(self, spreadsheet_id: str, ranges: Sequence[str], fields: str) -> dict:
        """includeGridData の読み取り（表示形式 numberFormat などを読むとき。担当 D）。"""
        params = [("ranges", r) for r in ranges] + [("includeGridData", "true"), ("fields", fields)]
        return self._call("GET", f"{API_BASE}/{spreadsheet_id}", params=params)

    def batch_get_values(self, spreadsheet_id: str, ranges: Sequence[str], render: str = "UNFORMATTED_VALUE") -> list[list]:
        """範囲ごとの値（行の 2 次元配列）のリスト。render は UNFORMATTED_VALUE（日時はシリアル値）か FORMATTED_VALUE。"""
        if render not in RENDER_OPTIONS:
            raise ValueError(f"render は {RENDER_OPTIONS} のどれか: {render}")
        params = [("ranges", r) for r in ranges] + [
            ("valueRenderOption", render), ("dateTimeRenderOption", "SERIAL_NUMBER"), ("majorDimension", "ROWS")]
        data = self._call("GET", f"{API_BASE}/{spreadsheet_id}/values:batchGet", params=params)
        value_ranges = data.get("valueRanges", [])
        return [vr.get("values", []) for vr in value_ranges] + [[] for _ in range(len(ranges) - len(value_ranges))]

    def get_values(self, spreadsheet_id: str, a1: str, render: str = "UNFORMATTED_VALUE") -> list:
        return self.batch_get_values(spreadsheet_id, [a1], render)[0]

    def read_both(self, spreadsheet_id: str, ranges: Sequence[str]) -> tuple[list[list], list[list]]:
        """UNFORMATTED と FORMATTED を一緒に読む（罠 11）。"""
        return (self.batch_get_values(spreadsheet_id, ranges, "UNFORMATTED_VALUE"),
                self.batch_get_values(spreadsheet_id, ranges, "FORMATTED_VALUE"))

    def sheet_id(self, spreadsheet_id: str, tab: str) -> int:
        """タブの sheetId（メタデータは 1 回読んで覚える）。無ければ SheetsError（タブを作らない）。"""
        meta = self._meta_cache.get(spreadsheet_id) or self.get_spreadsheet(spreadsheet_id)
        sid = sheet_id_by_title(meta, tab)
        if sid is None:
            raise SheetsError(f"タブ「{tab}」がありません（タブは作りません）。")
        return sid

    # ----- 書く（安全装置を通す） -----

    def _guard(self, spreadsheet_id: str, tab: str, purpose: str, method: str) -> None:
        safety.assert_writable(spreadsheet_id, tab, purpose, self.app_settings)
        safety.assert_write_method(method, tab)

    def batch_update(self, spreadsheet_id: str, requests: Sequence[Mapping], *, tab: str, purpose: str) -> dict:
        """spreadsheets.batchUpdate（updateCells と appendDimension だけ。すべて tab のシート）。"""
        safety.assert_writable(spreadsheet_id, tab, purpose, self.app_settings)
        safety.assert_write_requests(requests, tab, self.sheet_id(spreadsheet_id, tab))
        return self._call("POST", f"{API_BASE}/{spreadsheet_id}:batchUpdate", body={"requests": list(requests)})

    def append_rows(self, spreadsheet_id: str, tab: str, n: int, *, purpose: str = "row_sync") -> dict:
        """タブの末尾に n 行を足す（appendDimension ROWS。行の数が足りないときだけ。A 区分だけ）。"""
        if n <= 0:
            raise ValueError("足す行の数は 1 以上")
        self._guard(spreadsheet_id, tab, purpose, "appendDimension")
        request = {"appendDimension": {"sheetId": self.sheet_id(spreadsheet_id, tab), "dimension": "ROWS",
                                       "length": int(n)}}
        return self.batch_update(spreadsheet_id, [request], tab=tab, purpose=purpose)

    def values_batch_update(self, spreadsheet_id: str, data: Sequence[Mapping], *, tab: str, purpose: str,
                            input_option: str = "RAW") -> dict:
        """values.batchUpdate。data は [{"range": "'タブ'!A1:C2", "values": [[...]]}]（すべて tab の範囲）。"""
        if input_option not in INPUT_OPTIONS:
            raise ValueError(f"input_option は {INPUT_OPTIONS} のどれか: {input_option}")
        self._guard(spreadsheet_id, tab, purpose, "values.batchUpdate")
        safety.assert_ranges_in_tab([d["range"] for d in data], tab)
        return self._call("POST", f"{API_BASE}/{spreadsheet_id}/values:batchUpdate",
                          body={"valueInputOption": input_option, "data": list(data)})

    def append_values(self, spreadsheet_id: str, tab: str, values: Sequence[Sequence[Any]], *,
                      purpose: str = "row_sync") -> int:
        """values.append（RAW・insertDataOption=OVERWRITE。範囲は '<tab>'!A:A）。書いた最初の行（1 始まり）を返す。

        Google が「表の最後の行の次」を決めて書くので、Sebastian の追記（INSERT_ROWS）と同時でも同じ行を取り合わない
        （決定 #26）。OVERWRITE は行を挿入しない（シートの末尾なら行が足される＝appendDimension と同じ）。
        文字は RAW なので式にならない。Requests の新しい行の requestId だけに使う（safety が照合する）。
        """
        rows = [list(r) for r in values]
        if not rows:
            raise ValueError("書く行がありません")
        for row in rows:
            for v in row:
                if not isinstance(v, str):
                    raise safety.UnsafeWriteError("values.append は文字（requestId）だけを書きます")
        self._guard(spreadsheet_id, tab, purpose, "values.append")
        rng = f"{quote_sheet(tab)}!A:A"
        safety.assert_ranges_in_tab([rng], tab)
        data = self._call("POST", f"{API_BASE}/{spreadsheet_id}/values/{quote_a1(rng)}:append",
                          params={"valueInputOption": "RAW", "insertDataOption": "OVERWRITE",
                                  "includeValuesInResponse": "false"},
                          body={"range": rng, "majorDimension": "ROWS", "values": rows})
        updated = str(((data or {}).get("updates") or {}).get("updatedRange") or "")
        return first_row_of_range(updated)

    def values_batch_clear(self, spreadsheet_id: str, ranges: Sequence[str], *, tab: str, purpose: str) -> dict:
        """values.batchClear（C 区分の余りの行の値だけ。行は消さない）。"""
        self._guard(spreadsheet_id, tab, purpose, "values.batchClear")
        safety.assert_ranges_in_tab(ranges, tab)
        return self._call("POST", f"{API_BASE}/{spreadsheet_id}/values:batchClear", body={"ranges": list(ranges)})


# ===== 小道具 =====


def quote_sheet(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def a1(tab: str, cells: str = "") -> str:
    """'タブ'!A1:B2 の形。cells が空ならタブ全体。"""
    return quote_sheet(tab) + (f"!{cells}" if cells else "")


def column_letter(index: int) -> str:
    """0 始まりの列番号 → A, B, ..., Z, AA。"""
    letters = ""
    n = index + 1
    while n:
        n, rest = divmod(n - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def sheet_titles(meta: dict) -> list[str]:
    return [s["properties"]["title"] for s in meta.get("sheets", [])]


def sheet_id_by_title(meta: dict, title: str) -> int | None:
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == title:
            return int(s["properties"]["sheetId"])
    return None


def spreadsheet_timezone(meta: dict) -> str:
    """properties.timeZone（DB_SHEET_TIMEZONE と照合する。gap-5 高-2）。"""
    return str((meta.get("properties") or {}).get("timeZone") or "")


def first_row_of_range(a1_range: str) -> int:
    """'Requests'!A2336:A2337 → 2336（values.append の返事の updatedRange）。読めなければ SheetsError。"""
    import re

    m = re.search(r"!\$?[A-Z]+\$?(\d+)", a1_range or "")
    if not m:
        raise SheetsError(f"values.append の返事の範囲が読めません: {a1_range!r}")
    return int(m.group(1))


def quote_a1(a1_range: str) -> str:
    """URL のパスに入れる A1 表記。"""
    return urllib.parse.quote(a1_range, safe="")
