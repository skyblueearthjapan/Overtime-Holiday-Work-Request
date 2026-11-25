"""試験用の偽の Google Sheets API（外へはつながない。担当 D の試験が使う）。

本物の SheetsApi（安全装置の照合を含む）に session_factory として渡す。
    fake = FakeSheets()
    book = fake.add_book("sid", tz="Asia/Tokyo")
    book.add_tab("Requests", [[...], ...], formats={(行, 列): {...}})
    api = fake.api(settings)

- セルは (値, 表示形式)。値は UNFORMATTED（日時はシリアル値）。FORMATTED_VALUE は表示形式から作る
  （日付の表示形式の無い数は '46352.29' のような数の文字になる）。
- 対応する呼び出し: スプレッドシートの情報（includeGridData を含む）・values:batchGet・:batchUpdate
  （updateCells・appendDimension だけ。ほかの要求は 400）・values:batchUpdate（RAW／USER_ENTERED。
  USER_ENTERED は数・日付の文字を読み替える。書く範囲が行数を超えれば行を足す＝本物と同じ）・values:batchClear。
- updateCells が行数を超えれば 400（本物と同じ。appendDimension で先に足す）。
"""
from __future__ import annotations

import json
import re
import urllib.parse
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from app.jsutil import js_string
from app.services.sheets_api import API_BASE, SheetsApi

_EPOCH = datetime(1899, 12, 30)


class FormulaCell(str):
    """USER_ENTERED で式として入ったセル（試験で「式になった」ことを見分ける印）。"""


def serial(text: str) -> float:
    """'yyyy-MM-dd HH:mm:ss[.fff]'・'yyyy-MM-dd'（シートの壁時計）→ シリアル値。"""
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            delta = datetime.strptime(text, fmt) - _EPOCH
            break
        except ValueError:
            continue
    else:
        raise ValueError(text)
    ms = delta.days * 86_400_000 + delta.seconds * 1000 + delta.microseconds // 1000
    return ms / 86_400_000


DATE_FMT = {"type": "DATE", "pattern": "yyyy-mm-dd"}
DATETIME_FMT = {"type": "DATE_TIME", "pattern": "m/d/yyyy h:mm:ss"}
TEXT_FMT = {"type": "TEXT", "pattern": "@"}


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def parse_a1(a1: str) -> tuple[str, int, int | None, int, int | None]:
    """'タブ'!A1:C3 → (タブ, 行0, 行0の終わり|None, 列0, 列0の終わり|None)。終わりは含む。"""
    m = re.match(r"^'((?:[^']|'')*)'(?:!(.*))?$", a1)
    if not m:
        m2 = re.match(r"^([^!]+)(?:!(.*))?$", a1)
        title, rest = m2.group(1), m2.group(2)
    else:
        title, rest = m.group(1).replace("''", "'"), m.group(2)
    if not rest:
        return title, 0, None, 0, None
    parts = rest.split(":")

    def cell(text):
        mm = re.match(r"^([A-Z]*)(\d*)$", text)
        assert mm, a1
        return (col_index(mm.group(1)) if mm.group(1) else None), (int(mm.group(2)) - 1 if mm.group(2) else None)

    c1, r1 = cell(parts[0])
    c2, r2 = cell(parts[1]) if len(parts) > 1 else (c1, r1)
    return title, (r1 or 0), r2, (c1 or 0), c2


def format_number(value: float, fmt: dict | None) -> str:
    if fmt and fmt.get("type") in ("DATE", "DATE_TIME", "TIME"):
        ms = int(round(float(value) * 86_400_000))
        from datetime import timedelta

        w = _EPOCH + timedelta(milliseconds=ms)
        out, last_hour = [], False
        for tok in re.findall(r"yyyy|mm|m|dd|d|hh|h|ss|.", fmt.get("pattern") or ""):
            if tok == "yyyy":
                out.append(f"{w.year:04d}")
            elif tok in ("mm", "m"):
                v = w.minute if last_hour else w.month
                out.append(f"{v:02d}" if tok == "mm" else str(v))
            elif tok == "dd":
                out.append(f"{w.day:02d}")
            elif tok == "d":
                out.append(str(w.day))
            elif tok == "hh":
                out.append(f"{w.hour:02d}")
            elif tok == "h":
                out.append(str(w.hour))
            elif tok == "ss":
                out.append(f"{w.second:02d}")
            else:
                out.append(tok)
            if tok in ("hh", "h"):
                last_hour = True
            elif tok in ("yyyy", "dd", "d"):
                last_hour = False
        return "".join(out)
    return js_string(value)


@dataclass
class Cell:
    value: Any = None
    fmt: dict | None = None

    def formatted(self):
        if self.value is None or self.value == "":
            return ""
        if isinstance(self.value, bool):
            return "TRUE" if self.value else "FALSE"
        if isinstance(self.value, (int, float)):
            return format_number(self.value, self.fmt)
        return str(self.value)


@dataclass
class FakeTab:
    title: str
    sheet_id: int
    rows: list[list[Cell]] = field(default_factory=list)
    row_count: int = 1000
    col_count: int = 26

    def cell(self, r: int, c: int) -> Cell:
        while len(self.rows) <= r:
            self.rows.append([])
        row = self.rows[r]
        while len(row) <= c:
            row.append(Cell())
        return row[c]

    def get(self, r: int, c: int) -> Cell:
        if r < len(self.rows) and c < len(self.rows[r]):
            return self.rows[r][c]
        return Cell()

    def last_row(self) -> int:
        """値のある最後の行（1 始まり。無ければ 0）。"""
        for r in range(len(self.rows) - 1, -1, -1):
            if any(not (c.value is None or c.value == "") for c in self.rows[r]):
                return r + 1
        return 0

    def values(self, unformatted=True) -> list[list]:
        """値の 2 次元配列（行末の空と末尾の空行を除く。見るための小道具）。"""
        return _trim([[(c.value if unformatted else c.formatted()) for c in row] for row in self.rows])

    def row_values(self, r1: int) -> list:
        """1 始まりの行の値（行末の空を除く）。"""
        if r1 - 1 >= len(self.rows):
            return []
        row = [c.value for c in self.rows[r1 - 1]]
        while row and (row[-1] is None or row[-1] == ""):
            row.pop()
        return row


def _blank(v):
    return v is None or v == ""


def _trim(grid: list[list]) -> list[list]:
    out = []
    for row in grid:
        row = [("" if v is None else v) for v in row]
        while row and _blank(row[-1]):
            row.pop()
        out.append(row)
    while out and not out[-1]:
        out.pop()
    return out


@dataclass
class FakeBook:
    sid: str
    tz: str = "Asia/Tokyo"
    tabs: dict[str, FakeTab] = field(default_factory=dict)
    title: str = ""  # スプレッドシートの名前（空なら sid）

    def add_tab(self, title: str, values: list[list], formats: dict | None = None, row_count: int | None = None,
                col_count: int | None = None) -> FakeTab:
        tab = FakeTab(title=title, sheet_id=100 + len(self.tabs))
        for r, row in enumerate(values):
            for c, v in enumerate(row):
                cell = tab.cell(r, c)
                cell.value = v
        for (r, c), fmt in (formats or {}).items():
            tab.cell(r, c).fmt = fmt
        tab.row_count = row_count if row_count is not None else max(len(values) + 5, 10)
        tab.col_count = col_count if col_count is not None else max([len(r) for r in values] + [1]) + 2
        self.tabs[title] = tab
        return tab


class FakeResp:
    def __init__(self, status: int, body: Any):
        self.status_code = status
        self._body = body
        self.text = json.dumps(body, ensure_ascii=False, default=str)
        self.content = self.text.encode()

    def json(self):
        return self._body


class FakeSheets:
    def __init__(self):
        self.books: dict[str, FakeBook] = {}
        self.calls: list[tuple[str, str, Any, Any]] = []
        self.fail_queue: list[int] = []  # 次の呼び出しを失敗させる状態コード
        self.fail_write_rows: set[tuple[str, int]] = set()  # (タブ, 1 始まりの行) に触れる updateCells を 400 に

    def add_book(self, sid: str, tz: str = "Asia/Tokyo") -> FakeBook:
        book = FakeBook(sid=sid, tz=tz)
        self.books[sid] = book
        return book

    def api(self, settings) -> SheetsApi:
        return SheetsApi(lambda: object(), session_factory=lambda creds: self, app_settings=settings,
                         sleep=lambda s: None)

    # ----- 呼び出しの記録 -----

    def writes(self) -> list[tuple[str, str, Any]]:
        return [(m, u, b) for (m, u, p, b) in self.calls if m == "POST"]

    def batch_requests(self) -> list[dict]:
        out = []
        for m, u, b in self.writes():
            if u.endswith(":batchUpdate") and "/values:" not in u:
                out.extend(b["requests"])
        return out

    # ----- 本物の AuthorizedSession の代わり -----

    def request(self, method, url, params=None, json=None, timeout=None):
        self.calls.append((method, url, params, json))
        if self.fail_queue:
            return FakeResp(self.fail_queue.pop(0), {"error": "injected"})
        assert url.startswith(API_BASE + "/"), url
        path = url[len(API_BASE) + 1:]
        try:
            if method == "GET" and "/" not in path and ":" not in path:
                return self._meta(urllib.parse.unquote(path), params)
            sid, _, rest = path.partition("/")
            if method == "GET" and rest == "values:batchGet":
                return self._batch_get(sid, params)
            if method == "POST" and rest == "values:batchUpdate":
                return self._values_update(sid, json)
            if method == "POST" and rest == "values:batchClear":
                return self._values_clear(sid, json)
            if method == "POST" and path.endswith(":batchUpdate"):
                return self._batch_update(path[: -len(":batchUpdate")], json)
        except _Http as e:
            return FakeResp(e.status, {"error": e.message})
        return FakeResp(404, {"error": f"unknown {method} {url}"})

    def _book(self, sid) -> FakeBook:
        if sid not in self.books:
            raise _Http(404, f"no spreadsheet {sid}")
        return self.books[sid]

    def _tab(self, book: FakeBook, title: str) -> FakeTab:
        if title not in book.tabs:
            raise _Http(400, f"Unable to parse range: {title}")
        return book.tabs[title]

    def _meta(self, sid, params):
        book = self._book(sid)
        plist = list(params.items()) if isinstance(params, dict) else list(params or [])
        sheets = []
        for i, tab in enumerate(book.tabs.values()):
            sheets.append({"properties": {"sheetId": tab.sheet_id, "title": tab.title, "index": i,
                                          "gridProperties": {"rowCount": tab.row_count,
                                                             "columnCount": tab.col_count}}})
        meta = {"properties": {"title": book.title or sid, "timeZone": book.tz, "locale": "en_US"}, "sheets": sheets}
        if ("includeGridData", "true") in plist:
            ranges = [v for k, v in plist if k == "ranges"]
            out = []
            for a1 in ranges:
                title, r0, r1, c0, c1 = parse_a1(a1)
                tab = self._tab(book, title)
                r1 = tab.row_count - 1 if r1 is None else r1
                c1 = tab.col_count - 1 if c1 is None else c1
                row_data = []
                for r in range(r0, min(r1, len(tab.rows) - 1) + 1):
                    vals = []
                    for c in range(c0, c1 + 1):
                        cell = tab.get(r, c)
                        cd: dict = {}
                        if not _blank(cell.value):
                            v = cell.value
                            key = ("boolValue" if isinstance(v, bool) else "numberValue"
                                   if isinstance(v, (int, float)) else "stringValue")
                            cd["userEnteredValue"] = {key: v}
                            cd["formattedValue"] = cell.formatted()
                        if cell.fmt:
                            cd["userEnteredFormat"] = {"numberFormat": dict(cell.fmt)}
                        vals.append(cd)
                    row_data.append({"values": vals})
                out.append({"properties": {"sheetId": tab.sheet_id, "title": tab.title},
                            "data": [{"startRow": r0, "startColumn": c0, "rowData": row_data}]})
            meta["sheets"] = out
        return FakeResp(200, meta)

    def _batch_get(self, sid, params):
        book = self._book(sid)
        plist = list(params or [])
        render = dict((k, v) for k, v in plist if k != "ranges").get("valueRenderOption", "FORMATTED_VALUE")
        ranges = [v for k, v in plist if k == "ranges"]
        out = []
        for a1 in ranges:
            title, r0, r1, c0, c1 = parse_a1(a1)
            tab = self._tab(book, title)
            r_end = min(len(tab.rows) - 1, r1 if r1 is not None else 10 ** 9)
            grid = []
            for r in range(r0, r_end + 1):
                row = tab.rows[r] if r < len(tab.rows) else []
                c_end = min(len(row) - 1, c1 if c1 is not None else 10 ** 9)
                vals = []
                for c in range(c0, c_end + 1):
                    cell = row[c]
                    vals.append(cell.formatted() if render == "FORMATTED_VALUE" else
                                ("" if cell.value is None else cell.value))
                grid.append(vals)
            out.append({"range": a1, "values": _trim(grid)})
        return FakeResp(200, {"valueRanges": out})

    def _user_entered(self, v):
        if not isinstance(v, str):
            return v, None
        if v.startswith("'"):
            return v[1:], None  # 先頭の ' は「文字として入れる」印（値には残らない）
        if v[:1] == "=" or (v[:1] in "+-" and not re.fullmatch(r"[+-]\d*\.?\d+", v.strip())):
            return FormulaCell(v if v[:1] == "=" else "=" + v), None  # 本物と同じく式になる
        s = v.strip()
        if s == "":
            return "", None
        if re.fullmatch(r"[+-]?\d+", s):
            return int(s), None
        if re.fullmatch(r"[+-]?\d*\.\d+", s):
            return float(s), None
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
            return serial(s), dict(DATE_FMT)
        if re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", s):
            return serial(s), {"type": "DATE_TIME", "pattern": "yyyy-mm-dd hh:mm:ss"}
        if s.upper() in ("TRUE", "FALSE"):
            return s.upper() == "TRUE", None
        return v, None

    def _values_update(self, sid, body):
        book = self._book(sid)
        option = body["valueInputOption"]
        for item in body["data"]:
            title, r0, _r1, c0, _c1 = parse_a1(item["range"])
            tab = self._tab(book, title)
            for dr, row in enumerate(item["values"]):
                for dc, v in enumerate(row):
                    r, c = r0 + dr, c0 + dc
                    if r >= tab.row_count:
                        tab.row_count = r + 1  # 本物と同じく行を足す
                    if c >= tab.col_count:
                        tab.col_count = c + 1
                    cell = tab.cell(r, c)
                    if option == "USER_ENTERED":
                        val, fmt = self._user_entered(v)
                        cell.value = val
                        if fmt:
                            cell.fmt = fmt
                    else:
                        cell.value = v
        return FakeResp(200, {"totalUpdatedCells": 1})

    def _values_clear(self, sid, body):
        book = self._book(sid)
        for a1 in body["ranges"]:
            title, r0, r1, c0, c1 = parse_a1(a1)
            tab = self._tab(book, title)
            for r in range(r0, min(len(tab.rows) - 1, r1 if r1 is not None else 10 ** 9) + 1):
                row = tab.rows[r]
                for c in range(c0, min(len(row) - 1, c1 if c1 is not None else 10 ** 9) + 1):
                    row[c].value = None
        return FakeResp(200, {})

    def _batch_update(self, sid, body):
        book = self._book(sid)
        by_id = {t.sheet_id: t for t in book.tabs.values()}
        # 先にすべて確かめる（本物と同じく、1 つでも不正なら何も書かない）
        plan = []
        for req in body["requests"]:
            (kind, b), = req.items()
            if kind == "appendDimension":
                tab = by_id[b["sheetId"]]
                if b["dimension"] != "ROWS":
                    raise _Http(400, "only rows")
                plan.append(("append", tab, b))
            elif kind == "updateCells":
                start = b.get("start") or {}
                rng = b.get("range") or {}
                tab = by_id[(start or rng)["sheetId"]]
                r0 = start.get("rowIndex", rng.get("startRowIndex", 0))
                c0 = start.get("columnIndex", rng.get("startColumnIndex", 0))
                rows = b.get("rows", [])
                for dr in range(len(rows)):
                    if (tab.title, r0 + dr + 1) in self.fail_write_rows:
                        raise _Http(400, f"injected failure at {tab.title}!{r0 + dr + 1}")
                plan.append(("cells", tab, (r0, c0, rows, b["fields"])))
            else:
                raise _Http(400, f"unsupported request {kind}")
        rows_after = {t.title: t.row_count for t in book.tabs.values()}
        for kind, tab, b in plan:
            if kind == "append":
                rows_after[tab.title] += int(b["length"])
            else:
                r0, _c0, rows, _f = b
                if r0 + len(rows) > rows_after[tab.title]:
                    raise _Http(400, f"Attempting to write row {r0 + len(rows)} beyond grid {tab.title}")
        for kind, tab, b in plan:
            if kind == "append":
                tab.row_count += int(b["length"])
                continue
            r0, c0, rows, fields = b
            for dr, row in enumerate(rows):
                for dc, cd in enumerate(row.get("values", [])):
                    cell = tab.cell(r0 + dr, c0 + dc)
                    if "userEnteredValue" in fields:
                        uv = cd.get("userEnteredValue")
                        cell.value = None if not uv else next(iter(uv.values()))
                    if "numberFormat" in fields:
                        cell.fmt = ((cd.get("userEnteredFormat") or {}).get("numberFormat")) or None
        return FakeResp(200, {"replies": []})


class _Http(Exception):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status
        self.message = message
