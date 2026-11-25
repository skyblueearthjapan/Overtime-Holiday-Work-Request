"""シートへの行単位の反映（決定 A・実装仕様 6-2〜6-6。担当 D）。

worker の行の反映のレーン（全体で 1 本。2 秒ごと）が run_once(session_factory, limit=50) を呼ぶ。

1. sheet_sync_state の desired_version > synced_version の行を選ぶ（blocked と再試行の待ちは除く）。
   新しい行（sheet_row が NULL）は表ごとに seq 順（前の新しい行が待ちなら後ろも待つ＝作った順を守る）、
   ほかは first_dirty_at 順。最大 limit 件。
2. DB の今の値を読み、synced_values（最後にシートへ書いた値）と比べて違う列だけを書く
   （GAS が書く PDF の 3 列を古い値で上書きしない。6-5）。
3. 書く前に見出し（WorkLogs は 2 行目）を読み、決まりの見出しが 1 つでも無ければレーンを止める（列を足さない）。
   スプレッドシートの時間帯が DB_SHEET_TIMEZONE と違えばレーンを止める。安全装置の拒否もレーンを止める。
   止めたことは sync_runs（kind='row_sync'）に残し、5 分ごとに確かめ直す（読むだけ）。
4. 今ある行は sheet_row に書く。書く前にその行の ID の列（Requests・WorkLogs は requestId、BatchLogs は
   実行日時とバッチ名）を読んで確かめ、違えば ID の列から探し直して sheet_row を直す。見つからなければ
   書かずに state='blocked'。新しい行は「最後のデータ行の次」に作った順で書く（行数が足りなければ appendDimension）。
   Requests の新しい行だけは、先に requestId を values.append（RAW・OVERWRITE）で書いて Google に行を決めさせ、
   残りの列をその行へ書く（reserve_new_rows。Sebastian が同じタブへ直接追記するのと行を取り合わない。決定 #26）。
   DB に無い行（Sebastian が足してまだ取り込んでいない行）は、どの DB の行の位置にもならないので書かない。
5. 1 回の spreadsheets.batchUpdate（updateCells だけ）でタブごとにまとめて書く。失敗したら行ごとに書き直し、
   失敗した行だけ 10 秒・30 秒・1 分・2 分・5 分、以後 5 分ごとに再試行する。
6. 書けたら synced_version＝読んだ時点の row_version、synced_values を更新する（GAS の仕事は requires で待つ）。

行・列・タブの挿入・削除・並べ替え・作り直しは出さない（SheetsApi と safety が拒む）。
"""
from __future__ import annotations

import logging
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import func, select, update
from sqlalchemy.orm import Session

from app import sheet_queue, timeutil
from app.jsutil import js_string, normalize_gas
from app.models import SYNCED_TABLES, SheetSyncState, SyncRun
from app.safety import UnsafeWriteError
from app.services import sheet_codec as codec
from app.services import write_gate
from app.services.sheets_api import SheetsApi, SheetsError, column_letter, quote_sheet, sheet_id_by_title, spreadsheet_timezone

log = logging.getLogger(__name__)

RETRY_DELAYS = (10, 30, 60, 120, 300)  # 秒。以後は 300 秒ごと
RECHECK_SECONDS = 300  # 止めたレーンを確かめ直す間隔
TZ_CHECK_SECONDS = 3600
BACKOFF_SECONDS = 30  # 読み取りの失敗・委任が無いなど、レーン全体の一時的な失敗の後に待つ
REF_WINDOW = 20  # 表示形式を写すために読む、最後のデータ行からの行数
REF_CACHE_SECONDS = 600
LATE_SECONDS = 300  # これより遅れた行は「遅れ」
GRID_FIELDS = "sheets(data(startRow,startColumn,rowData(values(userEnteredValue,userEnteredFormat(numberFormat)))))"
RUN_KIND = "row_sync"


GATE_RECHECK_SECONDS = 30  # 書き込みの関所で止めたときは早めに確かめ直す


class LaneHalt(RuntimeError):
    """レーンを止める理由（見出しの違い・時間帯の違い・安全装置の拒否）。再試行しない。"""


class GateClosed(LaneHalt):
    """書き込みの関所（write_gate）で止めた。"""


# ===== レーンの状態（worker の 1 つのプロセスの中だけ。止めたことは sync_runs にも残す） =====


@dataclass
class _Lane:
    halted: str = ""
    recheck_at: datetime | None = None
    tz_checked_at: datetime | None = None
    refs: dict[str, tuple[datetime, dict]] = field(default_factory=dict)
    backoff_until: datetime | None = None


_lane = _Lane()


def reset_lane() -> None:
    """試験用・手での再開用（止めた状態と覚えた表示形式を捨てる）。"""
    global _lane
    _lane = _Lane()


def lane_state() -> dict:
    return {"halted": _lane.halted, "recheckAt": _lane.recheck_at.isoformat() if _lane.recheck_at else None}


def _default_api(session_factory: Callable[[], Session], app_settings: Any) -> SheetsApi:
    with session_factory() as db:
        return SheetsApi.for_db(db, app_settings)


# 試験で差し替える（偽の Sheets API を返す関数）
api_factory: Callable[[Callable[[], Session], Any], SheetsApi] = _default_api


def _settings(app_settings: Any):
    if app_settings is not None:
        return app_settings
    from app.config import get_settings

    return get_settings()


# ===== タブの今の形 =====


@dataclass
class TabContext:
    spec: codec.TableSpec
    sheet_id: int
    col: dict[str, int]  # 見出し → 0 始まりの列
    id_rows: dict[int, tuple]  # 1 始まりの行 → ID の値（比べる形）
    last_data_row: int
    row_count: int
    col_count: int
    refs: dict[str, dict] | None = None  # 見出し → {"kind", "format"}（最後のデータ行あたりの型と表示形式）

    def rows_with_id(self, key: tuple) -> list[int]:
        return sorted(r for r, v in self.id_rows.items() if _same_id(self.spec, v, key))


def _id_cell(spec: codec.TableSpec, header: str, value: Any, tz) -> Any:
    rule = spec.rule(header)
    if rule.kind == "datetime":
        if codec.is_number(value):
            return codec.serial_to_datetime(value, tz)
        return codec.parse_jst_datetime(js_string(value)) if not codec.is_blank(value) else None
    return normalize_gas(value)


def db_id(spec: codec.TableSpec, values: dict) -> tuple:
    """DB の行（見出し → 論理の値）から ID の形。"""
    out = []
    for h in spec.id_headers:
        v = values.get(h)
        out.append(v if spec.rule(h).kind == "datetime" else normalize_gas(v))
    return tuple(out)


def _same_id(spec: codec.TableSpec, a: tuple, b: tuple) -> bool:
    for h, x, y in zip(spec.id_headers, a, b):
        rule = spec.rule(h)
        if rule.kind == "datetime":
            if x is None or y is None:
                if not (x is None and y is None):
                    return False
                continue
            if abs((x - y).total_seconds()) >= 1.0:
                return False
        elif x != y:
            return False
    return True


def _blank_id(key: tuple) -> bool:
    return all(v is None or v == "" for v in key)


def check_timezone(api: SheetsApi, sid: str, settings) -> dict:
    meta = api.get_spreadsheet(sid)
    tz = spreadsheet_timezone(meta)
    if tz != settings.db_sheet_timezone:
        raise LaneHalt(f"スプレッドシートの時間帯（{tz or '空'}）が DB_SHEET_TIMEZONE（{settings.db_sheet_timezone}）と"
                       "違うため、行の反映を止めました。")
    return meta


def load_tab(api: SheetsApi, sid: str, spec: codec.TableSpec, meta: dict, tz, *, need_tail: bool) -> TabContext:
    """見出しを確かめ、ID の列と最後のデータ行を読む。見出しが欠けていれば LaneHalt。"""
    sheet_id = sheet_id_by_title(meta, spec.tab)
    if sheet_id is None:
        raise LaneHalt(f"タブ「{spec.tab}」がありません（タブは作りません）。行の反映を止めました。")
    grid = {}
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == spec.tab:
            grid = s["properties"].get("gridProperties") or {}
    row_count = int(grid.get("rowCount") or 0)
    col_count = int(grid.get("columnCount") or 26)
    q = quote_sheet(spec.tab)
    header_rows = api.get_values(sid, f"{q}!{spec.header_row}:{spec.header_row}", "UNFORMATTED_VALUE")
    header = header_rows[0] if header_rows else []
    missing = codec.missing_headers(spec, header)
    if missing:
        raise LaneHalt(f"{spec.tab}: 決まりの見出し {missing} がありません（列は足しません）。行の反映を止めました。")
    col = {h: i for h, i in codec.header_index(header).items() if h in spec.headers}
    start = spec.data_start_row
    ranges = [f"{q}!{column_letter(col[h])}{start}:{column_letter(col[h])}" for h in spec.id_headers]
    columns = api.batch_get_values(sid, ranges, "UNFORMATTED_VALUE")
    id_rows: dict[int, tuple] = {}
    height = max([len(c) for c in columns] + [0])
    last_id_row = spec.header_row
    for i in range(height):
        values = tuple(_id_cell(spec, h, (columns[k][i][0] if i < len(columns[k]) and columns[k][i] else ""), tz)
                       for k, h in enumerate(spec.id_headers))
        id_rows[start + i] = values
        if not _blank_id(values):
            last_id_row = start + i
    last_data_row = last_id_row
    if need_tail and row_count > last_id_row:
        tail = api.get_values(sid, f"{q}!A{last_id_row + 1}:{column_letter(max(col_count, 1) - 1)}{row_count}",
                              "UNFORMATTED_VALUE")
        for i, row in enumerate(tail):
            if any(not codec.is_blank(v) for v in row):
                last_data_row = last_id_row + 1 + i
    return TabContext(spec=spec, sheet_id=sheet_id, col=col, id_rows=id_rows, last_data_row=last_data_row,
                      row_count=row_count, col_count=col_count)


def load_refs(api: SheetsApi, sid: str, ctx: TabContext, now: datetime) -> dict[str, dict]:
    """最後のデータ行あたりの、日付・日時の列の型と表示形式（新しい行の auto と表示形式に使う）。"""
    cached = _lane.refs.get(ctx.spec.tab)
    if cached and (now - cached[0]).total_seconds() < REF_CACHE_SECONDS:
        ctx.refs = cached[1]
        return ctx.refs
    refs: dict[str, dict] = {}
    if ctx.last_data_row >= ctx.spec.data_start_row:
        lo = max(ctx.spec.data_start_row, ctx.last_data_row - REF_WINDOW + 1)
        last_col = column_letter(max(ctx.col.values()))
        meta = api.get_spreadsheet_row_metadata(sid, [f"{quote_sheet(ctx.spec.tab)}!A{lo}:{last_col}{ctx.last_data_row}"],
                                                GRID_FIELDS)
        rows = []
        for s in meta.get("sheets", []):
            for d in s.get("data", []):
                rows.extend(r.get("values", []) for r in d.get("rowData", []))
        for rule in ctx.spec.columns:
            if not rule.date_like:
                continue
            c = ctx.col[rule.header]
            ref: dict = {}
            for vals in reversed(rows):
                cd = vals[c] if c < len(vals) else {}
                uv = cd.get("userEnteredValue") or {}
                if not uv:
                    continue
                kind = codec.SERIAL if "numberValue" in uv else codec.STRING
                ref.setdefault("kind", kind)
                nf = (cd.get("userEnteredFormat") or {}).get("numberFormat") or {}
                if kind == codec.SERIAL and nf.get("type") in codec.DATE_FORMAT_TYPES:
                    ref["format"] = {"type": nf.get("type"), "pattern": nf.get("pattern", "")}
                    break
            refs[rule.header] = ref
    _lane.refs[ctx.spec.tab] = (now, refs)
    ctx.refs = refs
    return refs


# ===== 書く要求 =====


@dataclass
class RowWork:
    table: str
    pk: str
    version: int
    values: dict  # 見出し → 論理の値（DB の今）
    kinds: dict  # 行の cell_kinds
    synced: dict  # synced_values（KINDS_KEY を含む）
    sheet_row: int | None
    seq: int
    first_dirty_at: datetime | None
    attempts: int
    target_row: int | None = None
    adopted: bool = False
    deferred: bool = False  # 前の新しい行が書けなかったので、この回は書かない（作った順を守る）
    written: dict | None = None  # 書いた列の synced_values（見出し → 値、KINDS_KEY → 型）
    headers: list[str] | None = None  # 書く見出し（None は差分で決める）
    error: str = ""
    blocked: str = ""


def changed_headers(spec: codec.TableSpec, work: RowWork) -> list[str]:
    """synced_values と違う列（新しい行は値のある列すべて）。"""
    synced_kinds = work.synced.get(codec.KINDS_KEY) or {}
    out = []
    for rule in spec.columns:
        cur = codec.canonical(rule, work.values.get(rule.header))
        if work.sheet_row is None and not work.adopted:
            if cur != "":
                out.append(rule.header)
            continue
        if (work.adopted or rule.header not in work.synced
                or codec.canonical_stored(rule, work.synced.get(rule.header)) != cur):
            out.append(rule.header)
            continue
        if rule.date_like and rule.kind != "jst_text" and cur != "":
            want = work.kinds.get(rule.header)
            have = synced_kinds.get(rule.header)
            if want in (codec.SERIAL, codec.STRING) and have in (codec.SERIAL, codec.STRING) and want != have:
                out.append(rule.header)
    return out


def resolve_kind(rule: codec.ColumnRule, kinds: dict, synced_kinds: dict, refs: dict | None) -> str:
    kind = codec.expected_kind(rule, kinds, synced_kinds)
    if kind == codec.AUTO:
        kind = ((refs or {}).get(rule.header) or {}).get("kind") or codec.SERIAL
    return kind


def cell_requests(ctx: TabContext, row: int, headers: list[str], values: dict, kinds: dict, synced_kinds: dict,
                  tz) -> tuple[list[dict], dict]:
    """1 行の書く列 → updateCells の要求（セルごと）と、書いた型（見出し → kind）。"""
    reqs, written_kinds = [], {}
    for h in headers:
        rule = ctx.spec.rule(h)
        kind = resolve_kind(rule, kinds, synced_kinds, ctx.refs) if rule.date_like else None
        fmt = ((ctx.refs or {}).get(h) or {}).get("format")
        enc = codec.encode_cell(rule, values.get(h), kind or codec.STRING, tz, fmt)
        if rule.date_like and not codec.is_blank(values.get(h)):
            written_kinds[h] = enc["kind"]
        reqs.append({"updateCells": {
            "start": {"sheetId": ctx.sheet_id, "rowIndex": row - 1, "columnIndex": ctx.col[h]},
            "rows": [{"values": [enc["cell"]]}],
            "fields": enc["fields"],
        }})
    return reqs, written_kinds


def _needs_refs(spec: codec.TableSpec, works: list[RowWork]) -> bool:
    for w in works:
        for h in (w.headers or []):
            rule = spec.rule(h)
            if rule.date_like and rule.kind != "jst_text" and not codec.is_blank(w.values.get(h)):
                return True
    return False


# ===== 1 回分 =====


def _model_values(spec: codec.TableSpec, row: Any) -> dict:
    return {c.header: codec.db_value(c, row) for c in spec.columns}


def _pk_of(table: str, row: Any) -> str:
    return str(row.request_id if table == "requests" else row.id)


def select_work(db: Session, limit: int, now: datetime) -> tuple[list[RowWork], list[tuple[str, str, str]]]:
    """反映する行（DB の今の値）と、行が見つからない状態（表・pk・理由）を選ぶ。"""
    states = list(db.scalars(select(SheetSyncState).where(
        SheetSyncState.desired_version > SheetSyncState.synced_version,
        SheetSyncState.state != "blocked")))
    missing: list[tuple[str, str, str]] = []
    new_by_table: dict[str, list[RowWork]] = {}
    existing: list[RowWork] = []
    ready: dict[tuple[str, str], bool] = {}
    for st in states:
        model = SYNCED_TABLES.get(st.table_name)
        if model is None:
            continue
        key = st.pk if st.table_name == "requests" else int(st.pk)
        row = db.get(model, key)
        if row is None:
            missing.append((st.table_name, st.pk, "DB に行がありません"))
            continue
        spec = codec.TABLES[st.table_name]
        work = RowWork(table=st.table_name, pk=st.pk, version=int(row.row_version), values=_model_values(spec, row),
                       kinds=dict(row.cell_kinds or {}), synced=dict(st.synced_values or {}), sheet_row=row.sheet_row,
                       seq=int(st.seq or 0), first_dirty_at=st.first_dirty_at, attempts=int(st.attempts or 0))
        ready[(st.table_name, st.pk)] = st.next_try_at is None or st.next_try_at <= now
        if row.sheet_row is None:
            new_by_table.setdefault(st.table_name, []).append(work)
        else:
            existing.append(work)
    chosen: list[RowWork] = []
    for table in codec.TABLE_ORDER:
        for w in sorted(new_by_table.get(table, []), key=lambda w: w.seq):
            if not ready[(w.table, w.pk)]:
                break  # 前の新しい行が待ちなら、後ろも待つ（作った順を守る）
            chosen.append(w)
    far = datetime.max.replace(tzinfo=timeutil.UTC)
    existing = [w for w in existing if ready[(w.table, w.pk)]]
    existing.sort(key=lambda w: (w.first_dirty_at or far, w.seq))
    chosen.extend(existing)
    return chosen[:limit], missing


def _claimed_rows(db: Session, table: str) -> dict[int, tuple[str, tuple]]:
    """DB で sheet_row を持つ行（行番号 → (pk, ID の形)）。"""
    model = SYNCED_TABLES[table]
    spec = codec.TABLES[table]
    out: dict[int, tuple[str, tuple]] = {}
    for row in db.scalars(select(model).where(model.sheet_row.is_not(None))):
        values = {h: codec.db_value(spec.rule(h), row) for h in spec.id_headers}
        out[int(row.sheet_row)] = (_pk_of(table, row), db_id(spec, values))
    return out


def place_rows(ctx: TabContext, works: list[RowWork], claimed: dict[int, tuple[str, tuple]]) -> None:
    """今ある行の位置を確かめ（違えば探し直す）、新しい行の位置を決める。

    探し直すとき、ほかの DB の行が持つ行番号でも、シートのその行の ID がその DB の行と違えば（人が行を挿入して
    ずれた）空いているとみなす。
    """
    spec = ctx.spec
    next_row = ctx.last_data_row + 1
    taken = dict(claimed)

    def owned_by_other(r: int, pk: str) -> bool:
        owner = taken.get(r)
        if owner is None or owner[0] == pk:
            return False
        return r not in ctx.id_rows or _same_id(spec, owner[1], ctx.id_rows[r])

    for w in works:
        key = db_id(spec, w.values)
        if w.sheet_row is not None and _blank_id(key) and w.sheet_row <= ctx.last_data_row:
            # ID が空の行は、行がずれても見分けられない（別の空の ID の行に書く恐れ）ので書かない（データ安全 低-2）
            w.blocked = f"{spec.tab} の {w.sheet_row} 行目は ID が空のため、位置を確かめられません（書きません）"
            continue
        if w.sheet_row is not None:
            if w.sheet_row in ctx.id_rows and _same_id(spec, ctx.id_rows[w.sheet_row], key):
                w.target_row = w.sheet_row
                continue
            if w.sheet_row > ctx.last_data_row and _blank_id(key):
                w.target_row = w.sheet_row
                continue
            found = [r for r in ctx.rows_with_id(key) if not owned_by_other(r, w.pk)] if not _blank_id(key) else []
            if not found:
                w.blocked = (f"{spec.tab} の {w.sheet_row} 行目の ID が DB と違い、ID の列にも見つかりません"
                             f"（{', '.join(js_string(k) for k in key)}）")
                continue
            w.target_row = found[0]
            taken[w.target_row] = (w.pk, key)
            continue
        # 新しい行: 書きかけで落ちた行（同じ ID でどの DB の行のものでもない行）があれば、それを使う
        found = [r for r in ctx.rows_with_id(key) if r not in taken] if not _blank_id(key) else []
        if found:
            w.target_row = found[0]
            w.adopted = True
        else:
            w.target_row = next_row
            next_row += 1
        taken[w.target_row] = (w.pk, key)


def reserve_new_rows(api: SheetsApi, sid: str, ctx: TabContext, works: list[RowWork]) -> int:
    """Requests の新しい行の位置を、requestId だけを values.append（RAW・OVERWRITE）で書いて Google に決めさせる。

    Sebastian（外のシステム）は同じタブへ values.append（INSERT_ROWS）で直接追記する（決定 #26）。「最後のデータ行の
    次」を読んでから updateCells で書くと、その間に Sebastian が同じ行へ追記したとき、その行を上書きしてしまう。
    append はシートの側で「表の次の行」を決めて書くので、2 つの追記が同じ行を取り合わない。残りの列は、この後の
    updateCells で同じ行へ書く。間で落ちても、次の回に place_rows が requestId で見つけて使う（adopted）。
    requestId が A 列でないときは今までどおり（最後のデータ行の次）。取った行の数を返す。
    """
    spec = ctx.spec
    if spec.table != "requests" or ctx.col.get(spec.id_headers[0]) != 0:
        return 0
    new = [w for w in works if not w.blocked and w.sheet_row is None and not w.adopted]
    ids = [js_string(w.values.get(spec.id_headers[0])) for w in new]
    if not new or any(i == "" for i in ids):
        return 0
    first = api.append_values(sid, spec.tab, [[i] for i in ids], purpose="row_sync")
    for k, w in enumerate(new):
        w.target_row = first + k
        ctx.id_rows[w.target_row] = (normalize_gas(ids[k]),)
    ctx.row_count = max(ctx.row_count, first + len(new) - 1)
    ctx.last_data_row = max(ctx.last_data_row, first + len(new) - 1)
    return len(new)


def _retry_delay(attempts: int) -> int:
    return RETRY_DELAYS[min(attempts, len(RETRY_DELAYS)) - 1] if attempts >= 1 else RETRY_DELAYS[0]


def _record_halt(session_factory, reason: str, recheck_seconds: int = RECHECK_SECONDS) -> None:
    now = timeutil.now_utc()
    first = _lane.halted != reason
    _lane.halted = reason
    _lane.recheck_at = now + timedelta(seconds=recheck_seconds)
    log.error("row_sync halted: %s", reason)
    if first:
        with session_factory() as db:
            db.add(SyncRun(kind=RUN_KIND, started_at=now, finished_at=now, ok=False,
                           summary_json={"halted": True, "reason": reason}))
            db.commit()


def _record_resume(session_factory) -> None:
    now = timeutil.now_utc()
    reason = _lane.halted
    _lane.halted = ""
    _lane.recheck_at = None
    with session_factory() as db:
        db.add(SyncRun(kind=RUN_KIND, started_at=now, finished_at=now, ok=True,
                       summary_json={"resumed": True, "previousReason": reason}))
        db.commit()


def write_rows(api: SheetsApi, sid: str, ctx: TabContext, works: list[RowWork], tz, *, purpose: str) -> None:
    """works（target_row と headers が決まった行）を書く。失敗した行は work.error に理由。UnsafeWriteError は投げる。"""
    plans: list[tuple[RowWork, list[dict], dict]] = []
    for w in works:
        if not w.headers:
            plans.append((w, [], {}))
            continue
        synced_kinds = w.synced.get(codec.KINDS_KEY) or {}
        reqs, kinds = cell_requests(ctx, w.target_row, w.headers, w.values, w.kinds, synced_kinds, tz)
        plans.append((w, reqs, kinds))
    need = max([w.target_row for w, r, _ in plans if r] + [0])
    if need > ctx.row_count:
        api.append_rows(sid, ctx.spec.tab, need - ctx.row_count, purpose="row_sync")
        ctx.row_count = need
    all_reqs = [r for _, reqs, _ in plans for r in reqs]
    ok_all = True
    if all_reqs:
        try:
            api.batch_update(sid, all_reqs, tab=ctx.spec.tab, purpose=purpose)
        except UnsafeWriteError:
            raise
        except SheetsError as error:
            ok_all = False
            log.warning("row_sync batch failed on %s, retrying row by row: %s", ctx.spec.tab, error)
    new_failed = False
    for w, reqs, kinds in plans:
        is_new = w.sheet_row is None and not w.adopted
        if not ok_all and reqs:
            if is_new and new_failed:
                w.deferred = True
                continue
            try:
                api.batch_update(sid, reqs, tab=ctx.spec.tab, purpose=purpose)
            except UnsafeWriteError:
                raise
            except SheetsError as error:
                w.error = str(error)[:1000]
                new_failed = new_failed or is_new
                continue
        snap = {h: codec.canonical(ctx.spec.rule(h), w.values.get(h)) for h in (w.headers or [])}
        synced_kinds = dict(w.synced.get(codec.KINDS_KEY) or {})
        synced_kinds.update(kinds)
        for h in (w.headers or []):
            if codec.is_blank(w.values.get(h)):
                synced_kinds.pop(h, None)
        new_synced = {k: v for k, v in w.synced.items() if k != codec.KINDS_KEY}
        new_synced.update(snap)
        new_synced[codec.KINDS_KEY] = synced_kinds
        w.synced = new_synced
        w.written = {h: snap[h] for h in snap}
        w.written[codec.KINDS_KEY] = {h: synced_kinds.get(h) for h in (w.headers or [])}


def merge_synced(current: dict | None, written: dict) -> dict:
    """今の synced_values に、書いた列だけを重ねる（その間に M が直した PDF の列などを消さない）。"""
    out = dict(current or {})
    kinds = dict(out.get(codec.KINDS_KEY) or {})
    for h, k in (written.get(codec.KINDS_KEY) or {}).items():
        if k is None:
            kinds.pop(h, None)
        else:
            kinds[h] = k
    out.update({h: v for h, v in written.items() if h != codec.KINDS_KEY})
    out[codec.KINDS_KEY] = kinds
    return out


def _save_results(session_factory, works: list[RowWork], missing: list[tuple[str, str, str]], now: datetime) -> dict:
    counts = {"written": 0, "failed": 0, "blocked": 0}
    with session_factory() as db:
        for table, pk, reason in missing:
            st = db.get(SheetSyncState, (table, pk))
            if st is not None:
                st.state, st.last_error = "blocked", reason
                counts["blocked"] += 1
        for w in works:
            st = db.get(SheetSyncState, (w.table, w.pk))
            if st is None:
                continue
            if w.deferred:
                continue
            if w.blocked:
                st.state, st.last_error = "blocked", w.blocked[:1000]
                counts["blocked"] += 1
                continue
            if w.error:
                st.attempts = int(st.attempts or 0) + 1
                st.state = "retry"
                st.last_error = w.error
                st.next_try_at = now + timedelta(seconds=_retry_delay(st.attempts))
                counts["failed"] += 1
                continue
            st.synced_version = max(int(st.synced_version or 0), w.version)
            st.synced_values = merge_synced(st.synced_values, w.written or {})
            st.state, st.attempts, st.last_error = "ok", 0, ""
            st.last_synced_at = now
            if st.desired_version <= st.synced_version:
                st.first_dirty_at = None
                st.next_try_at = None
            else:
                st.next_try_at = now
            if w.target_row is not None and w.target_row != w.sheet_row:
                model = SYNCED_TABLES[w.table]
                pk_col = model.request_id if w.table == "requests" else model.id
                key = w.pk if w.table == "requests" else int(w.pk)
                db.execute(update(model).where(pk_col == key).values(sheet_row=w.target_row))
            counts["written"] += 1
        db.commit()
    return counts


def run_once(session_factory: Callable[[], Session], limit: int = 50, *, api: SheetsApi | None = None,
             app_settings: Any = None) -> dict:
    """行の反映を 1 回。戻り値 {ok, written, failed, blocked, halted, maxLatencySeconds}。"""
    settings = _settings(app_settings)
    now = timeutil.now_utc()
    if _lane.halted and _lane.recheck_at and now < _lane.recheck_at:
        return {"ok": False, "halted": _lane.halted, "written": 0}
    if _lane.backoff_until and now < _lane.backoff_until:
        return {"ok": False, "error": "一時的な失敗の後の待ち", "written": 0}
    was_halted = bool(_lane.halted)
    with session_factory() as db:
        works, missing = select_work(db, limit, now)
    need_tz = _lane.tz_checked_at is None or (now - _lane.tz_checked_at).total_seconds() >= TZ_CHECK_SECONDS
    if not works and not missing and not need_tz and not was_halted:
        return {"ok": True, "written": 0, "failed": 0, "blocked": 0}
    sid = settings.db_spreadsheet_id
    tz = settings.db_sheet_tz
    result: dict = {"ok": True, "written": 0, "failed": 0, "blocked": 0}
    processed: list[RowWork] = []  # 書き終えた表の行（途中で止めても、書いた分は記録する）
    try:
        if works:
            # 書く前の関所（データ安全 高-3）: GAS が移転モードでなければ書かない
            with session_factory() as db:
                allowed, reason = write_gate.check(db, settings, now)
            if not allowed:
                raise GateClosed("書き込みの関所: " + reason)
        if api is None:
            api = api_factory(session_factory, settings)
        meta = check_timezone(api, sid, settings)
        _lane.tz_checked_at = now
        by_table: dict[str, list[RowWork]] = {}
        for w in works:
            by_table.setdefault(w.table, []).append(w)
        for table in codec.TABLE_ORDER:
            group = by_table.get(table)
            if not group:
                continue
            spec = codec.TABLES[table]
            ctx = load_tab(api, sid, spec, meta, tz, need_tail=any(w.sheet_row is None for w in group))
            with session_factory() as db:
                claimed = _claimed_rows(db, table)
            place_rows(ctx, group, claimed)
            reserve_new_rows(api, sid, ctx, group)
            todo = [w for w in group if not w.blocked]
            for w in todo:
                w.headers = changed_headers(spec, w)
            if _needs_refs(spec, todo):
                load_refs(api, sid, ctx, now)
            write_rows(api, sid, ctx, todo, tz, purpose="row_sync")
            processed.extend(group)
    except (LaneHalt, UnsafeWriteError) as halt:
        _record_halt(session_factory, str(halt),
                     GATE_RECHECK_SECONDS if isinstance(halt, GateClosed) else RECHECK_SECONDS)
        done = _save_results(session_factory, processed, [], timeutil.now_utc()) if processed else {"written": 0}
        return {"ok": False, "halted": str(halt), "written": done["written"]}
    except SheetsError as error:
        # 読み取りの失敗などはレーン全体の一時的な失敗（行の試行回数を増やさない）
        log.warning("row_sync: Sheets API failed: %s", error)
        _lane.backoff_until = now + timedelta(seconds=BACKOFF_SECONDS)
        done = _save_results(session_factory, processed, [], timeutil.now_utc()) if processed else {"written": 0}
        return {"ok": False, "error": str(error)[:500], "written": done["written"]}
    except Exception as error:  # noqa: BLE001 - 委任が無いなど。次の回にもう一度
        log.exception("row_sync failed")
        _lane.backoff_until = now + timedelta(seconds=BACKOFF_SECONDS)
        done = _save_results(session_factory, processed, [], timeutil.now_utc()) if processed else {"written": 0}
        return {"ok": False, "error": f"{type(error).__name__}: {error}"[:500], "written": done["written"]}
    if was_halted:
        _record_resume(session_factory)
    counts = _save_results(session_factory, works, missing, timeutil.now_utc())
    result.update(counts)
    late = [(now - w.first_dirty_at).total_seconds() for w in works
            if w.first_dirty_at and not (w.error or w.blocked or w.deferred)]
    result["maxLatencySeconds"] = max(late) if late else 0
    return result


# ===== 管理者の欄（ops_status・healthz）と手での操作 =====


def status(db: Session) -> dict:
    """反映の待ち・遅れ・止めた行・レーンの停止（管理者の欄）。"""
    now = timeutil.now_utc()
    dirty = SheetSyncState.desired_version > SheetSyncState.synced_version
    pending = int(db.scalar(select(func.count()).select_from(SheetSyncState).where(dirty)) or 0)
    oldest = db.scalar(select(func.min(SheetSyncState.first_dirty_at)).where(dirty))
    blocked = [{"table": s.table_name, "pk": s.pk, "error": s.last_error}
               for s in db.scalars(select(SheetSyncState).where(SheetSyncState.state == "blocked").limit(50))]
    retry = int(db.scalar(select(func.count()).select_from(SheetSyncState).where(SheetSyncState.state == "retry")) or 0)
    last = db.scalars(select(SyncRun).where(SyncRun.kind == RUN_KIND).order_by(SyncRun.id.desc()).limit(1)).first()
    halted = bool(last and not last.ok and (last.summary_json or {}).get("halted"))
    lag = (now - oldest).total_seconds() if oldest else 0
    return {"pending": pending, "retry": retry, "blocked": blocked, "oldestDirtySeconds": lag,
            "late": lag > LATE_SECONDS, "halted": halted,
            "haltReason": (last.summary_json or {}).get("reason", "") if halted else ""}


def unblock(db: Session, table: str, pk: str) -> bool:
    """blocked の行をもう一度試す（管理者が行を直した後。commit は呼ぶ側）。"""
    st = db.get(SheetSyncState, (table, str(pk)))
    if st is None or st.state != "blocked":
        return False
    st.state, st.attempts, st.last_error = "ok", 0, ""
    st.next_try_at = timeutil.now_utc()
    return True


def reappend_blocked(db: Session, table: str, pk: str) -> bool:
    """blocked の行（人がシートの行を消したなど）を、シートの末尾に新しい行として書き直す印にする（データ安全 中-2）。

    sheet_row を空にし、synced_values を空にして版を進める。次の行の反映が末尾に全部の列を書く。commit は呼ぶ側。
    戻しの前（計画書 C-2）に blocked を 0 にするために使う。
    """
    st = db.get(SheetSyncState, (table, str(pk)))
    if st is None or st.state != "blocked":
        return False
    model = SYNCED_TABLES[table]
    row = db.get(model, pk if table == "requests" else int(pk))
    if row is None:
        return False
    row.sheet_row = None
    st.state, st.attempts, st.last_error = "ok", 0, ""
    st.synced_values = {}
    st.next_try_at = timeutil.now_utc()
    db.flush()
    sheet_queue.mark_dirty(db, table, str(pk))
    return True


def main(argv: list[str] | None = None) -> int:
    """管理者の手での操作（VPS: docker compose exec worker python -m app.services.row_sync <操作> ...）。

    status                  反映待ち・blocked の一覧
    unblock  <table> <pk>   blocked の行をもう一度試す（行を直した後）
    reappend <table> <pk>   blocked の行を末尾に新しい行として書き直す（行が消された後）
    """
    import argparse
    import json

    from app.db import get_sessionmaker

    parser = argparse.ArgumentParser(prog="python -m app.services.row_sync")
    parser.add_argument("op", choices=["status", "unblock", "reappend"])
    parser.add_argument("table", nargs="?", choices=list(SYNCED_TABLES))
    parser.add_argument("pk", nargs="?")
    args = parser.parse_args(argv)
    factory = get_sessionmaker()
    with factory() as db:
        if args.op == "status":
            print(json.dumps(status(db), ensure_ascii=False, indent=1, default=str))
            return 0
        if not args.table or not args.pk:
            parser.error("table と pk が要ります")
        fn = unblock if args.op == "unblock" else reappend_blocked
        done = fn(db, args.table, args.pk)
        db.commit()
    print("ok" if done else "blocked の行が見つかりません")
    return 0 if done else 1


def register_worker() -> None:
    """行の反映のレーンは worker が run_once を直接呼ぶ（app.worker._resolve_row_sync）。登録するものは無い。"""


if __name__ == "__main__":
    import sys

    sys.exit(main())
