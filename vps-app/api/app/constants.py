"""業務の決まった値（src/config.gs の定数の写し）と、DB の enum の値。

画面の選択肢（REASONS・OT_HOURS・HD_HOURS）はテンプレート変数 REASONS_JSON などで渡す（実装仕様 3-3）。
"""
from __future__ import annotations

# src/config.gs:42-49
REASONS: tuple[str, ...] = (
    "急ぎ作業を要するため",
    "納期遅延解消のため",
    "マシントラブル対応のため",
    "業者対応のため",
    "顧客対応のため",
    "その他（→補足理由必須）",
)
# src/config.gs:51-52
OT_HOURS: tuple[str, ...] = ("0.5", "1.0", "1.5", "2.0", "2.5", "3.0", "3.5", "4.0")
HD_HOURS: tuple[str, ...] = ("半日", "1日")

# src/config.gs:55-62（休憩の時刻窓。(開始時, 開始分, 終了時, 終了分)。9-1 #23 で Sebastian もこれを使う）
BREAK_WINDOWS: tuple[tuple[int, int, int, int], ...] = (
    (10, 0, 10, 10),
    (12, 15, 13, 0),
    (15, 0, 15, 10),
    (17, 10, 17, 20),
    (19, 20, 19, 30),
    (21, 30, 21, 40),
)

# 残業の開始は対象日の 17:20 JST 固定（実装仕様 8-2）
OVERTIME_START_HH_MM = (17, 20)

# ===== シートのタブ名（実装仕様 6-1） =====
TAB_REQUESTS = "Requests"
TAB_WORKLOGS = "WorkLogs"
TAB_BATCHLOGS = "BatchLogs"
TAB_SETTINGS = "Settings"
TAB_DEPT_APPROVERS = "DeptApprovers"
TAB_APPROVER_MAP = "ApproverMap"

# ===== DB の enum（CheckConstraint） =====
REQUEST_TYPES: tuple[str, ...] = ("overtime", "holiday")
REQUEST_STATUSES: tuple[str, ...] = ("submitted", "approved", "canceled")

# WorkLogs の updatedBy の決まった値（同じ文字で書く）
UPDATED_BY_FIXED: tuple[str, ...] = ("formSubmit", "sebastian-agent", "unknown", "delete", "admin-manual",
                                     "break-recalc")

# 仕事（jobs）
JOB_STATES: tuple[str, ...] = ("pending", "running", "done", "failed", "skipped")
JOB_LANES: tuple[str, ...] = ("gas", "default")
PRIORITY_NORMAL = 0
PRIORITY_SEBASTIAN = 10

# 行の反映（sheet_sync_state）
SYNC_STATES: tuple[str, ...] = ("ok", "retry", "blocked")

# 冪等キー
IDEMPOTENCY_STATES: tuple[str, ...] = ("running", "done")
