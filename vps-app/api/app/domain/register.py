"""38 関数の登録（rpc.load_domain_handlers が register_all(rpc.register) を呼ぶ）。"""
from __future__ import annotations

from collections.abc import Callable

from . import approval, dashboards, masters, ops, requests, secondary, summaries, worklog

MODULES = (masters, dashboards, requests, worklog, approval, secondary, summaries, ops)


def all_handlers() -> dict[str, Callable]:
    out: dict[str, Callable] = {}
    for module in MODULES:
        for name, fn in module.HANDLERS.items():
            if name in out:
                raise RuntimeError(f"二重に登録しています: {name}")
            out[name] = fn
    return out


def register_all(register: Callable) -> list[str]:
    handlers = all_handlers()
    for name, fn in handlers.items():
        register(name, fn)
    return list(handlers)
