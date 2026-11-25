"""`python -m app.worker` で起動したとき、各担当の予定が実際に回る一覧に入ること（2026-11-26 の不具合の再発防止）。"""
import runpy

import pytest


def test_module_entrypoint_delegates_to_app_worker_main(monkeypatch):
    from app import worker

    called = {}

    def fake_main(argv=None):
        called["module"] = worker.__name__
        called["schedules"] = sorted(worker._schedules)

    monkeypatch.setattr(worker, "main", fake_main)
    runpy.run_module("app.worker", run_name="__main__")
    assert called.get("module") == "app.worker"


def test_registrations_land_in_app_worker_schedules():
    from app import worker

    worker.load_registrations()
    names = set(worker._schedules)
    for expected in ("gas_ping", "sheet_intake"):
        assert expected in names, f"{expected} が app.worker._schedules にありません: {sorted(names)}"
