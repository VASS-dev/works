"""Smoke-тесты роутов через TestClient.

Проверяем: базовые страницы отдают 200, формы и API возвращают ожидаемый код,
экспорт xlsx корректно формируется. Не проверяем контент глубоко — это защита
от регрессий при рефакторинге.
"""

from __future__ import annotations

import os
from datetime import date

# Полный bypass BasicAuth на время тестов:
os.environ.pop("APP_PASSWORD", None)

import pytest
from fastapi.testclient import TestClient

from app import main as app_module
from app.db import init_db, make_session_factory
from app.models import Contract, OrderLine, PlanEntry, WorkType


@pytest.fixture()
def client(tmp_path):
    engine = init_db(tmp_path / "t.db", drop_first=True)
    Session = make_session_factory(engine)

    with Session() as s:
        c = Contract(number="Т-001", counterparty="ООО Тест", project="Проект А")
        wt = WorkType(name="Монтаж", unit="шт")
        s.add_all([c, wt])
        s.flush()
        ol = OrderLine(
            contract_id=c.id, work_type_id=wt.id,
            description="Работа 1",
            sum_ordered=1000.0, qty_ordered=10.0,
            sum_done_snapshot=0.0, qty_done_snapshot=0.0,
            sum_remaining_snapshot=1000.0, qty_remaining_snapshot=10.0,
        )
        s.add(ol)
        s.commit()
        ctx = {"contract_id": c.id, "order_line_id": ol.id}

    original = app_module.SessionLocal
    app_module.SessionLocal = Session

    def _override():
        s = Session()
        try:
            yield s
        finally:
            s.close()

    app_module.app.dependency_overrides[app_module.get_session] = _override

    with TestClient(app_module.app) as c:
        c.ctx = ctx  # type: ignore[attr-defined]
        yield c

    app_module.app.dependency_overrides.clear()
    app_module.SessionLocal = original


def test_healthz(client):
    r = client.get("/healthz")
    assert r.status_code == 200
    assert r.json() == {"ok": True}


def test_index_page(client):
    r = client.get("/")
    assert r.status_code == 200
    assert "Работа 1" in r.text


def test_contracts_page(client):
    r = client.get("/contracts")
    assert r.status_code == 200
    assert "Т-001" in r.text


def test_execution_page(client):
    r = client.get("/execution")
    assert r.status_code == 200


def test_dashboard_page(client):
    r = client.get("/dashboard")
    assert r.status_code == 200


def test_upsert_plan_creates_plan_entry(client):
    ol_id = client.ctx["order_line_id"]
    period = date.today().strftime("%Y-%m")
    r = client.post("/api/plan", data={
        "order_line_id": ol_id,
        "period": period,
        "plan_qty": "3.5",
        "n": 3,
    })
    assert r.status_code == 200


def test_export_contracts_xlsx(client):
    r = client.get("/export/contracts.xlsx")
    assert r.status_code == 200
    assert r.headers["content-type"].startswith("application/vnd.openxmlformats")
    assert len(r.content) > 500


def test_export_plan_xlsx(client):
    r = client.get("/export/plan.xlsx")
    assert r.status_code == 200
    assert len(r.content) > 500


def test_export_execution_xlsx(client):
    r = client.get("/export/execution.xlsx")
    assert r.status_code == 200


def test_export_dashboard_xlsx(client):
    r = client.get("/export/dashboard.xlsx")
    assert r.status_code == 200


def test_404_for_unknown_plan(client):
    r = client.post("/api/plan", data={
        "order_line_id": 99999,
        "period": "2026-04",
        "plan_qty": "1",
        "n": 3,
    })
    assert r.status_code == 404
