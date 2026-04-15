from __future__ import annotations

import base64
import hmac
from datetime import date, datetime
from pathlib import Path
from typing import Optional

from fastapi import Depends, FastAPI, Form, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import func, or_
from sqlalchemy.orm import Session, selectinload
from starlette.middleware.base import BaseHTTPMiddleware

from app import config
from app.db import init_db, make_session_factory
from app.models import ActLine, Contract, ExecutionEntry, OrderLine, PlanEntry, WorkType
from sqlalchemy.orm import joinedload
from app.services.balances import balance, balance_bulk
from app.services.pricing import unit_price
from app.services.statuses import ExecutionStatus, PlanStatus
from app.services import export as export_svc
from app.services import forms as forms_svc

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = Jinja2Templates(directory=str(ROOT / "app" / "templates"))

config.validate()
engine = init_db()
SessionLocal = make_session_factory(engine)

app = FastAPI(title="Планирование работ")
app.mount("/static", StaticFiles(directory=str(ROOT / "app" / "static")), name="static")


class BasicAuthMiddleware(BaseHTTPMiddleware):
    """Одна общая пара логин/пароль. Если APP_PASSWORD пуст — пропускаем (dev)."""

    EXEMPT_PREFIXES = ("/static", "/healthz")

    async def dispatch(self, request: Request, call_next):
        if not config.APP_PASSWORD:
            return await call_next(request)
        if any(request.url.path.startswith(p) for p in self.EXEMPT_PREFIXES):
            return await call_next(request)
        header = request.headers.get("authorization", "")
        if header.lower().startswith("basic "):
            try:
                decoded = base64.b64decode(header[6:]).decode("utf-8", "replace")
                user, _, pwd = decoded.partition(":")
                if hmac.compare_digest(user, config.APP_USER) and hmac.compare_digest(pwd, config.APP_PASSWORD):
                    return await call_next(request)
            except Exception:
                pass
        return Response(
            status_code=401,
            headers={"WWW-Authenticate": 'Basic realm="WorksPlanning"'},
            content="Auth required",
        )


class OriginCheckMiddleware(BaseHTTPMiddleware):
    """CSRF-защита по Origin/Referer для state-changing запросов.

    Браузер шлёт Origin при всех cross-origin POST/PUT/DELETE. Если он не
    совпадает с нашим Host или списком ALLOWED_ORIGINS — отклоняем. Для
    same-origin submit Origin == scheme+Host (или отсутствует в некоторых
    старых случаях — тогда проверяем Referer).
    """

    UNSAFE_METHODS = {"POST", "PUT", "PATCH", "DELETE"}

    async def dispatch(self, request: Request, call_next):
        if request.method not in self.UNSAFE_METHODS:
            return await call_next(request)
        origin = request.headers.get("origin") or ""
        referer = request.headers.get("referer") or ""
        host = request.headers.get("host") or ""
        allowed = set(config.ALLOWED_ORIGINS)
        if host:
            allowed.add(f"http://{host}")
            allowed.add(f"https://{host}")
        source = origin or referer
        if source:
            if not any(source.startswith(a) for a in allowed):
                return Response(status_code=403, content="Origin not allowed")
        return await call_next(request)


app.add_middleware(OriginCheckMiddleware)
app.add_middleware(BasicAuthMiddleware)


@app.get("/healthz")
def healthz():
    from sqlalchemy import text
    try:
        s = SessionLocal()
        try:
            s.execute(text("SELECT 1"))
        finally:
            s.close()
    except Exception as exc:
        return JSONResponse({"ok": False, "error": str(exc)[:200]}, status_code=500)
    return JSONResponse({"ok": True})


def _xlsx_response(data: bytes, filename: str) -> Response:
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/export/contracts.xlsx")
def export_contracts_xlsx(
    contract_id: Optional[str] = None,
    q: Optional[str] = None,
):
    s = SessionLocal()
    try:
        cid = int(contract_id) if contract_id else None
        data = export_svc.export_contracts(s, contract_id=cid, q=q)
    finally:
        s.close()
    return _xlsx_response(data, "contracts.xlsx")


@app.get("/export/plan.xlsx")
def export_plan_xlsx(
    contract_id: Optional[str] = None,
    q: Optional[str] = None,
    start: Optional[str] = None,
    n: int = Query(3, ge=1, le=12),
):
    s = SessionLocal()
    try:
        cid = int(contract_id) if contract_id else None
        periods = next_periods(start or None, n=n)
        data = export_svc.export_plan(s, periods=periods, contract_id=cid, q=q)
    finally:
        s.close()
    return _xlsx_response(data, "plan.xlsx")


@app.get("/export/execution.xlsx")
def export_execution_xlsx(
    period: Optional[str] = None,
    contract_id: Optional[str] = None,
):
    s = SessionLocal()
    try:
        cid = int(contract_id) if contract_id else None
        p = period or date.today().strftime("%Y-%m")
        data = export_svc.export_execution(s, period=p, contract_id=cid)
    finally:
        s.close()
    return _xlsx_response(data, f"execution-{p}.xlsx")


@app.get("/export/dashboard.xlsx")
def export_dashboard_xlsx(
    start: Optional[str] = None,
    n: int = Query(3, ge=1, le=24),
    all: int = 0,
):
    s = SessionLocal()
    try:
        periods_filter = None if all else next_periods(start, n)
        data = export_svc.export_dashboard(s, periods_filter=periods_filter)
    finally:
        s.close()
    return _xlsx_response(data, "dashboard.xlsx")


# ── Страница форм ────────────────────────────────────────────────────────────

@app.get("/forms", response_class=HTMLResponse)
def forms_page(request: Request):
    s = SessionLocal()
    try:
        contracts = s.query(Contract).order_by(Contract.counterparty).all()
        today_period = date.today().strftime("%Y-%m")
    finally:
        s.close()
    return TEMPLATES.TemplateResponse(request, "forms.html", {
        "contracts": contracts,
        "today_period": today_period,
    })


@app.get("/export/form/work_order")
def export_form_work_order(
    period: str = Query(...),
    contract_id: str = Query(...),
):
    s = SessionLocal()
    try:
        cid = int(contract_id)
        data = forms_svc.export_work_order(s, period=period, contract_id=cid)
        c = s.get(Contract, cid)
        name = (c.contract_no or c.number or "zadanie").replace("/", "-")
    finally:
        s.close()
    return _xlsx_response(data, f"zadanie-{name}-{period}.xlsx")


@app.get("/export/form/monthly_plan")
def export_form_monthly_plan(
    period: str = Query(...),
    contract_id: Optional[str] = None,
):
    s = SessionLocal()
    try:
        cid = int(contract_id) if contract_id else None
        data = forms_svc.export_monthly_plan_form(s, period=period, contract_id=cid)
    finally:
        s.close()
    return _xlsx_response(data, f"plan-{period}.xlsx")


@app.get("/export/form/progress_report")
def export_form_progress_report(
    period: str = Query(...),
    contract_id: Optional[str] = None,
):
    s = SessionLocal()
    try:
        cid = int(contract_id) if contract_id else None
        data = forms_svc.export_progress_report(s, period=period, contract_id=cid)
    finally:
        s.close()
    return _xlsx_response(data, f"otchet-{period}.xlsx")


def get_session() -> Session:
    s = SessionLocal()
    try:
        yield s
    finally:
        s.close()


def shift_period(period: str, delta: int) -> str:
    y, m = map(int, period.split("-"))
    idx = y * 12 + (m - 1) + delta
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def next_periods(start: Optional[str] = None, n: int = 3) -> list[str]:
    """Список YYYY-MM для n ближайших месяцев, начиная со start (или сегодня)."""
    if start:
        y, m = map(int, start.split("-"))
    else:
        t = date.today()
        y, m = t.year, t.month
    out = []
    for _ in range(n):
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


def ru_month(period: str) -> str:
    y, m = period.split("-")
    names = ["", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
             "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
    return f"{names[int(m)]} {y}"


def load_rows(session: Session, periods: list[str], contract_id: Optional[int], q: Optional[str], hide_closed: bool = True):
    qset = (
        session.query(OrderLine)
        .options(selectinload(OrderLine.contract), selectinload(OrderLine.work_type), selectinload(OrderLine.plan_entries))
    )
    if contract_id:
        qset = qset.filter(OrderLine.contract_id == contract_id)
    if q:
        like = f"%{q}%"
        qset = qset.join(Contract).join(WorkType).filter(
            or_(
                OrderLine.description.ilike(like),
                WorkType.name.ilike(like),
                Contract.number.ilike(like),
                Contract.counterparty.ilike(like),
                Contract.project.ilike(like),
            )
        )
    order_lines = qset.order_by(OrderLine.contract_id, OrderLine.id).all()

    plans_by_key = {(p.order_line_id, p.period): p for ol in order_lines for p in ol.plan_entries}

    plan_ids = [p.id for p in plans_by_key.values()]
    ex_by_plan: dict[int, ExecutionEntry] = {}
    if plan_ids:
        ex_by_plan = {
            ex.plan_entry_id: ex
            for ex in session.query(ExecutionEntry).filter(ExecutionEntry.plan_entry_id.in_(plan_ids)).all()
        }

    balances_map = balance_bulk(session, [ol.id for ol in order_lines])
    groups: dict[int, dict] = {}
    hidden = 0
    for ol in order_lines:
        b = balances_map.get(ol.id) or balance(session, ol.id)
        closed = b.qty_remaining is not None and b.qty_remaining <= 0
        g = groups.setdefault(ol.contract_id, {
            "contract": ol.contract, "rows": [],
            "ordered_rub": 0.0, "done_rub": 0.0, "signed_rub": 0.0,
        })
        g["ordered_rub"] += float(ol.sum_ordered or 0.0)
        g["done_rub"] += float(b.sum_done or 0.0)
        unit = unit_price(ol)
        plans = [plans_by_key.get((ol.id, p)) for p in periods]
        plan_values = [(pe.plan_qty if pe else None) for pe in plans]
        plan_statuses = [(pe.status if pe else PlanStatus.DRAFT) for pe in plans]
        fact_values = [(ex_by_plan.get(pe.id).qty_fact if pe and ex_by_plan.get(pe.id) else None) for pe in plans]
        fact_signed = [(ex_by_plan.get(pe.id).status == ExecutionStatus.SIGNED if pe and ex_by_plan.get(pe.id) else False) for pe in plans]
        plan_sum = sum((pv or 0) for pv, fs in zip(plan_values, fact_signed) if not fs)
        signed_rub = sum(((fv or 0) * unit) for fv, fs in zip(fact_values, fact_signed) if fs)
        g["signed_rub"] += signed_rub
        if hide_closed and closed:
            hidden += 1
            continue
        g["rows"].append({
            "ol": ol,
            "balance": b,
            "plan_values": plan_values,
            "plan_statuses": plan_statuses,
            "fact_values": fact_values,
            "fact_signed": fact_signed,
            "plan_sum": plan_sum,
        })
    result = [g for g in groups.values() if g["rows"]]
    result.sort(key=lambda g: g["ordered_rub"], reverse=True)
    return result, hidden


def _period_counts(session: Session, periods: list[str]) -> dict[str, dict]:
    """Для каждой period возвращает сводку: draft / approved."""
    out = {p: {"draft": 0, "approved": 0, "total": 0} for p in periods}
    rows = (
        session.query(PlanEntry.period, PlanEntry.status, func.count())
        .filter(PlanEntry.period.in_(periods))
        .group_by(PlanEntry.period, PlanEntry.status)
        .all()
    )
    for period, status, n in rows:
        out[period][status] = n
        out[period]["total"] += n
    return out


@app.get("/", response_class=HTMLResponse)
def index(
    request: Request,
    contract_id: Optional[str] = None,
    q: Optional[str] = None,
    start: Optional[str] = None,
    n: int = Query(3, ge=1, le=12),
    show_closed: int = 0,
    session: Session = Depends(get_session),
):
    contract_id_int = int(contract_id) if contract_id else None
    start = start or None
    if start and not start.strip():
        start = None
    periods = next_periods(start, n=n)
    hide_closed = not bool(show_closed)
    today_period = date.today().strftime("%Y-%m")
    start_period = periods[0]
    prev_start = shift_period(start_period, -1)
    next_start = shift_period(start_period, 1)
    contracts = session.query(Contract).order_by(Contract.number).all()
    groups, hidden_closed = load_rows(session, periods, contract_id_int, q, hide_closed=hide_closed)

    total_ordered = session.query(func.coalesce(func.sum(OrderLine.sum_ordered), 0.0)).scalar() or 0.0
    total_remaining = session.query(func.coalesce(func.sum(OrderLine.sum_remaining_snapshot), 0.0)).scalar() or 0.0
    total_rows = sum(len(g["rows"]) for g in groups)
    period_stats = _period_counts(session, periods)

    return TEMPLATES.TemplateResponse(
        request,
        "index.html",
        {
            "contracts": contracts,
            "groups": groups,
            "periods": periods,
            "period_labels": [ru_month(p) for p in periods],
            "contract_id": contract_id_int,
            "q": q or "",
            "start": start or "",
            "n": n,
            "today_period": today_period,
            "prev_start": prev_start,
            "next_start": next_start,
            "total_ordered": total_ordered,
            "total_remaining": total_remaining,
            "total_rows": total_rows,
            "period_stats": period_stats,
            "hide_closed": hide_closed,
            "hidden_closed": hidden_closed,
        },
    )


@app.post("/api/plan", response_class=HTMLResponse)
def upsert_plan(
    request: Request,
    order_line_id: int = Form(...),
    period: str = Form(...),
    plan_qty: str = Form(""),
    start: Optional[str] = Form(None),
    n: int = Form(3),
    session: Session = Depends(get_session),
):
    n = max(1, min(n, 12))
    ol = session.get(OrderLine, order_line_id)
    if ol is None:
        raise HTTPException(404, "order_line not found")

    qty_val: Optional[float]
    s = plan_qty.replace(",", ".").strip()
    qty_val = float(s) if s else None

    pe = (
        session.query(PlanEntry)
        .filter_by(order_line_id=order_line_id, period=period)
        .one_or_none()
    )
    if pe is not None and pe.status == PlanStatus.APPROVED:
        raise HTTPException(409, "План утверждён, сначала снимите утверждение.")
    if pe is None and qty_val is not None:
        pe = PlanEntry(order_line_id=order_line_id, period=period, plan_qty=qty_val)
        session.add(pe)
    elif pe is not None and qty_val is None:
        session.delete(pe)
    elif pe is not None:
        pe.plan_qty = qty_val
    session.commit()

    periods = next_periods(start, n=n)
    plans = {
        p.period: p for p in session.query(PlanEntry).filter_by(order_line_id=order_line_id).all()
    }
    plan_values = [(plans.get(p).plan_qty if plans.get(p) else None) for p in periods]
    plan_statuses = [(plans.get(p).status if plans.get(p) else "draft") for p in periods]
    plan_ids = [plans[p].id for p in periods if plans.get(p)]
    ex_by_plan = {}
    if plan_ids:
        ex_by_plan = {
            ex.plan_entry_id: ex
            for ex in session.query(ExecutionEntry).filter(ExecutionEntry.plan_entry_id.in_(plan_ids)).all()
        }
    fact_values = [(ex_by_plan.get(plans[p].id).qty_fact if plans.get(p) and ex_by_plan.get(plans[p].id) else None) for p in periods]
    fact_signed = [(ex_by_plan.get(plans[p].id).status == ExecutionStatus.SIGNED if plans.get(p) and ex_by_plan.get(plans[p].id) else False) for p in periods]
    plan_sum = sum((pv or 0) for pv, fs in zip(plan_values, fact_signed) if not fs)
    b = balance(session, order_line_id)

    return TEMPLATES.TemplateResponse(
        request,
        "partials/row.html",
        {
            "row": {"ol": ol, "balance": b, "plan_values": plan_values, "plan_statuses": plan_statuses, "fact_values": fact_values, "fact_signed": fact_signed, "plan_sum": plan_sum},
            "periods": periods,
        },
    )


@app.post("/api/plan/approve")
def approve_period(
    period: str = Form(...),
    approved_by: Optional[str] = Form("planner"),
    session: Session = Depends(get_session),
):
    now = datetime.utcnow()
    to_approve = (
        session.query(PlanEntry)
        .filter(PlanEntry.period == period, PlanEntry.status == PlanStatus.DRAFT)
        .all()
    )
    for pe in to_approve:
        pe.status = PlanStatus.APPROVED
        pe.approved_at = now
        pe.approved_by = approved_by
        exists = session.query(ExecutionEntry).filter_by(plan_entry_id=pe.id).one_or_none()
        if exists is None:
            session.add(ExecutionEntry(plan_entry_id=pe.id))
    session.commit()
    return {"approved": len(to_approve), "period": period}


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(
    request: Request,
    start: Optional[str] = None,
    n: int = Query(3, ge=1, le=24),
    all: int = 0,
    session: Session = Depends(get_session),
):
    _unit = unit_price
    periods_filter: Optional[list[str]] = None
    if not all:
        periods_filter = next_periods(start, n)

    qry = (
        session.query(PlanEntry, OrderLine, Contract, ExecutionEntry)
        .join(OrderLine, PlanEntry.order_line_id == OrderLine.id)
        .join(Contract, OrderLine.contract_id == Contract.id)
        .outerjoin(ExecutionEntry, ExecutionEntry.plan_entry_id == PlanEntry.id)
    )
    if periods_filter is not None:
        qry = qry.filter(PlanEntry.period.in_(periods_filter))
    rows = qry.all()

    totals = {
        "plan_draft_sum": 0.0, "plan_approved_sum": 0.0,
        "fact_sum": 0.0, "signed_sum": 0.0,
        "plan_approved_qty": 0.0, "fact_qty": 0.0, "signed_qty": 0.0,
    }
    by_period: dict[str, dict] = {}
    by_contract: dict[int, dict] = {}

    for pe, ol, c, ex in rows:
        unit = _unit(ol)
        plan_qty = pe.plan_qty or 0.0
        plan_sum = (pe.plan_sum or 0.0) or plan_qty * unit
        fact_qty = (ex.qty_fact or 0.0) if ex else 0.0
        fact_sum = (ex.sum_fact or 0.0) if ex else 0.0
        signed = bool(ex and ex.status == ExecutionStatus.SIGNED)
        approved = pe.status == PlanStatus.APPROVED

        p = by_period.setdefault(pe.period, {
            "period": pe.period, "label": ru_month(pe.period),
            "plan_draft_sum": 0.0, "plan_approved_sum": 0.0,
            "fact_sum": 0.0, "signed_sum": 0.0,
            "plan_draft_n": 0, "plan_approved_n": 0, "signed_n": 0,
        })
        k = by_contract.setdefault(c.id, {
            "contract": c,
            "plan_sum": 0.0, "fact_sum": 0.0, "signed_sum": 0.0,
            "rows": 0,
        })

        if approved:
            totals["plan_approved_sum"] += plan_sum
            totals["plan_approved_qty"] += plan_qty
            p["plan_approved_sum"] += plan_sum
            p["plan_approved_n"] += 1
            k["plan_sum"] += plan_sum
        else:
            totals["plan_draft_sum"] += plan_sum
            p["plan_draft_sum"] += plan_sum
            p["plan_draft_n"] += 1
            k["plan_sum"] += plan_sum

        totals["fact_sum"] += fact_sum
        totals["fact_qty"] += fact_qty
        p["fact_sum"] += fact_sum
        k["fact_sum"] += fact_sum

        if signed:
            totals["signed_sum"] += fact_sum
            totals["signed_qty"] += fact_qty
            p["signed_sum"] += fact_sum
            p["signed_n"] += 1
            k["signed_sum"] += fact_sum

        k["rows"] += 1

    if periods_filter is not None:
        for p in periods_filter:
            by_period.setdefault(p, {
                "period": p, "label": ru_month(p),
                "plan_draft_sum": 0.0, "plan_approved_sum": 0.0,
                "fact_sum": 0.0, "signed_sum": 0.0,
                "plan_draft_n": 0, "plan_approved_n": 0, "signed_n": 0,
            })

    periods_sorted = sorted(by_period.values(), key=lambda x: x["period"])
    contracts_sorted = sorted(
        by_contract.values(),
        key=lambda x: x["plan_sum"],
        reverse=True,
    )

    total_ordered = float(session.query(func.coalesce(func.sum(OrderLine.sum_ordered), 0.0)).scalar() or 0.0)
    total_acts = float(session.query(func.coalesce(func.sum(ActLine.sum), 0.0)).scalar() or 0.0)
    total_signed_exec = float(
        session.query(func.coalesce(func.sum(ExecutionEntry.sum_fact), 0.0))
        .filter(ExecutionEntry.status == ExecutionStatus.SIGNED, ExecutionEntry.act_line_id.is_(None))
        .scalar() or 0.0
    )
    snapshot_done = float(session.query(func.coalesce(func.sum(OrderLine.sum_done_snapshot), 0.0)).scalar() or 0.0)
    total_done = max(total_acts + total_signed_exec, snapshot_done)
    total_contract_remaining = total_ordered - total_done
    contracts_count = session.query(Contract).count()

    start_effective = (periods_filter[0] if periods_filter else (periods_sorted[0]["period"] if periods_sorted else date.today().strftime("%Y-%m")))
    return TEMPLATES.TemplateResponse(
        request, "dashboard.html",
        {
            "totals": totals,
            "periods": periods_sorted,
            "contracts": contracts_sorted,
            "start": start_effective,
            "n": n,
            "all_periods": bool(all),
            "prev_start": shift_period(start_effective, -1),
            "next_start": shift_period(start_effective, 1),
            "today_period": date.today().strftime("%Y-%m"),
            "total_ordered": total_ordered,
            "total_done": total_done,
            "total_contract_remaining": total_contract_remaining,
            "contracts_count": contracts_count,
        },
    )


@app.get("/contracts", response_class=HTMLResponse)
def contracts_page(
    request: Request,
    contract_id: Optional[str] = None,
    q: Optional[str] = None,
    session: Session = Depends(get_session),
):
    contract_id_int = int(contract_id) if contract_id else None

    cset = session.query(Contract)
    if contract_id_int:
        cset = cset.filter(Contract.id == contract_id_int)
    if q:
        like = f"%{q}%"
        cset = cset.filter(or_(
            Contract.number.ilike(like),
            Contract.contract_no.ilike(like),
            Contract.counterparty.ilike(like),
            Contract.project.ilike(like),
            Contract.order_label.ilike(like),
        ))
    contracts_full = cset.options(
        selectinload(Contract.order_lines).selectinload(OrderLine.work_type)
    ).all()

    all_ol_ids = [ol.id for c in contracts_full for ol in c.order_lines]
    balances_map = balance_bulk(session, all_ol_ids)

    plan_qty_rows = session.query(
        PlanEntry.order_line_id,
        func.coalesce(func.sum(PlanEntry.plan_qty), 0.0),
    ).filter(PlanEntry.order_line_id.in_(all_ol_ids)).group_by(PlanEntry.order_line_id).all() if all_ol_ids else []
    plan_qty_map = {r[0]: float(r[1]) for r in plan_qty_rows}

    signed_rows = session.query(
        PlanEntry.order_line_id,
        func.coalesce(func.sum(ExecutionEntry.qty_fact), 0.0),
    ).join(ExecutionEntry, ExecutionEntry.plan_entry_id == PlanEntry.id).filter(
        PlanEntry.order_line_id.in_(all_ol_ids),
        ExecutionEntry.status == ExecutionStatus.SIGNED,
        ExecutionEntry.act_line_id.is_(None),
    ).group_by(PlanEntry.order_line_id).all() if all_ol_ids else []
    signed_qty_map = {r[0]: float(r[1]) for r in signed_rows}

    groups = []
    total_ordered = 0.0
    total_done = 0.0
    total_remaining = 0.0
    for c in contracts_full:
        rows = []
        ord_sum = 0.0
        done_sum = 0.0
        rem_sum = 0.0
        closed_n = 0
        plan_sum_total = 0.0
        signed_sum_total = 0.0
        for ol in c.order_lines:
            b = balances_map.get(ol.id) or balance(session, ol.id)
            u = unit_price(ol)
            plan_rub = plan_qty_map.get(ol.id, 0.0) * u
            signed_rub = signed_qty_map.get(ol.id, 0.0) * u
            rows.append({"ol": ol, "balance": b, "plan_rub": plan_rub, "signed_rub": signed_rub})
            ord_sum += float(ol.sum_ordered or 0.0)
            done_sum += float(b.sum_done or 0.0)
            plan_sum_total += plan_rub
            signed_sum_total += signed_rub
            if b.sum_remaining is not None:
                rem_sum += float(b.sum_remaining)
            if b.qty_remaining is not None and b.qty_remaining <= 0:
                closed_n += 1
        rows.sort(key=lambda r: (r["balance"].qty_remaining is not None and r["balance"].qty_remaining <= 0, -(r["ol"].sum_ordered or 0)))
        groups.append({
            "contract": c,
            "rows": rows,
            "ordered_rub": ord_sum,
            "done_rub": done_sum,
            "remaining_rub": rem_sum,
            "closed_n": closed_n,
            "plan_rub": plan_sum_total,
            "signed_rub": signed_sum_total,
        })
        total_ordered += ord_sum
        total_done += done_sum
        total_remaining += rem_sum

    groups.sort(key=lambda g: g["ordered_rub"], reverse=True)
    all_contracts = session.query(Contract).order_by(Contract.number).all()
    return TEMPLATES.TemplateResponse(
        request, "contracts.html",
        {
            "groups": groups,
            "contracts": all_contracts,
            "contract_id": contract_id_int,
            "q": q or "",
            "total_ordered": total_ordered,
            "total_done": total_done,
            "total_remaining": total_remaining,
        },
    )


@app.get("/execution", response_class=HTMLResponse)
def execution_page(
    request: Request,
    period: Optional[str] = None,
    contract_id: Optional[str] = None,
    session: Session = Depends(get_session),
):
    contract_id_int = int(contract_id) if contract_id else None
    period = period or date.today().strftime("%Y-%m")

    qset = (
        session.query(ExecutionEntry)
        .join(PlanEntry, ExecutionEntry.plan_entry_id == PlanEntry.id)
        .join(OrderLine, PlanEntry.order_line_id == OrderLine.id)
        .filter(PlanEntry.period == period)
        .options(
            selectinload(ExecutionEntry.plan_entry)
            .selectinload(PlanEntry.order_line)
            .selectinload(OrderLine.contract)
        )
    )
    if contract_id_int:
        qset = qset.filter(OrderLine.contract_id == contract_id_int)
    entries = qset.order_by(OrderLine.contract_id, OrderLine.id).all()

    _unit_price = unit_price

    groups: dict[int, dict] = {}
    totals = {"plan_qty": 0.0, "plan_sum": 0.0, "fact_qty": 0.0, "fact_sum": 0.0}
    for ex in entries:
        ol = ex.plan_entry.order_line
        g = groups.setdefault(ol.contract_id, {
            "contract": ol.contract, "rows": [],
            "plan_sum": 0.0, "plan_qty": 0.0, "fact_sum": 0.0, "fact_qty": 0.0,
        })
        plan_qty = ex.plan_entry.plan_qty or 0.0
        unit = _unit_price(ol)
        plan_sum = (ex.plan_entry.plan_sum or 0.0) or plan_qty * unit
        fact_qty = ex.qty_fact or 0.0
        fact_sum = ex.sum_fact or 0.0
        g["rows"].append({"ex": ex, "plan_sum": plan_sum, "unit_price": unit})
        g["plan_qty"] += plan_qty
        g["plan_sum"] += plan_sum
        g["fact_qty"] += fact_qty
        g["fact_sum"] += fact_sum
        totals["plan_qty"] += plan_qty
        totals["plan_sum"] += plan_sum
        totals["fact_qty"] += fact_qty
        totals["fact_sum"] += fact_sum

    contracts = session.query(Contract).order_by(Contract.number).all()

    return TEMPLATES.TemplateResponse(
        request,
        "execution.html",
        {
            "period": period,
            "period_label": ru_month(period),
            "groups": list(groups.values()),
            "contracts": contracts,
            "contract_id": contract_id_int,
            "total_rows": sum(len(g["rows"]) for g in groups.values()),
            "totals": totals,
        },
    )


@app.post("/api/execution", response_class=HTMLResponse)
def upsert_execution(
    request: Request,
    execution_id: int = Form(...),
    qty_fact: str = Form(""),
    note: str = Form(""),
    session: Session = Depends(get_session),
):
    ex = session.get(ExecutionEntry, execution_id)
    if ex is None:
        raise HTTPException(404, "execution not found")
    if ex.status == ExecutionStatus.SIGNED:
        raise HTTPException(409, "Запись подписана актом, правки невозможны.")

    def _num(s: str) -> Optional[float]:
        s = s.replace(",", ".").strip()
        return float(s) if s else None

    ol = ex.plan_entry.order_line
    unit = unit_price(ol)

    qty = _num(qty_fact)
    ex.qty_fact = qty
    ex.sum_fact = (qty * unit) if (qty is not None and unit) else None
    ex.note = note.strip() or None
    session.commit()

    plan_sum = (ex.plan_entry.plan_sum or 0.0) or (ex.plan_entry.plan_qty or 0.0) * unit

    return TEMPLATES.TemplateResponse(
        request,
        "partials/execution_row.html",
        {"ex": ex, "plan_sum": plan_sum},
    )


@app.post("/api/execution/sign", response_class=HTMLResponse)
def sign_execution(
    request: Request,
    execution_id: int = Form(...),
    signed_by: Optional[str] = Form("planner"),
    session: Session = Depends(get_session),
):
    ex = session.get(ExecutionEntry, execution_id)
    if ex is None:
        raise HTTPException(404, "execution not found")
    ex.status = ExecutionStatus.SIGNED
    ex.signed_at = datetime.utcnow()
    ex.signed_by = signed_by
    session.commit()

    ol = ex.plan_entry.order_line
    unit = unit_price(ol)
    plan_sum = (ex.plan_entry.plan_sum or 0.0) or (ex.plan_entry.plan_qty or 0.0) * unit
    return TEMPLATES.TemplateResponse(
        request, "partials/execution_row.html", {"ex": ex, "plan_sum": plan_sum},
    )


@app.post("/api/execution/unsign", response_class=HTMLResponse)
def unsign_execution(
    request: Request,
    execution_id: int = Form(...),
    session: Session = Depends(get_session),
):
    ex = session.get(ExecutionEntry, execution_id)
    if ex is None:
        raise HTTPException(404, "execution not found")
    ex.status = ExecutionStatus.DRAFT
    ex.signed_at = None
    ex.signed_by = None
    ex.act_line_id = None
    session.commit()

    ol = ex.plan_entry.order_line
    unit = unit_price(ol)
    plan_sum = (ex.plan_entry.plan_sum or 0.0) or (ex.plan_entry.plan_qty or 0.0) * unit
    return TEMPLATES.TemplateResponse(
        request, "partials/execution_row.html", {"ex": ex, "plan_sum": plan_sum},
    )


@app.post("/api/execution/sign_bulk")
def sign_execution_bulk(
    period: str = Form(...),
    contract_id: Optional[int] = Form(None),
    signed_by: Optional[str] = Form("planner"),
    session: Session = Depends(get_session),
):
    now = datetime.utcnow()
    qset = (
        session.query(ExecutionEntry)
        .join(PlanEntry, ExecutionEntry.plan_entry_id == PlanEntry.id)
        .join(OrderLine, PlanEntry.order_line_id == OrderLine.id)
        .filter(PlanEntry.period == period, ExecutionEntry.status == PlanStatus.DRAFT)
    )
    if contract_id:
        qset = qset.filter(OrderLine.contract_id == contract_id)
    count = 0
    for ex in qset.all():
        ex.status = ExecutionStatus.SIGNED
        ex.signed_at = now
        ex.signed_by = signed_by
        count += 1
    session.commit()
    return {"signed": count, "period": period, "contract_id": contract_id}


@app.post("/api/plan/unfreeze")
def unfreeze_period(
    period: str = Form(...),
    session: Session = Depends(get_session),
):
    count = (
        session.query(PlanEntry)
        .filter(PlanEntry.period == period, PlanEntry.status == PlanStatus.APPROVED)
        .update({"status": PlanStatus.DRAFT, "approved_at": None, "approved_by": None},
                synchronize_session=False)
    )
    session.commit()
    return {"unfrozen": count, "period": period}
