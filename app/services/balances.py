"""Расчёт остатка и факта по order_line.

Принцип: фактически выполненный объём = сумма act_line по order_line.
Остаток = ordered - done. Если актов нет (Фаза 0, демо-данные) — используем
snapshot из выгрузки как fallback.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Optional

from sqlalchemy import func
from sqlalchemy.orm import Session

from app.models import Act, ActLine, ExecutionEntry, OrderLine, PlanEntry


@dataclass
class Balance:
    order_line_id: int
    qty_ordered: Optional[float]
    qty_done: float
    qty_remaining: Optional[float]
    sum_ordered: Optional[float]
    sum_done: float
    sum_remaining: Optional[float]
    from_acts: bool  # True если done посчитан по актам, False — взят из snapshot


def _period_bounds(period: str) -> tuple[date, date]:
    """period = 'YYYY-MM' → (first_day, first_day_of_next_month)."""
    y, m = period.split("-")
    y, m = int(y), int(m)
    start = date(y, m, 1)
    if m == 12:
        end = date(y + 1, 1, 1)
    else:
        end = date(y, m + 1, 1)
    return start, end


def balance(session: Session, order_line_id: int) -> Balance:
    ol = session.get(OrderLine, order_line_id)
    if ol is None:
        raise ValueError(f"OrderLine {order_line_id} not found")

    qty_act, sum_act, cnt_act = session.query(
        func.coalesce(func.sum(ActLine.qty), 0.0),
        func.coalesce(func.sum(ActLine.sum), 0.0),
        func.count(ActLine.id),
    ).filter(ActLine.order_line_id == order_line_id).one()

    qty_signed, sum_signed, cnt_signed = session.query(
        func.coalesce(func.sum(ExecutionEntry.qty_fact), 0.0),
        func.coalesce(func.sum(ExecutionEntry.sum_fact), 0.0),
        func.count(ExecutionEntry.id),
    ).join(PlanEntry, ExecutionEntry.plan_entry_id == PlanEntry.id).filter(
        PlanEntry.order_line_id == order_line_id,
        ExecutionEntry.status == "signed",
        ExecutionEntry.act_line_id.is_(None),
    ).one()

    qty_sum = float(qty_act) + float(qty_signed)
    sum_sum = float(sum_act) + float(sum_signed)
    cnt = cnt_act + cnt_signed

    # Если по строке уже есть какая-либо активность (план, исполнение, акт) —
    # доверяем вычисленному значению и НЕ используем snapshot из 1С.
    # Иначе снятие подписи у executions "вспоминало" бы старое выполнение из снапшота.
    has_plan = session.query(PlanEntry.id).filter(PlanEntry.order_line_id == order_line_id).first() is not None
    has_exec = session.query(ExecutionEntry.id).join(
        PlanEntry, ExecutionEntry.plan_entry_id == PlanEntry.id
    ).filter(PlanEntry.order_line_id == order_line_id).first() is not None
    managed = has_plan or has_exec or cnt_act > 0

    if cnt > 0 or managed:
        qty_remaining = None if ol.qty_ordered is None else ol.qty_ordered - qty_sum
        sum_remaining = None if ol.sum_ordered is None else ol.sum_ordered - sum_sum
        return Balance(
            order_line_id=ol.id,
            qty_ordered=ol.qty_ordered,
            qty_done=qty_sum,
            qty_remaining=qty_remaining,
            sum_ordered=ol.sum_ordered,
            sum_done=sum_sum,
            sum_remaining=sum_remaining,
            from_acts=True,
        )

    # fallback на snapshot
    return Balance(
        order_line_id=ol.id,
        qty_ordered=ol.qty_ordered,
        qty_done=float(ol.qty_done_snapshot or 0.0),
        qty_remaining=ol.qty_remaining_snapshot,
        sum_ordered=ol.sum_ordered,
        sum_done=float(ol.sum_done_snapshot or 0.0),
        sum_remaining=ol.sum_remaining_snapshot,
        from_acts=False,
    )


def fact(session: Session, order_line_id: int, period: str) -> tuple[float, float]:
    """Вернуть (qty, sum) факта за period YYYY-MM по order_line."""
    start, end = _period_bounds(period)
    qty_sum, sum_sum = session.query(
        func.coalesce(func.sum(ActLine.qty), 0.0),
        func.coalesce(func.sum(ActLine.sum), 0.0),
    ).join(Act, Act.id == ActLine.act_id).filter(
        ActLine.order_line_id == order_line_id,
        Act.act_date >= start,
        Act.act_date < end,
    ).one()
    return float(qty_sum), float(sum_sum)
