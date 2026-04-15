"""Единая точка для расчёта удельной цены единицы по OrderLine.

До этого формула `sum_ordered / qty_ordered` встречалась в 10+ местах —
в main.py, export.py и Jinja-шаблонах. Унифицировано здесь, чтобы при
смене логики (например, учёт скидки) менять в одном месте.
"""
from __future__ import annotations

from app.models import OrderLine


def unit_price(ol: OrderLine) -> float:
    """Цена за единицу работы (₽/шт, ₽/м² и т.п.).

    Возвращает 0.0, если любая из сторон неизвестна или нулевая —
    вызывающие коды везде ожидают число, не None.
    """
    if ol.qty_ordered and ol.sum_ordered:
        return ol.sum_ordered / ol.qty_ordered
    return 0.0


def plan_rub(plan_qty: float | None, ol: OrderLine) -> float:
    if not plan_qty:
        return 0.0
    return float(plan_qty) * unit_price(ol)
