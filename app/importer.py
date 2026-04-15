"""Парсер выгрузки из 1С вида `Выгрузка.xlsx`.

Лист `TDSheet` с колонками:
 1. "Номер, Заказ покупателя, Контрагент, Проект" (склейка четырёх полей через запятую)
 2. Номенклатура (вид работ)
 3. Содержание (описание работы)
 4. Заказано (сумма по договору, руб)
 5. Выполнено (руб)
 6. Осталось выполнить (руб)
 7. Заказано (количество)
 8. Отгружено (количество)
 9. Осталось отгрузить (количество)

В файле может быть 2 строки шапки — данные начинаются с 3-й строки (если первая ячейка выглядит как "НФФР-..." или "Заказ покупателя..." — это данные).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Iterator, Optional

import openpyxl
from sqlalchemy.orm import Session

from app.models import Contract, ImportBatch, OrderLine, WorkType


CONTRACT_NUMBER_RE = re.compile(r"^[А-ЯA-Z]+-\d+", re.UNICODE)
ORDER_DATE_RE = re.compile(r"от\s+(\d{2}[./]\d{2}[./]\d{4})")
EXCLUDED_WORK_TYPES_RE = re.compile(r"^Поставка\s+ЛО", re.IGNORECASE)


@dataclass
class ParsedRow:
    contract_number: str
    order_label: Optional[str]
    order_date: Optional[date]
    counterparty: Optional[str]
    project: Optional[str]
    work_type_name: str
    description: str
    sum_ordered: Optional[float]
    sum_done: Optional[float]
    sum_remaining: Optional[float]
    qty_ordered: Optional[float]
    qty_done: Optional[float]
    qty_remaining: Optional[float]


def _split_contract_cell(cell: str) -> tuple[str, Optional[str], Optional[date], Optional[str], Optional[str]]:
    """Разбить склейку "НФФР-003678, Заказ покупателя 3678 от 17.11.2025 , Контрагент, Адрес с запятыми".

    Проект (адрес) может содержать запятые, поэтому делим только по первым 3 запятым.
    """
    s = (cell or "").strip()
    parts = s.split(",", 3)
    parts = [p.strip() for p in parts]
    while len(parts) < 4:
        parts.append("")
    number, order_label, counterparty, project = parts
    order_date = None
    if order_label:
        m = ORDER_DATE_RE.search(order_label)
        if m:
            raw = m.group(1).replace(".", "/")
            for fmt in ("%d/%m/%Y", "%m/%d/%Y"):
                try:
                    order_date = datetime.strptime(raw, fmt).date()
                    break
                except ValueError:
                    continue
    return number, order_label or None, order_date, counterparty or None, project or None


def _as_float(v) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def parse_workbook(path: Path | str, sheet: str = "TDSheet") -> Iterator[ParsedRow]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        first = str(row[0]).strip()
        if not CONTRACT_NUMBER_RE.match(first):
            continue
        (number, order_label, order_date, counterparty, project) = _split_contract_cell(first)
        work_type_name = (str(row[1]).strip() if row[1] else "")
        description = (str(row[2]).strip() if row[2] else "")
        if not work_type_name or not description:
            continue
        if EXCLUDED_WORK_TYPES_RE.match(work_type_name):
            continue
        yield ParsedRow(
            contract_number=number,
            order_label=order_label,
            order_date=order_date,
            counterparty=counterparty,
            project=project,
            work_type_name=work_type_name,
            description=description,
            sum_ordered=_as_float(row[3]),
            sum_done=_as_float(row[4]),
            sum_remaining=_as_float(row[5]),
            qty_ordered=_as_float(row[6]),
            qty_done=_as_float(row[7]),
            qty_remaining=_as_float(row[8]),
        )


def _get_or_create_contract(session: Session, row: ParsedRow, cache: dict) -> tuple[Contract, bool]:
    """Вернуть (contract, is_new). Если договор уже есть — НЕ обновляем его (наша БД авторитетна)."""
    cached = cache.get(row.contract_number)
    if cached is not None:
        return cached
    c = session.query(Contract).filter_by(number=row.contract_number).one_or_none()
    is_new = False
    if c is None:
        c = Contract(
            number=row.contract_number,
            order_label=row.order_label,
            order_date=row.order_date,
            counterparty=row.counterparty,
            project=row.project,
        )
        session.add(c)
        session.flush()
        is_new = True
    cache[row.contract_number] = (c, is_new)
    return c, is_new


def _get_or_create_work_type(session: Session, name: str, cache: dict) -> WorkType:
    wt = cache.get(name)
    if wt is not None:
        return wt
    wt = session.query(WorkType).filter_by(name=name).one_or_none()
    if wt is None:
        wt = WorkType(name=name)
        session.add(wt)
        session.flush()
    cache[name] = wt
    return wt


def import_orders(
    session: Session,
    path: Path | str,
    uploaded_by: Optional[str] = None,
    sheet: str = "TDSheet",
) -> ImportBatch:
    """Импортирует выгрузку заказов. Возвращает ImportBatch с метриками."""
    path = Path(path)
    batch = ImportBatch(
        filename=path.name,
        uploaded_by=uploaded_by,
        kind="orders",
    )
    session.add(batch)
    session.flush()

    contract_cache: dict = {}
    work_type_cache: dict = {}
    parsed = 0
    skipped_existing = 0
    added_lines = 0

    for row in parse_workbook(path, sheet=sheet):
        parsed += 1
        contract, is_new = _get_or_create_contract(session, row, contract_cache)
        # Существующие договоры не трогаем — план/факт ведём в нашей БД.
        if not is_new:
            skipped_existing += 1
            continue
        work_type = _get_or_create_work_type(session, row.work_type_name, work_type_cache)
        session.add(OrderLine(
            contract_id=contract.id,
            work_type_id=work_type.id,
            description=row.description,
            sum_ordered=row.sum_ordered,
            sum_done_snapshot=row.sum_done,
            sum_remaining_snapshot=row.sum_remaining,
            qty_ordered=row.qty_ordered,
            qty_done_snapshot=row.qty_done,
            qty_remaining_snapshot=row.qty_remaining,
            import_batch_id=batch.id,
        ))
        added_lines += 1

    batch.rows_parsed = parsed
    batch.rows_matched = added_lines
    batch.rows_unmatched = skipped_existing
    batch.notes = f"added_lines={added_lines}, skipped_existing={skipped_existing}"
    session.commit()
    return batch
