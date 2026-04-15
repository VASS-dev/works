"""Парсер файла `Договоры текущие.xlsx` — полный реестр действующих договоров.

Структура листа TDSheet (шапки в первых 5 строках):
- Строка-заголовок договора в col 0: "Заказ покупателя N от DD.MM.YYYY , Контрагент, Проект/адрес..., № <Номер договора> от DD.MM.YYYY"
  col 4: количество (итого по договору)
- Следующая строка: col 4 = сумма по договору (итого)
- Далее парами идут работы:
    первая строка — col 0: описание (Содержание), col 3: Номенклатура, col 4: qty_ordered
    вторая строка — col 4: sum_ordered

Логика импорта:
- Сопоставляем договор по номеру заказа (Заказ покупателя N) с существующим Contract.order_label.
- НОВЫЕ договоры не создаём (требование: только дополнить текущий датафрейм).
- Обновляем Contract.contract_no / contract_date.
- Для каждой работы: если такая OrderLine уже есть — оставляем, только обновляем sum_ordered/qty_ordered.
  Если нет — добавляем, считая её "закрытой" (qty_done_snapshot = qty_ordered, qty_remaining_snapshot = 0, то же для sum_*).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterator, Optional

import openpyxl
from sqlalchemy.orm import Session

from app.models import Contract, ImportBatch, OrderLine, WorkType


ORDER_NUM_RE = re.compile(r"Заказ покупателя\s+(\d+)\s+от\s+(\d{2}[./]\d{2}[./]\d{4})")
CONTRACT_NO_RE = re.compile(r"№\s*(.+?)\s+от\s+(\d{2}[./]\d{2}[./]\d{4})\s*$")

# Номенклатуры, которые не ведём (поставка лифтового оборудования — отдельный учёт)
EXCLUDED_WORK_TYPES_RE = re.compile(r"^Поставка\s+ЛО", re.IGNORECASE)


@dataclass
class ParsedContract:
    order_number: str              # "3678"
    order_date: Optional[date]
    raw_header: str
    counterparty: Optional[str]
    project: Optional[str]
    contract_no: Optional[str]
    contract_date: Optional[date]
    lines: list["ParsedLine"]


@dataclass
class ParsedLine:
    description: str
    work_type_name: str
    qty_ordered: Optional[float]
    sum_ordered: Optional[float]


def _parse_date(s: str) -> Optional[date]:
    s = s.replace(".", "/")
    for fmt in ("%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _as_float(v) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _split_header(header: str) -> tuple[Optional[str], Optional[date], Optional[str], Optional[str], Optional[str], Optional[date]]:
    """Вернуть (order_number, order_date, counterparty, project, contract_no, contract_date)."""
    s = header.strip()
    m = ORDER_NUM_RE.match(s)
    order_number = m.group(1) if m else None
    order_date = _parse_date(m.group(2)) if m else None

    contract_no = None
    contract_date = None
    mc = CONTRACT_NO_RE.search(s)
    if mc:
        contract_no = mc.group(1).strip()
        contract_date = _parse_date(mc.group(2))
        s = s[:mc.start()].rstrip(", ")

    # убираем префикс "Заказ покупателя N от dd.mm.yyyy ,"
    if m:
        s = s[m.end():].lstrip(" ,")

    parts = [p.strip() for p in s.split(",", 1)]
    counterparty = parts[0] if parts and parts[0] else None
    project = parts[1].strip() if len(parts) > 1 else None
    return order_number, order_date, counterparty, project, contract_no, contract_date


def parse_full_contracts(path: Path | str, sheet: str = "TDSheet") -> Iterator[ParsedContract]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    current: Optional[ParsedContract] = None
    mode = "idle"
    pending_line: Optional[ParsedLine] = None

    for row in ws.iter_rows(values_only=True):
        c0 = (str(row[0]).strip() if row and row[0] else "")
        c3 = (str(row[3]).strip() if len(row) > 3 and row[3] else "")
        c4 = row[4] if len(row) > 4 else None

        if "Заказ покупателя" in c0 and ORDER_NUM_RE.match(c0):
            if current is not None:
                yield current
            order_number, order_date, counterparty, project, contract_no, contract_date = _split_header(c0)
            current = ParsedContract(
                order_number=order_number or "",
                order_date=order_date,
                raw_header=c0,
                counterparty=counterparty,
                project=project,
                contract_no=contract_no,
                contract_date=contract_date,
                lines=[],
            )
            mode = "contract_qty_row"
            continue

        if current is None:
            continue

        if mode == "contract_qty_row":
            mode = "contract_sum_row"
            continue
        if mode == "contract_sum_row":
            mode = "work_qty_row"
            continue

        if mode == "work_qty_row":
            if not (c0 or c3):
                continue
            pending_line = ParsedLine(
                description=c0,
                work_type_name=c3,
                qty_ordered=_as_float(c4),
                sum_ordered=None,
            )
            mode = "work_sum_row"
            continue

        if mode == "work_sum_row":
            if pending_line is not None:
                pending_line.sum_ordered = _as_float(c4)
                if pending_line.description and pending_line.work_type_name:
                    current.lines.append(pending_line)
                pending_line = None
            mode = "work_qty_row"
            continue

    if current is not None:
        yield current


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


def import_full_contracts(
    session: Session,
    path: Path | str,
    uploaded_by: Optional[str] = None,
    sheet: str = "TDSheet",
) -> dict:
    """Импорт файла 'Договоры текущие.xlsx'.

    Не создаёт новых Contract: только дополняет существующие (matched по номеру заказа).
    Новые OrderLine помечаются как уже закрытые (qty_remaining = 0).
    """
    path = Path(path)
    batch = ImportBatch(filename=path.name, uploaded_by=uploaded_by, kind="full_contracts")
    session.add(batch)
    session.flush()

    work_type_cache: dict = {}
    stats = {
        "contracts_seen": 0,
        "contracts_matched": 0,
        "contracts_skipped_new": 0,
        "lines_parsed": 0,
        "lines_updated": 0,
        "lines_added_closed": 0,
    }

    # индекс существующих Contract по номеру заказа (из order_label: "Заказ покупателя 3678 от ...")
    by_order_num: dict[str, Contract] = {}
    for c in session.query(Contract).all():
        if not c.order_label:
            continue
        m = re.search(r"Заказ покупателя\s+(\d+)", c.order_label)
        if m:
            by_order_num[m.group(1)] = c

    for parsed in parse_full_contracts(path, sheet=sheet):
        stats["contracts_seen"] += 1
        contract = by_order_num.get(parsed.order_number)
        if contract is None:
            stats["contracts_skipped_new"] += 1
            continue
        stats["contracts_matched"] += 1
        if parsed.contract_no and not contract.contract_no:
            contract.contract_no = parsed.contract_no
        if parsed.contract_date and not contract.contract_date:
            contract.contract_date = parsed.contract_date
        if parsed.counterparty and not contract.counterparty:
            contract.counterparty = parsed.counterparty
        if parsed.project and not contract.project:
            contract.project = parsed.project

        for line in parsed.lines:
            stats["lines_parsed"] += 1
            if EXCLUDED_WORK_TYPES_RE.match(line.work_type_name):
                stats.setdefault("lines_excluded", 0)
                stats["lines_excluded"] += 1
                continue
            wt = _get_or_create_work_type(session, line.work_type_name, work_type_cache)
            existing = (
                session.query(OrderLine)
                .filter_by(contract_id=contract.id, work_type_id=wt.id, description=line.description)
                .one_or_none()
            )
            if existing is not None:
                if line.qty_ordered is not None and not existing.qty_ordered:
                    existing.qty_ordered = line.qty_ordered
                if line.sum_ordered is not None and not existing.sum_ordered:
                    existing.sum_ordered = line.sum_ordered
                stats["lines_updated"] += 1
            else:
                session.add(OrderLine(
                    contract_id=contract.id,
                    work_type_id=wt.id,
                    description=line.description,
                    qty_ordered=line.qty_ordered,
                    sum_ordered=line.sum_ordered,
                    qty_done_snapshot=line.qty_ordered,
                    sum_done_snapshot=line.sum_ordered,
                    qty_remaining_snapshot=0.0,
                    sum_remaining_snapshot=0.0,
                    import_batch_id=batch.id,
                ))
                stats["lines_added_closed"] += 1

    batch.rows_parsed = stats["lines_parsed"]
    batch.rows_matched = stats["lines_updated"] + stats["lines_added_closed"]
    batch.rows_unmatched = 0
    batch.notes = (
        f"matched={stats['contracts_matched']}, skipped_new={stats['contracts_skipped_new']}, "
        f"lines_added_closed={stats['lines_added_closed']}, lines_updated={stats['lines_updated']}"
    )
    session.commit()
    return stats
