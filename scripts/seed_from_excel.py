"""Наполнить SQLite из выгрузки 1С.

Usage:
    python scripts/seed_from_excel.py Выгрузка.xlsx [--reset]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# чтобы можно было запускать скрипт без `pip install -e .`
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from sqlalchemy import func
from app.db import init_db, make_session_factory
from app.importer import import_orders
from app.models import Act, ActLine, Contract, OrderLine, WorkType


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("excel", help="Путь к .xlsx выгрузки 1С")
    ap.add_argument("--reset", action="store_true", help="Пересоздать БД с нуля")
    ap.add_argument("--sheet", default="TDSheet")
    ap.add_argument("--by", default=None, help="Имя пользователя (для import_batch.uploaded_by)")
    args = ap.parse_args()

    excel_path = Path(args.excel)
    if not excel_path.exists():
        print(f"[ERR] файл не найден: {excel_path}", file=sys.stderr)
        return 2

    engine = init_db(drop_first=args.reset)
    SessionLocal = make_session_factory(engine)

    with SessionLocal() as session:
        batch = import_orders(session, excel_path, uploaded_by=args.by, sheet=args.sheet)
        print(f"[OK] Импорт завершён: batch_id={batch.id}, файл={batch.filename}")
        print(f"     строк распарсено: {batch.rows_parsed}, обновлено существующих: {batch.rows_matched}")

        n_contracts = session.query(func.count(Contract.id)).scalar()
        n_work_types = session.query(func.count(WorkType.id)).scalar()
        n_order_lines = session.query(func.count(OrderLine.id)).scalar()
        n_acts = session.query(func.count(Act.id)).scalar()
        n_act_lines = session.query(func.count(ActLine.id)).scalar()
        sum_ordered = session.query(func.coalesce(func.sum(OrderLine.sum_ordered), 0.0)).scalar() or 0.0
        sum_remaining = session.query(func.coalesce(func.sum(OrderLine.sum_remaining_snapshot), 0.0)).scalar() or 0.0

        print()
        print("Сводка по БД:")
        print(f"  Договоров      : {n_contracts}")
        print(f"  Видов работ    : {n_work_types}")
        print(f"  Строк заказов  : {n_order_lines}")
        print(f"  Актов          : {n_acts}")
        print(f"  Строк актов    : {n_act_lines}")
        print(f"  Σ Заказано, ₽  : {sum_ordered:,.2f}")
        print(f"  Σ Остаток, ₽   : {sum_remaining:,.2f}")

        print()
        print("Топ-5 договоров по остатку:")
        rows = (
            session.query(
                Contract.number,
                Contract.counterparty,
                func.count(OrderLine.id),
                func.coalesce(func.sum(OrderLine.sum_remaining_snapshot), 0.0),
            )
            .join(OrderLine, OrderLine.contract_id == Contract.id)
            .group_by(Contract.id)
            .order_by(func.sum(OrderLine.sum_remaining_snapshot).desc())
            .limit(5)
            .all()
        )
        for num, cp, nlines, rem in rows:
            cp_short = (cp or "—")[:40]
            print(f"  {num:<16} {cp_short:<40} строк={nlines:>3}  остаток={rem:>18,.2f}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
