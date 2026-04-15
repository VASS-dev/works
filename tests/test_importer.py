from datetime import date
from pathlib import Path

import pytest

from app.db import init_db, make_session_factory
from app.importer import _split_contract_cell, import_orders, parse_workbook
from app.models import Contract, OrderLine, WorkType

EXCEL = Path(__file__).resolve().parent.parent / "Выгрузка.xlsx"


def test_split_contract_cell_basic():
    cell = "НФФР-003678, Заказ покупателя 3678 от 17.11.2025 , ЛСР. СТРОИТЕЛЬСТВО-СЗ ООО, Санкт-Петербург, муниципальный округ Полюстрово, Муринская дорога, земельный участок 5"
    number, label, dt, cp, project = _split_contract_cell(cell)
    assert number == "НФФР-003678"
    assert label.startswith("Заказ покупателя 3678")
    assert dt == date(2025, 11, 17)
    assert cp == "ЛСР. СТРОИТЕЛЬСТВО-СЗ ООО"
    assert project.startswith("Санкт-Петербург")
    assert "земельный участок 5" in project


def test_split_without_project():
    cell = "АБВ-001, Заказ 1, Контрагент N, "
    number, label, dt, cp, project = _split_contract_cell(cell)
    assert number == "АБВ-001"
    assert cp == "Контрагент N"
    assert project is None


@pytest.mark.skipif(not EXCEL.exists(), reason="Выгрузка.xlsx отсутствует")
def test_parse_workbook_yields_288_rows():
    # 290 строк всего, 2 строки шапки → 288 строк данных (плюс-минус пустые)
    rows = list(parse_workbook(EXCEL))
    assert 280 <= len(rows) <= 290
    first = rows[0]
    assert first.contract_number.startswith(("НФФР", "СМ", "С"))
    assert first.work_type_name
    assert first.description


@pytest.mark.skipif(not EXCEL.exists(), reason="Выгрузка.xlsx отсутствует")
def test_import_orders_into_sqlite(tmp_path):
    db_file = tmp_path / "test.db"
    engine = init_db(db_file, drop_first=True)
    Session = make_session_factory(engine)
    with Session() as s:
        batch = import_orders(s, EXCEL)
        assert batch.rows_parsed > 0

        n_contracts = s.query(Contract).count()
        n_lines = s.query(OrderLine).count()
        n_wt = s.query(WorkType).count()

        assert n_contracts > 0
        assert n_lines == batch.rows_parsed
        assert n_wt > 0

        # естественный ключ: (contract, work_type, description) уникален
        dupes = (
            s.query(OrderLine.contract_id, OrderLine.work_type_id, OrderLine.description)
            .group_by(OrderLine.contract_id, OrderLine.work_type_id, OrderLine.description)
            .having(__import__("sqlalchemy").func.count() > 1)
            .all()
        )
        assert dupes == []


@pytest.mark.skipif(not EXCEL.exists(), reason="Выгрузка.xlsx отсутствует")
def test_reimport_is_idempotent(tmp_path):
    db_file = tmp_path / "test.db"
    engine = init_db(db_file, drop_first=True)
    Session = make_session_factory(engine)
    with Session() as s:
        import_orders(s, EXCEL)
        n1 = s.query(OrderLine).count()
    with Session() as s:
        import_orders(s, EXCEL)
        n2 = s.query(OrderLine).count()
    assert n1 == n2
