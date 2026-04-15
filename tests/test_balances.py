from datetime import date

from app.db import init_db, make_session_factory
from app.models import Act, ActLine, Contract, OrderLine, WorkType
from app.services.balances import balance, fact


def _setup(tmp_path):
    engine = init_db(tmp_path / "b.db", drop_first=True)
    Session = make_session_factory(engine)
    with Session() as s:
        c = Contract(number="ТЕСТ-001", counterparty="ООО Тест", project="Объект 1")
        wt = WorkType(name="Монтаж", unit="шт")
        s.add_all([c, wt])
        s.flush()
        ol = OrderLine(
            contract_id=c.id,
            work_type_id=wt.id,
            description="Монтаж оборудования",
            sum_ordered=1000.0,
            qty_ordered=10.0,
            sum_done_snapshot=0.0,
            qty_done_snapshot=0.0,
            sum_remaining_snapshot=1000.0,
            qty_remaining_snapshot=10.0,
        )
        s.add(ol)
        s.commit()
        return Session, c.id, ol.id


def test_balance_fallback_snapshot_when_no_acts(tmp_path):
    Session, _, ol_id = _setup(tmp_path)
    with Session() as s:
        b = balance(s, ol_id)
        assert b.from_acts is False
        assert b.qty_remaining == 10.0
        assert b.sum_remaining == 1000.0
        assert b.qty_done == 0.0


def test_balance_from_acts(tmp_path):
    Session, contract_id, ol_id = _setup(tmp_path)
    with Session() as s:
        a1 = Act(contract_id=contract_id, number="А-1", act_date=date(2026, 4, 10), total_sum=300)
        s.add(a1); s.flush()
        s.add(ActLine(act_id=a1.id, order_line_id=ol_id, qty=3, sum=300))
        a2 = Act(contract_id=contract_id, number="А-2", act_date=date(2026, 5, 5), total_sum=200)
        s.add(a2); s.flush()
        s.add(ActLine(act_id=a2.id, order_line_id=ol_id, qty=2, sum=200))
        s.commit()

        b = balance(s, ol_id)
        assert b.from_acts is True
        assert b.qty_done == 5.0
        assert b.qty_remaining == 5.0
        assert b.sum_done == 500.0
        assert b.sum_remaining == 500.0


def test_fact_by_period(tmp_path):
    Session, contract_id, ol_id = _setup(tmp_path)
    with Session() as s:
        a1 = Act(contract_id=contract_id, number="А-1", act_date=date(2026, 4, 10), total_sum=300)
        s.add(a1); s.flush()
        s.add(ActLine(act_id=a1.id, order_line_id=ol_id, qty=3, sum=300))
        a2 = Act(contract_id=contract_id, number="А-2", act_date=date(2026, 5, 5), total_sum=200)
        s.add(a2); s.flush()
        s.add(ActLine(act_id=a2.id, order_line_id=ol_id, qty=2, sum=200))
        s.commit()

        q_apr, s_apr = fact(s, ol_id, "2026-04")
        assert q_apr == 3.0 and s_apr == 300.0
        q_may, s_may = fact(s, ol_id, "2026-05")
        assert q_may == 2.0 and s_may == 200.0
        q_jun, s_jun = fact(s, ol_id, "2026-06")
        assert q_jun == 0.0 and s_jun == 0.0
