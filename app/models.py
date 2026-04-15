from __future__ import annotations

from datetime import date, datetime
from typing import Optional

from sqlalchemy import ForeignKey, String, Float, Integer, Date, DateTime, Text, Boolean, UniqueConstraint
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


class Base(DeclarativeBase):
    pass


class Contract(Base):
    __tablename__ = "contract"

    id: Mapped[int] = mapped_column(primary_key=True)
    number: Mapped[str] = mapped_column(String(64), unique=True, index=True)  # НФФР-003678
    order_label: Mapped[Optional[str]] = mapped_column(String(256))           # "Заказ покупателя 3678 от 17.11.2025"
    order_date: Mapped[Optional[date]] = mapped_column(Date)
    counterparty: Mapped[Optional[str]] = mapped_column(String(256))          # Контрагент
    project: Mapped[Optional[str]] = mapped_column(Text)                      # Проект/адрес (может содержать запятые)
    contract_no: Mapped[Optional[str]] = mapped_column(String(128))           # № договора (внешний, напр. "4-ЦГ-132/18/215-ГП")
    contract_date: Mapped[Optional[date]] = mapped_column(Date)
    status: Mapped[Optional[str]] = mapped_column(String(64))
    source_id_1c: Mapped[Optional[str]] = mapped_column(String(128))

    order_lines: Mapped[list["OrderLine"]] = relationship(back_populates="contract", cascade="all, delete-orphan")
    acts: Mapped[list["Act"]] = relationship(back_populates="contract", cascade="all, delete-orphan")


class WorkType(Base):
    __tablename__ = "work_type"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(256), unique=True, index=True)   # Номенклатура
    unit: Mapped[Optional[str]] = mapped_column(String(32))                    # ед. изм. (если выяснится)


class OrderLine(Base):
    """Строка заказа = конкретная работа по договору."""
    __tablename__ = "order_line"
    __table_args__ = (
        UniqueConstraint("contract_id", "work_type_id", "description", name="uq_order_line_natural"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    contract_id: Mapped[int] = mapped_column(ForeignKey("contract.id"), index=True)
    work_type_id: Mapped[int] = mapped_column(ForeignKey("work_type.id"), index=True)
    description: Mapped[str] = mapped_column(Text)                            # Содержание

    # Денежные метрики из 1С: "Заказано / Выполнено / Осталось" (руб)
    sum_ordered: Mapped[Optional[float]] = mapped_column(Float)
    sum_done_snapshot: Mapped[Optional[float]] = mapped_column(Float)         # из выгрузки, для сверки
    sum_remaining_snapshot: Mapped[Optional[float]] = mapped_column(Float)

    # Натуральные метрики: "Заказано / Отгружено / Осталось" (шт/м2/...)
    qty_ordered: Mapped[Optional[float]] = mapped_column(Float)
    qty_done_snapshot: Mapped[Optional[float]] = mapped_column(Float)
    qty_remaining_snapshot: Mapped[Optional[float]] = mapped_column(Float)

    is_active: Mapped[bool] = mapped_column(Boolean, default=True)
    source_id_1c: Mapped[Optional[str]] = mapped_column(String(128))
    import_batch_id: Mapped[Optional[int]] = mapped_column(ForeignKey("import_batch.id"))

    contract: Mapped["Contract"] = relationship(back_populates="order_lines")
    work_type: Mapped["WorkType"] = relationship()
    act_lines: Mapped[list["ActLine"]] = relationship(back_populates="order_line", cascade="all, delete-orphan")
    plan_entries: Mapped[list["PlanEntry"]] = relationship(back_populates="order_line", cascade="all, delete-orphan")


class Act(Base):
    __tablename__ = "act"

    id: Mapped[int] = mapped_column(primary_key=True)
    contract_id: Mapped[int] = mapped_column(ForeignKey("contract.id"), index=True)
    number: Mapped[Optional[str]] = mapped_column(String(64))
    act_date: Mapped[date] = mapped_column(Date, index=True)
    total_sum: Mapped[Optional[float]] = mapped_column(Float)
    source_id_1c: Mapped[Optional[str]] = mapped_column(String(128))
    import_batch_id: Mapped[Optional[int]] = mapped_column(ForeignKey("import_batch.id"))

    contract: Mapped["Contract"] = relationship(back_populates="acts")
    lines: Mapped[list["ActLine"]] = relationship(back_populates="act", cascade="all, delete-orphan")


class ActLine(Base):
    __tablename__ = "act_line"

    id: Mapped[int] = mapped_column(primary_key=True)
    act_id: Mapped[int] = mapped_column(ForeignKey("act.id"), index=True)
    order_line_id: Mapped[int] = mapped_column(ForeignKey("order_line.id"), index=True)
    qty: Mapped[Optional[float]] = mapped_column(Float)
    sum: Mapped[Optional[float]] = mapped_column(Float)
    description: Mapped[Optional[str]] = mapped_column(Text)
    source_id_1c: Mapped[Optional[str]] = mapped_column(String(128))

    act: Mapped["Act"] = relationship(back_populates="lines")
    order_line: Mapped["OrderLine"] = relationship(back_populates="act_lines")


class PlanEntry(Base):
    __tablename__ = "plan_entry"
    __table_args__ = (
        UniqueConstraint("order_line_id", "period", name="uq_plan_orderline_period"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    order_line_id: Mapped[int] = mapped_column(ForeignKey("order_line.id"), index=True)
    period: Mapped[str] = mapped_column(String(7), index=True)                # YYYY-MM
    plan_qty: Mapped[Optional[float]] = mapped_column(Float)
    plan_sum: Mapped[Optional[float]] = mapped_column(Float)
    note: Mapped[Optional[str]] = mapped_column(Text)
    status: Mapped[str] = mapped_column(String(16), default="draft")          # draft / approved
    approved_at: Mapped[Optional[datetime]] = mapped_column(DateTime)
    approved_by: Mapped[Optional[str]] = mapped_column(String(64))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    author: Mapped[Optional[str]] = mapped_column(String(64))

    order_line: Mapped["OrderLine"] = relationship(back_populates="plan_entries")


class ExecutionEntry(Base):
    """Плановый акт: строка исполнения по утверждённому плану."""
    __tablename__ = "execution_entry"
    __table_args__ = (
        UniqueConstraint("plan_entry_id", name="uq_execution_plan"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    plan_entry_id: Mapped[int] = mapped_column(ForeignKey("plan_entry.id"), index=True)
    qty_fact: Mapped[Optional[float]] = mapped_column(Float)
    sum_fact: Mapped[Optional[float]] = mapped_column(Float)
    status: Mapped[str] = mapped_column(String(16), default="draft")  # draft / signed
    act_line_id: Mapped[Optional[int]] = mapped_column(ForeignKey("act_line.id"))
    signed_at: Mapped[Optional[datetime]] = mapped_column(DateTime)
    signed_by: Mapped[Optional[str]] = mapped_column(String(64))
    note: Mapped[Optional[str]] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    plan_entry: Mapped["PlanEntry"] = relationship()


class PlanHistory(Base):
    __tablename__ = "plan_history"

    id: Mapped[int] = mapped_column(primary_key=True)
    plan_entry_id: Mapped[int] = mapped_column(ForeignKey("plan_entry.id"), index=True)
    old_qty: Mapped[Optional[float]] = mapped_column(Float)
    new_qty: Mapped[Optional[float]] = mapped_column(Float)
    old_sum: Mapped[Optional[float]] = mapped_column(Float)
    new_sum: Mapped[Optional[float]] = mapped_column(Float)
    changed_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    changed_by: Mapped[Optional[str]] = mapped_column(String(64))


class ImportBatch(Base):
    __tablename__ = "import_batch"

    id: Mapped[int] = mapped_column(primary_key=True)
    filename: Mapped[str] = mapped_column(String(256))
    uploaded_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    uploaded_by: Mapped[Optional[str]] = mapped_column(String(64))
    kind: Mapped[str] = mapped_column(String(32))  # orders / acts / snapshot
    rows_parsed: Mapped[int] = mapped_column(Integer, default=0)
    rows_matched: Mapped[int] = mapped_column(Integer, default=0)
    rows_unmatched: Mapped[int] = mapped_column(Integer, default=0)
    notes: Mapped[Optional[str]] = mapped_column(Text)
