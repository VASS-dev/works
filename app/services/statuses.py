"""Константы статусов для PlanEntry и ExecutionEntry.

Храним как простые строковые атрибуты класса (не Enum), потому что:
- столбец status в БД — VARCHAR, а не Enum-тип;
- Jinja-шаблоны сравнивают со строками напрямую;
- отсутствует риск импорт-циклов и магии __members__.

Важно: значения совпадают с тем, что уже лежит в проде — ничего не мигрируется.
"""
from __future__ import annotations


class PlanStatus:
    DRAFT = "draft"
    APPROVED = "approved"

    ALL = (DRAFT, APPROVED)


class ExecutionStatus:
    DRAFT = "draft"
    SIGNED = "signed"

    ALL = (DRAFT, SIGNED)
