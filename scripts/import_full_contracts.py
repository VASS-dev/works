"""Импорт файла 'Договоры текущие.xlsx' в текущую БД.

Использование:
    python scripts/import_full_contracts.py "Договоры текущие.xlsx"
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.db import init_db, make_session_factory
from app.importer_full import import_full_contracts


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: import_full_contracts.py <path-to-xlsx>")
        return 2
    path = Path(sys.argv[1])
    if not path.exists():
        print(f"not found: {path}")
        return 2
    engine = init_db()
    Session = make_session_factory(engine)
    with Session() as session:
        stats = import_full_contracts(session, path)
    for k, v in stats.items():
        print(f"  {k}: {v}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
