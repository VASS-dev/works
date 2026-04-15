"""Конфигурация приложения через переменные окружения.

В dev-режиме достаточно дефолтов; в Koyeb/Docker переменные задаются через UI/env.
"""
from __future__ import annotations

import os
import secrets
from pathlib import Path


def _default_db_url() -> str:
    # В контейнере /data — смонтированный том; локально — ./data/worksplanning.db.
    if os.path.isdir("/data"):
        return "sqlite:////data/worksplanning.db"
    root = Path(__file__).resolve().parent.parent
    return f"sqlite:///{root / 'data' / 'worksplanning.db'}"


DB_URL: str = os.environ.get("DB_URL") or _default_db_url()
APP_USER: str = os.environ.get("APP_USER", "team")
APP_PASSWORD: str | None = os.environ.get("APP_PASSWORD")  # если None — auth отключен (dev)
SESSION_SECRET: str = os.environ.get("SESSION_SECRET") or secrets.token_hex(16)

# Явный переключатель dev-режима. В dev можно работать без пароля.
# В проде ВСЕГДА ставьте ENV=prod, иначе приложение потребует APP_PASSWORD.
ENV: str = os.environ.get("ENV", "dev").lower()

# Домены, с которых разрешены state-changing запросы (CSRF-защита через Origin).
# Пустая строка = разрешать все same-origin (по Host-заголовку).
ALLOWED_ORIGINS: list[str] = [
    o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "").split(",") if o.strip()
]


def validate() -> None:
    """Проверка обязательных переменных в prod-режиме."""
    if ENV == "prod" and not APP_PASSWORD:
        raise RuntimeError(
            "ENV=prod, но APP_PASSWORD не задан — приложение было бы открыто всем. "
            "Установите APP_PASSWORD в окружении или ENV=dev для локальной работы."
        )
