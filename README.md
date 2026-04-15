# WorksPlanning

FastAPI + SQLite + HTMX. Планирование работ, план/факт, договоры.

## Локальный запуск

```bash
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Без `APP_PASSWORD` авторизация отключена (удобно в dev). Приложение — `http://localhost:8000`.

## Конфигурация (переменные окружения)

Смотри `.env.example`. Ключевые:

- `DB_URL` — строка подключения. По умолчанию SQLite в `./data/` (dev) или `/data/` (контейнер).
- `APP_USER`, `APP_PASSWORD` — пара для BasicAuth. Если `APP_PASSWORD` пуст — auth off.

## Деплой на Koyeb (бесплатно)

1. Запушить репозиторий на GitHub.
2. Koyeb → **Create App** → Deploy from GitHub → ветка `main`.
3. Builder = Dockerfile (авто).
4. Instance = **Nano (Free)**, регион Frankfurt.
5. **Volumes**: создать `worksplanning-data`, mount path `/data`.
6. **Env vars**: `APP_PASSWORD=<пароль>`, `APP_USER=team`.
7. **Health check**: HTTP `/healthz`.
8. Deploy. URL: `https://<slug>.koyeb.app`.

Подробности — в `.claude/plans/quirky-booping-sphinx.md`.
