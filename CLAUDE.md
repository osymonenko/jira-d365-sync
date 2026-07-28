# d365-time-entry

Заполнение недельного табеля: Jira → Excel (Python) и Excel → Time Entries в Microsoft Dynamics 365 (Playwright). Инструкция для пользователя — [README.md](README.md), здесь только техническая часть.
Статус (2026-07-28): слой Jira → Excel рабочий. В D365 реализован Stage 1 (Quick Create, дата, duration, поиск/создание задачи). Stage 2 ещё не реализован — будет через клик по ячейкам Weekly Time Entries grid.

## Commands

| Команда | Назначение |
|---|---|
| `start.bat` | WinForms GUI ([jira-sync.ps1](jira-sync.ps1)) поверх Python- и Node-слоёв |
| `jira-sync.bat` | тот же GUI, но не скрывает ошибки запуска PowerShell |
| `python scripts/jira-sync.py --command sync --file data\timesheet.xlsx --weeks 2026-06-15` | Jira → Excel (`test`, `read-weeks`, `sync`, `fill-standard`; `--month`, `--dry-run`) |
| `npm start -- --file data/timesheet.xlsx` | Прогон Stage 1 (создание задач + первого дня) |
| `npm start -- --file data/timesheet.xlsx --preflight-only` | Только preflight (Excel parse + URL/browser checks) |
| `npm start -- --file ... --week 2026-05-04` | Фильтр по неделе (понедельник, YYYY-MM-DD) |
| `npm run typecheck` | `tsc --noEmit` |
| `npm run lint` / `npm run format` | eslint / prettier по `src` |
| `npx playwright test` | Прогнать спеки из `tests/` (Test Explorer / `--ui` / `--debug`) |
| `npx playwright test --ui` | UI Mode: пошаговый trace, time-travel debugging |
| `npx playwright test --debug` | Inspector с pause-on-step |

## Architecture

- **Stage 1** (`fillEntryWithTaskLookup`) — для каждой уникальной задачи открывает Quick Create: Time Entry, заполняет дату/часы первого дня, ищет задачу. Если найдена — выбирает из dropdown и сохраняет. Если нет — жмёт New → Quick Create: Imported Project Task → сохраняет задачу → возвращается в основную панель → Save & Close.
- **Stage 2** — TODO. Старый подход (открытие Quick Create для каждого оставшегося дня) удалён. Новый подход: после Stage 1 задачи уже существуют как строки в "All Weekly Time Entries" grid. Нужно для каждой задачи и каждого её доп. дня кликнуть по ячейке (пересечение task × day) и ввести часы (`0.5` и т.д.).
- **Несколько недель** — ручной перезапуск с `--week <YYYY-MM-DD>` для каждой.

Модули:
- [src/index.ts](src/index.ts) — CLI (commander), оркестрирует Stage 1, retry-логика (`withRetry`, 2 попытки, 1.5с задержка).
- [src/d365.ts](src/d365.ts) — Playwright-клиент `D365Client`. Селекторы в `SELECTORS` — множественные OR-fallback'и, потому что D365 UI часто меняется. Экспортирует `D365Client.fromPage(page, url)` для использования в Playwright-спеках без повторной инициализации браузера.
- [src/excel.ts](src/excel.ts) — `parseTimesheet`, `firstEntriesPerTask`, `remainingEntries`, `uniqueTasks`. Формат Excel: row 0 = header, далее блоки `[weekStart, weekEnd, "", ...]` за которыми идут task-rows `[null, null, taskName, Mon, Tue, Wed, Thu, Fri]`. `remainingEntries` пока не используется в CLI — нужна будущему Stage 2.
- [src/preflight.ts](src/preflight.ts) — три чека: Excel-файл существует и парсится, `D365_URL` задан, browser доступен.
- [src/types.ts](src/types.ts) — `WeekBlock`, `TaskRow`, `TimeEntry`, `DayKey`.
- [jira-sync.ps1](jira-sync.ps1) — WinForms GUI (единственный; `gui.ps1` удалён). Запускает `python`/`node` как дочерний процесс, читает stdout/stderr через PowerShell runspaces (`Start-PyProc`), очередь + UI-таймер, sentinel `__EXIT__:N`. Settings-диалог пишет `.env` (`Save-JiraEnv`, без BOM) и `config/*.json` (`Write-JsonConfig`).
- [scripts/jira-sync.py](scripts/jira-sync.py) — Jira → Excel: 4 команды (`test`, `read-weeks`, `sync`, `fill-standard`), 6 JQL-шаблонов (`DEFAULT_JQL` + оверлей из `config/jql_queries.json`), правила названий (`DEFAULT_NAME_RULES`, режимы `count`/`priority`/`per_issue`), запись через openpyxl с сохранением гиперссылок и пересчётом `check sum`.
- [scripts/standard_tasks.py](scripts/standard_tasks.py) — чистая логика регулярных задач: `rows_for_week`, `is_sprint_end_week` (цикл от `SPRINT_ANCHOR`, длина из `SPRINT_LENGTH_WEEKS`), `DAY_COL` (Mon=D … Fri=H), QA-заглушки только для будущих недель.
- [scripts/month_filter.py](scripts/month_filter.py) — `clamp_week_to_month`, `month_bounds` для `--month`.
- `config/jql_queries.json`, `config/standard_tasks.json` — gitignored, создаются из Settings. Отсутствие/битый JSON = встроенные дефолты (см. `load_jql`, `load_name_rules`, `load_schedule`).

### Тестовая инфраструктура для отладки

Production-CLI (`npm start`) и Playwright Test (`npx playwright test`) делят один и тот же `D365Client` — нулевое дублирование селекторов. Спеки используются для пошаговой отладки в VS Code Test Explorer.

- [playwright.config.ts](playwright.config.ts) — `msedge` channel, `headless: false`, `workers: 1`, `trace: 'on'`.
- [tests/fixtures.ts](tests/fixtures.ts) — worker-scoped persistent Edge context на тот же `./edge-profile` (SSO куки шарятся с CLI). **Нельзя запускать одновременно `npm start` и `npx playwright test` — один лок профиля.**
- [tests/d365-flow.spec.ts](tests/d365-flow.spec.ts) — discrete `test()` блоки: login → navigate → openNewTimeEntry → fillEntryWithTaskLookup. Каждый шаг = отдельная строка в Test Explorer, можно ставить брейкпоинты внутри `D365Client` методов.
- **Pick Locator** в VS Code Playwright Test расширении работает на открытом Edge — клик по элементу в браузере = готовый локатор в IDE.

## Browser modes (`BROWSER_MODE` env)

| Mode | Когда использовать |
|---|---|
| `chrome-profile` (default в `.env.example` и в коде при пустом `BROWSER_MODE`) | Персистентный Edge (`channel: 'msedge'`) с профилем в `USER_DATA_DIR`. Куки SSO сохраняются между прогонами. **Рекомендованный.** Браузер — только Edge, это требование компании. |
| `cdp` | Подключение к уже-запущенному Chrome через `--remote-debugging-port=9222` (`CDP_URL`). Запускать Chrome вручную — кнопки в GUI для этого больше нет. |
| `persistent` | Bundled Chromium с профилем в `USER_DATA_DIR`. Требует `npx playwright install chromium`. |

## AMC Bridge SSO

D365 в AMC Bridge требует рабочей сессии Microsoft SSO в профиле браузера. При пустом профиле — редирект-цикл на login: нужно один раз залогиниться вручную в открывшемся окне Edge, дальше куки живут в `USER_DATA_DIR`. Автоматизация продолжается, как только видит навбар D365 (авто-detect, timeout 5 минут).

Кнопки "Sync Chrome Profile" / "Launch Debug Chrome" из старого `gui.ps1` в `jira-sync.ps1` не переносились.

## Conventions

- **Даты для D365**: формат `M/D/YYYY` (см. `toD365Date` в [src/excel.ts:9-11](src/excel.ts#L9-L11)). Без leading zeros.
- **Excel serial → Date**: `(serial - 25569) * 86400 * 1000` (offset для 1900 leap-year bug Lotus 1-2-3 совместимости).
- **Часы для D365 Payable Duration**: открываем dropdown и кликаем точную опцию по label'у. Маппинг: `<1h → "N minutes"` (15/30/45), `=1h → "1 hour"`, `>1h → "N.N hours"` или `"N hours"` (1.5, 2, 2.25, ... до 8). См. `hoursToD365DurationOption` в [src/d365.ts](src/d365.ts).
- **CLI флаги**: `--stage1-only` оставлен для обратной совместимости с GUI (alias дефолтного поведения, поскольку Stage 2 не реализован). `--stage2-only` отвергается с ошибкой.
- **Selectors**: множественные локаторы через запятую (`'A, B, C'` = first match) — Playwright OR. При сломе одного fallback'а — добавлять новый, не удалять старые.

## Gotchas

- `__AWAIT_LOGIN__` sentinel в stdout ([src/d365.ts](src/d365.ts)) задумывался как сигнал GUI показать кнопку "I'm logged in - Continue"; в `jira-sync.ps1` такой кнопки нет, поэтому строка просто попадает в лог, а продолжение идёт через авто-detect навбара. Ручной ack (`CONTINUE` в stdin) в коде остался и работает при запуске из консоли.
- В `chrome-profile` mode `userDataDir` принудительно резолвится через `path.resolve` — относительные пути в `.env` работают.
- При `cdp` mode код пытается переиспользовать существующую вкладку с D365-хостом; если её нет — берёт любую "обычную" вкладку (не chrome:// и не extension://) или открывает новую. Это намеренно — позволяет запустить тулзу не пугая открытые в Chrome табы.
- Excel парсер чувствителен к структуре: тaskName *должен* быть в col C, дни Mon-Fri в col D-H. Sheet — первый.
- `weekFilter` сравнивается с `weekStart + 1 day` (строка [src/excel.ts:62-64](src/excel.ts#L62-L64)) с tolerance ±1 день — потому что в Excel `weekStart` это воскресенье, а пользователь передаёт понедельник.

## Files / dirs

- `data/` — Excel-файлы (gitignored).
- `config/` — `jql_queries.json`, `standard_tasks.json` (gitignored, пишутся из Settings).
- `logs/<runId>/` — `run.log`, Playwright trace, скриншоты ошибок (gitignored).
- `browser-profile/`, `edge-profile/` — профили браузера (gitignored). Текущий `.env` использует `./edge-profile` — тот же, что и Playwright-фикстуры, отсюда запрет на одновременный запуск CLI и спеков.
- `.env` / `.env.example` — `D365_URL`, `BROWSER_MODE`, `CDP_URL`, `USER_DATA_DIR`, `COPY_TO_BILLABLE_DURATION`, `EXCEL_FILE`, `SPRINT_ANCHOR`, `SPRINT_LENGTH_WEEKS`, `JIRA_*`.

## Development

- TS strict mode, target ES2022, commonjs module ([tsconfig.json](tsconfig.json)).
- eslint ([eslint.config.js](eslint.config.js)) + prettier ([.prettierrc.json](.prettierrc.json)) по `src` — `npm run lint`, `npm run format`.
- Тесты TS: `npm test` (vitest, [src/excel.test.ts](src/excel.test.ts)) — `parseTimesheet`, `uniqueTasks`, `firstEntriesPerTask`, `remainingEntries`.
- Тесты Python: автономные скрипты `scripts/test_*.py` (без pytest, печатают `ALL PASS`) — запускать из `scripts/`, т.к. импортируют модули напрямую. Есть и два headless GUI-теста на PowerShell (`scripts/test_*_gui.ps1`) через `PerformClick()`.
- Python 3.10+ (аннотации `str | None`) + `openpyxl`; сетевые вызовы — на `urllib`, без `requests`.
