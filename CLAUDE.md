# d365-time-entry

Заполнение недельного табеля: Jira → Excel (Python) и Excel → Time Entries в Microsoft Dynamics 365 (Playwright). Инструкция для пользователя — [README.md](README.md), здесь только техническая часть.
Статус (2026-09-22): слой Jira → Excel рабочий (`sync`, `fill-standard`, `balance`). Заливка в D365 идёт через Web API (`--mode api`, дефолт) — один проход по всем дням, без Quick Create и без переключения недель. Старый UI-путь остался под `--mode ui` как запасной.

## Commands

| Команда | Назначение |
|---|---|
| `start.bat` | WinForms GUI ([jira-sync.ps1](jira-sync.ps1)) поверх Python- и Node-слоёв |
| `jira-sync.bat` | тот же GUI, но не скрывает ошибки запуска PowerShell |
| `python scripts/jira-sync.py --command sync --file data\timesheet.xlsx --weeks 2026-06-15` | Jira → Excel (`test`, `read-weeks`, `sync`, `fill-standard`, `balance`; `--month`, `--dry-run`) |
| `python scripts/jira-sync.py --command balance --file data\timesheet.xlsx --month 2026-09` | Добить пустые строки задач часами до 8 ч/день (`--reset`, `--hours-per-day`, `--dry-run`) |
| `.\submit.bat --file data\timesheet.xlsx` | Заливка всего файла в D365 (создание задач + все дни) |
| `.\submit.bat --file data\timesheet.xlsx --week 2026-09-21` | Заливка одной недели (понедельник, YYYY-MM-DD) |
| `.\submit.bat --file data\timesheet.xlsx --preflight-only` | Только preflight (Excel parse + URL/browser checks) |
| `.\submit.bat --file ... --mode ui` | Запасной путь через Quick Create (только первый день каждой задачи) |
| `npm run typecheck` | `tsc --noEmit` |
| `node --require ts-node/register/transpile-only scripts/d365-query.ts <queries.json> <outDir>` | Read-only запросник к Web API для разведки |
| `npm run lint` / `npm run format` | eslint / prettier по `src` |
| `npx playwright test` | Прогнать спеки из `tests/` (Test Explorer / `--ui` / `--debug`) |
| `npx playwright test --ui` | UI Mode: пошаговый trace, time-travel debugging |
| `npx playwright test --debug` | Inspector с pause-on-step |

## Architecture

- **`--mode api` (дефолт)** — `runApiMode` в [src/index.ts](src/index.ts) поверх [src/d365-api.ts](src/d365-api.ts). Браузер поднимается только чтобы получить авторизованную сессию; дальше всё через `fetch` к `/api/data/v9.2/` из контекста страницы. Один проход по всем записям Excel: деления на Stage 1 / Stage 2 нет — дата это поле записи, а не то, какую неделю показывает грид. Переключать недели тоже не нужно, поэтому `--week` здесь лишь фильтр Excel.
- **Payable + Billable** — на каждую ячейку Excel создаются ДВЕ записи `msdyn_timeentry`, различающиеся `amc_hourstype` (`100000000` Payable / `100000001` Billable). В этом тенанте это отдельные строки, а не два поля одной записи. Плагин двойник не создаёт — обе постим явно.
- **`--mode ui`** — старый путь через Quick Create (`fillEntryWithTaskLookup`): для каждой уникальной задачи открывает панель, заполняет дату/часы первого дня, ищет задачу, при отсутствии создаёт через под-панель Imported Project Task. Заполняет только первый день каждой задачи (Stage 2 в UI так и не был написан), поэтому годится лишь как запасной вариант, если Web API закроют.

Модули:
- [src/index.ts](src/index.ts) — CLI (commander). `--mode api|ui`; `runApiMode` — заливка через Web API с пропуском уже существующих записей; UI-путь сохраняет retry-логику (`withRetry`, 2 попытки, 1.5с задержка).
- [src/d365-api.ts](src/d365-api.ts) — `D365Api`: OData-запросы из страницы. Payload создания выверен бисектом ([scripts/d365-bisect-create.ts](scripts/d365-bisect-create.ts)): **обязателен `msdyn_date`**, без него тенант отвечает `0x80040265` про незавершённый апгрейд Field Service — вводящее в заблуждение сообщение, настоящая причина в его хвосте (`missing the required field Date`). Всё остальное тенант дозаполняет сам: `msdyn_bookableresource`, `msdyn_timeentrysettingId`, `msdyn_type=Work`, `msdyn_entrystatus=Draft`, `amc_payableduration`, `amc_enablecutoff`. Проект и команда берутся из последней собственной записи (`resolveContext`), а не из `.env` — GUID в конфиге протух бы при переводе на другой проект. Задачи живут в `amc_importedprojecttasks` (поля `amc_name`, `amc_project`, `amc_tasktype=100000000`).
- [src/d365.ts](src/d365.ts) — Playwright-клиент `D365Client`. Селекторы в `SELECTORS` — множественные OR-fallback'и, потому что D365 UI часто меняется. Экспортирует `D365Client.fromPage(page, url)` для использования в Playwright-спеках без повторной инициализации браузера.
- [src/excel.ts](src/excel.ts) — `parseTimesheet`, `firstEntriesPerTask`, `remainingEntries`, `uniqueTasks`. Формат Excel: row 0 = header, далее блоки `[weekStart, weekEnd, "", ...]` за которыми идут task-rows `[null, null, taskName, Mon, Tue, Wed, Thu, Fri]`. `remainingEntries` в CLI не используется: режим api берёт все записи сразу, а не «первый день + остальные».
- [src/preflight.ts](src/preflight.ts) — три чека: Excel-файл существует и парсится, `D365_URL` задан, browser доступен.
- [src/types.ts](src/types.ts) — `WeekBlock`, `TaskRow`, `TimeEntry`, `DayKey`.
- [jira-sync.ps1](jira-sync.ps1) — WinForms GUI (единственный; `gui.ps1` удалён). Запускает `python`/`node` как дочерний процесс, читает stdout/stderr через PowerShell runspaces (`Start-PyProc`), очередь + UI-таймер, sentinel `__EXIT__:N`. Settings-диалог пишет `.env` (`Save-JiraEnv`, без BOM) и `config/*.json` (`Write-JsonConfig`).
- [scripts/jira-sync.py](scripts/jira-sync.py) — Jira → Excel: 5 команд (`test`, `read-weeks`, `sync`, `fill-standard`, `balance`), 6 JQL-шаблонов (`DEFAULT_JQL` + оверлей из `config/jql_queries.json`), правила названий (`DEFAULT_NAME_RULES`, режимы `count`/`priority`/`per_issue`), запись через openpyxl с сохранением гиперссылок и пересчётом `check sum`. `_merge_duplicate_task_rows` (зовётся в начале `balance`) схлопывает строки с одинаковым именем внутри недели в одну: Jira отдаёт несколько тикетов под одним именем («Automation test maintenance» ×4), а в табеле это одна работа. Часы переливаются в первую строку, дубли удаляются через `_delete_rows_preserving_hyperlinks` — зеркало insert-хелпера, потому что openpyxl при сдвиге строк оставляет гиперссылки на старых координатах. Кратность уходит в раскладку: слитая задача получает долю по числу вобранных тикетов и вправе занять несколько дней.
- [scripts/standard_tasks.py](scripts/standard_tasks.py) — чистая логика регулярных задач: `rows_for_week`, `is_sprint_end_week` (цикл от `SPRINT_ANCHOR`, длина из `SPRINT_LENGTH_WEEKS`), `DAY_COL` (Mon=D … Fri=H), QA-заглушки только для будущих недель.
- [scripts/balance_hours.py](scripts/balance_hours.py) — чистая логика добивки часов: `allocate` (ровное деление остатка между задачами), `spread` (круговая раскладка по дням), `pack` (укладка подряд), `plan_week`. Считает в четвертях часа (int), чтобы не ловить float-ошибки, но выдаёт только кратное 0.5 ч (`STEP_Q`) — дробные 0.25/0.75 в табеле неудобны. Задачи из `SPREAD_PREFIXES` (`Bug verification`, `Investigation issue`) размазываются кусками по 0.5 ч через всю неделю (`spread_all` — очередь общая на все такие задачи, иначе остаток, скопившийся в одном дне, падал бы целиком на одну из них). Остальные кладутся ОДНИМ блоком в один день: переносов «часть часов в среду, остаток в четверг» быть не должно, поэтому дни сначала получают задачи (`split_counts`), потом бюджет (`distribute`), и только потом часы делятся внутри дня. Задачи со словом `automation` весят вдвое (`AUTOMATION_WEIGHT`) — излишек сверх минимума делится по весам. Исключение из правила «один день на задачу» — **будущая неделя** (`spread_all_tasks`, включается когда понедельник блока позже сегодня): там нет отчёта о сделанном, только план, поэтому каждая задача размазывается по всем дням, иначе получаются неправдоподобные глыбы по 6.5 ч. `normalize_task_name` срезает хвосты-счётчики (`Investigation issue 2`, `Bug verification P1-2, P2-4`, `… 2 test items`), но не трогает имена, оканчивающиеся идентификатором (`… story ID GT2-1662`) — иначе слиплись бы разные истории. Регулярные задачи из `schedule` и строки с уже проставленными часами не трогаются — поэтому повторный запуск `balance` не задваивает; чтобы пересчитать заново, нужен `--reset` (стирает раскладку в строках задач, регулярные не трогает). Кнопка «Balance 8h ⚖» в GUI спрашивает про `--reset` диалогом.
- [scripts/month_filter.py](scripts/month_filter.py) — `clamp_week_to_month`, `month_bounds` для `--month`.
- `config/jql_queries.json`, `config/standard_tasks.json` — gitignored, создаются из Settings. Отсутствие/битый JSON = встроенные дефолты (см. `load_jql`, `load_name_rules`, `load_schedule`).

### Тестовая инфраструктура для отладки

Production-CLI (`npm start`) и Playwright Test (`npx playwright test`) делят один и тот же `D365Client` — нулевое дублирование селекторов. Спеки используются для пошаговой отладки в VS Code Test Explorer.

- [playwright.config.ts](playwright.config.ts) — `msedge` channel, `headless: false`, `workers: 1`, `trace: 'on'`.
- [tests/fixtures.ts](tests/fixtures.ts) — worker-scoped persistent Edge context на тот же `./edge-profile` (SSO куки шарятся с CLI). **Нельзя запускать одновременно `npm start` и `npx playwright test` — один лок профиля.**
- [tests/d365-flow.spec.ts](tests/d365-flow.spec.ts) — discrete `test()` блоки: login → navigate → openNewTimeEntry → fillEntryWithTaskLookup. Каждый шаг = отдельная строка в Test Explorer, можно ставить брейкпоинты внутри `D365Client` методов.
- **Pick Locator** в VS Code Playwright Test расширении работает на открытом Edge — клик по элементу в браузере = готовый локатор в IDE.

### Preflight в Python-слое

`sync`, `fill-standard` и `balance` первым делом зовут `check_excel_writable` (файл есть → нет `~$<имя>.xlsx` → открывается на запись → парсится → есть блоки недель), а `sync`/`test` ещё и `check_network` (TCP-пробник хоста Jira, 4 с). Обе падают с `exit 1`. Смысл — не считать минуту JQL-запросов, чтобы упереться в залоченный Excel на последнем шаге. `wb.save()` при `PermissionError` тоже завершается `exit 1`: GUI красит статус по коду выхода, и раньше провал сохранения выглядел зелёным «Sync complete».

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
- **CLI флаги**: `--stage1-only` оставлен для обратной совместимости (alias дефолтного поведения). `--stage2-only` отвергается с ошибкой: в режиме api отдельной второй стадии нет.
- **Запуск Node-слоя из терминала**: только [submit.bat](submit.bat) (или `node --require ts-node/register/transpile-only src/index.ts ...`). `npm start -- --file ...` работает в Git Bash, но в **PowerShell** npm не доносит флаги: до ts-node доезжают только значения (`src/index.ts data/timesheet.xlsx 2026-09-28`), и commander падает на `required option '-f, --file <path>' not specified`. Выглядит как ошибка пользователя, хотя команда набрана верно. GUI не задет — он и так зовёт `node` напрямую.
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
- Тесты Python: автономные скрипты `scripts/test_*.py` (без pytest, печатают `ALL PASS`). Большинство запускается из `scripts/` (импортируют модули напрямую), но `test_name_rules`, `test_jql_config`, `test_standard_config`, `test_generate_week_rows_names` собирают путь от текущей директории и требуют запуска из корня репозитория (`python scripts/test_name_rules.py`). Есть и два headless GUI-теста на PowerShell (`scripts/test_*_gui.ps1`) через `PerformClick()`.
- Python 3.10+ (аннотации `str | None`) + `openpyxl`; сетевые вызовы — на `urllib`, без `requests`.
