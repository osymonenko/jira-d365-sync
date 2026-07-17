# d365-time-entry

Автоматическое заполнение Time Entries в Microsoft Dynamics 365 из Excel-табеля через Playwright.
Статус (2026-05-20): Stage 1 в работе (открывает Quick Create, заполняет дату, выбирает duration). Stage 2 ещё не реализован — будет через клик по ячейкам Weekly Time Entries grid.

## Commands

| Команда | Назначение |
|---|---|
| `npm start -- --file data/timesheet.xlsx` | Прогон Stage 1 (создание задач + первого дня) |
| `npm start -- --file data/timesheet.xlsx --preflight-only` | Только preflight (Excel parse + URL/browser checks) |
| `npm start -- --file ... --week 2026-05-04` | Фильтр по неделе (понедельник, YYYY-MM-DD) |
| `npm run typecheck` | `tsc --noEmit` |
| `start.bat` | WinForms GUI ([gui.ps1](gui.ps1)) поверх CLI |
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
- [gui.ps1](gui.ps1) — WinForms GUI: запускает node-процесс, читает stdout/stderr через PowerShell runspaces, перехватывает sentinel'ы `__AWAIT_LOGIN__` и `__EXIT__:N`.

### Тестовая инфраструктура для отладки

Production-CLI (`npm start`) и Playwright Test (`npx playwright test`) делят один и тот же `D365Client` — нулевое дублирование селекторов. Спеки используются для пошаговой отладки в VS Code Test Explorer.

- [playwright.config.ts](playwright.config.ts) — `msedge` channel, `headless: false`, `workers: 1`, `trace: 'on'`.
- [tests/fixtures.ts](tests/fixtures.ts) — worker-scoped persistent Edge context на тот же `./edge-profile` (SSO куки шарятся с CLI). **Нельзя запускать одновременно `npm start` и `npx playwright test` — один лок профиля.**
- [tests/d365-flow.spec.ts](tests/d365-flow.spec.ts) — discrete `test()` блоки: login → navigate → openNewTimeEntry → fillEntryWithTaskLookup. Каждый шаг = отдельная строка в Test Explorer, можно ставить брейкпоинты внутри `D365Client` методов.
- **Pick Locator** в VS Code Playwright Test расширении работает на открытом Edge — клик по элементу в браузере = готовый локатор в IDE.

## Browser modes (`BROWSER_MODE` env)

| Mode | Когда использовать |
|---|---|
| `chrome-profile` | Персистентный Edge с профилем в `./browser-profile/`. Куки сохраняются между прогонами. **Рекомендованный для регулярного использования.** |
| `cdp` | Подключение к уже-запущенному Chrome через `--remote-debugging-port=9222`. Стартуется кнопкой "Launch Debug Chrome" в GUI. Использует реальный профиль `%LOCALAPPDATA%\Google\Chrome\User Data`. |
| `persistent` (default в `.env.example`) | Bundled Chromium с профилем в `./browser-profile/`. |

## AMC Bridge SSO

D365 в AMC Bridge требует Microsoft SSO extension `ppnbnpeolgkicgegkbkbjmhlideopiji`. Без него — бесконечный редирект на login. GUI кнопка "Sync Chrome Profile" ([gui.ps1:415-425](gui.ps1#L415-L425)) копирует это расширение + куки + Login Data из реального Chrome-профиля в `./browser-profile/`.

При первом запуске: закрыть весь Chrome → нажать "Sync Chrome Profile" → нажать "Run All".

## Conventions

- **Даты для D365**: формат `M/D/YYYY` (см. `toD365Date` в [src/excel.ts:9-11](src/excel.ts#L9-L11)). Без leading zeros.
- **Excel serial → Date**: `(serial - 25569) * 86400 * 1000` (offset для 1900 leap-year bug Lotus 1-2-3 совместимости).
- **Часы для D365 Payable Duration**: открываем dropdown и кликаем точную опцию по label'у. Маппинг: `<1h → "N minutes"` (15/30/45), `=1h → "1 hour"`, `>1h → "N.N hours"` или `"N hours"` (1.5, 2, 2.25, ... до 8). См. `hoursToD365DurationOption` в [src/d365.ts](src/d365.ts).
- **CLI флаги**: `--stage1-only` оставлен для обратной совместимости с GUI (alias дефолтного поведения, поскольку Stage 2 не реализован). `--stage2-only` отвергается с ошибкой.
- **Selectors**: множественные локаторы через запятую (`'A, B, C'` = first match) — Playwright OR. При сломе одного fallback'а — добавлять новый, не удалять старые.

## Gotchas

- `__AWAIT_LOGIN__` sentinel в stdout (см. [src/d365.ts:123](src/d365.ts#L123)) — сигнал GUI показать кнопку "I'm logged in - Continue". GUI пишет `CONTINUE` в stdin → код [src/d365.ts:38-48](src/d365.ts#L38-L48) расценивает это как manual ack и продолжает. Параллельно идёт авто-detect навбара (timeout 5 минут).
- В `chrome-profile` mode `userDataDir` принудительно резолвится через `path.resolve` — относительные пути в `.env` работают.
- При `cdp` mode код пытается переиспользовать существующую вкладку с D365-хостом; если её нет — берёт любую "обычную" вкладку (не chrome:// и не extension://) или открывает новую. Это намеренно — позволяет запустить тулзу не пугая открытые в Chrome табы.
- Excel парсер чувствителен к структуре: тaskName *должен* быть в col C, дни Mon-Fri в col D-H. Sheet — первый.
- `weekFilter` сравнивается с `weekStart + 1 day` (строка [src/excel.ts:62-64](src/excel.ts#L62-L64)) с tolerance ±1 день — потому что в Excel `weekStart` это воскресенье, а пользователь передаёт понедельник.

## Files / dirs

- `data/` — Excel-файлы (gitignored).
- `browser-profile/` — Edge/Chromium профиль (gitignored).
- `edge-profile/` — отдельный Edge-профиль (зачем — не задокументировано, проверить).
- `.env` / `.env.example` — `D365_URL`, `BROWSER_MODE`, `CDP_URL`, `USER_DATA_DIR`.

## Development

- TS strict mode, target ES2022, commonjs module ([tsconfig.json](tsconfig.json)).
- Линтер/форматтер пока не настроены.
- Тесты: `npm test` (vitest). Покрыт `parseTimesheet`, `uniqueTasks`, `firstEntriesPerTask`, `remainingEntries` — 12 тестов.
