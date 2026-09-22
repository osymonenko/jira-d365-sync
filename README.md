# Jira → Timesheet → D365

Заполнение недельного табеля почти без ручной работы: инструмент вытягивает активность из Jira
и набор регулярных («стандартных») задач в Excel-табель, а затем создаёт из этого табеля
Time Entries в Microsoft Dynamics 365.

Всё делается из одного окна (`start.bat`), но каждый слой можно запускать и отдельно из консоли.

| Слой | Технология | Файл | Что делает |
|---|---|---|---|
| GUI | PowerShell + WinForms | [jira-sync.ps1](jira-sync.ps1) | Настройки, выбор недель, кнопки операций, цветной лог |
| Jira → Excel | Python | [scripts/jira-sync.py](scripts/jira-sync.py) | 6 JQL-запросов на неделю + регулярные задачи → строки табеля, пересчёт `check sum` |
| Excel → D365 | Node + TypeScript + Playwright | [src/](src/) | Создаёт Time Entries в D365, при необходимости — саму задачу (Imported Project Task) |

---

## Что нужно установить

| Требование | Зачем | Проверка |
|---|---|---|
| Windows + PowerShell 5.1 | GUI на WinForms | входит в Windows |
| **Python 3.10+** и пакет `openpyxl` | слой Jira → Excel | `python --version`, `pip install openpyxl` |
| **Node.js 18+** (LTS) и `npm install` в папке проекта | слой Excel → D365 | `node --version` |
| **Microsoft Edge** | Playwright работает через канал `msedge` (требование компании — только Edge) | установлен в системе |
| Jira API-токен | доступ к Jira REST API | [создать токен](https://id.atlassian.com/manage-profile/security/api-tokens) |
| Доступ к D365 (Microsoft SSO) | создание Time Entries | вход в браузере |
| Excel-файл табеля | источник и приёмник данных | см. [Формат Excel](#формат-excel) |

Если Python или Node не установлены (или не в `PATH`), GUI не падает с невнятной ошибкой,
а показывает popup с точной инструкцией, что поставить и какую команду выполнить.

## Установка

```powershell
git clone https://github.com/osymonenko/jira-d365-sync.git
cd jira-d365-sync
npm install                # зависимости Node (playwright, ts-node, xlsx, commander)
pip install openpyxl       # зависимость Python
copy .env.example .env     # базовый конфиг; остальное заполняется через GUI
```

`npm install` обязателен: `node_modules/` не входит в репозиторий и не попадает в ZIP-скачивание.

## Запуск

| Файл | Поведение |
|---|---|
| `start.bat` | обычный запуск GUI (окно PowerShell скрыто) |
| `jira-sync.bat` | то же, но при падении показывает код ошибки и ждёт нажатия клавиши — для диагностики |

## Настройка (кнопка **Settings**)

Все значения сохраняются в `.env` (в репозиторий не попадает) и в `config/*.json`.

**Connection** — D365 URL, Jira URL, API Token (`Show token` показывает его), Project key, Email,
Account ID. Кнопка **Test Connection** проверяет авторизацию и печатает список доступных проектов.
Account ID можно оставить пустым — тогда используется владелец токена (`/rest/api/3/myself`).

**Standard tasks** — таблица регулярных задач: название, часы по дням Mon–Fri, частота
(`weekly` / `sprint-end` / `placeholder`). Здесь же **Sprint length (weeks)** и **Sprint end** —
любая пятница, которая точно была концом спринта; циклы отсчитываются от неё в обе стороны.
Без заполненного Sprint end задачи с частотой `sprint-end` пропускаются.

**Jira queries** — 6 JQL-шаблонов и правила формирования названий строк в табеле.
Поддерживаются подстановки `{project}`, `{account_id}`, `{ws}`, `{we}` (начало/конец недели).
Режимы названий: `count` (одна строка со счётчиком), `priority` (одна строка с разбивкой P1/P2/P3),
`per_issue` (строка на каждую задачу; доступны `{key}`, `{summary}`, `{number}`).
Пустой/битый конфиг не ломает прогон — используются встроенные значения по умолчанию.

**Other** — путь к Excel-файлу, ссылка на QAE Reporting Rules и чекбокс
«Set 'Copy to Billable Duration' to Yes» для создаваемых в D365 записей.

### Ключи `.env`, которых нет в GUI

| Ключ | Значение |
|---|---|
| `BROWSER_MODE` | `chrome-profile` — Edge с профилем в `USER_DATA_DIR` (**рекомендуется**); `cdp` — подключение к уже запущенному Chrome на `--remote-debugging-port=9222`; `persistent` — встроенный Chromium (нужен `npx playwright install chromium`) |
| `USER_DATA_DIR` | папка профиля браузера, например `./edge-profile`. Куки SSO сохраняются между прогонами |
| `CDP_URL` | адрес отладочного порта для режима `cdp` (по умолчанию `http://localhost:9222`) |

## Порядок работы

1. **Read File** — читает Excel и заполняет правую панель неделями (в логе видно, какие задачи уже есть в каждой неделе).
2. Выбрать, что обрабатывать: галочки нужных недель либо **Month** в выпадающем списке.
   Выбранный месяц имеет приоритет: недели на его границах обрезаются по месяцу.
3. **Standard ↓** — вписывает регулярные задачи (митинги, отчёты, sprint review) и QA-заглушки
   в выбранные недели, снизу дописывает строку `check sum` с формулами `=SUM()` по дням.
4. **Jira ↓** — прогоняет 6 JQL-запросов на каждую выбранную неделю и вставляет полученные строки
   перед `check sum`.
5. **Balance 8h ⚖** — добивает пустые строки задач часами так, чтобы в `check sum` под каждым днём
   вышло 8. Спрашивает, пересчитать раскладку с нуля или только дозаполнить пустые строки.
   Регулярные задачи и уже проставленные вами часы не трогает.
6. **Submit to D365 →** — открывает Edge, идёт в D365 и создаёт Time Entries сразу по всем дням
   каждой задачи. Если задачи в D365 нет — создаёт её. На каждую ячейку табеля появляются две
   записи: Payable и Billable. Всё создаётся в статусе **Draft** — проверьте и нажмите Submit в D365.

Дополнительные кнопки: **Test** — проверка связи с Jira, **Open File** — открыть табель в Excel,
**⛔** — убить текущий процесс, **Copy log** — скопировать лог целиком.

Полезное:

- Повторный прогон безопасен: строки с уже существующими названиями пропускаются (регистр не важен).
- Excel-файл должен быть **закрыт** — иначе запись не пройдёт (в логе будет
  `[ERROR] Cannot save — close the file in Excel first`).
- При первом заходе в D365 нужно вручную залогиниться в открывшемся окне Edge; автоматизация
  продолжится сама, как только увидит навбар D365 (ожидание до 5 минут).
- **Submit to D365** обрабатывает одну неделю за прогон. Если отмечено несколько — берётся первая,
  GUI об этом предупреждает; для остальных нужно перезапустить.
- Статус `No entries for the selected week` означает, что в выбранной неделе нет часов, — это не успех и не ошибка.

## Формат Excel

Читается **первый лист**. Структура, на которую опираются оба слоя:

```
 A            B            C                          D    E    F    G    H
 ─────────────────────────────────────────────────────────────────────────────
 2026-06-14   2026-06-20   Internal Daily meeting     0.5  0.5  0.5  0.5  0.5   ← блок недели
                           Bug verification P1-2, P2-1      1
                           check sum                  0.5  1.5  0.5  0.5  0.5   ← закрывает блок
 2026-06-21   2026-06-27   ...                                                  ← следующий блок
```

- Строка со **датой в колонке A** открывает блок недели (в Excel это воскресенье), колонка B — конец недели.
- Строки задач: **колонка C** — название, **колонки D–H** — часы Mon–Fri (`0.5`, `1`, `2` …).
- Строка `check sum` в колонке C закрывает блок; новые строки вставляются перед ней.
- В интерфейсе и в CLI недели указываются **понедельником** (`YYYY-MM-DD`) — сопоставление
  с воскресеньем из Excel идёт с допуском ±1 день.

## Работа из консоли (без GUI)

```powershell
# Jira / Excel (Python)
python scripts/jira-sync.py --command test --file data\timesheet.xlsx
python scripts/jira-sync.py --command read-weeks --file data\timesheet.xlsx      # список недель в JSON
python scripts/jira-sync.py --command sync --file data\timesheet.xlsx --weeks 2026-06-15 2026-06-22
python scripts/jira-sync.py --command sync --file data\timesheet.xlsx --month 2026-06 --dry-run
python scripts/jira-sync.py --command fill-standard --file data\timesheet.xlsx --month 2026-06
python scripts/jira-sync.py --command balance --file data\timesheet.xlsx --month 2026-06          # добить до 8 ч/день
python scripts/jira-sync.py --command balance --file data\timesheet.xlsx --month 2026-06 --reset  # пересчитать с нуля

# D365 (Node)
.\submit.bat --file data\timesheet.xlsx --list-tasks        # только разбор Excel
.\submit.bat --file data\timesheet.xlsx --preflight-only    # проверки без браузера
.\submit.bat --file data\timesheet.xlsx --week 2026-06-15   # прогон одной недели
.\submit.bat --file data\timesheet.xlsx                     # весь файл
```

Запускайте именно `submit.bat`, а не `npm start -- --file ...`: в PowerShell npm теряет имена флагов по дороге, и скрипт ругается `required option '-f, --file <path>' not specified`, хотя команда набрана правильно.

`--dry-run` (только Python) показывает, что было бы вставлено, ничего не записывая в файл.

## Логи и отладка

- Каждый прогон D365 пишет `logs/<runId>/run.log`, Playwright-trace и скриншоты ошибок.
  Trace открывается через `npx playwright show-trace logs/<runId>/trace.zip`.
- В GUI лог раскрашен по префиксам (`[OK]`, `[WARN]`, `[ERR]`, `[SKIP]`, `[INFO]`) — **Copy log**
  копирует его целиком.
- Пошаговая отладка UI-сценариев: `npx playwright test --ui` (спеки в [tests/](tests/) используют
  тот же `D365Client`, что и продакшн-CLI).
- **Нельзя** одновременно запускать заливку (`.\submit.bat`) и `npx playwright test` — они делят один профиль браузера.

## Тесты

```powershell
npm test                                  # vitest: разбор Excel (src/excel.ts)
npm run typecheck                         # tsc --noEmit
cd scripts; python test_standard_tasks.py # автономные Python-скрипты (scripts/test_*.py)
npx playwright test                       # сценарии D365 в реальном Edge
```

## Ограничения

- Записи создаются в статусе **Draft** — проверить и нажать **Submit** в D365 нужно самому.
- Повторная заливка не исправляет уже созданные записи: они пропускаются как существующие.
  Если часы в табеле поменялись, удалите записи за эти дни в D365 и залейте неделю заново.
- Заливка идёт через Web API D365. Если тенант его закроет, остаётся запасной путь
  `--mode ui` (Quick Create), но он заполняет только первый день каждой задачи.
- В PowerShell запускайте `.\submit.bat`, а не `npm start -- --file ...`: npm теряет имена
  флагов, и скрипт ругается `required option '-f, --file <path>' not specified`.
- Для SSO в AMC Bridge профиль браузера должен содержать рабочую сессию Microsoft; при пустом
  профиле возможен цикл редиректов на страницу логина — залогиньтесь в открытом окне вручную.
- `data/`, `config/`, `.env`, профили браузера и `logs/` в репозиторий не попадают: там реальные
  данные учёта времени и секреты.

## Структура репозитория

```
jira-sync.ps1        GUI (WinForms)
start.bat            запуск GUI
jira-sync.bat        запуск GUI с диагностикой ошибок
scripts/             Python: jira-sync.py, standard_tasks.py, month_filter.py + тесты
src/                 TypeScript: index.ts (CLI), d365.ts (Playwright), excel.ts, preflight.ts
tests/               Playwright-спеки для отладки сценариев D365
config/              jql_queries.json, standard_tasks.json (создаются из Settings)
docs/superpowers/    спеки и планы доработок
data/                Excel-табели
logs/                логи прогонов, trace, скриншоты
```

Технические детали для доработки — в [CLAUDE.md](CLAUDE.md).
