# Jira-sync Month Filter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Позволить пользователю указывать месяц (`YYYY-MM`), чтобы Jira-sync обрезал окно JQL для переходных недель до границ месяца — без двойного учёта уже отчитанных дней предыдущего месяца.

**Architecture:** Чистая логика обрезки выносится в отдельный модуль `scripts/month_filter.py` (юнит-тестируемый без сети). `scripts/jira-sync.py` получает флаг `--month`, фильтрует и обрезает недели через этот модуль. GUI `jira-sync.ps1` получает выпадающий список "Month:", который передаёт `--month` в python.

**Tech Stack:** Python 3.13 (stdlib `datetime`, `argparse`), openpyxl 3.1.5 (уже установлен), PowerShell WinForms.

## Global Constraints

- **Проект НЕ под git** (`Is a git repository: false`) → шаги «commit» заменены на «сохранить файл + прогнать проверку». Не выполнять `git` команды.
- **Область — только Jira-sync.** Не трогать `src/` (TS / D365 submit).
- **Формат месяца:** ровно `YYYY-MM` (напр. `2026-07`).
- **Запуск python:** GUI зовёт `python scripts\jira-sync.py ...` с WorkingDirectory = корень проекта; `sys.path[0]` = `scripts/`, поэтому `from month_filter import ...` работает.
- **Кодировка python-файлов:** UTF-8 (в проекте уже так).
- **Ключ вставки строк в Excel остаётся реальный `week_start`** — обрезается только окно JQL, не позиция строк.

---

### Task 1: Чистый модуль обрезки недели по месяцу

**Files:**
- Create: `scripts/month_filter.py`
- Test: `scripts/test_month_filter.py`

**Interfaces:**
- Produces:
  - `month_bounds(month: str) -> tuple[date, date]` — первый и последний день месяца; кидает `ValueError` при формате ≠ `YYYY-MM`.
  - `clamp_week_to_month(week_start: date, week_end: date, month: str) -> tuple[date, date] | None` — обрезанное окно `[max(week_start, month_first), min(week_end, month_last)]`, либо `None` если неделя не пересекается с месяцем.

- [ ] **Step 1: Написать падающий тест**

Create `scripts/test_month_filter.py`:

```python
from datetime import date
from month_filter import clamp_week_to_month, month_bounds


def check(name, got, exp):
    assert got == exp, f"{name}: got {got}, expected {exp}"
    print(f"  OK {name}")


# month_bounds
check("bounds July", month_bounds("2026-07"), (date(2026, 7, 1), date(2026, 7, 31)))
check("bounds Feb leap", month_bounds("2024-02"), (date(2024, 2, 1), date(2024, 2, 29)))

# whole week inside the month -> unchanged
check("inside", clamp_week_to_month(date(2026, 7, 6), date(2026, 7, 12), "2026-07"),
      (date(2026, 7, 6), date(2026, 7, 12)))

# boundary week at start of month (Sun 29 Jun - 5 Jul) -> 1-5 Jul
check("start boundary", clamp_week_to_month(date(2026, 6, 29), date(2026, 7, 5), "2026-07"),
      (date(2026, 7, 1), date(2026, 7, 5)))

# boundary week at end of month (27 Jul - 2 Aug) -> 27-31 Jul
check("end boundary", clamp_week_to_month(date(2026, 7, 27), date(2026, 8, 2), "2026-07"),
      (date(2026, 7, 27), date(2026, 7, 31)))

# week fully outside the month -> None
check("outside", clamp_week_to_month(date(2026, 6, 1), date(2026, 6, 7), "2026-07"), None)

# invalid month string -> ValueError
try:
    month_bounds("July")
    raise SystemExit("FAIL: expected ValueError for 'July'")
except ValueError:
    print("  OK invalid month raises")

print("ALL PASS")
```

- [ ] **Step 2: Запустить тест — убедиться, что падает**

Run: `python scripts/test_month_filter.py`
Expected: FAIL — `ModuleNotFoundError: No module named 'month_filter'`

- [ ] **Step 3: Написать модуль**

Create `scripts/month_filter.py`:

```python
"""Pure date-window helpers for month-scoped Jira sync. No I/O, no network."""

from datetime import date, timedelta


def month_bounds(month: str) -> tuple[date, date]:
    """First and last calendar day of a YYYY-MM month. Raises ValueError otherwise."""
    try:
        year_s, mon_s = month.split("-")
        year, mon = int(year_s), int(mon_s)
        first = date(year, mon, 1)
    except (ValueError, AttributeError):
        raise ValueError(f"--month must be YYYY-MM, got: {month!r}")
    next_month = date(year + 1, 1, 1) if mon == 12 else date(year, mon + 1, 1)
    return first, next_month - timedelta(days=1)


def clamp_week_to_month(week_start: date, week_end: date, month: str):
    """Intersect [week_start, week_end] with the month.

    Returns the clamped (start, end) tuple, or None if the week does not
    overlap the month at all.
    """
    first, last = month_bounds(month)
    if week_end < first or week_start > last:
        return None
    return (max(week_start, first), min(week_end, last))
```

- [ ] **Step 4: Запустить тест — убедиться, что проходит**

Run: `python scripts/test_month_filter.py`
Expected: PASS — печатает список `OK ...` и `ALL PASS`, exit code 0.

- [ ] **Step 5: Сохранить (git недоступен — коммит не делаем)**

Файлы `scripts/month_filter.py` и `scripts/test_month_filter.py` сохранены.

---

### Task 2: Флаг `--month` в jira-sync.py

**Files:**
- Modify: `scripts/jira-sync.py` — импорт (после строки 17), `cmd_sync` (~351-412), argparse в `main` (~419-425)

**Interfaces:**
- Consumes: `clamp_week_to_month`, `month_bounds` из Task 1.
- Produces: CLI-флаг `--month YYYY-MM` для команды `sync`.

- [ ] **Step 1: Добавить импорт модуля**

В `scripts/jira-sync.py` после блока `from pathlib import Path` (строка 17) добавить:

```python
from month_filter import clamp_week_to_month, month_bounds
```

- [ ] **Step 2: Зарегистрировать аргумент `--month`**

В `main()` (`scripts/jira-sync.py`), рядом с `--weeks` (после строки 423) добавить:

```python
    parser.add_argument("--month", help="Clamp week windows to this month (YYYY-MM)")
```

- [ ] **Step 3: Валидация месяца в начале `cmd_sync`**

В `cmd_sync`, сразу после успешного получения `account_id` (после строки 366 `print(f"[INFO] Account ID: {account_id}", flush=True)`), добавить:

```python
    if args.month:
        try:
            month_bounds(args.month)
        except ValueError as e:
            print(f"[ERROR] {e}", flush=True)
            sys.exit(1)
```

- [ ] **Step 4: Обрезать окно JQL в цикле недель**

Заменить тело цикла `for week in weeks_info:` (строки 391-402) на:

```python
    insertions: dict[str, list[tuple[str, str]]] = {}
    for week in weeks_info:
        ws_key = fmt(week["week_start"])
        we = week["week_end"] or (week["week_start"] + timedelta(days=6))
        q_start, q_end = week["week_start"], we
        if args.month:
            clamped = clamp_week_to_month(week["week_start"], we, args.month)
            if clamped is None:
                print(f"[SKIP] {ws_key}: outside month {args.month}", flush=True)
                continue
            q_start, q_end = clamped
        label = f"{ws_key} - {fmt(we)}"
        if (q_start, q_end) != (week["week_start"], we):
            label += f"  (clamped to {fmt(q_start)} - {fmt(q_end)})"
        print(f"\n[INFO] Week {label}", flush=True)
        try:
            rows = generate_week_rows(base_url, email, token, account_id, project, q_start, q_end)
        except Exception as e:
            print(f"[ERROR] {e}", flush=True)
            continue
        print(f"  -> {len(rows)} row(s) to insert", flush=True)
        if rows:
            insertions[ws_key] = rows
```

(Ключ `insertions[ws_key]` = реальный `week_start`, поэтому строки лягут в правильный блок недели. В `generate_week_rows` уходит обрезанное окно `q_start`/`q_end`.)

- [ ] **Step 5: Проверить, что флаг зарегистрирован**

Run: `python scripts/jira-sync.py --command sync --help`
Expected: в выводе присутствует `--month` с описанием `Clamp week windows to this month (YYYY-MM)`.

- [ ] **Step 6: Проверить валидацию месяца**

Run: `python scripts/jira-sync.py --command sync --file data/timesheet.xlsx --month July --dry-run`
Expected: печатает `[ERROR] --month must be YYYY-MM, got: 'July'` и завершается с кодом 1 (до сетевых запросов валидация не дойдёт — проверка стоит после auth; если auth-креды валидны, увидишь ошибку месяца после строки Account ID).

- [ ] **Step 7: Проверить обрезку на реальном файле (нужна сеть + Jira-креды из .env)**

Run: `python scripts/jira-sync.py --command sync --file data/timesheet.xlsx --month 2026-07 --dry-run`
Expected: недели полностью вне июля печатают `[SKIP] ...: outside month 2026-07`; переходная неделя печатает `(clamped to 2026-07-01 - ...)`; Excel не изменяется (`[DRY RUN]`).

Если нет доступа к сети/Jira — этот шаг пропустить, обрезка уже покрыта юнит-тестами Task 1.

- [ ] **Step 8: Сохранить (git недоступен — коммит не делаем)**

---

### Task 3: Выпадающий список "Month:" в GUI

**Files:**
- Modify: `jira-sync.ps1` — панель недель (`$pnlWeeks`, ~145-214), `Populate-Weeks` (~178-211), `$btnSync.Add_Click` (~569-584)

**Interfaces:**
- Consumes: python-флаг `--month` из Task 2; JSON недель из `read-weeks` (поля `start`, `end` как `YYYY-MM-DD`).
- Produces: элемент `$cmbMonth` (ComboBox) и `$script:monthCodes` (hashtable display→`YYYY-MM`).

- [ ] **Step 1: Добавить label + ComboBox "Month:" в панель недель**

В `jira-sync.ps1`, после блока создания `$lnkNone` (после строки 174, перед `$script:weekCheckboxes = @()`), добавить:

```powershell
$lblMonth = New-Object System.Windows.Forms.Label
$lblMonth.Text = 'Month:'
$lblMonth.Location = New-Object System.Drawing.Point(122,13)
$lblMonth.Size = New-Object System.Drawing.Size(44,18)
$lblMonth.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblMonth.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
[void]$pnlWeeks.Controls.Add($lblMonth)

$cmbMonth = New-Object System.Windows.Forms.ComboBox
$cmbMonth.Location = New-Object System.Drawing.Point(166,10)
$cmbMonth.Size = New-Object System.Drawing.Size(128,22)
$cmbMonth.DropDownStyle = 'DropDownList'
$cmbMonth.Font = New-Object System.Drawing.Font('Segoe UI',9)
[void]$cmbMonth.Items.Add('All')
$cmbMonth.SelectedIndex = 0
[void]$pnlWeeks.Controls.Add($cmbMonth)

$script:monthCodes = @{}
```

- [ ] **Step 2: Сдвинуть чекбоксы недель правее ComboBox**

В функции `Populate-Weeks` заменить строку 186:

```powershell
    $colW = 142; $startX = 122; $cols = [Math]::Floor((900 - $startX) / $colW)
```

на:

```powershell
    $colW = 142; $startX = 306; $cols = [Math]::Floor((900 - $startX) / $colW)
```

(Чекбоксы теперь начинаются после ComboBox; `cols` пересчитывается динамически — при 306 помещается 4 в ряд.)

- [ ] **Step 3: Заполнять список месяцев при чтении файла**

В функции `Populate-Weeks`, сразу после строки 183 (`if (-not $weeksJson -or $weeksJson.Count -eq 0) { return }`), добавить заполнение месяцев:

```powershell
    # Rebuild month dropdown from the months each week touches (start + end).
    $cmbMonth.Items.Clear()
    [void]$cmbMonth.Items.Add('All')
    $script:monthCodes = @{}
    $codes = New-Object 'System.Collections.Generic.SortedSet[string]'
    foreach ($w in $weeksJson) {
        [void]$codes.Add($w.start.Substring(0,7))
        [void]$codes.Add($w.end.Substring(0,7))
    }
    foreach ($code in $codes) {
        $display = [datetime]::ParseExact("$code-01",'yyyy-MM-dd',$null).ToString('MMMM yyyy',[System.Globalization.CultureInfo]::InvariantCulture)
        $script:monthCodes[$display] = $code
        [void]$cmbMonth.Items.Add($display)
    }
    $cmbMonth.SelectedIndex = 0
```

- [ ] **Step 4: Передавать `--month` из кнопки Sync**

Заменить тело `$btnSync.Add_Click` (строки 569-584) на:

```powershell
$btnSync.Add_Click({
    $monthCode = $null
    if ($cmbMonth.SelectedIndex -gt 0) { $monthCode = $script:monthCodes[[string]$cmbMonth.SelectedItem] }

    $selected = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
    if (-not $monthCode -and $script:weekCheckboxes.Count -gt 0 -and $selected.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No weeks selected.', 'Sync')
        return
    }
    $txtLog.Clear()
    Set-Status 'Syncing from Jira...' ([System.Drawing.Color]::DodgerBlue)

    # A chosen month drives week selection itself — the week checkboxes are ignored for that run.
    $extraArgs = ''
    if ($monthCode) {
        $extraArgs = ' --month ' + $monthCode
    } elseif ($selected.Count -gt 0) {
        $extraArgs = ' --weeks ' + ($selected -join ' ')
    }
    Start-PyProc ('scripts\jira-sync.py --command sync --file "' + $txtFile.Text + '"' + $extraArgs) {
        param($code)
        if ($code -eq 0) { Set-Status 'Sync complete' ([System.Drawing.Color]::LimeGreen) }
        else              { Set-Status ('Sync failed (exit '+$code+')') ([System.Drawing.Color]::OrangeRed) }
    }
})
```

- [ ] **Step 5: Ручная проверка GUI**

1. Запусти `jira-sync.bat`.
2. Укажи Excel-файл → нажми **Read File**.
   Expected: список "Month:" заполнен (`All`, `June 2026`, `July 2026`, …); чекбоксы недель видны и сдвинуты вправо, не наезжают на ComboBox.
3. Выбери `July 2026` → нажми **Sync Jira -> Excel**.
   Expected: в лог уходит команда с ` --month 2026-07`; переходная неделя пишет `(clamped to 2026-07-01 - ...)`, недели вне июля — `[SKIP]`.
4. Верни `All` → нажми **Sync**.
   Expected: поведение как раньше — по чекбоксам недель (`--weeks ...`).

- [ ] **Step 6: Сохранить (git недоступен — коммит не делаем)**

---

## Self-Review

**1. Spec coverage:**
- Обрезка окна JQL по месяцу → Task 1 (`clamp_week_to_month`) + Task 2 (Step 4). ✓
- Аргумент `--month YYYY-MM` + валидация → Task 2 (Steps 2, 3). ✓
- Ключ вставки строк = реальный `week_start` → Task 2 (Step 4, комментарий). ✓
- Взаимодействие `--month`/`--weeks` (GUI при месяце не шлёт `--weeks`) → Task 3 (Step 4). ✓
- GUI ComboBox "Month:" (`All` + месяцы из файла, заполнение при Read File) → Task 3 (Steps 1, 3). ✓
- При выбранном месяце чекбоксы недель игнорируются → Task 3 (Step 4). ✓
- Будущие дни отсекаются естественно → доп. кода нет (зафиксировано в спеке). ✓
- Чистая функция + юнит-тесты → Task 1. ✓

**2. Placeholder scan:** плейсхолдеров нет — весь код приведён целиком.

**3. Type consistency:** `month_bounds`/`clamp_week_to_month` определены в Task 1 и используются с теми же сигнатурами в Task 2. `$script:monthCodes` (Task 3 Step 1) заполняется в Step 3 и читается в Step 4 — имена совпадают. `$cmbMonth` создан в Step 1, используется в Steps 3-4. ✓
