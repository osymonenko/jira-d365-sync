# Standard Tasks Generator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Pre-fill `timesheet.xlsx` with recurring standard tasks (with hours) plus QA placeholder rows for future weeks, driven entirely from the GUI.

**Architecture:** A pure-logic Python module (`standard_tasks.py`) computes which task rows a given week needs; `jira-sync.py` gains a shared insertion helper (Jira rows and standard rows both flow through it) and a new `fill-standard` command; the GUI gets a "Standard → Excel" button and a `SPRINT_ANCHOR` setting.

**Tech Stack:** Python 3 + openpyxl (Excel), PowerShell WinForms (GUI). No new dependencies.

## Global Constraints

- Excel day columns: **Mon=4, Tue=5, Wed=6, Thu=7, Fri=8**. Column 3 = task name. `week_start` stored in Excel is the **Sunday**; work days are `week_start + 1 … +5`.
- Only fill **existing** week blocks; never create blocks. Missing/unmatched weeks → `[SKIP]`.
- Dedup standard rows **case-insensitively** by task name against the week's existing names.
- Hours are 15-minute-grid values ≤ 8h (0.5 / 1.0 / 2.0 here).
- Python tests are plain scripts run as `python scripts/<test>.py`, using a `check()`/assert helper and printing `ALL PASS` (see `scripts/test_month_filter.py`). No pytest.
- Task names are written to Excel verbatim from the schedule; D365 lookup is case-insensitive so casing is cosmetic.
- Sprint-end weeks: a week whose Friday `f` satisfies `(f − SPRINT_ANCHOR).days % 14 == 0`.
- QA placeholders added iff `week_friday > today`.

---

### Task 1: `standard_tasks.py` — schedule + per-week row logic

Pure logic, no openpyxl/network. Fully unit-tested.

**Files:**
- Create: `scripts/standard_tasks.py`
- Test: `scripts/test_standard_tasks.py`

**Interfaces:**
- Consumes: `clamp_week_to_month(week_start, week_end, month)` from `month_filter` (returns clamped `(date, date)` or `None`).
- Produces:
  - `rows_for_week(week_start: date, week_end: date, month: str|None, sprint_anchor: date|None, today: date) -> list[dict]` where each dict is `{"name": str, "hours_by_col": dict[int, float]}` (empty `hours_by_col` = QA placeholder).
  - `is_sprint_end_week(week_start: date, sprint_anchor: date) -> bool`
  - Constants `SCHEDULE`, `QA_PLACEHOLDERS`, `DAY_COL`, `DAY_OFFSET`.

- [ ] **Step 1: Write the failing test**

Create `scripts/test_standard_tasks.py`:

```python
from datetime import date
from standard_tasks import rows_for_week, is_sprint_end_week


def check(name, got, exp):
    assert got == exp, f"{name}: got {got}, expected {exp}"
    print(f"  OK {name}")


def by_name(rows):
    return {r["name"]: r["hours_by_col"] for r in rows}


ANCHOR = date(2026, 7, 17)           # a Friday, sprint end
PAST = date(2026, 7, 10)             # QA trigger: friday(2026-07-17) > 2026-07-10 -> True

# --- weekly tasks present every week (week 2026-07-12 Sun .. 2026-07-18) ---
rows = by_name(rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, ANCHOR, PAST))
check("daily cols", rows["Internal Daily meeting"], {4: 0.5, 5: 0.5, 6: 0.5, 7: 0.5, 8: 0.5})
check("bug triage Tue", rows["Internal bug triage"], {5: 0.5})
check("external Tue+Wed", rows["External customer meeting"], {5: 1.0, 6: 1.0})
check("weekly report Fri", rows["Weekly project report"], {8: 1.0})

# --- sprint-end tasks: present on anchor week ---
check("sprint review present", "Internal sprint review" in rows, True)
check("summary present", "Summary report creation" in rows, True)
check("summary hours", rows["Summary report creation"], {8: 2.0})

# --- sprint-end absent on the off-sprint week (2026-07-19, friday 2026-07-24) ---
off = by_name(rows_for_week(date(2026, 7, 19), date(2026, 7, 25), None, ANCHOR, PAST))
check("sprint review absent off-week", "Internal sprint review" in off, False)
check("summary absent off-week", "Summary report creation" in off, False)

# --- sprint-end present again two weeks later (2026-07-26, friday 2026-07-31) ---
nxt = by_name(rows_for_week(date(2026, 7, 26), date(2026, 8, 1), None, ANCHOR, PAST))
check("sprint review +14", "Internal sprint review" in nxt, True)

# --- no anchor -> sprint-end tasks skipped, weekly ones stay ---
noanchor = by_name(rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, None, PAST))
check("no-anchor skips sprint", "Internal sprint review" in noanchor, False)
check("no-anchor keeps daily", "Internal Daily meeting" in noanchor, True)

# --- QA placeholders appear when week friday > today ---
qa = rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, ANCHOR, PAST)
qa_names = {r["name"] for r in qa if not r["hours_by_col"]}
check("QA present when future", qa_names,
      {"Bug verification", "Functional testing", "Automation test maintenance", "Investigation issue"})

# --- QA absent when friday <= today (today = the friday itself) ---
qa_none = rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, ANCHOR, date(2026, 7, 17))
check("QA absent when not future", any(not r["hours_by_col"] for r in qa_none), False)

# --- month clamp drops out-of-month work days (week 2026-06-28 .. 07-04, July) ---
clamped = by_name(rows_for_week(date(2026, 6, 28), date(2026, 7, 4), "2026-07", ANCHOR, PAST))
check("daily clamped to July", clamped["Internal Daily meeting"], {6: 0.5, 7: 0.5, 8: 0.5})

# --- week fully outside month -> no rows ---
outside = rows_for_week(date(2026, 6, 1), date(2026, 6, 7), "2026-07", ANCHOR, PAST)
check("outside month empty", outside, [])

# --- is_sprint_end_week direct ---
check("anchor is sprint end", is_sprint_end_week(date(2026, 7, 12), ANCHOR), True)
check("off week not sprint end", is_sprint_end_week(date(2026, 7, 19), ANCHOR), False)

print("ALL PASS")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python scripts/test_standard_tasks.py`
Expected: FAIL — `ModuleNotFoundError: No module named 'standard_tasks'`

- [ ] **Step 3: Write minimal implementation**

Create `scripts/standard_tasks.py`:

```python
"""Recurring standard-task schedule and per-week row generation.

Pure logic: no openpyxl, no network, no system clock (today is passed in).
"""

from datetime import date, timedelta

from month_filter import clamp_week_to_month

# Excel day columns (Mon=D .. Fri=H). The first hours column is Monday.
DAY_COL = {"Mon": 4, "Tue": 5, "Wed": 6, "Thu": 7, "Fri": 8}
# Offset in days from the Sunday week_start stored in Excel.
DAY_OFFSET = {"Mon": 1, "Tue": 2, "Wed": 3, "Thu": 4, "Fri": 5}

SCHEDULE = [
    {"name": "Internal Daily meeting",    "hours": 0.5, "days": ["Mon", "Tue", "Wed", "Thu", "Fri"], "freq": "weekly"},
    {"name": "Internal bug triage",       "hours": 0.5, "days": ["Tue"],          "freq": "weekly"},
    {"name": "External customer meeting", "hours": 1.0, "days": ["Tue", "Wed"],   "freq": "weekly"},
    {"name": "Weekly project report",     "hours": 1.0, "days": ["Fri"],          "freq": "weekly"},
    {"name": "Internal sprint review",    "hours": 0.5, "days": ["Fri"],          "freq": "sprint-end"},
    {"name": "Summary report creation",   "hours": 2.0, "days": ["Fri"],          "freq": "sprint-end"},
]

QA_PLACEHOLDERS = [
    "Bug verification",
    "Functional testing",
    "Automation test maintenance",
    "Investigation issue",
]


def _friday(week_start: date) -> date:
    return week_start + timedelta(days=DAY_OFFSET["Fri"])


def is_sprint_end_week(week_start: date, sprint_anchor: date) -> bool:
    """A week is a sprint end iff its Friday is a whole number of 2-week cycles
    away from the anchor Friday."""
    return (_friday(week_start) - sprint_anchor).days % 14 == 0


def rows_for_week(week_start, week_end, month, sprint_anchor, today):
    """Standard task rows for one week block.

    Returns list of {"name": str, "hours_by_col": {col: hours}}.
    QA placeholders have an empty hours_by_col. Returns [] if the week is
    entirely outside the selected month.
    """
    if month:
        clamped = clamp_week_to_month(week_start, week_end, month)
        if clamped is None:
            return []
        win_start, win_end = clamped
    else:
        win_start, win_end = week_start, week_end

    rows = []
    for task in SCHEDULE:
        if task["freq"] == "sprint-end":
            if sprint_anchor is None or not is_sprint_end_week(week_start, sprint_anchor):
                continue
        hours_by_col = {}
        for day in task["days"]:
            d = week_start + timedelta(days=DAY_OFFSET[day])
            if win_start <= d <= win_end:
                hours_by_col[DAY_COL[day]] = task["hours"]
        if hours_by_col:
            rows.append({"name": task["name"], "hours_by_col": hours_by_col})

    # QA placeholders only when the week still has a future work day.
    if _friday(week_start) > today:
        for name in QA_PLACEHOLDERS:
            rows.append({"name": name, "hours_by_col": {}})

    return rows
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python scripts/test_standard_tasks.py`
Expected: PASS — every `OK` line then `ALL PASS`

- [ ] **Step 5: Commit**

```bash
git add scripts/standard_tasks.py scripts/test_standard_tasks.py
git commit -m "feat: standard-task schedule + per-week row logic"
```

---

### Task 2: `jira-sync.py` — shared insertion, standard writer, `fill-standard` command

Refactor the Jira insertion loop into a writer-callback helper (guarded by the existing `test_excel_insert.py`), add a standard-row writer that fills hours, and wire the new command.

**Files:**
- Modify: `scripts/jira-sync.py` (`update_excel` refactor; add `_apply_insertions`, `insert_standard_rows`, `cmd_fill_standard`; extend argparse + `main`; import `rows_for_week`)
- Test: `scripts/test_standard_insert.py` (new)

**Interfaces:**
- Consumes: `find_weeks`, `_insert_rows_preserving_hyperlinks`, `fmt`, `load_env`, `month_bounds` (existing in `jira-sync.py`); `rows_for_week` (Task 1).
- Produces:
  - `_apply_insertions(ws, insertions, write_row) -> int` — bottom-up insert with case-insensitive dedup; `write_row(row_idx, payload)` fills cells for one row; each payload has a `"name"` key.
  - `insert_standard_rows(file_path, insertions) -> int` — `insertions[key]` = list of `{"name", "hours_by_col"}`.
  - `cmd_fill_standard(args)` — command handler.

- [ ] **Step 1: Write the failing test**

Create `scripts/test_standard_insert.py`:

```python
"""Network-free test for standard-row insertion (hours land in day columns)."""

import importlib.util
import pathlib
import tempfile
from datetime import datetime

import openpyxl

_p = pathlib.Path(__file__).parent / "jira-sync.py"
_spec = importlib.util.spec_from_file_location("jira_sync", _p)
jira_sync = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(jira_sync)


def build_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(1, 1).value = "weekstart"
    ws.cell(1, 3).value = "Task"
    ws.cell(2, 1).value = datetime(2026, 7, 12)   # Sunday
    ws.cell(2, 2).value = datetime(2026, 7, 18)
    ws.cell(2, 3).value = "Internal Daily meeting"  # pre-existing (dedup, diff case)
    ws.cell(3, 3).value = "check sum"
    wb.save(path)


def main():
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    path = tmp.name
    build_workbook(path)

    insertions = {
        "2026-07-12": [
            {"name": "internal daily meeting", "hours_by_col": {4: 0.5}},  # dup (case-insensitive) -> skipped
            {"name": "Weekly project report", "hours_by_col": {8: 1.0}},
            {"name": "Bug verification", "hours_by_col": {}},              # QA placeholder, no hours
        ]
    }
    total = jira_sync.insert_standard_rows(path, insertions)
    assert total == 2, f"expected 2 inserted (dup skipped), got {total}"
    print("  OK dedup skipped the case-different duplicate")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    # Collect name -> {col: value} for column C rows and their day columns.
    found = {}
    for r in range(1, ws.max_row + 1):
        name = ws.cell(r, 3).value
        if not name:
            continue
        found[name] = {c: ws.cell(r, c).value for c in (4, 5, 6, 7, 8) if ws.cell(r, c).value is not None}
    wb.close()

    assert found.get("Weekly project report") == {8: 1}, found.get("Weekly project report")
    print("  OK weekly report hours landed in Friday column (8)")
    assert "Bug verification" in found and found["Bug verification"] == {}, found.get("Bug verification")
    print("  OK QA placeholder inserted with no hours")

    print("ALL PASS")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python scripts/test_standard_insert.py`
Expected: FAIL — `AttributeError: module 'jira_sync' has no attribute 'insert_standard_rows'`

- [ ] **Step 3a: Refactor `update_excel` onto a shared helper**

In `scripts/jira-sync.py`, replace the whole `update_excel` function (currently lines ~307-343) with:

```python
def _apply_insertions(ws, insertions: dict, write_row) -> int:
    """Insert rows bottom-up into existing week blocks.

    insertions[key] is a list of payload dicts (each with a "name"). Rows whose
    name already exists in the block (case-insensitive) are skipped. For each
    surviving row, write_row(row_idx, payload) fills its cells.
    """
    weeks = find_weeks(ws)
    total = 0
    for week in reversed(weeks):
        key = fmt(week["week_start"])
        if key not in insertions:
            continue
        existing_lower = {n.lower() for n in week["existing_names"]}
        new_rows = [p for p in insertions[key] if p["name"].lower() not in existing_lower]
        if not new_rows:
            print(f"[SKIP] {key}: all rows already present", flush=True)
            continue
        check_row = week["check_sum_row"]
        anchor = check_row if check_row is not None else week["end_row"]
        _insert_rows_preserving_hyperlinks(ws, anchor, len(new_rows))
        for i, payload in enumerate(new_rows):
            write_row(anchor + i, payload)
        total += len(new_rows)
        note = "" if check_row is not None else " (no check sum row; appended to block end)"
        print(f"[OK]   {key}: {len(new_rows)} rows inserted{note}", flush=True)
    return total


def update_excel(file_path: str, insertions: dict) -> int:
    wb = openpyxl.load_workbook(file_path)
    ws = wb.worksheets[0]
    link_font = openpyxl.styles.Font(color="0563C1", underline="single")
    payloads = {k: [{"name": n, "url": u} for (n, u) in rows] for k, rows in insertions.items()}

    def write_row(row_idx, p):
        cell = ws.cell(row=row_idx, column=3)
        cell.value = p["name"]
        cell.hyperlink = p["url"]
        cell.font = link_font

    total = _apply_insertions(ws, payloads, write_row)
    try:
        wb.save(file_path)
    except PermissionError:
        print(f"[ERROR] Cannot save — close the file in Excel first: {file_path}", flush=True)
        return 0
    return total


def insert_standard_rows(file_path: str, insertions: dict) -> int:
    """insertions[key] = list of {"name": str, "hours_by_col": {col: hours}}."""
    wb = openpyxl.load_workbook(file_path)
    ws = wb.worksheets[0]

    def write_row(row_idx, p):
        ws.cell(row=row_idx, column=3).value = p["name"]
        for col, hours in p["hours_by_col"].items():
            ws.cell(row=row_idx, column=col).value = int(hours) if float(hours).is_integer() else hours

    total = _apply_insertions(ws, insertions, write_row)
    try:
        wb.save(file_path)
    except PermissionError:
        print(f"[ERROR] Cannot save — close the file in Excel first: {file_path}", flush=True)
        return 0
    return total
```

- [ ] **Step 3b: Add the import**

At the top of `scripts/jira-sync.py`, next to `from month_filter import clamp_week_to_month, month_bounds` (line ~19), add:

```python
from standard_tasks import rows_for_week
```

- [ ] **Step 3c: Add the command handler**

Add this function after `cmd_sync` (before the `# Main` section, ~line 470):

```python
def cmd_fill_standard(args):
    root = Path(args.file).resolve().parent.parent
    env = load_env(root)

    if args.month:
        try:
            month_bounds(args.month)
        except ValueError as e:
            print(f"[ERROR] {e}", flush=True)
            sys.exit(1)

    sprint_anchor = None
    raw_anchor = env.get("SPRINT_ANCHOR", "").strip()
    if raw_anchor:
        try:
            sprint_anchor = datetime.strptime(raw_anchor, "%Y-%m-%d").date()
            if sprint_anchor.weekday() != 4:  # 0=Mon .. 4=Fri
                snapped = sprint_anchor + timedelta(days=(4 - sprint_anchor.weekday()))
                print(f"[WARN] SPRINT_ANCHOR {raw_anchor} is not a Friday; using {fmt(snapped)}", flush=True)
                sprint_anchor = snapped
        except ValueError:
            print(f"[WARN] SPRINT_ANCHOR {raw_anchor!r} invalid (want YYYY-MM-DD); sprint-end tasks skipped", flush=True)
            sprint_anchor = None
    else:
        print("[WARN] SPRINT_ANCHOR not set; sprint-end tasks (sprint review, summary report) skipped", flush=True)

    wb_read = openpyxl.load_workbook(args.file, data_only=True)
    weeks_info = find_weeks(wb_read.worksheets[0])
    wb_read.close()
    if not weeks_info:
        print("[ERROR] No week blocks found in Excel", flush=True)
        sys.exit(1)

    if args.weeks:
        selected = set(args.weeks)
        weeks_info = [
            w for w in weeks_info
            if any(abs(((w["week_start"] + timedelta(days=1)) - datetime.strptime(m, "%Y-%m-%d").date()).days) <= 1
                   for m in selected)
        ]
        if not weeks_info:
            print("[ERROR] None of the specified weeks found in Excel", flush=True)
            sys.exit(1)

    today = date.today()
    insertions = {}
    for week in weeks_info:
        ws_key = fmt(week["week_start"])
        we = week["week_end"] or (week["week_start"] + timedelta(days=6))
        rows = rows_for_week(week["week_start"], we, args.month, sprint_anchor, today)
        if not rows:
            print(f"[SKIP] {ws_key}: no standard rows for this week", flush=True)
            continue
        insertions[ws_key] = rows
        names = ", ".join(r["name"] for r in rows)
        print(f"[INFO] {ws_key}: {len(rows)} row(s) -> {names}", flush=True)

    if insertions:
        total = insert_standard_rows(args.file, insertions)
        print(f"\n[DONE] {total} rows inserted.", flush=True)
    else:
        print("\n[DONE] Nothing to insert.", flush=True)
```

- [ ] **Step 3d: Register the command in argparse + `main`**

Change the `--command` choices line (~478) from:

```python
    parser.add_argument("--command", choices=["test", "read-weeks", "sync"], default="sync")
```

to:

```python
    parser.add_argument("--command", choices=["test", "read-weeks", "sync", "fill-standard"], default="sync")
```

Then in `main`, change the dispatch tail (~485-493) from:

```python
    if args.command == "test":
        cmd_test(args)
    elif args.command == "read-weeks":
        cmd_read_weeks(args)
    else:
        if not args.file:
            print("[ERROR] --file required for sync", flush=True)
            sys.exit(1)
        cmd_sync(args)
```

to:

```python
    if args.command == "test":
        cmd_test(args)
    elif args.command == "read-weeks":
        cmd_read_weeks(args)
    elif args.command == "fill-standard":
        if not args.file:
            print("[ERROR] --file required for fill-standard", flush=True)
            sys.exit(1)
        cmd_fill_standard(args)
    else:
        if not args.file:
            print("[ERROR] --file required for sync", flush=True)
            sys.exit(1)
        cmd_sync(args)
```

- [ ] **Step 4: Run tests to verify they pass (new + Jira regression)**

Run: `python scripts/test_standard_insert.py`
Expected: PASS — `ALL PASS`

Run: `python scripts/test_excel_insert.py`
Expected: PASS — `ALL PASS` (Jira insertion + hyperlink preservation still work after the refactor)

Run: `python scripts/test_standard_tasks.py`
Expected: PASS — `ALL PASS`

- [ ] **Step 5: Commit**

```bash
git add scripts/jira-sync.py scripts/test_standard_insert.py
git commit -m "feat: fill-standard command + shared Excel insertion helper"
```

---

### Task 3: GUI Settings — `SPRINT_ANCHOR` field

Add a "Sprint end (last Fri)" field to the Settings dialog, persisted to `.env`.

**Files:**
- Modify: `jira-sync.ps1` (`Save-JiraEnv` signature + keys; Settings dialog field + layout shift; two `Save-JiraEnv` call sites)
- Modify: `.env.example`

**Interfaces:**
- Consumes: existing `Add-Row`, `Read-EnvFile`, `Save-JiraEnv` patterns.
- Produces: `.env` key `SPRINT_ANCHOR` (YYYY-MM-DD).

- [ ] **Step 1: Extend `Save-JiraEnv`**

In `jira-sync.ps1`, change (currently ~line 19-21):

```powershell
function Save-JiraEnv($url, $email, $token, $project, $accountId, $excelFile) {
    $keys = @('JIRA_URL','JIRA_EMAIL','JIRA_API_TOKEN','JIRA_PROJECT','JIRA_ACCOUNT_ID','EXCEL_FILE')
    $vals = @{ JIRA_URL=$url; JIRA_EMAIL=$email; JIRA_API_TOKEN=$token; JIRA_PROJECT=$project; JIRA_ACCOUNT_ID=$accountId; EXCEL_FILE=$excelFile }
```

to:

```powershell
function Save-JiraEnv($url, $email, $token, $project, $accountId, $excelFile, $sprintAnchor) {
    $keys = @('JIRA_URL','JIRA_EMAIL','JIRA_API_TOKEN','JIRA_PROJECT','JIRA_ACCOUNT_ID','EXCEL_FILE','SPRINT_ANCHOR')
    $vals = @{ JIRA_URL=$url; JIRA_EMAIL=$email; JIRA_API_TOKEN=$token; JIRA_PROJECT=$project; JIRA_ACCOUNT_ID=$accountId; EXCEL_FILE=$excelFile; SPRINT_ANCHOR=$sprintAnchor }
```

- [ ] **Step 2: Add the field + hint and shift the Excel row down**

In the Settings dialog, after the Account ID hint block (`$lblAcctHint` … `Controls.Add($lblAcctHint)`), insert a Sprint-anchor row and shift the Excel file row + everything below. Replace the block that currently starts at `# ---- Excel file row ----` down to the `$btnExBrowse` creation with:

```powershell
    $tAnchor = Add-Row $dlg 'Sprint end:' 219
    $tAnchor.Text = if ($d['SPRINT_ANCHOR']) { $d['SPRINT_ANCHOR'] } else { '' }
    $lblAnchorHint = New-Object System.Windows.Forms.Label
    $lblAnchorHint.Text = 'Any sprint-end Friday (YYYY-MM-DD) - 2-week cycles counted from here'
    $lblAnchorHint.Location = New-Object System.Drawing.Point(114,244)
    $lblAnchorHint.Size = New-Object System.Drawing.Size(390,16)
    $lblAnchorHint.ForeColor = [System.Drawing.Color]::Gray
    $lblAnchorHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$dlg.Controls.Add($lblAnchorHint)

    # ---- Excel file row ----
    $lblEx = New-Object System.Windows.Forms.Label
    $lblEx.Text = 'Excel file:'; $lblEx.Location = New-Object System.Drawing.Point(16,269)
    $lblEx.Size = New-Object System.Drawing.Size(90,20); $lblEx.TextAlign = 'MiddleRight'
    [void]$dlg.Controls.Add($lblEx)
    $tExcel = New-Object System.Windows.Forms.TextBox
    $tExcel.Location = New-Object System.Drawing.Point(114,266)
    $tExcel.Size = New-Object System.Drawing.Size(354,24)
    $tExcel.Text = $txtFile.Text
    [void]$dlg.Controls.Add($tExcel)
    $btnExBrowse = New-Object System.Windows.Forms.Button
    $btnExBrowse.Text = '...'; $btnExBrowse.Location = New-Object System.Drawing.Point(474,266)
```

- [ ] **Step 3: Shift the checkbox, status label, buttons, and dialog height**

Apply these coordinate changes in the same dialog:

- Dialog size (`$dlg.Size = New-Object System.Drawing.Size(560,410)`) → `Size(560,470)`
- `$chkSh.Location` `Point(114,254)` → `Point(114,301)`
- `$lblTest.Location` `Point(16,257)` → `Point(16,304)`
- `$btnT.Location` `Point(16,316)` → `Point(16,363)`
- `$btnSv.Location` `Point(410,316)` → `Point(410,363)`

- [ ] **Step 4: Pass the anchor in both `Save-JiraEnv` calls**

The Test button handler:

```powershell
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text
```

The Save & Close handler:

```powershell
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text
```

- [ ] **Step 5: Document the key in `.env.example`**

After the `EXCEL_FILE=` block, add:

```
# A known sprint-end Friday (YYYY-MM-DD). 2-week sprint cycles are counted from
# it in both directions. Empty = sprint-end tasks are skipped by fill-standard.
SPRINT_ANCHOR=
```

- [ ] **Step 6: Verify syntax + render**

Run: `powershell -NoProfile -Command "$e=$null;[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

Then launch the GUI, open Settings, confirm the **Sprint end** field appears with its hint and the Excel row + buttons are not overlapping. Enter a Friday, Save & Close, and confirm `.env` gained `SPRINT_ANCHOR=<date>`.

- [ ] **Step 7: Commit**

```bash
git add jira-sync.ps1 .env.example
git commit -m "feat: SPRINT_ANCHOR setting in GUI"
```

---

### Task 4: GUI — "Standard → Excel" button

Add the button that triggers `fill-standard` from the selected month + checked weeks, widening the window to fit it.

**Files:**
- Modify: `jira-sync.ps1` (widen form + panels; add button; wire click; adjust week-column width calc)

**Interfaces:**
- Consumes: `New-Btn`, `Start-PyProc`, `$cmbMonth`, `$script:monthCodes`, `$script:weekCheckboxes`, `$txtFile`, `Set-Status`, `Append-Log` (existing).
- Produces: runs `scripts\jira-sync.py --command fill-standard --file <path> [--month <YYYY-MM>] [--weeks <mondays>]`.

- [ ] **Step 1: Widen the form and panels**

Change the form size lines (~43-45):

```powershell
$form.Size = New-Object System.Drawing.Size(1060, 680)
$form.MinimumSize = New-Object System.Drawing.Size(1060, 400)
$form.MaximumSize = New-Object System.Drawing.Size(1060, 2000)
```

Change every panel width from `900` to `1060` in these five lines (`$pnlTop`, `$pnlBtns`, `$pnlWeeks`, `$divider`, `$pnlStatus` size definitions):

```powershell
$pnlTop.Size    = New-Object System.Drawing.Size(1060,46)
$pnlBtns.Size   = New-Object System.Drawing.Size(1060,52)
$pnlWeeks.Size  = New-Object System.Drawing.Size(1060,40)
$divider.Size   = New-Object System.Drawing.Size(1060,1)
$pnlStatus.Size = New-Object System.Drawing.Size(1060,28)
```

- [ ] **Step 2: Move Stop and add the new button**

The Stop button currently: `$btnStop = New-Btn $pnlBtns 'Stop' 866 8 22 36 140 30 30`. Change its X to 1024 and add the new button before it:

```powershell
$btnFillStd = New-Btn $pnlBtns 'Standard -> Excel' 866 8 150 36 120 80 160
$btnStop    = New-Btn $pnlBtns 'Stop' 1024 8 22 36 140 30 30
```

(Remove the old `$btnStop = New-Btn ... 866 ...` line — it is replaced by the pair above.)

- [ ] **Step 3: Fix the week-checkbox column width calc**

In `Populate-Weeks`, change (~line 198):

```powershell
    $colW = 142; $startX = 306; $cols = [Math]::Floor((900 - $startX) / $colW)
```

to:

```powershell
    $colW = 142; $startX = 306; $cols = [Math]::Floor((1060 - $startX) / $colW)
```

- [ ] **Step 4: Wire the button click**

After the `$btnSync.Add_Click({ ... })` block, add:

```powershell
$btnFillStd.Add_Click({
    if (-not (Test-Path $txtFile.Text)) {
        [void][System.Windows.Forms.MessageBox]::Show('Excel file not found: ' + $txtFile.Text, 'Standard')
        return
    }
    $monthCode = $null
    if ($cmbMonth.SelectedIndex -gt 0) { $monthCode = $script:monthCodes[[string]$cmbMonth.SelectedItem] }
    $selected = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
    if (-not $monthCode -and $selected.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Select a month or check at least one week.', 'Standard')
        return
    }
    $txtLog.Clear()
    Set-Status 'Filling standard tasks...' ([System.Drawing.Color]::DodgerBlue)
    $extraArgs = ''
    if ($monthCode)            { $extraArgs += ' --month ' + $monthCode }
    if ($selected.Count -gt 0) { $extraArgs += ' --weeks ' + ($selected -join ' ') }
    Start-PyProc ('scripts\jira-sync.py --command fill-standard --file "' + $txtFile.Text + '"' + $extraArgs) {
        param($code)
        if ($code -eq 0) { Set-Status 'Standard tasks filled' ([System.Drawing.Color]::LimeGreen) }
        else             { Set-Status ('Fill failed (exit ' + $code + ')') ([System.Drawing.Color]::OrangeRed) }
    }
})
```

- [ ] **Step 5: Verify syntax + end-to-end render**

Run: `powershell -NoProfile -Command "$e=$null;[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

Then launch the GUI: confirm the **Standard → Excel** button appears in the button strip and Stop is at the right edge, click Read File to load weeks, check a future week (or pick a month), click **Standard → Excel**, and confirm the log shows `[INFO] <week>: N row(s) -> ...` and `[DONE] N rows inserted`, and that opening the Excel shows the standard tasks with hours (and QA placeholder names for future weeks).

- [ ] **Step 6: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: Standard -> Excel button in GUI"
```

---

## Notes for the implementer

- Run all three Python tests from the repo root (`python scripts/test_*.py`) — `scripts/` is `sys.path[0]` so the `from month_filter import ...` / `from standard_tasks import ...` imports resolve.
- The GUI tasks (3, 4) can only be fully verified by launching `jira-sync.ps1` — the PowerShell parse check catches syntax errors but not layout. Eyeball the dialog/button positions.
- Do not touch the D365 automation (`src/`) or the Jira query logic (`generate_week_rows`) — this feature is Excel-only.
