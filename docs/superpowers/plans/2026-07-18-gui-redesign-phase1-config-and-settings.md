# GUI Redesign Phase 1 — Editable Config + Tabbed Settings

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Externalize the standard-task schedule and the 6 JQL queries to editable JSON configs the Python reads, and rebuild the Settings dialog as a 3-tab layout (Connection / Standard tasks / Jira queries) with editing + a per-query Copy button.

**Architecture:** Two config files under `config/` (gitignored user data). `standard_tasks.py` and `jira-sync.py` load them with fallback to built-in defaults. The WinForms Settings dialog (`jira-sync.ps1` `Show-Settings`) becomes a `TabControl`; a `DataGridView` edits the schedule and multiline textboxes edit the JQL, all persisted on Save.

**Tech Stack:** Python 3 + openpyxl, PowerShell WinForms. No new dependencies.

**Spec:** `docs/superpowers/specs/2026-07-18-gui-redesign-editable-config-design.md`

## Global Constraints

- Platform is WinForms PowerShell — no web/HTML. Improve UX by hand.
- Configs live in `config/` at the repo root: `config/standard_tasks.json`, `config/jql_queries.json`. Gitignored (user data). Python ships built-in defaults so a fresh clone works without them.
- Missing/invalid config → log a `[WARN]` and use built-in defaults; never crash.
- JQL query slot **keys are fixed** (they bind to Python post-processing): `investigation`, `bug_verification`, `story_creation`, `functional_testing`, `regression_testing`, `other_qa`. Only each slot's JQL text is user-editable.
- Do NOT change `generate_week_rows` result-processing (priority buckets, name selection) — only where the JQL text comes from.
- Excel day columns: Mon=4..Fri=8. `freq` ∈ `weekly | sprint-end`; placeholders are names-only (model as `freq = placeholder` in the grid).
- Copy button places a fully-substituted JQL (`{project}`,`{account_id}` from Connection fields; `{ws}`,`{we}` from the first checked week in the main window, else the current week) on the clipboard.
- Python tests are plain scripts (`python scripts/<test>.py`, assert + print `ALL PASS`, no pytest).
- Config path from a `scripts/*.py` module: `pathlib.Path(__file__).resolve().parent.parent / "config"`.

---

### Task 1: Externalize the standard schedule

**Files:**
- Modify: `scripts/standard_tasks.py`
- Modify: `scripts/jira-sync.py` (`cmd_fill_standard` uses the loaded schedule)
- Modify: `.gitignore` (ignore `config/`)
- Test: `scripts/test_standard_config.py` (new)

**Interfaces:**
- Consumes: existing `SCHEDULE`, `QA_PLACEHOLDERS`, `rows_for_week` in `standard_tasks.py`.
- Produces:
  - `load_schedule() -> tuple[list[dict], list[str]]` — returns `(schedule, placeholders)` from `config/standard_tasks.json`, or the built-in defaults on missing/invalid.
  - `rows_for_week(..., schedule=None, placeholders=None)` — new optional params; `None` means use the module defaults (keeps existing 5-arg test calls working).

- [ ] **Step 1: Write the failing test**

Create `scripts/test_standard_config.py`:

```python
import json, tempfile, pathlib, importlib.util
from datetime import date

spec = importlib.util.spec_from_file_location("standard_tasks", "scripts/standard_tasks.py")
st = importlib.util.module_from_spec(spec); spec.loader.exec_module(st)


def load_from(tmp_config_dir):
    # Point the module at a temp config dir by monkeypatching its resolver.
    st._CONFIG_DIR = pathlib.Path(tmp_config_dir)
    return st.load_schedule()


tmp = tempfile.mkdtemp()
cfgdir = pathlib.Path(tmp) / "config"; cfgdir.mkdir()

# 1. missing file -> built-in defaults
sched, ph = load_from(cfgdir)
assert any(t["name"] == "Internal Daily meeting" for t in sched), "defaults missing"
assert "Bug verification" in ph
print("  OK missing config -> defaults")

# 2. valid file -> parsed
(cfgdir / "standard_tasks.json").write_text(json.dumps({
    "schedule": [{"name": "Custom task", "hours": 1.0, "days": ["Mon"], "freq": "weekly"}],
    "placeholders": ["QA X"]
}), encoding="utf-8")
sched, ph = load_from(cfgdir)
assert [t["name"] for t in sched] == ["Custom task"], sched
assert ph == ["QA X"], ph
print("  OK valid config parsed")

# 3. malformed file -> defaults
(cfgdir / "standard_tasks.json").write_text("{ not json", encoding="utf-8")
sched, ph = load_from(cfgdir)
assert any(t["name"] == "Internal Daily meeting" for t in sched)
print("  OK malformed config -> defaults")

# 4. rows_for_week honors a passed schedule
rows = st.rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, None, date(2026, 7, 1),
                        schedule=[{"name": "Only Mon", "hours": 2.0, "days": ["Mon"], "freq": "weekly"}],
                        placeholders=[])
assert [r["name"] for r in rows] == ["Only Mon"], rows
assert rows[0]["hours_by_col"] == {4: 2.0}, rows[0]
print("  OK rows_for_week uses passed schedule")

print("ALL PASS")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python scripts/test_standard_config.py`
Expected: FAIL — `AttributeError: module 'standard_tasks' has no attribute 'load_schedule'`

- [ ] **Step 3: Implement in `standard_tasks.py`**

At the top, after the existing imports (`from datetime import ...`, `from month_filter import ...`), add:

```python
import json
import pathlib

_CONFIG_DIR = pathlib.Path(__file__).resolve().parent.parent / "config"
```

Rename nothing; keep `SCHEDULE` and `QA_PLACEHOLDERS` as the built-in defaults. After the `QA_PLACEHOLDERS = [...]` block add:

```python
def load_schedule():
    """Return (schedule, placeholders) from config/standard_tasks.json, or the
    built-in defaults if the file is missing or invalid."""
    path = _CONFIG_DIR / "standard_tasks.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        schedule = data["schedule"]
        placeholders = data["placeholders"]
        if not isinstance(schedule, list) or not isinstance(placeholders, list):
            raise ValueError("schedule/placeholders must be lists")
        return schedule, placeholders
    except FileNotFoundError:
        return SCHEDULE, QA_PLACEHOLDERS
    except Exception as e:
        print(f"[WARN] Invalid config/standard_tasks.json ({e}); using defaults", flush=True)
        return SCHEDULE, QA_PLACEHOLDERS
```

Change `rows_for_week` to accept the schedule/placeholders. Replace its signature and the two references to `SCHEDULE`/`QA_PLACEHOLDERS`:

```python
def rows_for_week(week_start, week_end, month, sprint_anchor, today,
                  schedule=None, placeholders=None):
    if schedule is None:
        schedule = SCHEDULE
    if placeholders is None:
        placeholders = QA_PLACEHOLDERS
```

Then in the body change `for task in SCHEDULE:` → `for task in schedule:` and `for name in QA_PLACEHOLDERS:` → `for name in placeholders:`.

- [ ] **Step 4: Wire `cmd_fill_standard` to the loaded schedule**

In `scripts/jira-sync.py`, change the import to also pull `load_schedule`:

```python
from standard_tasks import rows_for_week, DAY_COL, load_schedule
```

In `cmd_fill_standard`, just before the `today = date.today()` line, add:

```python
    schedule, placeholders = load_schedule()
```

and change the `rows = rows_for_week(...)` call to pass them:

```python
        rows = rows_for_week(week["week_start"], we, args.month, sprint_anchor, today,
                             schedule=schedule, placeholders=placeholders)
```

- [ ] **Step 5: Ignore `config/`**

Add to `.gitignore` (new line):

```
config/
```

- [ ] **Step 6: Run tests**

Run: `python scripts/test_standard_config.py`
Expected: PASS — `ALL PASS`

Run: `python scripts/test_standard_tasks.py`
Expected: PASS — `ALL PASS` (unchanged 5-arg calls still use defaults)

Run: `python scripts/test_standard_insert.py`
Expected: PASS — `ALL PASS`

- [ ] **Step 7: Commit**

```bash
git add scripts/standard_tasks.py scripts/jira-sync.py scripts/test_standard_config.py .gitignore
git commit -m "feat: externalize standard schedule to config/standard_tasks.json"
```

---

### Task 2: Externalize the JQL queries

**Files:**
- Modify: `scripts/jira-sync.py` (`generate_week_rows`, add `DEFAULT_JQL` + `load_jql`)
- Test: `scripts/test_jql_config.py` (new)

**Interfaces:**
- Consumes: `generate_week_rows`, `search_jira`.
- Produces:
  - `DEFAULT_JQL: dict[str, str]` — the six templates by fixed key, using `{project} {account_id} {ws} {we}` `str.format` placeholders.
  - `load_jql() -> dict[str, str]` — merges `config/jql_queries.json` over `DEFAULT_JQL` per key; missing/invalid → all defaults.

- [ ] **Step 1: Write the failing test**

Create `scripts/test_jql_config.py`:

```python
import json, tempfile, pathlib, importlib.util

_p = pathlib.Path("scripts/jira-sync.py")
spec = importlib.util.spec_from_file_location("jira_sync", _p)
js = importlib.util.module_from_spec(spec); spec.loader.exec_module(js)

tmp = tempfile.mkdtemp()
cfgdir = pathlib.Path(tmp) / "config"; cfgdir.mkdir()
js._CONFIG_DIR = cfgdir

# 1. missing file -> all defaults, all six keys present
q = js.load_jql()
assert set(q) == {"investigation", "bug_verification", "story_creation",
                  "functional_testing", "regression_testing", "other_qa"}, set(q)
print("  OK missing config -> six default templates")

# 2. override one key, others stay default
(cfgdir / "jql_queries.json").write_text(json.dumps({
    "queries": [{"key": "investigation", "label": "x", "jql": "project = {project} custom {ws}"}]
}), encoding="utf-8")
q = js.load_jql()
assert q["investigation"] == "project = {project} custom {ws}", q["investigation"]
assert "issuetype = Bug" in q["bug_verification"], "non-overridden default lost"
print("  OK per-key override, others default")

# 3. templates substitute cleanly
resolved = q["investigation"].format(project="GT2", account_id="ACC", ws="2026-07-12", we="2026-07-18")
assert resolved == "project = GT2 custom 2026-07-12", resolved
print("  OK template substitutes {project}/{ws}")

print("ALL PASS")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python scripts/test_jql_config.py`
Expected: FAIL — `AttributeError: module 'jira_sync' has no attribute 'load_jql'`

- [ ] **Step 3: Add `DEFAULT_JQL` + `load_jql`**

In `scripts/jira-sync.py`, near the top after the `_CONFIG_DIR`-style constants (add one if absent), define the config dir and defaults. After the existing imports add (if not already present from Task 1's sibling module — this module needs its own):

```python
import pathlib as _pathlib
_CONFIG_DIR = _pathlib.Path(__file__).resolve().parent.parent / "config"

DEFAULT_JQL = {
    "investigation":
        'project = {project} AND issuetype = Bug '
        'AND created >= "{ws}" AND created <= "{we}" '
        'AND (creator = {account_id} OR reporter = {account_id}) '
        'ORDER BY priority DESC, issuetype ASC, key ASC',
    "bug_verification":
        'project = {project} AND issuetype = Bug '
        'AND status CHANGED TO "Done" BY {account_id} DURING ("{ws}","{we}") '
        'ORDER BY priority DESC, issuetype ASC, key ASC',
    "story_creation":
        'project = {project} AND issuetype = Story '
        'AND created >= "{ws}" AND created <= "{we}" '
        'AND creator = {account_id} AND parent = {project}-80 '
        'ORDER BY status DESC, issuetype ASC, key ASC',
    "functional_testing":
        'project = {project} AND issuetype = Story AND ('
        'status CHANGED FROM "Ready for QA" BY {account_id} DURING ("{ws}", "{we}") OR '
        'status CHANGED FROM "IN QA" BY {account_id} DURING ("{ws}", "{we}")) '
        'ORDER BY key ASC',
    "regression_testing":
        'project = {project} AND status CHANGED TO Done BY {account_id} DURING ("{ws}", "{we}") '
        'AND parent = {project}-73 AND summary ~ "Regression" '
        'ORDER BY status DESC, issuetype ASC, key ASC',
    "other_qa":
        'project = {project} AND status CHANGED TO Done BY {account_id} DURING ("{ws}", "{we}") '
        'AND parent = {project}-73 '
        'AND summary !~ "Smoke" AND summary !~ "Regression" AND summary !~ "Functional." '
        'ORDER BY status DESC, issuetype ASC, key ASC',
}


def load_jql() -> dict:
    """Return {key: jql_template}, overlaying config/jql_queries.json onto the
    built-in DEFAULT_JQL per key. Missing/invalid config -> all defaults."""
    templates = dict(DEFAULT_JQL)
    path = _CONFIG_DIR / "jql_queries.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        for entry in data["queries"]:
            key, jql = entry["key"], entry["jql"]
            if key in templates and isinstance(jql, str) and jql.strip():
                templates[key] = jql
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"[WARN] Invalid config/jql_queries.json ({e}); using defaults", flush=True)
    return templates
```

(`json` is already imported at the top of the file.)

- [ ] **Step 4: Use the templates in `generate_week_rows`**

At the start of `generate_week_rows`, after `ws, we = fmt(week_start), fmt(week_end)`, add:

```python
    templates = load_jql()

    def build(key):
        return templates[key].format(project=project, account_id=account_id, ws=ws, we=we)
```

Then replace each of the six `jql = (f'...')` blocks with a `build(...)` call, keeping every surrounding line (the `search_jira`, `print`, and row-append logic) untouched:

- Query 1: `jql = build("investigation")`
- Query 2: `jql = build("bug_verification")`
- Query 3: `jql = build("story_creation")`
- Query 4: `jql = build("functional_testing")`
- Query 5: `jql = build("regression_testing")`
- Query 6: `jql = build("other_qa")`

- [ ] **Step 5: Run tests**

Run: `python scripts/test_jql_config.py`
Expected: PASS — `ALL PASS`

Run: `python scripts/test_excel_insert.py` and `python scripts/test_standard_tasks.py` and `python scripts/test_standard_config.py`
Expected: each PASS — `ALL PASS` (no behavior change to those paths)

- [ ] **Step 6: Commit**

```bash
git add scripts/jira-sync.py scripts/test_jql_config.py
git commit -m "feat: externalize JQL queries to config/jql_queries.json"
```

---

### Task 3: Settings → TabControl with a Connection tab

**Files:**
- Modify: `jira-sync.ps1` (`Show-Settings`)

**Interfaces:**
- Consumes: existing `Add-Row`, `Read-EnvFile`, `Save-JiraEnv`, `$txtFile`, `$script:weekCheckboxes`, `$cmbMonth`, `$script:monthCodes`.
- Produces: a `$dlg` with a `TabControl` `$tabs`; the Connection tab hosts the existing fields; footer buttons (Test Connection, Save & Close) live on `$dlg` below the tabs. The tab pages for Standard tasks and Jira queries are added in Tasks 4–5.

- [ ] **Step 1: Resize the dialog and add the TabControl**

In `Show-Settings`, change the dialog size line to give room for tabs:

```powershell
    $dlg.Size = New-Object System.Drawing.Size(620,540)
```

Immediately after the `$dlg.BackColor = ...` line (before `function Add-Row`), add the tab control and a Connection page, and make `Add-Row` target a parent panel:

```powershell
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(8,8)
    $tabs.Size = New-Object System.Drawing.Size(596,430)
    [void]$dlg.Controls.Add($tabs)

    $tabConn = New-Object System.Windows.Forms.TabPage; $tabConn.Text = 'Connection'
    $tabStd  = New-Object System.Windows.Forms.TabPage; $tabStd.Text  = 'Standard tasks'
    $tabJql  = New-Object System.Windows.Forms.TabPage; $tabJql.Text  = 'Jira queries'
    $tabConn.BackColor = [System.Drawing.Color]::White
    $tabStd.BackColor  = [System.Drawing.Color]::White
    $tabJql.BackColor  = [System.Drawing.Color]::White
    $tabs.TabPages.AddRange(@($tabConn, $tabStd, $tabJql))
```

- [ ] **Step 2: Move the Connection fields onto `$tabConn`**

Change `Add-Row` so its `$parent` is used (it already takes `$parent`); update the existing Connection field creation calls to pass `$tabConn` instead of `$dlg`, and keep their same Y coordinates:

```powershell
    $d = Read-EnvFile
    $tUrl  = Add-Row $tabConn 'Jira URL:'    20
    $tMail = Add-Row $tabConn 'Email:'       58
    $tTok  = Add-Row $tabConn 'API Token:'   96 $true
    $tProj = Add-Row $tabConn 'Project:'    134
    $tAcct = Add-Row $tabConn 'Account ID:' 172
    $tAnchor = Add-Row $tabConn 'Sprint end:' 219
```

Re-parent every remaining Connection control (`$lblAcctHint`, `$lblAnchorHint`, the Excel-file label/textbox/browse `$lblEx`/`$tExcel`/`$btnExBrowse`, and `$chkSh`) by replacing their `[void]$dlg.Controls.Add(...)` with `[void]$tabConn.Controls.Add(...)`. Keep their existing coordinates.

- [ ] **Step 3: Move the footer buttons below the tabs**

The Test/Save buttons and the `$lblTest` status label move onto `$dlg` (not a tab), positioned under the TabControl. Change:

```powershell
    $lblTest.Location = New-Object System.Drawing.Point(16,446)
    ...
    $btnT.Location  = New-Object System.Drawing.Point(16,470)
    $btnSv.Location = New-Object System.Drawing.Point(470,470)
```

and ensure `$lblTest`, `$btnT`, `$btnSv` are added with `[void]$dlg.Controls.Add(...)` (already the case). Leave the `$btnT.Add_Click` and `$btnSv.Add_Click` bodies unchanged for now.

- [ ] **Step 4: Verify parse + render**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

Launch `jira-sync.ps1`, open Settings: three tabs appear; the Connection tab shows all existing fields correctly; Test Connection + Save & Close sit below the tabs and still work.

- [ ] **Step 5: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: settings dialog uses a TabControl; Connection tab"
```

---

### Task 4: Standard tasks tab (editable grid)

**Files:**
- Modify: `jira-sync.ps1` (`Show-Settings` — populate `$tabStd`; extend the Save handler)

**Interfaces:**
- Consumes: `$tabStd`, `$scriptDir`, the Save & Close handler.
- Produces: `$gridStd` (a `DataGridView`) and a `Save-StandardConfig` helper writing `config/standard_tasks.json`.

- [ ] **Step 1: Add a config read/write helper (script scope)**

Near the top of `jira-sync.ps1` (after `Save-JiraEnv`), add:

```powershell
$configDir = Join-Path $scriptDir 'config'

function Read-JsonConfig($name) {
    $p = Join-Path $configDir $name
    if (Test-Path $p) {
        try { return (Get-Content $p -Raw -Encoding utf8 | ConvertFrom-Json) } catch { return $null }
    }
    return $null
}

function Write-JsonConfig($name, $obj) {
    if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir | Out-Null }
    $obj | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $configDir $name) -Encoding utf8
}
```

- [ ] **Step 2: Build the grid on `$tabStd`**

In `Show-Settings`, after the tabs are created, add:

```powershell
    $gridStd = New-Object System.Windows.Forms.DataGridView
    $gridStd.Location = New-Object System.Drawing.Point(8,8)
    $gridStd.Size = New-Object System.Drawing.Size(572,350)
    $gridStd.AllowUserToAddRows = $true
    $gridStd.AllowUserToDeleteRows = $true
    $gridStd.AutoSizeColumnsMode = 'Fill'
    [void]$tabStd.Controls.Add($gridStd)
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colName.HeaderText = 'Task'; $colName.FillWeight = 200
    [void]$gridStd.Columns.Add($colName)
    foreach ($dh in 'Mon','Tue','Wed','Thu','Fri') {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.HeaderText = $dh; $col.FillWeight = 45
        [void]$gridStd.Columns.Add($col)
    }
    $colFreq = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colFreq.HeaderText = 'Frequency'; $colFreq.FillWeight = 90
    [void]$colFreq.Items.AddRange(@('weekly','sprint-end','placeholder'))
    [void]$gridStd.Columns.Add($colFreq)

    $lblStdHint = New-Object System.Windows.Forms.Label
    $lblStdHint.Text = 'Hours per day in Mon-Fri; blank = not that day. placeholder rows carry a name only.'
    $lblStdHint.Location = New-Object System.Drawing.Point(8,362); $lblStdHint.Size = New-Object System.Drawing.Size(572,18)
    $lblStdHint.ForeColor = [System.Drawing.Color]::Gray; $lblStdHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$tabStd.Controls.Add($lblStdHint)
```

- [ ] **Step 3: Populate the grid from config-or-defaults**

Standard defaults must match `standard_tasks.py`. Add, after the grid is built:

```powershell
    $stdCfg = Read-JsonConfig 'standard_tasks.json'
    $defaultSchedule = @(
        @{name='Internal Daily meeting'; hours=0.5; days=@('Mon','Tue','Wed','Thu','Fri'); freq='weekly'},
        @{name='Internal bug triage'; hours=0.5; days=@('Tue'); freq='weekly'},
        @{name='External customer meeting'; hours=1.0; days=@('Tue','Wed'); freq='weekly'},
        @{name='Weekly project report'; hours=1.0; days=@('Fri'); freq='weekly'},
        @{name='Internal sprint review'; hours=0.5; days=@('Fri'); freq='sprint-end'},
        @{name='External sprint review'; hours=1.0; days=@('Fri'); freq='sprint-end'},
        @{name='Summary report creation'; hours=2.0; days=@('Fri'); freq='sprint-end'}
    )
    $defaultPlaceholders = @('Bug verification','Functional testing','Automation test maintenance','Investigation issue')

    $schedule = if ($stdCfg -and $stdCfg.schedule) { $stdCfg.schedule } else { $defaultSchedule }
    $placeholders = if ($stdCfg -and $stdCfg.placeholders) { $stdCfg.placeholders } else { $defaultPlaceholders }

    $dayIndex = @{ Mon=1; Tue=2; Wed=3; Thu=4; Fri=5 }
    foreach ($t in $schedule) {
        $cells = @($t.name, '', '', '', '', '', $t.freq)
        foreach ($dn in $t.days) { $cells[$dayIndex[$dn]] = [string]$t.hours }
        [void]$gridStd.Rows.Add($cells)
    }
    foreach ($ph in $placeholders) {
        [void]$gridStd.Rows.Add(@($ph, '', '', '', '', '', 'placeholder'))
    }
```

- [ ] **Step 4: Persist the grid on Save**

Add a helper (near `Show-Settings` or inside it before the buttons):

```powershell
    $script:SaveStandardFromGrid = {
        $sched = @(); $ph = @()
        foreach ($row in $gridStd.Rows) {
            if ($row.IsNewRow) { continue }
            $name = [string]$row.Cells[0].Value
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $freq = [string]$row.Cells[6].Value
            if ($freq -eq 'placeholder') { $ph += $name.Trim(); continue }
            $days = @(); $hours = $null
            foreach ($dn in 'Mon','Tue','Wed','Thu','Fri') {
                $v = [string]$row.Cells[$dayIndex[$dn]].Value
                if (-not [string]::IsNullOrWhiteSpace($v)) { $days += $dn; $hours = [double]$v }
            }
            if ($days.Count -eq 0) { continue }
            $sched += @{ name=$name.Trim(); hours=$hours; days=$days; freq=$(if ($freq) { $freq } else { 'weekly' }) }
        }
        Write-JsonConfig 'standard_tasks.json' @{ schedule=$sched; placeholders=$ph }
    }
```

Then in the `$btnSv.Add_Click` body, after the existing `Save-JiraEnv ...` line, add:

```powershell
        & $script:SaveStandardFromGrid
```

- [ ] **Step 5: Verify parse + render**

Run the parser check (same command as Task 3 Step 4) → `OK`.
Launch Settings → Standard tasks tab shows the schedule + placeholders; edit a value, add a row, Save & Close; reopen and confirm `config/standard_tasks.json` reflects the edits and the grid reloads them.

- [ ] **Step 6: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: editable standard-tasks grid in Settings"
```

---

### Task 5: Jira queries tab (editable JQL + Copy)

**Files:**
- Modify: `jira-sync.ps1` (`Show-Settings` — populate `$tabJql`; extend Save)

**Interfaces:**
- Consumes: `$tabJql`, `Read-JsonConfig`/`Write-JsonConfig`, `$tUrl`/`$tProj`/`$tAcct` (Connection fields), `$script:weekCheckboxes`.
- Produces: per-slot textboxes `$script:jqlBoxes` (hashtable key→TextBox) and `$script:SaveJqlConfig` writing `config/jql_queries.json`.

- [ ] **Step 1: Define the six slots + defaults, build the editor rows**

In `Show-Settings`, after the tabs exist, add:

```powershell
    $jqlSlots = @(
        @{ key='investigation';      label='1. Investigation issues' },
        @{ key='bug_verification';   label='2. Bug verification' },
        @{ key='story_creation';     label='3. User story creation' },
        @{ key='functional_testing'; label='4. Functional testing' },
        @{ key='regression_testing'; label='5. Regression testing' },
        @{ key='other_qa';           label='6. Other QA activities' }
    )
    $defaultJql = @{
        investigation      = 'project = {project} AND issuetype = Bug AND created >= "{ws}" AND created <= "{we}" AND (creator = {account_id} OR reporter = {account_id}) ORDER BY priority DESC, issuetype ASC, key ASC'
        bug_verification   = 'project = {project} AND issuetype = Bug AND status CHANGED TO "Done" BY {account_id} DURING ("{ws}","{we}") ORDER BY priority DESC, issuetype ASC, key ASC'
        story_creation     = 'project = {project} AND issuetype = Story AND created >= "{ws}" AND created <= "{we}" AND creator = {account_id} AND parent = {project}-80 ORDER BY status DESC, issuetype ASC, key ASC'
        functional_testing = 'project = {project} AND issuetype = Story AND (status CHANGED FROM "Ready for QA" BY {account_id} DURING ("{ws}", "{we}") OR status CHANGED FROM "IN QA" BY {account_id} DURING ("{ws}", "{we}")) ORDER BY key ASC'
        regression_testing = 'project = {project} AND status CHANGED TO Done BY {account_id} DURING ("{ws}", "{we}") AND parent = {project}-73 AND summary ~ "Regression" ORDER BY status DESC, issuetype ASC, key ASC'
        other_qa           = 'project = {project} AND status CHANGED TO Done BY {account_id} DURING ("{ws}", "{we}") AND parent = {project}-73 AND summary !~ "Smoke" AND summary !~ "Regression" AND summary !~ "Functional." ORDER BY status DESC, issuetype ASC, key ASC'
    }
    $jqlCfg = Read-JsonConfig 'jql_queries.json'
    $jqlValues = @{}
    foreach ($k in $defaultJql.Keys) { $jqlValues[$k] = $defaultJql[$k] }
    if ($jqlCfg -and $jqlCfg.queries) {
        foreach ($q in $jqlCfg.queries) { if ($jqlValues.ContainsKey($q.key) -and $q.jql) { $jqlValues[$q.key] = $q.jql } }
    }

    $panelJql = New-Object System.Windows.Forms.Panel
    $panelJql.Location = New-Object System.Drawing.Point(0,0)
    $panelJql.Dock = 'Fill'; $panelJql.AutoScroll = $true
    [void]$tabJql.Controls.Add($panelJql)

    $script:jqlBoxes = @{}
    $y = 8
    foreach ($slot in $jqlSlots) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $slot.label; $lbl.Location = New-Object System.Drawing.Point(8,$y)
        $lbl.Size = New-Object System.Drawing.Size(400,16); $lbl.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
        [void]$panelJql.Controls.Add($lbl)
        $box = New-Object System.Windows.Forms.TextBox
        $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
        $box.Location = New-Object System.Drawing.Point(8,($y+18)); $box.Size = New-Object System.Drawing.Size(490,46)
        $box.Font = New-Object System.Drawing.Font('Consolas',8)
        $box.Text = [string]$jqlValues[$slot.key]
        [void]$panelJql.Controls.Add($box)
        $script:jqlBoxes[$slot.key] = $box

        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(502,($y+18)); $btnCopy.Size = New-Object System.Drawing.Size(64,46)
        $btnCopy.FlatStyle = 'Flat'; $btnCopy.Tag = $slot.key
        $btnCopy.Add_Click({
            $key = $this.Tag
            $tpl = [string]$script:jqlBoxes[$key].Text
            $proj = if ($tProj.Text) { $tProj.Text } else { 'GT2' }
            $acct = $tAcct.Text
            $checked = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
            if ($checked.Count -gt 0) {
                $ws = [string]$checked[0]
                $we = ([datetime]::ParseExact($ws,'yyyy-MM-dd',$null).AddDays(6)).ToString('yyyy-MM-dd')
            } else {
                $today = Get-Date
                $sunday = $today.AddDays(-[int]$today.DayOfWeek)
                $ws = $sunday.ToString('yyyy-MM-dd'); $we = $sunday.AddDays(6).ToString('yyyy-MM-dd')
            }
            $resolved = $tpl.Replace('{project}',$proj).Replace('{account_id}',$acct).Replace('{ws}',$ws).Replace('{we}',$we)
            [System.Windows.Forms.Clipboard]::SetText($resolved)
        })
        [void]$panelJql.Controls.Add($btnCopy)
        $y += 74
    }
```

- [ ] **Step 2: Persist the JQL editors on Save**

Add near the other save helper:

```powershell
    $script:SaveJqlConfig = {
        $queries = @()
        foreach ($slot in $jqlSlots) {
            $queries += @{ key=$slot.key; label=$slot.label; jql=[string]$script:jqlBoxes[$slot.key].Text }
        }
        Write-JsonConfig 'jql_queries.json' @{ queries=$queries }
    }
```

Then in `$btnSv.Add_Click`, after the `& $script:SaveStandardFromGrid` line, add:

```powershell
        & $script:SaveJqlConfig
```

- [ ] **Step 3: Verify parse + render + copy**

Parser check → `OK`.
Launch Settings → Jira queries tab shows six editable boxes with Copy buttons. Check a week in the main window, open Settings, click Copy on query 1, paste into a text editor: it should be a fully-substituted JQL (real project, account_id, and the selected week's dates, no `{}` left). Edit a query, Save & Close, reopen → the edit persisted in `config/jql_queries.json`.

- [ ] **Step 4: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: editable JQL queries tab with resolved Copy"
```

---

## Notes for the implementer

- Run Python tests from the repo root (`python scripts/test_*.py`); `scripts/` is on `sys.path[0]`.
- The GUI tasks (3–5) can only be fully verified by launching `jira-sync.ps1` — the parser check catches syntax, not layout or runtime WinForms errors. Eyeball each tab and exercise Save/reload and Copy.
- Keep the Python `DEFAULT_JQL`/`SCHEDULE` and the PowerShell `$defaultJql`/`$defaultSchedule` in sync — they encode the same defaults in two languages by necessity (Python for headless CLI, PowerShell for the offline GUI). If one changes, change both.
