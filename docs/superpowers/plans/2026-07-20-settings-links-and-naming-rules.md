# Settings Links Tab + Configurable Jira-Query Naming Rules Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Links" Settings tab (Excel file path + QAE reporting-rules link), and make the row-naming logic for 5 of the 6 built-in Jira-sync queries user-configurable (aggregation mode + name template) directly from the "Jira queries" Settings tab, while leaving `other_qa`'s hardcoded keyword classification untouched.

**Architecture:** A small, pure "naming rules engine" added to `scripts/jira-sync.py` (mode-based context builders + template rendering with safe fallback) replaces 5 of the 6 hardcoded naming blocks in `generate_week_rows`. `jira-sync.ps1`'s existing `New-JqlRow`/Settings-dialog JQL editor is extended so its title textbox *is* the naming template, with a new mode dropdown next to it; both persist to a new `name_rules` key in `config/jql_queries.json`, independent of the pre-existing (and now-unused-by-Python) `queries[].label` field to avoid a silent migration regression.

**Tech Stack:** PowerShell + WinForms (`jira-sync.ps1`), Python 3 stdlib + openpyxl (`scripts/jira-sync.py`), plain-script tests (no pytest) run via `python scripts/test_*.py` and `powershell -File scripts/test_*.ps1`.

## Global Constraints

- Spec: `docs/superpowers/specs/2026-07-20-settings-links-and-naming-rules-design.md` — follow it exactly; this plan implements it task-by-task.
- Default behavior for existing installs (including the current `config/jql_queries.json` on disk, which has `label` fields with ordinal text like `"1. Investigation issues"`) must be **unchanged** until the user edits a template or mode in Settings.
- `other_qa` keyword classification (`_AUTOMATION_RULES`, `choose_automation_name`) stays hardcoded — not exposed as an editable rules grid.
- Custom queries (`+ Add query`, `extra_*` keys) stay clipboard-copy-only — the sync never executes them; no naming-rule UI for them.
- No "Open file" button for the Excel path (explicitly declined by user).
- All new UI copy is in English, matching the existing Settings dialog's language.
- No pytest / vitest framework changes — new Python tests follow the existing plain-script convention (see `scripts/test_jql_config.py`): assert + print, `sys.exit`/non-zero exit only via explicit `AssertionError`, run directly with `python scripts/test_name.py`.

---

### Task 1: Python naming-rules engine (pure functions)

**Files:**
- Modify: `scripts/jira-sync.py` — insert new code block immediately before `def generate_week_rows(` (currently line 219; insert after the `_COUNT_NAMES = {...}` block that ends at line 205, before the blank lines preceding `def generate_week_rows`)
- Test: `scripts/test_name_rules.py` (new)

**Interfaces:**
- Consumes: `priority_bucket(name: str) -> str` (existing, jira-sync.py:169-175), `extract_last_number(text: str) -> str | None` (existing, jira-sync.py:178-182), `json`, `_CONFIG_DIR` (existing module globals)
- Produces: `DEFAULT_NAME_RULES: dict`, `load_name_rules() -> dict`, `render_name(template: str, ctx: dict, default_template: str) -> str`, `build_names(mode: str, issues: list, template: str, default_template: str, key: str | None = None) -> list[str]` — all consumed by Task 2.

- [ ] **Step 1: Write the failing test**

Create `scripts/test_name_rules.py`:

```python
import json, tempfile, pathlib, importlib.util

_p = pathlib.Path("scripts/jira-sync.py")
spec = importlib.util.spec_from_file_location("jira_sync_nr", _p)
js = importlib.util.module_from_spec(spec); spec.loader.exec_module(js)

tmp = tempfile.mkdtemp()
cfgdir = pathlib.Path(tmp) / "config"; cfgdir.mkdir()
js._CONFIG_DIR = cfgdir


def issue(key, summary, priority="Medium"):
    return {"key": key, "fields": {"summary": summary, "priority": {"name": priority}}}


# 1. load_name_rules(): missing config -> defaults for all 5 keys
rules = js.load_name_rules()
assert set(rules) == {"investigation", "bug_verification", "story_creation",
                      "functional_testing", "regression_testing"}, set(rules)
assert rules["bug_verification"] == {"mode": "priority", "template": "Bug verification {buckets}"}, rules["bug_verification"]
print("  OK load_name_rules: missing config -> five defaults")

# 2. load_name_rules(): partial override, others stay default
(cfgdir / "jql_queries.json").write_text(json.dumps({
    "name_rules": {"investigation": {"mode": "count", "template": "Investigations: {count}"}}
}), encoding="utf-8")
rules = js.load_name_rules()
assert rules["investigation"] == {"mode": "count", "template": "Investigations: {count}"}, rules["investigation"]
assert rules["story_creation"] == js.DEFAULT_NAME_RULES["story_creation"], rules["story_creation"]
print("  OK load_name_rules: per-key override, others default")

# 3. load_name_rules(): invalid mode / blank template are ignored, default kept
(cfgdir / "jql_queries.json").write_text(json.dumps({
    "name_rules": {"story_creation": {"mode": "bogus", "template": "   "}}
}), encoding="utf-8")
rules = js.load_name_rules()
assert rules["story_creation"] == js.DEFAULT_NAME_RULES["story_creation"], rules["story_creation"]
print("  OK load_name_rules: invalid mode / blank template fall back to default")

# 4. build_names: count mode
names = js.build_names("count", [issue("A-1", "x"), issue("A-2", "y")], "Total {count}", "Total {count}")
assert names == ["Total 2"], names
names = js.build_names("count", [], "Total {count}", "Total {count}")
assert names == [], names
print("  OK build_names: count mode")

# 5. build_names: priority mode, zero buckets omitted from {buckets}, raw counts available
issues = [issue("A-1", "x", "Highest"), issue("A-2", "y", "Medium"), issue("A-3", "z", "Medium")]
names = js.build_names("priority", issues, "Bug verification {buckets}", "Bug verification {buckets}")
assert names == ["Bug verification P1-1, P2-2"], names
names = js.build_names("priority", issues, "P1={P1} P2={P2} P3={P3} n={count}", "Bug verification {buckets}")
assert names == ["P1=1 P2=2 P3=0 n=3"], names
print("  OK build_names: priority mode")

# 6. build_names: per_issue mode, one name per issue, {number} extracted from summary
issues = [issue("A-1", "Regression (16/0/0/16)"), issue("A-2", "Regression no number")]
names = js.build_names("per_issue", issues, "Regression testing {number} test items", "Regression testing {number} test items")
assert names[0] == "Regression testing 16 test items", names[0]
print("  OK build_names: per_issue mode extracts {number}")

# 7. build_names: regression_testing special-case fallback when {number} is empty
names = js.build_names("per_issue", issues, "Regression testing {number} test items",
                        "Regression testing {number} test items", key="regression_testing")
assert names[1] == "Regression testing [REVIEW] A-2", names[1]
print("  OK build_names: regression_testing falls back to [REVIEW] when no number")

# 8. render_name: invalid placeholder falls back to default_template, doesn't raise
name = js.render_name("Bad {nope}", {"count": 3}, "Total {count}")
assert name == "Total 3", name
print("  OK render_name: invalid placeholder falls back to default")

print("ALL PASS")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python scripts/test_name_rules.py`
Expected: `AttributeError: module 'jira_sync_nr' has no attribute 'load_name_rules'` (or similar — the functions don't exist yet)

- [ ] **Step 3: Write minimal implementation**

In `scripts/jira-sync.py`, insert this block immediately before `def generate_week_rows(` (i.e. right after the existing `_COUNT_NAMES = {...}` block, keeping one blank line of separation before and after):

```python
DEFAULT_NAME_RULES = {
    "investigation":      {"mode": "count",     "template": "Investigation issue {count}"},
    "bug_verification":   {"mode": "priority",  "template": "Bug verification {buckets}"},
    "story_creation":     {"mode": "count",     "template": "User story creation {count}"},
    "functional_testing": {"mode": "per_issue", "template": "Functional testing of story ID {key}"},
    "regression_testing": {"mode": "per_issue", "template": "Regression testing {number} test items"},
}

_VALID_MODES = {"count", "priority", "per_issue"}


def load_name_rules() -> dict:
    """Return {key: {"mode", "template"}} for the 5 templatable query keys,
    overlaying config/jql_queries.json's "name_rules" onto DEFAULT_NAME_RULES.
    An invalid mode or a blank template for a key keeps that field's default;
    missing/invalid config keeps everything default."""
    rules = {k: dict(v) for k, v in DEFAULT_NAME_RULES.items()}
    path = _CONFIG_DIR / "jql_queries.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        for key, cfg in (data.get("name_rules") or {}).items():
            if key not in rules or not isinstance(cfg, dict):
                continue
            mode = cfg.get("mode")
            if mode in _VALID_MODES:
                rules[key]["mode"] = mode
            template = cfg.get("template")
            if isinstance(template, str) and template.strip():
                rules[key]["template"] = template
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"[WARN] Invalid config/jql_queries.json name_rules ({e}); using defaults", flush=True)
    return rules


def _count_ctx(issues: list) -> dict:
    return {"count": len(issues)}


def _priority_ctx(issues: list) -> dict:
    buckets = {"P1": 0, "P2": 0, "P3": 0}
    for iss in issues:
        p = iss["fields"].get("priority", {}).get("name", "Medium")
        buckets[priority_bucket(p)] += 1
    parts = [f"{k}-{v}" for k, v in buckets.items() if v > 0]
    return {**buckets, "count": len(issues), "buckets": ", ".join(parts)}


def _per_issue_ctx(iss: dict) -> dict:
    summary = iss["fields"]["summary"]
    return {"key": iss["key"], "summary": summary, "number": extract_last_number(summary) or ""}


def render_name(template: str, ctx: dict, default_template: str) -> str:
    """Render `template` against `ctx`; on a bad/unknown placeholder, log a
    warning and fall back to `default_template` (always safe for `ctx`)."""
    try:
        return template.format(**ctx)
    except (KeyError, IndexError, ValueError) as e:
        print(f"[WARN] Invalid name_template ({e}); using default", flush=True)
        return default_template.format(**ctx)


def build_names(mode: str, issues: list, template: str, default_template: str, key: str | None = None) -> list[str]:
    """Return one name per Excel row to insert: a single name for "count"/
    "priority" modes (empty list if no issues), or one name per issue for
    "per_issue". For key == "regression_testing" specifically, a per-issue
    result with no extractable {number} is overridden with the fixed
    [REVIEW] fallback, matching the tool's pre-existing behavior."""
    if not issues:
        return []
    if mode == "count":
        return [render_name(template, _count_ctx(issues), default_template)]
    if mode == "priority":
        return [render_name(template, _priority_ctx(issues), default_template)]
    if mode == "per_issue":
        names = []
        for iss in issues:
            ctx = _per_issue_ctx(iss)
            name = render_name(template, ctx, default_template)
            if key == "regression_testing" and not ctx["number"]:
                name = f"Regression testing [REVIEW] {iss['key']}"
            names.append(name)
        return names
    raise ValueError(f"unknown naming mode {mode!r}")

```

- [ ] **Step 4: Run test to verify it passes**

Run: `python scripts/test_name_rules.py`
Expected: 8 `OK` lines then `ALL PASS`

- [ ] **Step 5: Commit**

```bash
git add scripts/jira-sync.py scripts/test_name_rules.py
git commit -m "feat: add configurable naming-rules engine (mode + template) to jira-sync.py"
```

---

### Task 2: Wire the naming-rules engine into `generate_week_rows`

**Files:**
- Modify: `scripts/jira-sync.py:219-304` (the `generate_week_rows` function body)
- Test: `scripts/test_generate_week_rows_names.py` (new)

**Interfaces:**
- Consumes: `load_name_rules()`, `build_names()`, `DEFAULT_NAME_RULES` (Task 1), existing `load_jql()`, `search_jira()`, `choose_automation_name()`
- Produces: `generate_week_rows(...)` unchanged signature and return type (`list[tuple[str, str]]`), same console log format, now driven by `name_rules` for 5 of its 6 queries.

- [ ] **Step 1: Write the failing test**

Create `scripts/test_generate_week_rows_names.py`:

```python
import json, tempfile, pathlib, importlib.util, collections
from datetime import date

_p = pathlib.Path("scripts/jira-sync.py")
spec = importlib.util.spec_from_file_location("jira_sync_gwr", _p)
js = importlib.util.module_from_spec(spec); spec.loader.exec_module(js)

tmp = tempfile.mkdtemp()
cfgdir = pathlib.Path(tmp) / "config"; cfgdir.mkdir()
js._CONFIG_DIR = cfgdir


def issue(key, summary, priority="Medium"):
    return {"key": key, "fields": {"summary": summary, "priority": {"name": priority}}}


def run(queued_results):
    queue = collections.deque(queued_results)
    js.search_jira = lambda base_url, email, token, jql: queue.popleft()
    return js.generate_week_rows(
        "https://x.atlassian.net", "e@x.com", "tok", "ACC", "GT2",
        date(2026, 7, 13), date(2026, 7, 19),
    )


# 1. Default name_rules reproduce the tool's pre-existing naming behavior.
rows = run([
    [],                                                        # investigation: 0 found
    [issue("GT2-1", "bug", "Highest"), issue("GT2-2", "bug2")],  # bug_verification: P1=1, P2=1
    [issue("GT2-3", "story")],                                # story_creation
    [issue("GT2-4", "func")],                                 # functional_testing
    [issue("GT2-5", "Regression (16/0/0/16)")],                # regression_testing
    [issue("GT2-6", "Automation test creation")],             # other_qa
])
names = [n for n, _ in rows]
assert names == [
    "Bug verification P1-1, P2-1",
    "User story creation 1",
    "Functional testing of story ID GT2-4",
    "Regression testing 16 test items",
    "Automation test creation",
], names
print("  OK default name_rules reproduce pre-existing naming")

# 2. A custom template (raw P1/P2/P3, zero buckets included) is honored.
(cfgdir / "jql_queries.json").write_text(json.dumps({
    "name_rules": {
        "bug_verification": {"mode": "priority", "template": "Bug verification P1-{P1}, P2-{P2}, P3-{P3}"}
    }
}), encoding="utf-8")
rows = run([[], [issue("GT2-1", "bug", "Highest")], [], [], [], []])
names = [n for n, _ in rows]
assert names == ["Bug verification P1-1, P2-0, P3-0"], names
print("  OK custom template honored, zero buckets included when explicit")

# 3. An invalid placeholder falls back to the default template for that key.
(cfgdir / "jql_queries.json").write_text(json.dumps({
    "name_rules": {"story_creation": {"mode": "count", "template": "User story creation {nope}"}}
}), encoding="utf-8")
rows = run([[], [], [issue("GT2-3", "story")], [], [], []])
names = [n for n, _ in rows]
assert names == ["User story creation 1"], names
print("  OK invalid placeholder falls back to default template")

print("ALL PASS")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python scripts/test_generate_week_rows_names.py`
Expected: `AssertionError` on the first `names == [...]` check (current hardcoded logic produces the same names today, so this specific test may actually pass already for scenario 1 — but scenarios 2 and 3 MUST fail, since `name_rules` isn't read yet). Confirm at least one `AssertionError` before proceeding.

- [ ] **Step 3: Write minimal implementation**

Replace the full body of `generate_week_rows` (`scripts/jira-sync.py:219-304`) with:

```python
def generate_week_rows(
    base_url: str, email: str, token: str, account_id: str,
    project: str, week_start: date, week_end: date,
) -> list[tuple[str, str]]:
    ws, we = fmt(week_start), fmt(week_end)
    rows: list[tuple[str, str]] = []

    templates = load_jql()
    name_rules = load_name_rules()

    def build(key):
        return templates[key].format(project=project, account_id=account_id, ws=ws, we=we)

    def search_url(jql: str) -> str:
        return f"{base_url}/issues/?jql={urllib.parse.quote(jql)}"

    def issue_url(key: str) -> str:
        return f"{base_url}/browse/{key}"

    def names_for(key: str, issues: list) -> list[str]:
        rule = name_rules[key]
        default_template = DEFAULT_NAME_RULES[key]["template"]
        return build_names(rule["mode"], issues, rule["template"], default_template, key=key)

    # 1. Investigation issue — bugs created this week that the user filed
    #    (creator) or is the reporter of. In Jira creator (who clicked "Create")
    #    and reporter (who the bug is attributed to) can differ, so match either.
    jql = build("investigation")
    issues = search_jira(base_url, email, token, jql)
    names = names_for("investigation", issues)
    if names:
        print(f"  [1/6] Investigation issues (bugs you created/reported this week): {len(issues)} found → \"{names[0]}\"", flush=True)
        rows.append((names[0], search_url(jql)))
    else:
        print(f"  [1/6] Investigation issues (bugs you created/reported this week): 0 found", flush=True)

    # 2. Bug verification — count by priority bucket
    jql = build("bug_verification")
    issues = search_jira(base_url, email, token, jql)
    names = names_for("bug_verification", issues)
    if names:
        print(f"  [2/6] Bug verification (closed by you): {len(issues)} found → \"{names[0]}\"", flush=True)
        rows.append((names[0], search_url(jql)))
    else:
        print(f"  [2/6] Bug verification (closed by you): 0 found", flush=True)

    # 3. User story creation — count of stories created under GT2-80
    jql = build("story_creation")
    issues = search_jira(base_url, email, token, jql)
    names = names_for("story_creation", issues)
    if names:
        print(f"  [3/6] User story creation: {len(issues)} found → \"{names[0]}\"", flush=True)
        rows.append((names[0], search_url(jql)))
    else:
        print(f"  [3/6] User story creation: 0 found", flush=True)

    # 4. Functional testing — one row per story key, direct issue link
    jql = build("functional_testing")
    issues = search_jira(base_url, email, token, jql)
    print(f"  [4/6] Functional testing stories: {len(issues)} found", flush=True)
    for iss, name in zip(issues, names_for("functional_testing", issues)):
        print(f"         {iss['key']} → \"{name}\"", flush=True)
        rows.append((name, issue_url(iss['key'])))

    # 5. Regression testing — last number from summary, direct issue link
    jql = build("regression_testing")
    issues = search_jira(base_url, email, token, jql)
    if len(issues) > 1:
        print(f"  [WARN] Regression: {len(issues)} items found (expected 1)", flush=True)
    print(f"  [5/6] Regression testing: {len(issues)} found", flush=True)
    for iss, name in zip(issues, names_for("regression_testing", issues)):
        print(f"         {iss['key']} → \"{name}\"", flush=True)
        rows.append((name, issue_url(iss['key'])))

    # 6. Other QA activities (GT2-73, non-smoke/regression/functional) — direct
    #    issue link. Naming stays hardcoded (keyword classification), not
    #    driven by name_rules — see choose_automation_name.
    jql = build("other_qa")
    issues = search_jira(base_url, email, token, jql)
    print(f"  [6/6] Other QA activities ({project}-73 subtasks): {len(issues)} found", flush=True)
    for iss in issues:
        name = choose_automation_name(iss["fields"]["summary"], iss["key"])
        print(f"         {iss['key']} → \"{name}\"", flush=True)
        rows.append((name, issue_url(iss['key'])))

    return rows
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python scripts/test_generate_week_rows_names.py`
Expected: 3 `OK` lines then `ALL PASS`

Also re-run Task 1's test and the pre-existing config test to confirm no regression:
Run: `python scripts/test_name_rules.py && python scripts/test_jql_config.py`
Expected: both print `ALL PASS`

- [ ] **Step 5: Commit**

```bash
git add scripts/jira-sync.py scripts/test_generate_week_rows_names.py
git commit -m "feat: drive 5 of 6 jira-sync row names from configurable name_rules"
```

---

### Task 3: `New-JqlRow` — name-based control lookup + optional mode dropdown

**Files:**
- Modify: `jira-sync.ps1:397-451` (`New-JqlRow` function)
- Test: `scripts/test_jql_row_gui.ps1` (rewrite — it's a standalone duplicate of `New-JqlRow`, not sourced from `jira-sync.ps1`, per existing repo convention for headless WinForms testing)

**Interfaces:**
- Consumes (from enclosing `Show-Settings` scope, added in Task 4 but referenced here — for this task's own test file, defined locally): `$modeDisplay: hashtable` (mode key → display string), `$modeHint: hashtable` (mode key → placeholder hint string), `$jqlTips: System.Windows.Forms.ToolTip`
- Produces: `New-JqlRow($key, $titleText, $jqlText, $titleEditable, $modeKey = $null)` — new optional 5th parameter; row's title control now named `'titleBox'`, JQL control named `'jqlBox'` (both found via `$row.Controls.Find(name, $false)`), consumed by Task 4's Save/Load logic and by the `other_qa` row's tooltip wiring.

- [ ] **Step 1: Write the failing test**

Replace the entire contents of `scripts/test_jql_row_gui.ps1` with:

```powershell
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:weekCheckboxes = @()
$tProj = New-Object System.Windows.Forms.TextBox; $tProj.Text = 'GT2'
$tAcct = New-Object System.Windows.Forms.TextBox; $tAcct.Text = 'ACC123'
$panelJqlFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$jqlTips = New-Object System.Windows.Forms.ToolTip

$modeDisplay = @{ count='Total count'; priority='Priority breakdown (P1-P3)'; per_issue='One row per issue' }
$modeValue   = @{ 'Total count'='count'; 'Priority breakdown (P1-P3)'='priority'; 'One row per issue'='per_issue' }
$modeHint = @{
    count     = 'Placeholders: {count}'
    priority  = 'Placeholders: {P1} {P2} {P3} {count} {buckets}'
    per_issue = 'Placeholders: {key} {summary} {number}'
}

function New-JqlRow($key, $titleText, $jqlText, $titleEditable, $modeKey = $null) {
    $row = New-Object System.Windows.Forms.Panel
    $row.Size = New-Object System.Drawing.Size(566,74)
    $row.Margin = New-Object System.Windows.Forms.Padding(4,4,4,0)
    $row.Tag = $key

    if ($titleEditable) {
        $titleCtl = New-Object System.Windows.Forms.TextBox
    } else {
        $titleCtl = New-Object System.Windows.Forms.Label
    }
    $titleCtl.Name = 'titleBox'
    $titleCtl.Text = $titleText; $titleCtl.Location = New-Object System.Drawing.Point(0,0)
    $titleWidth = if ($modeKey) { 260 } else { 420 }
    $titleCtl.Size = New-Object System.Drawing.Size($titleWidth,18)
    $titleCtl.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
    [void]$row.Controls.Add($titleCtl)

    if ($modeKey) {
        $combo = New-Object System.Windows.Forms.ComboBox
        $combo.Name = 'modeCombo'
        $combo.DropDownStyle = 'DropDownList'
        $combo.Location = New-Object System.Drawing.Point(264,0)
        $combo.Size = New-Object System.Drawing.Size(156,20)
        $combo.Font = New-Object System.Drawing.Font('Segoe UI',8)
        [void]$combo.Items.AddRange(@('Total count','Priority breakdown (P1-P3)','One row per issue'))
        $combo.SelectedItem = $modeDisplay[$modeKey]
        $jqlTips.SetToolTip($titleCtl, $modeHint[$modeKey])
        $jqlTips.SetToolTip($combo, [string]$combo.SelectedItem)
        $combo.Add_SelectedIndexChanged({
            $newMode = $modeValue[[string]$this.SelectedItem]
            $jqlTips.SetToolTip($this.Parent.Controls.Find('titleBox',$false)[0], $modeHint[$newMode])
            $jqlTips.SetToolTip($this, [string]$this.SelectedItem)
        })
        [void]$row.Controls.Add($combo)
    }

    $box = New-Object System.Windows.Forms.TextBox
    $box.Name = 'jqlBox'
    $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
    $box.Location = New-Object System.Drawing.Point(0,20); $box.Size = New-Object System.Drawing.Size(420,44)
    $box.Font = New-Object System.Drawing.Font('Consolas',8)
    $box.Text = $jqlText
    [void]$row.Controls.Add($box)

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(424,20); $btnCopy.Size = New-Object System.Drawing.Size(60,21)
    $btnCopy.FlatStyle = 'Flat'
    $btnCopy.Add_Click({
        $tpl = [string]$this.Parent.Controls.Find('jqlBox',$false)[0].Text
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
    [void]$row.Controls.Add($btnCopy)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = 'Delete'; $btnDel.Location = New-Object System.Drawing.Point(424,44); $btnDel.Size = New-Object System.Drawing.Size(60,20)
    $btnDel.FlatStyle = 'Flat'; $btnDel.ForeColor = [System.Drawing.Color]::FromArgb(255,180,40,40)
    $btnDel.Add_Click({
        $this.Parent.Parent.Controls.Remove($this.Parent)
    })
    [void]$row.Controls.Add($btnDel)

    [void]$panelJqlFlow.Controls.Add($row)
    return $row
}

$script:failures = @()
function Check($name, $cond) {
    if ($cond) { Write-Host "  OK $name" } else { $script:failures += $name; Write-Host "  FAIL $name" }
}

function Get-Copy($row) { ($row.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Text -eq 'Copy' })[0] }
function Get-Delete($row) { ($row.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Text -eq 'Delete' })[0] }

# --- fixed-shaped row (titleEditable=$false, no mode), e.g. other_qa ---
$fixedRow = New-JqlRow 'other_qa' 'Other QA activities (auto-named by keyword rules)' 'project = {project} AND creator = {account_id}' $false
Check "fixed row has no mode combo" ($fixedRow.Controls.Find('modeCombo',$false).Count -eq 0)
[System.Windows.Forms.Clipboard]::SetText('')
(Get-Copy $fixedRow).PerformClick()
Check "fixed row Copy substitutes project/account" ([System.Windows.Forms.Clipboard]::GetText() -eq 'project = GT2 AND creator = ACC123')

# --- custom-shaped row (titleEditable=$true, no mode) ---
$customRow = New-JqlRow 'extra_ab12cd34' 'My query' 'project = {project} custom' $true
Check "custom row has no mode combo" ($customRow.Controls.Find('modeCombo',$false).Count -eq 0)
[System.Windows.Forms.Clipboard]::SetText('')
(Get-Copy $customRow).PerformClick()
Check "custom row Copy substitutes project" ([System.Windows.Forms.Clipboard]::GetText() -eq 'project = GT2 custom')

# --- templatable row (titleEditable=$true, modeKey set), e.g. bug_verification ---
$tplRow = New-JqlRow 'bug_verification' 'Bug verification {buckets}' 'project = {project} custom' $true 'priority'
$combo = $tplRow.Controls.Find('modeCombo',$false)[0]
Check "templatable row has mode combo pre-selected" ([string]$combo.SelectedItem -eq 'Priority breakdown (P1-P3)')
Check "templatable row title tooltip matches priority hint" ($jqlTips.GetToolTip($tplRow.Controls.Find('titleBox',$false)[0]) -eq 'Placeholders: {P1} {P2} {P3} {count} {buckets}')
$combo.SelectedItem = 'Total count'
Check "changing mode updates title tooltip" ($jqlTips.GetToolTip($tplRow.Controls.Find('titleBox',$false)[0]) -eq 'Placeholders: {count}')
[System.Windows.Forms.Clipboard]::SetText('')
(Get-Copy $tplRow).PerformClick()
Check "templatable row Copy still substitutes JQL box, not the template text" ([System.Windows.Forms.Clipboard]::GetText() -eq 'project = GT2 custom')

# --- Delete removes the row from the panel, for all three shapes ---
foreach ($r in @($fixedRow, $customRow, $tplRow)) {
    $countBefore = $panelJqlFlow.Controls.Count
    (Get-Delete $r).PerformClick()
    Check "Delete removes row '$($r.Tag)' from the panel" (
        $panelJqlFlow.Controls.Count -eq ($countBefore - 1) -and -not $panelJqlFlow.Controls.Contains($r)
    )
}

if ($script:failures.Count -gt 0) {
    Write-Host "FAILURES: $($script:failures -join ', ')"
    exit 1
} else {
    Write-Host "ALL PASS"
    exit 0
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `powershell -File scripts/test_jql_row_gui.ps1`
Expected: FAIL — this test file's `New-JqlRow` is the *new* shape (it's a self-contained duplicate, so it actually can't "fail to compile"); to genuinely verify the test is meaningful, temporarily confirm the old `jira-sync.ps1` (not yet touched) still has the pre-Task-3 `New-JqlRow` with positional-only Controls — this test file doesn't touch `jira-sync.ps1`, so there's nothing to regress yet. Instead, sanity-check by running it now: it should already print `ALL PASS`, since this test file is self-contained. This is expected — the "failing" state for this task is `jira-sync.ps1` not yet matching this shape (Step 3 brings it in sync). Proceed to Step 3.

- [ ] **Step 3: Update `jira-sync.ps1` to match**

Replace `jira-sync.ps1:397-451` (the current `New-JqlRow` function) with:

```powershell
    function New-JqlRow($key, $titleText, $jqlText, $titleEditable, $modeKey = $null) {
        $row = New-Object System.Windows.Forms.Panel
        $row.Size = New-Object System.Drawing.Size(566,74)
        $row.Margin = New-Object System.Windows.Forms.Padding(4,4,4,0)
        $row.Tag = $key

        if ($titleEditable) {
            $titleCtl = New-Object System.Windows.Forms.TextBox
        } else {
            $titleCtl = New-Object System.Windows.Forms.Label
        }
        $titleCtl.Name = 'titleBox'
        $titleCtl.Text = $titleText; $titleCtl.Location = New-Object System.Drawing.Point(0,0)
        $titleWidth = if ($modeKey) { 260 } else { 420 }
        $titleCtl.Size = New-Object System.Drawing.Size($titleWidth,18)
        $titleCtl.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
        [void]$row.Controls.Add($titleCtl)

        if ($modeKey) {
            $combo = New-Object System.Windows.Forms.ComboBox
            $combo.Name = 'modeCombo'
            $combo.DropDownStyle = 'DropDownList'
            $combo.Location = New-Object System.Drawing.Point(264,0)
            $combo.Size = New-Object System.Drawing.Size(156,20)
            $combo.Font = New-Object System.Drawing.Font('Segoe UI',8)
            [void]$combo.Items.AddRange(@('Total count','Priority breakdown (P1-P3)','One row per issue'))
            $combo.SelectedItem = $modeDisplay[$modeKey]
            $jqlTips.SetToolTip($titleCtl, $modeHint[$modeKey])
            $jqlTips.SetToolTip($combo, [string]$combo.SelectedItem)
            $combo.Add_SelectedIndexChanged({
                $newMode = $modeValue[[string]$this.SelectedItem]
                $jqlTips.SetToolTip($this.Parent.Controls.Find('titleBox',$false)[0], $modeHint[$newMode])
                $jqlTips.SetToolTip($this, [string]$this.SelectedItem)
            })
            [void]$row.Controls.Add($combo)
        }

        $box = New-Object System.Windows.Forms.TextBox
        $box.Name = 'jqlBox'
        $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
        $box.Location = New-Object System.Drawing.Point(0,20); $box.Size = New-Object System.Drawing.Size(420,44)
        $box.Font = New-Object System.Drawing.Font('Consolas',8)
        $box.Text = $jqlText
        [void]$row.Controls.Add($box)

        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(424,20); $btnCopy.Size = New-Object System.Drawing.Size(60,21)
        $btnCopy.FlatStyle = 'Flat'
        $btnCopy.Add_Click({
            $tpl = [string]$this.Parent.Controls.Find('jqlBox',$false)[0].Text
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
        [void]$row.Controls.Add($btnCopy)

        $btnDel = New-Object System.Windows.Forms.Button
        $btnDel.Text = 'Delete'; $btnDel.Location = New-Object System.Drawing.Point(424,44); $btnDel.Size = New-Object System.Drawing.Size(60,20)
        $btnDel.FlatStyle = 'Flat'; $btnDel.ForeColor = [System.Drawing.Color]::FromArgb(255,180,40,40)
        $btnDel.Add_Click({
            $this.Parent.Parent.Controls.Remove($this.Parent)
        })
        [void]$row.Controls.Add($btnDel)

        [void]$panelJqlFlow.Controls.Add($row)
        return $row
    }
```

Note: this references `$modeDisplay`, `$modeHint`, `$modeValue`, and `$jqlTips` from the enclosing `Show-Settings` scope — these are added in Task 4, which must land before `jira-sync.ps1` can actually open the Settings dialog without erroring. This task alone leaves `jira-sync.ps1` non-functional for `Show-Settings` until Task 4 completes; that's expected and resolved by the next task's commit.

- [ ] **Step 4: Run test to verify it passes**

Run: `powershell -File scripts/test_jql_row_gui.ps1`
Expected: all `OK` lines then `ALL PASS`, exit code 0

- [ ] **Step 5: Commit**

```bash
git add jira-sync.ps1 scripts/test_jql_row_gui.ps1
git commit -m "feat: New-JqlRow supports optional mode dropdown, name-based control lookup"
```

---

### Task 4: Wire mode/template load-save into `Show-Settings`, describe `other_qa`

**Files:**
- Modify: `jira-sync.ps1:366-388` (jqlSlots/defaultJql/jqlValues setup — add mode maps, name-rule defaults/overlay, tooltip, other_qa description)
- Modify: `jira-sync.ps1:458-471` (the row-creation loop over `$jqlSlots`)
- Modify: `jira-sync.ps1:606-619` (`$script:SaveJqlConfig`)

**Interfaces:**
- Consumes: `New-JqlRow` with 5th `$modeKey` param (Task 3), `Read-JsonConfig`/`Write-JsonConfig` (existing, jira-sync.ps1:38-50)
- Produces: `config/jql_queries.json` gains a `name_rules` key on save; Settings dialog now shows editable naming templates + mode dropdowns for 5 queries, and a descriptive (non-editable) label for `other_qa`.

- [ ] **Step 1: Replace the jqlSlots/defaultJql/jqlValues setup block**

Replace `jira-sync.ps1:366-388`:

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
```

with:

```powershell
    $jqlSlots = @(
        @{ key='investigation' },
        @{ key='bug_verification' },
        @{ key='story_creation' },
        @{ key='functional_testing' },
        @{ key='regression_testing' },
        @{ key='other_qa' }
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

    $modeDisplay = @{ count='Total count'; priority='Priority breakdown (P1-P3)'; per_issue='One row per issue' }
    $modeValue   = @{ 'Total count'='count'; 'Priority breakdown (P1-P3)'='priority'; 'One row per issue'='per_issue' }
    $modeHint = @{
        count     = 'Placeholders: {count}'
        priority  = 'Placeholders: {P1} {P2} {P3} {count} {buckets}'
        per_issue = 'Placeholders: {key} {summary} {number}'
    }

    $defaultNameRules = @{
        investigation      = @{ mode='count';     template='Investigation issue {count}' }
        bug_verification   = @{ mode='priority';  template='Bug verification {buckets}' }
        story_creation     = @{ mode='count';     template='User story creation {count}' }
        functional_testing = @{ mode='per_issue'; template='Functional testing of story ID {key}' }
        regression_testing = @{ mode='per_issue'; template='Regression testing {number} test items' }
    }
    $nameRuleValues = @{}
    foreach ($k in $defaultNameRules.Keys) { $nameRuleValues[$k] = @{ mode=$defaultNameRules[$k].mode; template=$defaultNameRules[$k].template } }
    if ($jqlCfg -and $jqlCfg.PSObject.Properties.Match('name_rules').Count -and $jqlCfg.name_rules) {
        foreach ($prop in $jqlCfg.name_rules.PSObject.Properties) {
            $k = $prop.Name
            if (-not $nameRuleValues.ContainsKey($k)) { continue }
            $cfg = $prop.Value
            $m = [string]$cfg.mode
            if ($modeDisplay.ContainsKey($m)) { $nameRuleValues[$k].mode = $m }
            if ($cfg.template) { $nameRuleValues[$k].template = [string]$cfg.template }
        }
    }

    $otherQaLabel = 'Other QA activities (auto-named by keyword rules)'
    $otherQaHint = 'Automation test creation/update/maintenance, Checklist creation/update, Backlog refinement, Debugging, Maintenance, Other project documentation work, or [REVIEW] fallback'

    $jqlTips = New-Object System.Windows.Forms.ToolTip
```

- [ ] **Step 2: Replace the row-creation loop**

Replace `jira-sync.ps1:458-461` (originally, now shifted down by the block inserted in Step 1 — locate by the surrounding `foreach ($slot in $jqlSlots)` loop that calls `New-JqlRow`):

```powershell
    foreach ($slot in $jqlSlots) {
        if ($hiddenBuiltin.ContainsKey($slot.key)) { continue }
        New-JqlRow $slot.key $slot.label ([string]$jqlValues[$slot.key]) $false
    }
```

with:

```powershell
    foreach ($slot in $jqlSlots) {
        if ($hiddenBuiltin.ContainsKey($slot.key)) { continue }
        if ($slot.key -eq 'other_qa') {
            $r = New-JqlRow $slot.key $otherQaLabel ([string]$jqlValues[$slot.key]) $false
            $lbl = $r.Controls.Find('titleBox',$false)[0]
            $lbl.AutoEllipsis = $true
            $jqlTips.SetToolTip($lbl, $otherQaHint)
        } else {
            $rule = $nameRuleValues[$slot.key]
            New-JqlRow $slot.key $rule.template ([string]$jqlValues[$slot.key]) $true $rule.mode
        }
    }
```

- [ ] **Step 3: Replace `$script:SaveJqlConfig`**

Replace `jira-sync.ps1:606-619`:

```powershell
    $script:SaveJqlConfig = {
        $queries = @()
        $fixedKeys = @($jqlSlots | ForEach-Object { $_.key })
        $presentFixed = @{}
        foreach ($row in $panelJqlFlow.Controls) {
            $title = [string]$row.Controls[0].Text
            $jql   = [string]$row.Controls[1].Text
            if ($fixedKeys -contains $row.Tag) { $presentFixed[$row.Tag] = $true }
            if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($jql)) { continue }
            $queries += @{ key=$row.Tag; label=$title; jql=$jql }
        }
        $hidden = @($fixedKeys | Where-Object { -not $presentFixed.ContainsKey($_) })
        Write-JsonConfig 'jql_queries.json' @{ queries=$queries; hidden_builtin=$hidden }
    }
```

with:

```powershell
    $script:SaveJqlConfig = {
        $queries = @()
        $nameRules = @{}
        $fixedKeys = @($jqlSlots | ForEach-Object { $_.key })
        $templatableKeys = @($fixedKeys | Where-Object { $_ -ne 'other_qa' })
        $presentFixed = @{}
        foreach ($row in $panelJqlFlow.Controls) {
            $titleCtl = $row.Controls.Find('titleBox',$false)[0]
            $jqlCtl   = $row.Controls.Find('jqlBox',$false)[0]
            $title = [string]$titleCtl.Text
            $jql   = [string]$jqlCtl.Text
            if ($fixedKeys -contains $row.Tag) { $presentFixed[$row.Tag] = $true }
            if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($jql)) { continue }
            $queries += @{ key=$row.Tag; label=$title; jql=$jql }
            if ($templatableKeys -contains $row.Tag) {
                $comboMatches = $row.Controls.Find('modeCombo',$false)
                if ($comboMatches.Count -gt 0) {
                    $mode = $modeValue[[string]$comboMatches[0].SelectedItem]
                    $nameRules[$row.Tag] = @{ mode=$mode; template=$title }
                }
            }
        }
        $hidden = @($fixedKeys | Where-Object { -not $presentFixed.ContainsKey($_) })
        Write-JsonConfig 'jql_queries.json' @{ queries=$queries; hidden_builtin=$hidden; name_rules=$nameRules }
    }
```

- [ ] **Step 4: Static syntax check**

Run: `powershell -Command "$null = [System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1', [ref]$null, [ref]$errors); if ($errors) { $errors } else { 'NO PARSE ERRORS' }"`
Expected: `NO PARSE ERRORS`

- [ ] **Step 5: Re-run the Task 3 headless test (still standalone, but confirms no regressions)**

Run: `powershell -File scripts/test_jql_row_gui.ps1`
Expected: `ALL PASS`

- [ ] **Step 6: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: load/save per-query naming mode+template in Settings, describe other_qa rules"
```

---

### Task 5: Links tab (Excel path + QAE rules link)

**Files:**
- Modify: `jira-sync.ps1:353-364` (tab creation/registration)
- Modify: `jira-sync.ps1:663-720` (move Excel row out of `tabConn`; add Links tab content)

**Interfaces:**
- Consumes: existing `$tExcel` variable (read by `Test Connection` and `Save & Close` — jira-sync.ps1:697, 728, 731), unchanged.
- Produces: new `$tabLinks` TabPage; no new functions.

- [ ] **Step 1: Register the new tab**

Replace `jira-sync.ps1:353-364`:

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

with:

```powershell
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(8,8)
    $tabs.Size = New-Object System.Drawing.Size(596,430)
    [void]$dlg.Controls.Add($tabs)

    $tabConn  = New-Object System.Windows.Forms.TabPage; $tabConn.Text  = 'Connection'
    $tabStd   = New-Object System.Windows.Forms.TabPage; $tabStd.Text   = 'Standard tasks'
    $tabJql   = New-Object System.Windows.Forms.TabPage; $tabJql.Text   = 'Jira queries'
    $tabLinks = New-Object System.Windows.Forms.TabPage; $tabLinks.Text = 'Links'
    $tabConn.BackColor  = [System.Drawing.Color]::White
    $tabStd.BackColor   = [System.Drawing.Color]::White
    $tabJql.BackColor   = [System.Drawing.Color]::White
    $tabLinks.BackColor = [System.Drawing.Color]::White
    $tabs.TabPages.AddRange(@($tabConn, $tabStd, $tabJql, $tabLinks))
```

- [ ] **Step 2: Move the Excel-file row into the Links tab, add the QAE rules link**

Replace `jira-sync.ps1:663-682`:

```powershell
    # ---- Excel file row ----
    $lblEx = New-Object System.Windows.Forms.Label
    $lblEx.Text = 'Excel file:'; $lblEx.Location = New-Object System.Drawing.Point(16,237)
    $lblEx.Size = New-Object System.Drawing.Size(90,20); $lblEx.TextAlign = 'MiddleRight'
    [void]$tabConn.Controls.Add($lblEx)
    $tExcel = New-Object System.Windows.Forms.TextBox
    $tExcel.Location = New-Object System.Drawing.Point(114,234)
    $tExcel.Size = New-Object System.Drawing.Size(354,24)
    $tExcel.Text = $txtFile.Text
    [void]$tabConn.Controls.Add($tExcel)
    $btnExBrowse = New-Object System.Windows.Forms.Button
    $btnExBrowse.Text = '...'; $btnExBrowse.Location = New-Object System.Drawing.Point(474,234)
    $btnExBrowse.Size = New-Object System.Drawing.Size(30,24); $btnExBrowse.FlatStyle = 'Flat'
    $btnExBrowse.Add_Click({
        $fd = New-Object System.Windows.Forms.OpenFileDialog
        $fd.Filter = 'Excel files (*.xlsx)|*.xlsx'
        $fd.InitialDirectory = Split-Path $tExcel.Text -Parent
        if ($fd.ShowDialog() -eq 'OK') { $tExcel.Text = $fd.FileName }
    })
    [void]$tabConn.Controls.Add($btnExBrowse)
```

with:

```powershell
    # ---- Excel file row (Links tab) ----
    $lblEx = New-Object System.Windows.Forms.Label
    $lblEx.Text = 'Excel file:'; $lblEx.Location = New-Object System.Drawing.Point(16,19)
    $lblEx.Size = New-Object System.Drawing.Size(90,20); $lblEx.TextAlign = 'MiddleRight'
    [void]$tabLinks.Controls.Add($lblEx)
    $tExcel = New-Object System.Windows.Forms.TextBox
    $tExcel.Location = New-Object System.Drawing.Point(114,16)
    $tExcel.Size = New-Object System.Drawing.Size(354,24)
    $tExcel.Text = $txtFile.Text
    [void]$tabLinks.Controls.Add($tExcel)
    $btnExBrowse = New-Object System.Windows.Forms.Button
    $btnExBrowse.Text = '...'; $btnExBrowse.Location = New-Object System.Drawing.Point(474,16)
    $btnExBrowse.Size = New-Object System.Drawing.Size(30,24); $btnExBrowse.FlatStyle = 'Flat'
    $btnExBrowse.Add_Click({
        $fd = New-Object System.Windows.Forms.OpenFileDialog
        $fd.Filter = 'Excel files (*.xlsx)|*.xlsx'
        $fd.InitialDirectory = Split-Path $tExcel.Text -Parent
        if ($fd.ShowDialog() -eq 'OK') { $tExcel.Text = $fd.FileName }
    })
    [void]$tabLinks.Controls.Add($btnExBrowse)

    # ---- QAE reporting rules link ----
    $lnkQae = New-Object System.Windows.Forms.LinkLabel
    $lnkQae.Text = 'QAE Reporting Rules (SharePoint)'
    $lnkQae.Location = New-Object System.Drawing.Point(16,56)
    $lnkQae.Size = New-Object System.Drawing.Size(320,20)
    $lnkQae.Add_LinkClicked({
        Start-Process 'https://sitrusllc.sharepoint.com/sites/amcwiki/CompanyRulesandPolicies/Pages/QADPOReportingRules.aspx'
    })
    [void]$tabLinks.Controls.Add($lnkQae)
```

- [ ] **Step 3: Tidy the now-empty gap in the Connection tab**

The `Test Connection` button and its status label sat below the old Excel row; move them up to close the gap. Replace `jira-sync.ps1:684-693` (the `$lblTest`/`$btnT` block — locate by `Text = 'Test Connection'`):

```powershell
    $lblTest = New-Object System.Windows.Forms.Label
    $lblTest.Location = New-Object System.Drawing.Point(16,272)
    $lblTest.Size = New-Object System.Drawing.Size(460,20)
    $lblTest.ForeColor = [System.Drawing.Color]::Gray
    [void]$tabConn.Controls.Add($lblTest)

    $btnT = New-Object System.Windows.Forms.Button
    $btnT.Text = 'Test Connection'
    $btnT.Location = New-Object System.Drawing.Point(16,296)
    $btnT.Size = New-Object System.Drawing.Size(140,32)
```

with:

```powershell
    $lblTest = New-Object System.Windows.Forms.Label
    $lblTest.Location = New-Object System.Drawing.Point(16,237)
    $lblTest.Size = New-Object System.Drawing.Size(460,20)
    $lblTest.ForeColor = [System.Drawing.Color]::Gray
    [void]$tabConn.Controls.Add($lblTest)

    $btnT = New-Object System.Windows.Forms.Button
    $btnT.Text = 'Test Connection'
    $btnT.Location = New-Object System.Drawing.Point(16,261)
    $btnT.Size = New-Object System.Drawing.Size(140,32)
```

- [ ] **Step 4: Static syntax check**

Run: `powershell -Command "$null = [System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1', [ref]$null, [ref]$errors); if ($errors) { $errors } else { 'NO PARSE ERRORS' }"`
Expected: `NO PARSE ERRORS`

- [ ] **Step 5: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: add Settings Links tab (Excel path + QAE reporting rules link)"
```

---

### Task 6: End-to-end manual verification

**Files:** none (verification only)

- [ ] **Step 1: Run the full automated test suite**

Run:
```bash
python scripts/test_name_rules.py
python scripts/test_generate_week_rows_names.py
python scripts/test_jql_config.py
powershell -File scripts/test_jql_row_gui.ps1
```
Expected: all four print `ALL PASS` (PowerShell script also exits 0).

- [ ] **Step 2: Launch the GUI and open Settings**

Run: `start.bat` (or `powershell -File jira-sync.ps1` directly)

- [ ] **Step 3: Verify the Links tab**
  - Open Settings → "Links" tab exists, positioned after "Jira queries".
  - Excel file path shows the current file, Browse (`...`) still opens a file picker and updates the path.
  - "QAE Reporting Rules (SharePoint)" link opens the SharePoint page in the default browser.
  - Switch to "Connection" tab: confirm the Excel row is gone and "Test Connection" still works (uses the path set in Links).

- [ ] **Step 4: Verify the Jira queries tab**
  - The 5 non-`other_qa` rows show an editable template textbox (not a static ordinal label) plus a mode dropdown pre-selected to each query's current mode (e.g. `bug_verification` → "Priority breakdown (P1-P3)").
  - Hovering the template textbox shows a tooltip listing that mode's placeholders; changing the dropdown updates the tooltip.
  - `other_qa`'s row shows the static description text (not editable) and its own tooltip with the full rule list.
  - Edit `investigation`'s template to `Investigation issue found: {count}`, click **Save & Close**, reopen Settings — the edited template and its mode persist.
  - Inspect `config/jql_queries.json` — confirm a new `name_rules` key exists with the edited entry, and the pre-existing `queries[].label` entries are untouched.

- [ ] **Step 5: Confirm no regression in a real (or dry-run) sync**

If a test Jira project/account is available, run `npm start`-equivalent for the Python sync (or `python scripts/jira-sync.py --command test --file <path>`) and confirm row names in the log match the configured templates. If no live Jira is available for this session, rely on Task 2's automated test coverage and note in the PR/commit description that live verification is pending.

- [ ] **Step 6: Final commit (only if Step 4 manual edits should NOT be kept as a permanent config change)**

If you edited `config/jql_queries.json` during manual verification and want to revert it to a clean state before finishing:

```bash
git status
git checkout -- config/jql_queries.json   # only if it's tracked and you want to discard the manual test edit
```

(`config/` may be gitignored — check `git status` output first; if it's untracked/ignored, no action needed.)

---

## Self-Review Notes

- **Spec coverage:** Links tab (Part 1) → Task 5. Naming engine + defaults table (Part 2) → Tasks 1–2. Settings UI (mode dropdown + template textbox + tooltips + other_qa description) → Tasks 3–4. Migration safety (`name_rules` independent of `label`) → Task 1's `load_name_rules` reads a different JSON key entirely; verified by Task 1 Step 1 test #1 (missing `name_rules` → defaults, ignoring the pre-existing `label` fields already on disk).
- **Placeholder scan:** no TBD/TODO; every step has literal code and exact commands.
- **Type/name consistency:** `New-JqlRow($key, $titleText, $jqlText, $titleEditable, $modeKey = $null)` signature is identical across Task 3's test file and its `jira-sync.ps1` counterpart; `build_names(mode, issues, template, default_template, key=None)` signature matches between Task 1's definition and Task 2's `names_for` call site; `DEFAULT_NAME_RULES` keys (`investigation`, `bug_verification`, `story_creation`, `functional_testing`, `regression_testing`) match `$defaultNameRules` keys in the PowerShell side exactly.
