# Custom JQL Query Slots (Add/Delete) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let the user add and delete an unlimited number of custom, Copy-only JQL query slots in the Settings → Jira queries tab, alongside the six fixed automated slots, with zero changes to the Python side.

**Architecture:** Refactor the Jira-queries tab's row rendering from a manual fixed-Y-coordinate layout to a `FlowLayoutPanel` of row-panels (Task 2, pure refactor, no behavior change), then add a "+ Add query" button and per-row Delete button on top of that structure (Task 3). A one-line Python test (Task 1) pins the existing `load_jql()` contract this design depends on: unrecognized keys in `config/jql_queries.json` are already silently ignored, so custom slots never reach `generate_week_rows`.

**Tech Stack:** PowerShell WinForms (`jira-sync.ps1`), Python 3 (`scripts/jira-sync.py`, test only).

**Spec:** `docs/superpowers/specs/2026-07-18-gui-redesign-custom-jql-queries-design.md`

## Global Constraints

- Zero changes to `scripts/jira-sync.py` production code — `load_jql()` is not modified; only its test gets a new case.
- Custom slot keys are generated once, at creation, as `extra_<8 lowercase hex>` via `[guid]::NewGuid().ToString('N').Substring(0,8)`.
- The six fixed slots stay exactly as Phase 1 shipped them: not deletable, keys/order/labels unchanged, only their JQL text stays user-editable.
- A new custom row starts with a blank title and blank JQL text.
- Delete removes a custom row immediately, no confirmation dialog.
- On Save, a custom row whose title AND JQL are both blank is dropped silently (not written to `config/jql_queries.json`).
- No limit on the number of custom slots. No reordering — custom rows render in creation order, after the fixed six.
- Platform is WinForms PowerShell — no web/HTML.

---

### Task 1: Pin the `load_jql()` "unknown key ignored" contract

**Files:**
- Modify: `scripts/test_jql_config.py:29-31` (insert a 4th case before the final `print("ALL PASS")`)

**Interfaces:**
- Consumes: `js.load_jql()` (already shipped, unchanged).
- Produces: nothing new — this is a regression test, not new functionality. `load_jql()` already ignores any `queries` entry whose `key` isn't one of the six fixed keys (`templates = dict(DEFAULT_JQL)`, then `if key in templates: ...`), so this test is expected to pass immediately with zero production-code changes. It exists to make that behavior an explicit, checked contract before the GUI starts relying on it.

- [ ] **Step 1: Add the test case**

Open `scripts/test_jql_config.py`. Find the existing block:

```python
resolved = q["investigation"].format(project="GT2", account_id="ACC", ws="2026-07-12", we="2026-07-18")
assert resolved == "project = GT2 custom 2026-07-12", resolved
print("  OK template substitutes {project}/{ws}")

print("ALL PASS")
```

Replace it with (adds a 4th case before the final print):

```python
resolved = q["investigation"].format(project="GT2", account_id="ACC", ws="2026-07-12", we="2026-07-18")
assert resolved == "project = GT2 custom 2026-07-12", resolved
print("  OK template substitutes {project}/{ws}")

# 4. an entry with an unrecognized key is ignored — this is the contract the
#    Settings GUI's custom (user-added) query slots rely on: they live in the
#    same config/jql_queries.json file, keyed differently, and must never
#    reach generate_week_rows or change what the six fixed keys resolve to.
(cfgdir / "jql_queries.json").write_text(json.dumps({
    "queries": [
        {"key": "investigation", "label": "x", "jql": "project = {project} custom {ws}"},
        {"key": "extra_ab12cd34", "label": "My custom query", "jql": "project = FOO"}
    ]
}), encoding="utf-8")
q = js.load_jql()
assert set(q) == {"investigation", "bug_verification", "story_creation",
                  "functional_testing", "regression_testing", "other_qa"}, set(q)
assert q["investigation"] == "project = {project} custom {ws}", q["investigation"]
print("  OK unrecognized key ignored, six-key contract holds")

print("ALL PASS")
```

- [ ] **Step 2: Run the test — expect an immediate PASS**

Run: `python scripts/test_jql_config.py`
Expected: `ALL PASS`, including the new line `  OK unrecognized key ignored, six-key contract holds`.

This is expected to pass on the first run with no production-code changes — `load_jql()` already behaves this way. If it does NOT pass, stop: that means `load_jql()`'s behavior differs from what this whole feature assumes, and the plan's design (Task 2/3) is unsound until that's resolved — escalate rather than "fixing" the test to match different behavior.

- [ ] **Step 3: Commit**

```bash
git add scripts/test_jql_config.py
git commit -m "test: lock in load_jql ignoring unrecognized keys (GUI custom-slot contract)"
```

---

### Task 2: Refactor the Jira-queries tab onto a FlowLayoutPanel (no behavior change)

**Files:**
- Modify: `jira-sync.ps1:368-410` (the `$panelJql` block and its `foreach ($slot in $jqlSlots)` loop)

**Interfaces:**
- Consumes: `$jqlSlots`, `$jqlValues`, `$tabJql`, `$tProj`, `$tAcct`, `$script:weekCheckboxes` — all already defined earlier in `Show-Settings` (lines 337-366), untouched by this task.
- Produces: `$panelJqlFlow` (the `FlowLayoutPanel`, replaces `$panelJql`), function `New-JqlFixedRow($slot)` that builds and adds one fixed-slot row to `$panelJqlFlow`. `$script:jqlBoxes` keeps its existing contract (dictionary keyed by slot key → the JQL `TextBox`) — `$script:SaveJqlConfig` (defined later, untouched by this task) still reads it exactly as before.

This task is a pure structural refactor: after it, the six rows render and behave identically to before (same fonts, same Copy substitution, same `$script:jqlBoxes` contract) — only the container changes, from manually-positioned children of a plain `Panel` to row-panels inside a `FlowLayoutPanel`. Task 3 builds custom-row add/delete on top of this structure; a `FlowLayoutPanel` re-flows automatically when children are added/removed, which a manual Y-coordinate loop cannot do without recomputing every row below the change.

- [ ] **Step 1: Replace the block**

Find the block starting at `$panelJql = New-Object System.Windows.Forms.Panel` (jira-sync.ps1:368) and ending at the `}` that closes `foreach ($slot in $jqlSlots) { ... }` (jira-sync.ps1:410). Replace the ENTIRE block with:

```powershell
    $panelJqlFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelJqlFlow.Location = New-Object System.Drawing.Point(0,0)
    $panelJqlFlow.Dock = 'Fill'
    $panelJqlFlow.FlowDirection = 'TopDown'
    $panelJqlFlow.WrapContents = $false
    $panelJqlFlow.AutoScroll = $true
    [void]$tabJql.Controls.Add($panelJqlFlow)

    $script:jqlBoxes = @{}

    function New-JqlFixedRow($slot) {
        $row = New-Object System.Windows.Forms.Panel
        $row.Size = New-Object System.Drawing.Size(566,74)
        $row.Margin = New-Object System.Windows.Forms.Padding(4,4,4,0)
        $row.Tag = $slot.key

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $slot.label; $lbl.Location = New-Object System.Drawing.Point(0,0)
        $lbl.Size = New-Object System.Drawing.Size(400,16); $lbl.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
        [void]$row.Controls.Add($lbl)

        $box = New-Object System.Windows.Forms.TextBox
        $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
        $box.Location = New-Object System.Drawing.Point(0,18); $box.Size = New-Object System.Drawing.Size(490,46)
        $box.Font = New-Object System.Drawing.Font('Consolas',8)
        $box.Text = [string]$jqlValues[$slot.key]
        [void]$row.Controls.Add($box)
        $script:jqlBoxes[$slot.key] = $box

        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(494,18); $btnCopy.Size = New-Object System.Drawing.Size(64,46)
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
        [void]$row.Controls.Add($btnCopy)

        [void]$panelJqlFlow.Controls.Add($row)
    }

    foreach ($slot in $jqlSlots) { New-JqlFixedRow $slot }
```

Note: `$row.Tag = $slot.key` is new (the original code had no per-row tag). It doesn't change any existing behavior — it's read by Task 3's Save logic to tell fixed rows apart from custom ones by inspecting every row in `$panelJqlFlow.Controls` uniformly.

The `$btnCopy.Add_Click` body is copied verbatim from the original — same `$this.Tag` + `$script:jqlBoxes[$key]` lookup pattern (proven safe in Phase 1's review: it avoids PowerShell's shared-loop-variable closure bug). Do not "simplify" it to close over `$box`/`$slot` directly even though this is now inside a function (where that would also be safe) — keep the diff minimal and match the already-reviewed pattern.

- [ ] **Step 2: Verify parse**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

- [ ] **Step 3: Manual verification (defer to human — no live GUI in this environment)**

Note in your report that this step is pending: launch `jira-sync.ps1`, open Settings → Jira queries. Confirm all six rows render exactly as before (same labels, same JQL text, same layout), Copy still produces a fully-substituted JQL for each, and Save & Close still persists edits to `config/jql_queries.json` — i.e., confirm zero behavior change from this refactor.

- [ ] **Step 4: Commit**

```bash
git add jira-sync.ps1
git commit -m "refactor: Jira queries tab rows onto a FlowLayoutPanel (no behavior change)"
```

---

### Task 3: Add/Delete custom query slots

**Files:**
- Modify: `jira-sync.ps1` (Jira-queries tab block from Task 2, plus the existing `$script:SaveJqlConfig` at `jira-sync.ps1:491-497`)

**Interfaces:**
- Consumes: `$panelJqlFlow`, `New-JqlFixedRow` (Task 2), `$jqlSlots`, `$jqlCfg` (already parsed earlier at line 361), `$tProj`, `$tAcct`, `$script:weekCheckboxes`.
- Produces: function `New-JqlCustomRow($key, $label, $jql)` (adds one custom row to `$panelJqlFlow`, with an editable title, editable JQL, Copy, and Delete); a "+ Add query" button that calls it with a freshly generated key and blank text; extended `$script:SaveJqlConfig` that also persists surviving custom rows.

- [ ] **Step 1: Add `New-JqlCustomRow` and seed existing custom rows from config**

Immediately after the `foreach ($slot in $jqlSlots) { New-JqlFixedRow $slot }` line added in Task 2, add:

```powershell
    function New-JqlCustomRow($key, $label, $jql) {
        $row = New-Object System.Windows.Forms.Panel
        $row.Size = New-Object System.Drawing.Size(566,74)
        $row.Margin = New-Object System.Windows.Forms.Padding(4,4,4,0)
        $row.Tag = $key

        $txtTitle = New-Object System.Windows.Forms.TextBox
        $txtTitle.Text = $label; $txtTitle.Location = New-Object System.Drawing.Point(0,0)
        $txtTitle.Size = New-Object System.Drawing.Size(420,18)
        $txtTitle.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
        [void]$row.Controls.Add($txtTitle)

        $box = New-Object System.Windows.Forms.TextBox
        $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
        $box.Location = New-Object System.Drawing.Point(0,20); $box.Size = New-Object System.Drawing.Size(420,44)
        $box.Font = New-Object System.Drawing.Font('Consolas',8)
        $box.Text = $jql
        [void]$row.Controls.Add($box)

        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(424,20); $btnCopy.Size = New-Object System.Drawing.Size(60,21)
        $btnCopy.FlatStyle = 'Flat'
        $btnCopy.Add_Click({
            $tpl = [string]$box.Text
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
            $panelJqlFlow.Controls.Remove($row)
        })
        [void]$row.Controls.Add($btnDel)

        [void]$panelJqlFlow.Controls.Add($row)
    }

    if ($jqlCfg -and $jqlCfg.queries) {
        $fixedKeySet = @{}
        foreach ($s in $jqlSlots) { $fixedKeySet[$s.key] = $true }
        foreach ($q in $jqlCfg.queries) {
            if (-not $fixedKeySet.ContainsKey($q.key)) {
                New-JqlCustomRow $q.key $q.label $q.jql
            }
        }
    }
```

`New-JqlCustomRow`'s `Add_Click` handlers close directly over `$box`/`$row` (not via `$this.Tag`) — this is safe here because each call to `New-JqlCustomRow` is its own function invocation with its own local scope (unlike a bare `foreach` loop body, which shares one scope across iterations and would need the `$this.Tag` indirection Task 2 kept for the fixed rows).

- [ ] **Step 2: Add the "+ Add query" button above the row list**

Immediately after the block from Step 1, add:

```powershell
    $panelJqlAddStrip = New-Object System.Windows.Forms.Panel
    $panelJqlAddStrip.Dock = 'Top'; $panelJqlAddStrip.Height = 32
    $btnAddJql = New-Object System.Windows.Forms.Button
    $btnAddJql.Text = '+ Add query'; $btnAddJql.Location = New-Object System.Drawing.Point(0,2)
    $btnAddJql.Size = New-Object System.Drawing.Size(110,26); $btnAddJql.FlatStyle = 'Flat'
    $btnAddJql.Add_Click({
        $newKey = 'extra_' + ([guid]::NewGuid().ToString('N').Substring(0,8))
        New-JqlCustomRow $newKey '' ''
    })
    [void]$panelJqlAddStrip.Controls.Add($btnAddJql)
    [void]$tabJql.Controls.Add($panelJqlAddStrip)
```

`$panelJqlAddStrip` (`Dock = 'Top'`) is added to `$tabJql.Controls` AFTER `$panelJqlFlow` (`Dock = 'Fill'`, added in Task 2) — WinForms docks controls in reverse of the order they were added to `.Controls` (the most-recently-added docked control claims its edge first), so this ordering is required for the add-strip to sit above the row list rather than being covered by it. Do not reorder these two `.Controls.Add` calls.

- [ ] **Step 3: Extend `$script:SaveJqlConfig` to persist custom rows**

Find (jira-sync.ps1:491-497):

```powershell
    $script:SaveJqlConfig = {
        $queries = @()
        foreach ($slot in $jqlSlots) {
            $queries += @{ key=$slot.key; label=$slot.label; jql=[string]$script:jqlBoxes[$slot.key].Text }
        }
        Write-JsonConfig 'jql_queries.json' @{ queries=$queries }
    }
```

Replace with:

```powershell
    $script:SaveJqlConfig = {
        $queries = @()
        foreach ($slot in $jqlSlots) {
            $queries += @{ key=$slot.key; label=$slot.label; jql=[string]$script:jqlBoxes[$slot.key].Text }
        }
        $fixedKeys = @($jqlSlots | ForEach-Object { $_.key })
        foreach ($row in $panelJqlFlow.Controls) {
            if ($fixedKeys -contains $row.Tag) { continue }
            $title = [string]$row.Controls[0].Text
            $jql   = [string]$row.Controls[1].Text
            if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($jql)) { continue }
            $queries += @{ key=$row.Tag; label=$title; jql=$jql }
        }
        Write-JsonConfig 'jql_queries.json' @{ queries=$queries }
    }
```

`$row.Controls[0]`/`[1]` rely on `New-JqlCustomRow` adding the title `TextBox` first and the JQL `TextBox` second (Step 1 above) — do not reorder the `.Controls.Add` calls inside `New-JqlCustomRow` without updating this indexing.

- [ ] **Step 4: Verify parse**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

- [ ] **Step 5: Manual verification (defer to human — no live GUI in this environment)**

Note in your report that this step is pending: launch `jira-sync.ps1`, open Settings → Jira queries.
- Confirm the "+ Add query" button sits above the row list, not overlapping it.
- Click it: a blank row appears at the bottom (blank title, blank JQL, Copy, Delete). Type a title and a JQL using `{project}`/`{account_id}`/`{ws}`/`{we}`; Copy places the fully-substituted text on the clipboard.
- Save & Close, then reopen Settings: the custom row is still there with the typed values.
- Click Delete on it, Save & Close, reopen: it's gone.
- Confirm the six fixed rows are unaffected throughout (no Delete button on them, edits to their JQL still save/reload).

- [ ] **Step 6: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: add/delete custom Copy-only JQL query slots in Settings"
```

---

## Notes for the implementer

- Tasks 2 and 3 are WinForms layout/behavior changes with no automated test coverage — the parser check catches syntax only. Both tasks explicitly defer live-render/interaction verification to a human; report it as pending, don't attempt to launch the GUI yourself.
- Task 1 is a "lock in already-correct behavior" test, not new development — it's expected to pass on the first run. If it doesn't, stop and escalate; that means `load_jql()` doesn't do what this whole plan assumes.
- Keep `New-JqlFixedRow`'s Copy handler using `$this.Tag` + `$script:jqlBoxes[$key]` (matches the already-reviewed Phase 1 pattern) even though closing directly over locals would also be safe inside a function — minimize the diff against proven-safe code.
