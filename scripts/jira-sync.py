#!/usr/bin/env python3
"""
Jira -> Excel timesheet sync.
Reads weeks from Sheet1, runs 6 JQL queries per week, inserts result rows
before the 'check sum' row. Skips rows that already exist (case-insensitive
name match). Also hosts the fill-standard command for recurring tasks.
"""

import argparse
import base64
import json
import os
import re
import sys
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

from month_filter import clamp_week_to_month, month_bounds
from standard_tasks import rows_for_week, DAY_COL, load_schedule, parse_sprint_length_weeks

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


# Reverse of DAY_COL (column index -> day name) for human-readable logging.
_COL_DAY = {col: day for day, col in DAY_COL.items()}


def _describe_hours(hours_by_col: dict) -> str:
    """Human-readable day/hours summary for one task row, e.g. 'Tue,Wed 1.0h'.
    Empty (QA placeholder) -> 'no hours (placeholder)'."""
    if not hours_by_col:
        return "placeholder, no hours"
    parts = []
    for col in sorted(hours_by_col):
        parts.append(f"{_COL_DAY.get(col, f'col{col}')} {hours_by_col[col]}h")
    return ", ".join(parts)

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.formatting.rule import CellIsRule
except ImportError:
    print("[ERROR] openpyxl not installed. Run: pip install openpyxl", flush=True)
    sys.exit(1)

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')


# ---------------------------------------------------------------------------
# Config / auth
# ---------------------------------------------------------------------------

def load_env(root: Path) -> dict:
    env = {}
    env_file = root / ".env"
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                env[k.strip()] = v.strip()
    for key in ("JIRA_URL", "JIRA_EMAIL", "JIRA_API_TOKEN", "JIRA_PROJECT", "JIRA_ACCOUNT_ID"):
        if key in os.environ:
            env[key] = os.environ[key]
    return env


def jira_get(url: str, email: str, token: str) -> dict:
    auth = base64.b64encode(f"{email}:{token}".encode()).decode()
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Basic {auth}", "Accept": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())


def get_account_id(base_url: str, email: str, token: str, override: str = "") -> str:
    # If JIRA_ACCOUNT_ID is set in settings, use it verbatim — this is the user
    # whose activity the JQL queries filter on (BY/creator = <accountId>). Leave it
    # empty to fall back to the token owner via /myself.
    if override.strip():
        return override.strip()
    return jira_get(f"{base_url}/rest/api/3/myself", email, token)["accountId"]


def search_jira(base_url: str, email: str, token: str, jql: str) -> list:
    auth = base64.b64encode(f"{email}:{token}".encode()).decode()
    url = f"{base_url}/rest/api/3/search/jql"
    payload = json.dumps({
        "jql": jql,
        "fields": ["key", "summary", "priority"],
        "maxResults": 100,
    }).encode()
    req = urllib.request.Request(
        url,
        data=payload,
        headers={
            "Authorization": f"Basic {auth}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode()).get("issues", [])


# ---------------------------------------------------------------------------
# Row generators
# ---------------------------------------------------------------------------

def fmt(d: date) -> str:
    return d.strftime("%Y-%m-%d")


def priority_bucket(name: str) -> str:
    n = name.lower()
    if n in ("highest", "high"):
        return "P1"
    if n == "medium":
        return "P2"
    return "P3"


def extract_last_number(text: str) -> str | None:
    """Last number from trailing parenthetical like '(16/0/0/16)' → '16'."""
    m = re.search(r"\([\d/]+\)", text)
    nums = re.findall(r"\d+", m.group() if m else text)
    return nums[-1] if nums else None


# Match order matters: more-specific patterns first
_AUTOMATION_RULES = [
    (["automat", "creat"],    "Automation test creation"),
    (["automat", "updat"],    "Automation test update"),
    (["automat", "environ"],  "Automation test environment setup"),
    (["automat"],             "Automation test maintenance"),
    (["checklist", "updat"],  "Checklist update"),
    (["checklist"],           "Checklist creation"),
    (["mainten"],             "Maintenance"),
    (["backlog", "refin"],    "Backlog refinement"),
    (["backlog", "groom"],    "Backlog grooming"),
    (["backlog"],             "Backlog refinement"),
    (["debug"],               "Debugging"),
    (["doc"],                 "Other project documentation work"),
]

_COUNT_NAMES = {
    "Automation test creation", "Automation test update",
    "Automation test maintenance", "Checklist creation",
    "Checklist update", "Maintenance",
}


def choose_automation_name(summary: str, key: str) -> str:
    s = summary.lower()
    count = extract_last_number(summary)
    for keywords, name in _AUTOMATION_RULES:
        if all(kw in s for kw in keywords):
            if name in _COUNT_NAMES and count:
                return f"{name} {count} test items"
            return name
    return f"[REVIEW] {summary[:50]} ({key})"


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


# ---------------------------------------------------------------------------
# Excel writing
# ---------------------------------------------------------------------------

def find_weeks(ws) -> list[dict]:
    weeks = []
    current = None
    for row_idx in range(1, ws.max_row + 2):
        cell_a = ws.cell(row=row_idx, column=1).value
        cell_c = ws.cell(row=row_idx, column=3).value
        if isinstance(cell_a, datetime):
            current = {
                "week_start": cell_a.date(),
                "week_end": None,
                "start_row": row_idx,
                "check_sum_row": None,
                "existing_names": set(),
            }
            cell_b = ws.cell(row=row_idx, column=2).value
            if isinstance(cell_b, datetime):
                current["week_end"] = cell_b.date()
            weeks.append(current)
        if current is None:
            continue
        if isinstance(cell_c, str) and cell_c.strip():
            name = cell_c.strip()
            if name.lower() == "check sum":
                current["check_sum_row"] = row_idx
                current = None
            else:
                current["existing_names"].add(name)

    # Insertion anchor for weeks WITHOUT a check sum row: the start of the next
    # week block (i.e. the end of this block), or one past the last row for the
    # final block. Rows inserted before this anchor land at the block's end.
    for i, week in enumerate(weeks):
        week["end_row"] = weeks[i + 1]["start_row"] if i + 1 < len(weeks) else ws.max_row + 1
    return weeks


def _insert_rows_preserving_hyperlinks(ws, idx: int, amount: int) -> None:
    """Insert `amount` blank rows at `idx`.

    openpyxl's insert_rows() shifts cell VALUES and styles but leaves
    hyperlinks anchored to their original cell coordinates, orphaning them.
    Capture the hyperlink targets at/below `idx`, clear them, insert, then
    re-apply each target at its shifted position.
    """
    moved = []  # (row, column, target)
    for row in ws.iter_rows(min_row=idx):
        for cell in row:
            if cell.hyperlink is not None:
                moved.append((cell.row, cell.column, cell.hyperlink.target))
    for r, c, _ in moved:
        ws.cell(row=r, column=c).hyperlink = None
    ws.insert_rows(idx, amount)
    for r, c, target in moved:
        ws.cell(row=r + amount, column=c).hyperlink = target


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
    _reconcile_check_sums(ws)
    try:
        wb.save(file_path)
    except PermissionError:
        print(f"[ERROR] Cannot save — close the file in Excel first: {file_path}", flush=True)
        return 0
    return total


_DAY_COLS = (4, 5, 6, 7, 8)  # Mon..Fri
_COL_LETTER = {4: "D", 5: "E", 6: "F", 7: "G", 8: "H"}

# Pastel accents for the check-sum row. openpyxl's Color pads a 6-digit RGB
# string with an "00" (fully transparent) alpha byte, not "FF" (opaque) —
# harmless for a cell's own .fill (Excel ignores alpha there) but conditional-
# formatting dxf fills honor it, so an unprefixed color is invisible in a CF
# rule. Always spell these out as 8-digit ARGB with an explicit FF alpha.
_PASTEL_BLUE = "FFBDD7EE"   # check-sum label fill
_PASTEL_GREEN = "FFC6EFCE"  # a day that totals exactly 8h
_PASTEL_RED = "FFFFC7CE"    # a day that totals under or over 8h
_BLACK = "FF000000"         # standard font color, always — only the fill changes


def _write_check_sum(ws, row: int, first_task_row: int, last_task_row: int) -> None:
    """Write a 'check sum' row: col C label (right-aligned, pastel blue) + per-day
    SUM formulas spanning the week's task rows. The day cells get the same
    pastel-blue fill as the label directly on the cell (guaranteed to render
    in any viewer, matching how the label itself already renders), plus
    conditional formatting on top for spreadsheet apps that support it (pastel
    green exactly at 8h, pastel red otherwise — overrides the blue when a
    viewer honors it). Font color always standard black. This row is ignored
    by the D365 importer (which skips any 'check sum' task name)."""
    label = ws.cell(row=row, column=3)
    label.value = "check sum"
    label.alignment = Alignment(horizontal="right")
    label.fill = PatternFill("solid", fgColor=_PASTEL_BLUE)
    for col in _DAY_COLS:
        letter = _COL_LETTER[col]
        cell = ws.cell(row=row, column=col)
        cell.value = f"=SUM({letter}{first_task_row}:{letter}{last_task_row})"
        cell.fill = PatternFill("solid", fgColor=_PASTEL_BLUE)
    # Live coloring of the day totals (recomputes as the user edits hours):
    #   = 8       -> pastel green fill
    #   < 8 or > 8 -> pastel red fill  (covers both cases)
    # Font stays standard black in both cases.
    day_range = f"{_COL_LETTER[4]}{row}:{_COL_LETTER[8]}{row}"
    try:
        del ws.conditional_formatting[day_range]  # avoid duplicate rules on re-run
    except KeyError:
        pass
    ws.conditional_formatting.add(day_range, CellIsRule(
        operator="equal", formula=["8"],
        fill=PatternFill("solid", fgColor=_PASTEL_GREEN), font=Font(color=_BLACK)))
    ws.conditional_formatting.add(day_range, CellIsRule(
        operator="notEqual", formula=["8"],
        fill=PatternFill("solid", fgColor=_PASTEL_RED), font=Font(color=_BLACK)))


_DAY_RANGE_RE = re.compile(r"^D(\d+):H\1$")


def _reconcile_check_sums(ws) -> None:
    """Rewrite every existing check-sum row's SUM formulas to match its block's
    CURRENT physical task rows.

    A check-sum formula's day-column references are baked in as plain text
    (e.g. "=SUM(G24:G31)"). Inserting rows anywhere above a block — whether by
    this same run's own overflow handling for an earlier week, or by a Jira
    sync adding rows to a week above this one — shifts the block's cells
    (including its check-sum row) down as a unit, but openpyxl's insert_rows
    never rewrites formula text, so the old absolute row numbers stay literal
    and silently start summing whatever now occupies those rows (usually the
    week above). Re-deriving every check-sum row from a fresh find_weeks() scan
    at the end of a run keeps them all honest regardless of what moved them.

    Row insertion also leaves the check-sum row's OWN conditional-formatting
    entry behind at its pre-shift range (sqref strings aren't shifted by
    insert_rows either) — so after a shift, the row that used to host a
    check-sum keeps its red/green coloring even though it's now an ordinary
    task row (or blank). Drop any "D<n>:H<n>"-shaped rule whose row isn't a
    current check-sum row before re-adding fresh ones for the rows that are.
    """
    current_check_rows = {w["check_sum_row"] for w in find_weeks(ws) if w["check_sum_row"] is not None}
    for cf in list(ws.conditional_formatting):
        sqref = str(cf.sqref)
        m = _DAY_RANGE_RE.match(sqref)
        if m and int(m.group(1)) not in current_check_rows:
            del ws.conditional_formatting[sqref]

    for week in find_weeks(ws):
        check_row = week["check_sum_row"]
        if check_row is None:
            continue
        occupied = [
            r for r in range(week["start_row"], week["end_row"])
            if r != check_row
            and ws.cell(r, 3).value not in (None, "")
            and str(ws.cell(r, 3).value).strip().lower() != "check sum"
        ]
        if not occupied:
            continue
        _write_check_sum(ws, check_row, min(occupied), max(occupied))


def insert_standard_rows(file_path: str, insertions: dict) -> int:
    """Fill standard task rows into existing week blocks.

    insertions[key] = list of {"name": str, "hours_by_col": {col: hours}}.

    Unlike the Jira path, standard rows are packed from the TOP of the block:
    the first task lands on the week's date row (column C), the rest follow on
    consecutive rows, and a 'check sum' row with per-day =SUM() formulas is
    written directly below the last task. Names already present in the block
    are skipped (case-insensitive). Processed bottom-up so that inserting extra
    rows (only when a block lacks blank space) never invalidates upper blocks.
    """
    wb = openpyxl.load_workbook(file_path)
    ws = wb.worksheets[0]
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

        start, end, check = week["start_row"], week["end_row"], week["check_sum_row"]

        # Task rows already in the block (excluding any check-sum row).
        occupied = [
            r for r in range(start, end)
            if r != check
            and ws.cell(r, 3).value not in (None, "")
            and str(ws.cell(r, 3).value).strip().lower() != "check sum"
        ]
        first_write = (max(occupied) + 1) if occupied else start  # first task on the date row
        sum_first = min(occupied) if occupied else start

        # Drop any pre-existing check-sum row so we don't leave a stale duplicate.
        if check is not None:
            for col in (3,) + _DAY_COLS:
                ws.cell(check, col).value = None

        last_task_row = first_write + len(new_rows) - 1
        check_row = last_task_row + 1

        # Make room only if the check-sum row would collide with the next block.
        overflow = check_row - (end - 1)
        if overflow > 0:
            _insert_rows_preserving_hyperlinks(ws, end, overflow)

        for i, p in enumerate(new_rows):
            r = first_write + i
            ws.cell(r, 3).value = p["name"]
            for col, hours in p["hours_by_col"].items():
                ws.cell(r, col).value = int(hours) if float(hours).is_integer() else hours

        _write_check_sum(ws, check_row, sum_first, last_task_row)
        total += len(new_rows)
        print(f"[OK]   {key}: {len(new_rows)} rows written (check sum at row {check_row})", flush=True)

    _reconcile_check_sums(ws)
    try:
        wb.save(file_path)
    except PermissionError:
        print(f"[ERROR] Cannot save — close the file in Excel first: {file_path}", flush=True)
        return 0
    return total


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def cmd_test(args):
    root = Path(args.file).resolve().parent.parent if args.file else Path.cwd()
    env = load_env(root)
    base_url = env.get("JIRA_URL", "").rstrip("/")
    email    = env.get("JIRA_EMAIL", "")
    token    = env.get("JIRA_API_TOKEN", "")
    if not all([base_url, email, token]):
        print("[ERROR] JIRA_URL, JIRA_EMAIL or JIRA_API_TOKEN missing in .env", flush=True)
        sys.exit(1)
    print(f"[INFO] Connecting to {base_url} as {email} ...", flush=True)
    try:
        account_id = get_account_id(base_url, email, token, env.get("JIRA_ACCOUNT_ID", ""))
        source = "settings" if env.get("JIRA_ACCOUNT_ID", "").strip() else "/myself"
        print(f"[OK]   Authenticated. Account ID: {account_id} (from {source})", flush=True)
        projects = jira_get(f"{base_url}/rest/api/3/project/search?maxResults=5", email, token)
        names = [p["key"] for p in projects.get("values", [])]
        print(f"[OK]   Projects accessible: {', '.join(names) or '(none)'}", flush=True)
        print("[OK]   Connection test passed.", flush=True)
    except Exception as e:
        print(f"[ERROR] {e}", flush=True)
        sys.exit(1)


def cmd_read_weeks(args):
    if not args.file:
        print("[ERROR] --file required", flush=True)
        sys.exit(1)
    wb = openpyxl.load_workbook(args.file, data_only=True)
    weeks = find_weeks(wb.worksheets[0])
    wb.close()
    output = []
    for w in weeks:
        output.append({
            "start": fmt(w["week_start"]),
            "end": fmt(w["week_end"]) if w["week_end"] else fmt(w["week_start"] + timedelta(days=6)),
            "tasks": sorted(w["existing_names"]),
        })
    print(json.dumps(output), flush=True)


def cmd_sync(args):
    root = Path(args.file).resolve().parent.parent
    env = load_env(root)
    base_url = env.get("JIRA_URL", "").rstrip("/")
    email    = env.get("JIRA_EMAIL", "")
    token    = env.get("JIRA_API_TOKEN", "")
    project  = env.get("JIRA_PROJECT", "GT2")

    if not all([base_url, email, token]):
        print("[ERROR] JIRA_URL, JIRA_EMAIL or JIRA_API_TOKEN missing in .env", flush=True)
        sys.exit(1)

    print(f"[INFO] Connecting to {base_url} as {email} ...", flush=True)
    try:
        account_id = get_account_id(base_url, email, token, env.get("JIRA_ACCOUNT_ID", ""))
        source = "settings" if env.get("JIRA_ACCOUNT_ID", "").strip() else "/myself"
        print(f"[INFO] Account ID: {account_id} (from {source})", flush=True)
    except Exception as e:
        print(f"[ERROR] Auth failed: {e}", flush=True)
        sys.exit(1)

    if args.month:
        try:
            month_bounds(args.month)
        except ValueError as e:
            print(f"[ERROR] {e}", flush=True)
            sys.exit(1)

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
            print(f"[ERROR] None of the specified weeks found in Excel", flush=True)
            sys.exit(1)

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

    if args.dry_run:
        print("\n[DRY RUN] Excel not modified.", flush=True)
        return

    if insertions:
        total = update_excel(args.file, insertions)
        print(f"\n[DONE] {total} rows inserted.", flush=True)
    else:
        print("\n[DONE] Nothing to insert.", flush=True)


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

    sprint_length_weeks, sprint_len_warning = parse_sprint_length_weeks(env.get("SPRINT_LENGTH_WEEKS", ""))
    if sprint_len_warning:
        print(sprint_len_warning, flush=True)
    cycle_days = sprint_length_weeks * 7

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

    schedule, placeholders = load_schedule()
    today = date.today()
    insertions = {}
    for week in weeks_info:
        ws_key = fmt(week["week_start"])
        we = week["week_end"] or (week["week_start"] + timedelta(days=6))
        rows = rows_for_week(week["week_start"], we, args.month, sprint_anchor, today,
                             schedule=schedule, placeholders=placeholders, cycle_days=cycle_days)
        if not rows:
            print(f"[SKIP] {ws_key}: no standard rows for this week", flush=True)
            continue
        insertions[ws_key] = rows
        print(f"[INFO] {ws_key}: {len(rows)} task(s):", flush=True)
        for r in rows:
            print(f"    • {r['name']}  ({_describe_hours(r['hours_by_col'])})", flush=True)

    if args.dry_run:
        print("\n[DRY RUN] Excel not modified.", flush=True)
        return

    if insertions:
        total = insert_standard_rows(args.file, insertions)
        print(f"\n[DONE] {total} rows inserted.", flush=True)
    else:
        print("\n[DONE] Nothing to insert.", flush=True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--command", choices=["test", "read-weeks", "sync", "fill-standard"], default="sync")
    parser.add_argument("--file")
    parser.add_argument("--weeks", nargs="*", help="Monday dates YYYY-MM-DD (space separated)")
    parser.add_argument("--month", help="Clamp week windows to this month (YYYY-MM)")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

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


if __name__ == "__main__":
    main()
