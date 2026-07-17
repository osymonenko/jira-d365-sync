#!/usr/bin/env python3
"""
Jira -> Excel timesheet sync.
Reads weeks from Sheet1, runs 6 JQL queries per week, inserts result rows
before the 'check sum' row. Skips rows that already exist (exact match).
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

try:
    import openpyxl
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
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                env[k.strip()] = v.strip()
    for key in ("JIRA_URL", "JIRA_EMAIL", "JIRA_API_TOKEN", "JIRA_PROJECT"):
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


def get_account_id(base_url: str, email: str, token: str) -> str:
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


def generate_week_rows(
    base_url: str, email: str, token: str, account_id: str,
    project: str, week_start: date, week_end: date,
) -> list[tuple[str, str]]:
    ws, we = fmt(week_start), fmt(week_end)
    rows: list[tuple[str, str]] = []

    def search_url(jql: str) -> str:
        return f"{base_url}/issues/?jql={urllib.parse.quote(jql)}"

    def issue_url(key: str) -> str:
        return f"{base_url}/browse/{key}"

    # 1. Investigation issue — total count of bugs created this week
    jql = (f'project = {project} AND issuetype = Bug '
           f'AND created >= "{ws}" AND created <= "{we}" '
           f'ORDER BY priority DESC, issuetype ASC, key ASC')
    issues = search_jira(base_url, email, token, jql)
    if issues:
        name = f"Investigation issue {len(issues)}"
        print(f"  [1/6] Investigation issues (bugs created this week): {len(issues)} found → \"{name}\"", flush=True)
        rows.append((name, search_url(jql)))
    else:
        print(f"  [1/6] Investigation issues (bugs created this week): 0 found", flush=True)

    # 2. Bug verification — count by priority bucket
    jql = (f'project = {project} AND issuetype = Bug '
           f'AND status CHANGED TO "Done" BY {account_id} DURING ("{ws}","{we}") '
           f'ORDER BY priority DESC, issuetype ASC, key ASC')
    issues = search_jira(base_url, email, token, jql)
    if issues:
        buckets: dict[str, int] = {"P1": 0, "P2": 0, "P3": 0}
        for iss in issues:
            p = iss["fields"].get("priority", {}).get("name", "Medium")
            buckets[priority_bucket(p)] += 1
        parts = [f"{k}-{v}" for k, v in buckets.items() if v > 0]
        name = f"Bug verification {', '.join(parts)}"
        print(f"  [2/6] Bug verification (closed by you): {len(issues)} found → \"{name}\"", flush=True)
        rows.append((name, search_url(jql)))
    else:
        print(f"  [2/6] Bug verification (closed by you): 0 found", flush=True)

    # 3. User story creation — count of stories created under GT2-80
    jql = (f'project = {project} AND issuetype = Story '
           f'AND created >= "{ws}" AND created <= "{we}" '
           f'AND creator = {account_id} AND parent = {project}-80 '
           f'ORDER BY status DESC, issuetype ASC, key ASC')
    issues = search_jira(base_url, email, token, jql)
    if issues:
        name = f"User story creation {len(issues)}"
        print(f"  [3/6] User story creation: {len(issues)} found → \"{name}\"", flush=True)
        rows.append((name, search_url(jql)))
    else:
        print(f"  [3/6] User story creation: 0 found", flush=True)

    # 4. Functional testing — one row per story key, direct issue link
    jql = (f'project = {project} AND issuetype = Story AND ('
           f'status CHANGED FROM "Ready for QA" BY {account_id} DURING ("{ws}", "{we}") OR '
           f'status CHANGED FROM "IN QA" BY {account_id} DURING ("{ws}", "{we}")) '
           f'ORDER BY key ASC')
    issues = search_jira(base_url, email, token, jql)
    print(f"  [4/6] Functional testing stories: {len(issues)} found", flush=True)
    for iss in issues:
        name = f"Functional testing of story ID {iss['key']}"
        print(f"         {iss['key']} → \"{name}\"", flush=True)
        rows.append((name, issue_url(iss['key'])))

    # 5. Regression testing — last number from summary, direct issue link
    jql = (f'project = {project} AND status CHANGED TO Done BY {account_id} DURING ("{ws}", "{we}") '
           f'AND parent = {project}-73 AND summary ~ "Regression" '
           f'ORDER BY status DESC, issuetype ASC, key ASC')
    issues = search_jira(base_url, email, token, jql)
    if len(issues) > 1:
        print(f"  [WARN] Regression: {len(issues)} items found (expected 1)", flush=True)
    print(f"  [5/6] Regression testing: {len(issues)} found", flush=True)
    for iss in issues:
        n = extract_last_number(iss["fields"]["summary"])
        name = f"Regression testing {n} test items" if n else f"Regression testing [REVIEW] {iss['key']}"
        print(f"         {iss['key']} → \"{name}\"", flush=True)
        rows.append((name, issue_url(iss['key'])))

    # 6. Other QA activities (GT2-73, non-smoke/regression/functional) — direct issue link
    jql = (f'project = {project} AND status CHANGED TO Done BY {account_id} DURING ("{ws}", "{we}") '
           f'AND parent = {project}-73 '
           f'AND summary !~ "Smoke" AND summary !~ "Regression" AND summary !~ "Functional." '
           f'ORDER BY status DESC, issuetype ASC, key ASC')
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


def update_excel(file_path: str, insertions: dict) -> int:
    wb = openpyxl.load_workbook(file_path)
    ws = wb.worksheets[0]
    weeks = find_weeks(ws)
    total = 0
    link_font = openpyxl.styles.Font(color="0563C1", underline="single")

    # Process bottom-up so earlier row indices stay valid
    for week in reversed(weeks):
        key = fmt(week["week_start"])
        if key not in insertions:
            continue
        new_rows = [(name, url) for name, url in insertions[key]
                    if name not in week["existing_names"]]
        if not new_rows:
            print(f"[SKIP] {key}: all rows already present", flush=True)
            continue
        # Prefer the check sum row as the anchor (insert above it). Weeks with
        # no check sum row fall back to the block end (start of the next week).
        check_row = week["check_sum_row"]
        anchor = check_row if check_row is not None else week["end_row"]
        _insert_rows_preserving_hyperlinks(ws, anchor, len(new_rows))
        for i, (name, url) in enumerate(new_rows):
            cell = ws.cell(row=anchor + i, column=3)
            cell.value = name
            cell.hyperlink = url
            cell.font = link_font
        total += len(new_rows)
        note = "" if check_row is not None else " (no check sum row; appended to block end)"
        print(f"[OK]   {key}: {len(new_rows)} rows inserted{note}", flush=True)

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
        account_id = get_account_id(base_url, email, token)
        print(f"[OK]   Authenticated. Account ID: {account_id}", flush=True)
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
        account_id = get_account_id(base_url, email, token)
        print(f"[INFO] Account ID: {account_id}", flush=True)
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


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--command", choices=["test", "read-weeks", "sync"], default="sync")
    parser.add_argument("--file")
    parser.add_argument("--weeks", nargs="*", help="Monday dates YYYY-MM-DD (space separated)")
    parser.add_argument("--month", help="Clamp week windows to this month (YYYY-MM)")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    if args.command == "test":
        cmd_test(args)
    elif args.command == "read-weeks":
        cmd_read_weeks(args)
    else:
        if not args.file:
            print("[ERROR] --file required for sync", flush=True)
            sys.exit(1)
        cmd_sync(args)


if __name__ == "__main__":
    main()
