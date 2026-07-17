"""Network-free test for standard-row insertion: tasks packed from the top of
the block (first task on the date row), a 'check sum' row with per-day SUM
formulas below them, and case-insensitive dedup."""

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
    # Block A: fresh week (date row 2, blank rows 3-12), no tasks, no check sum.
    ws.cell(2, 1).value = datetime(2026, 7, 12)   # Sunday
    ws.cell(2, 2).value = datetime(2026, 7, 18)
    # Block B: date row 13 with ONE pre-existing task at row 14 (dedup + append).
    ws.cell(13, 1).value = datetime(2026, 7, 19)
    ws.cell(13, 2).value = datetime(2026, 7, 25)
    ws.cell(14, 3).value = "External customer meeting"
    wb.save(path)


def main():
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    path = tmp.name
    build_workbook(path)

    insertions = {
        "2026-07-12": [
            {"name": "Internal Daily meeting", "hours_by_col": {4: 0.5, 5: 0.5, 6: 0.5, 7: 0.5, 8: 0.5}},
            {"name": "Weekly project report", "hours_by_col": {8: 1.0}},
            {"name": "Bug verification", "hours_by_col": {}},  # placeholder, no hours
        ],
        "2026-07-19": [
            {"name": "external customer meeting", "hours_by_col": {5: 1.0, 6: 1.0}},  # dup (case) -> skip
            {"name": "Internal bug triage", "hours_by_col": {5: 0.5}},
        ],
    }
    total = jira_sync.insert_standard_rows(path, insertions)
    assert total == 4, f"expected 4 written (1 dup skipped), got {total}"
    print("  OK 4 rows written, case-different duplicate skipped")

    wb = openpyxl.load_workbook(path)  # NOT data_only -> formulas readable as text
    ws = wb.worksheets[0]

    def c(r, col):
        return ws.cell(r, col).value

    # --- Block A: first task lands ON the date row (row 2), packed downward ---
    assert c(2, 3) == "Internal Daily meeting", c(2, 3)
    assert [c(2, col) for col in (4, 5, 6, 7, 8)] == [0.5, 0.5, 0.5, 0.5, 0.5]
    print("  OK block A: first task on the date row (row 2) with hours")
    assert c(3, 3) == "Weekly project report" and c(3, 8) == 1, (c(3, 3), c(3, 8))
    assert c(4, 3) == "Bug verification" and c(4, 4) is None, (c(4, 3), c(4, 4))
    print("  OK block A: subsequent tasks packed on rows 3-4, placeholder has no hours")
    # check sum on row 5, per-day SUM spanning the task rows 2..4
    assert c(5, 3) == "check sum", c(5, 3)
    assert c(5, 4) == "=SUM(D2:D4)", c(5, 4)
    assert c(5, 8) == "=SUM(H2:H4)", c(5, 8)
    print("  OK block A: check sum row with =SUM(D2:D4)..=SUM(H2:H4)")

    # --- Block B: dup skipped, new task appended after the existing one ---
    assert c(14, 3) == "External customer meeting", c(14, 3)   # pre-existing untouched
    assert c(15, 3) == "Internal bug triage" and c(15, 5) == 0.5, (c(15, 3), c(15, 5))
    print("  OK block B: dup skipped, new task appended at row 15")
    assert c(16, 3) == "check sum", c(16, 3)
    assert c(16, 4) == "=SUM(D14:D15)", c(16, 4)   # SUM spans existing (14) + new (15)
    print("  OK block B: check sum spans existing + new task rows (D14:D15)")

    wb.close()
    print("ALL PASS")


if __name__ == "__main__":
    main()
