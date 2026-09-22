"""Network-free test for _merge_duplicate_task_rows: Jira отдаёт несколько
тикетов под одним именем, в табеле это должна быть одна строка."""

import importlib.util
import pathlib
from datetime import datetime

import openpyxl

_p = pathlib.Path(__file__).parent / "jira-sync.py"
_spec = importlib.util.spec_from_file_location("jira_sync", _p)
jira_sync = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(jira_sync)

FIXED = {"internal daily meeting"}
failures = []


def check(name, cond, detail=""):
    if cond:
        print(f"  PASS  {name}")
    else:
        print(f"  FAIL  {name}  {detail}")
        failures.append(name)


def build():
    """Две недели. В первой — тройной дубль с часами и гиперссылками, во второй
    задача с тем же именем (другая неделя — НЕ дубль) и дубль регулярной задачи
    (её трогать нельзя)."""
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.cell(1, 1).value = datetime(2026, 9, 6)
    ws.cell(1, 2).value = datetime(2026, 9, 12)
    ws.cell(1, 3).value = "Internal Daily meeting"
    ws.cell(1, 4).value = 0.5
    ws.cell(2, 3).value = "Internal Daily meeting"   # регулярная — не сливаем
    ws.cell(2, 5).value = 0.5
    ws.cell(3, 3).value = "Automation test maintenance"
    ws.cell(3, 4).value = 1.0
    ws.cell(3, 3).hyperlink = "https://jira/GT2-1"
    ws.cell(4, 3).value = "automation test maintenance"  # другой регистр
    ws.cell(4, 5).value = 2.0
    ws.cell(4, 3).hyperlink = "https://jira/GT2-2"
    ws.cell(5, 3).value = "Automation test maintenance "  # хвостовой пробел
    ws.cell(5, 6).value = 0.5
    ws.cell(6, 3).value = "Functional testing"
    ws.cell(6, 4).value = 3.0
    ws.cell(6, 3).hyperlink = "https://jira/GT2-9"
    ws.cell(7, 3).value = "check sum"

    ws.cell(8, 1).value = datetime(2026, 9, 13)
    ws.cell(8, 2).value = datetime(2026, 9, 19)
    ws.cell(8, 3).value = "Automation test maintenance"  # другая неделя
    ws.cell(8, 4).value = 4.0
    ws.cell(9, 3).value = "check sum"
    return wb, ws


wb, ws = build()
rows_before = ws.max_row
counts, deleted = jira_sync._merge_duplicate_task_rows(ws, FIXED)

check("deleted two duplicate rows", deleted == 2, f"{deleted}")
check("multiplicity recorded", counts == {(datetime(2026, 9, 6).date(), "automation test maintenance"): 3},
      f"{counts}")

names = [str(ws.cell(r, 3).value or "").strip().lower() for r in range(1, ws.max_row + 1)]
# openpyxl после delete_rows иногда оставляет висеть пустую строку в dimensions,
# поэтому считаем непустые имена, а не max_row.
filled = [n for n in names if n]
check("two rows fewer with content", len(filled) == rows_before - 2, f"{filled}")
week1 = names[:names.index("check sum") + 1]
check("one automation row left in week 1", week1.count("automation test maintenance") == 1, f"{week1}")
check("regular task duplicates untouched", names.count("internal daily meeting") == 2, f"{names}")
check("other week keeps its own row", names.count("automation test maintenance") == 2, f"{names}")

merged_row = next(r for r in range(1, ws.max_row + 1)
                  if str(ws.cell(r, 3).value or "").strip().lower() == "automation test maintenance")
hours = [ws.cell(merged_row, c).value for c in range(4, 9)]
check("hours summed into the surviving row", hours[:3] == [1.0, 2.0, 0.5], f"{hours}")
check("surviving row keeps its hyperlink",
      ws.cell(merged_row, 3).hyperlink is not None
      and ws.cell(merged_row, 3).hyperlink.target == "https://jira/GT2-1",
      f"{ws.cell(merged_row, 3).hyperlink}")

# Ссылка строки, уехавшей вверх после удаления, должна остаться при ней.
ft_row = next(r for r in range(1, ws.max_row + 1)
              if str(ws.cell(r, 3).value or "").strip() == "Functional testing")
check("hyperlink below the deletion shifted with its row",
      ws.cell(ft_row, 3).hyperlink is not None
      and ws.cell(ft_row, 3).hyperlink.target == "https://jira/GT2-9",
      f"row {ft_row}: {ws.cell(ft_row, 3).hyperlink}")

# Повторный прогон уже слитого файла ничего не меняет.
counts2, deleted2 = jira_sync._merge_duplicate_task_rows(ws, FIXED)
check("second pass is a no-op", deleted2 == 0 and counts2 == {}, f"{deleted2} {counts2}")

print()
if failures:
    print(f"{len(failures)} FAILURE(S): {', '.join(failures)}")
    raise SystemExit(1)
print("ALL PASS")
