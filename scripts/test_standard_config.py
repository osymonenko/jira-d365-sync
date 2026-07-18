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
