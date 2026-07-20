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
