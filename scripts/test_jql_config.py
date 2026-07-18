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
