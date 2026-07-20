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
