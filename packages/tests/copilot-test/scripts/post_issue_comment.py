#!/usr/bin/env python3
"""Post ATK Copilot test results as a GitHub issue comment.

Reads results.json, test.log, screenshots count, GIF embed URL, and posts
a formatted comment to the GitHub issue via gh CLI.

Required env vars:
  GH_TOKEN, ISSUE, REPO, RUN_URL, GIF_EMBED_URL, ARTIFACT_URL,
  TEST_OUTPUT_DIR, TEST_PLAN_PATH
"""
import os, sys, json, subprocess

issue          = os.environ['ISSUE']
repo           = os.environ['REPO']
run_url        = os.environ['RUN_URL']
gif_embed_url  = os.environ.get('GIF_EMBED_URL', '').strip()
artifact_url   = os.environ.get('ARTIFACT_URL', '').strip()
test_output    = os.environ.get('TEST_OUTPUT_DIR', '/tmp/atk-test-output')
test_plan_path = os.environ.get('TEST_PLAN_PATH', 'packages/tests/copilot-test/test-plans/simple-bot/simple-bot.md')

# Read results
results_file = os.path.join(test_output, 'results.json')
try:
    with open(results_file) as f:
        results = json.load(f)
except Exception:
    results = {'passed': 0, 'failed': 0, 'steps': []}

passed = results.get('passed', 0)
failed = results.get('failed', 0)
steps  = results.get('steps', [])
status = 'PASSED' if failed == 0 else 'FAILED'
badge  = '✅ PASSED' if status == 'PASSED' else '❌ FAILED'

# Count screenshots
screenshots_dir = os.path.join(test_output, 'screenshots')
try:
    shots = len([f for f in os.listdir(screenshots_dir) if f.endswith('.png')])
except Exception:
    shots = 0

# Read test log
log_file = os.path.join(test_output, 'test.log')
try:
    with open(log_file) as f:
        all_lines = f.readlines()
    log = ''.join(all_lines[-40:])
except Exception:
    log = 'no log'

# Read test plan
try:
    with open(test_plan_path) as f:
        test_plan = ''.join(f.readlines()[:80])
except Exception:
    test_plan = '*(test plan not found)*'

# Build step table
if steps:
    rows = ['| Step | Status |', '|------|--------|']
    for s in steps:
        icon = '✅' if s.get('status') in ('pass', 'passed') else '❌'
        rows.append(f'| {s.get("name", "?")} | {icon} |')
    steps_table = '\n'.join(rows)
else:
    steps_table = '*(no step details)*'

# Build GIF / screenshots section
if gif_embed_url:
    gif_section = f'\n### Screenshots\n\n![test-run]({gif_embed_url})\n'
elif shots > 0 and artifact_url:
    gif_section = f'\n### Screenshots\n\n[Download screenshots]({artifact_url})\n'
else:
    gif_section = ''

body = f"""<!-- atk-copilot-test-results -->
### 🤖 ATK Copilot Test Results — {badge}

| | |
|---|---|
| **Issue** | #{issue} |
| **Passed** | {passed} |
| **Failed** | {failed} |
| **Screenshots** | {shots} |
| **Run** | [View logs]({run_url}) |
{gif_section}
### Test Steps

{steps_table}

<details>
<summary>Test Plan</summary>

{test_plan}

</details>

<details>
<summary>Test log (last 40 lines)</summary>

```
{log}
```
</details>"""

result = subprocess.run(
    ['gh', 'issue', 'comment', issue, '--repo', repo, '--body', body],
    capture_output=True, text=True, env={**os.environ}
)
print(result.stdout)
if result.returncode != 0:
    print('Error:', result.stderr, file=sys.stderr)
    sys.exit(result.returncode)
