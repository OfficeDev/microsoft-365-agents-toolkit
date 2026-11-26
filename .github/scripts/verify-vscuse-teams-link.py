import json
import os
import subprocess
import sys
import time
import urllib.parse
import urllib.request
import uuid

HOME_PAGE = "https://teams.microsoft.com/v2/"
SIGN_IN = "step_browserM365PasswordSignIn_confirmStaySignedIn_c152093f5_10_3"
APP_READY = "step_assertReady_assertReady_c152093f5_10_4"
APP_OPENED = "step_addAndOpenApp_open_c152093f5_11_1"
CHAT_READY = "step_addAndOpenApp_assertReady_c152093f5_11_1"


def deployed_ids():
    result = subprocess.run(
        [
            "docker",
            "exec",
            "--user",
            "vscode",
            "vscuse-container",
            "python3",
            "-c",
            """
from pathlib import Path
import re
import uuid

root = Path('/home/vscode/AgentsToolkitProjects')
ids = []
for file in root.glob('vscuse_app_*/env/.env.dev'):
    launch = file.parents[1] / '.vscode' / 'launch.json'
    _, marker, chrome = launch.read_text().partition('"name": "Launch Remote (Chrome)"')
    if not marker:
        continue
    profile = chrome.split('"presentation"', 1)[0]
    variable = re.search(r'teams\\.microsoft\\.com/l/app/\\$\\{\\{(APP_ID|TEAMS_APP_ID)\\}\\}', profile)
    if not variable:
        continue
    values = {}
    for line in file.read_text().splitlines():
        name, _, value = line.partition('=')
        if name in (variable[1], 'BOT_ID') and value.strip():
            values[name] = uuid.UUID(value.strip().strip('"').strip("'"))
    if variable[1] in values and 'BOT_ID' in values:
        ids.append((values[variable[1]], values['BOT_ID']))
if len(ids) != 1:
    raise SystemExit('Expected one provisioned Teams app for the remote Chrome profile.')
print(*ids[0])
""",
        ],
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        timeout=20,
    )
    app, bot = result.stdout.split()
    return uuid.UUID(app), uuid.UUID(bot)


def interact(action, payload):
    request = urllib.request.Request(
        f"http://localhost:6081/api/v1/interaction/{action}",
        data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=10) as response:
        result = json.load(response)
    if result.get("success") is False or result.get("error"):
        raise RuntimeError(f"vscuse interaction failed: {action}")


def navigate(url, description):
    interact("keyboard_shortcut", {"keys": "ctrl+l"})
    interact("type_text", {"text": url})
    interact("key_press", {"key": "enter"})
    print(f"Teams verification: submitted {description}", flush=True)


def navigate_to_teams(path, params, description):
    query = urllib.parse.urlencode(params)
    navigate(f"https://teams.microsoft.com{path}?{query}", description)


def complete_navigation(path, params, description):
    navigate(HOME_PAGE, f"Teams home before {description}")
    time.sleep(20)
    navigate_to_teams(path, params, description)


def reopen_app():
    app, _ = deployed_ids()
    tenant = uuid.UUID(os.environ["M365_TENANT_ID"])
    complete_navigation(
        f"/l/app/{app}",
        {"tenantId": str(tenant), "installAppPackage": "true"},
        "the deployed app deep link",
    )


def open_bot_chat():
    _, bot = deployed_ids()
    tenant = uuid.UUID(os.environ["M365_TENANT_ID"])
    complete_navigation(
        "/l/chat/0/0", {"users": f"28:{bot}", "tenantId": str(tenant)}, "the installed bot chat deep link"
    )


def prepare_verification_plan(source, destination):
    with open(source, encoding="utf-8") as file:
        plan = json.load(file)
    steps = {step["step_id"]: step for step in plan["steps"]}
    for trigger, assertion in ((SIGN_IN, APP_READY), (APP_OPENED, CHAT_READY)):
        if trigger not in steps[assertion]["depends_on"]:
            raise RuntimeError(f"Unexpected dependency before Teams assertion: {assertion}")
        tags = steps[assertion]["tags"]
        tags[:] = [tag for tag in tags if not tag.startswith("step_retry_timeout:")]
        tags.append("step_retry_timeout: 300")
    with open(destination, "w", encoding="utf-8") as file:
        json.dump(plan, file)


def observe_run():
    actions = {
        f"[SUCCESS] Step {SIGN_IN}": reopen_app,
        f"[SUCCESS] Step {APP_OPENED}": open_bot_chat,
    }
    for line in sys.stdin:
        print(line, end="", flush=True)
        marker = next((marker for marker in actions if marker in line), None)
        if marker:
            actions.pop(marker)()


if __name__ == "__main__":
    if len(sys.argv) == 4 and sys.argv[1] == "--prepare":
        prepare_verification_plan(*sys.argv[2:])
    elif len(sys.argv) == 1:
        observe_run()
    else:
        raise SystemExit("Expected no arguments or --prepare SOURCE DESTINATION.")
