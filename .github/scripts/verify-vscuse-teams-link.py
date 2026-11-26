import json
import os
import subprocess
import sys
import time
import urllib.parse
import urllib.request
import uuid


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


def navigate(path, params, description):
    query = urllib.parse.urlencode(params)
    url = f"https://teams.microsoft.com{path}?{query}"
    interact("keyboard_shortcut", {"keys": "ctrl+l"})
    interact("type_text", {"text": url})
    interact("key_press", {"key": "enter"})
    print(f"Teams verification: submitted {description}", flush=True)


def reopen_app():
    app, _ = deployed_ids()
    tenant = uuid.UUID(os.environ["M365_TENANT_ID"])
    navigate(
        f"/l/app/{app}",
        {"tenantId": str(tenant), "installAppPackage": "true"},
        "the deployed app deep link after sign-in",
    )


def open_bot_chat():
    _, bot = deployed_ids()
    tenant = uuid.UUID(os.environ["M365_TENANT_ID"])
    time.sleep(2)
    navigate(
        "/l/chat/0/0",
        {"users": f"28:{bot}", "tenantId": str(tenant)},
        "the installed bot chat deep link",
    )


if __name__ == "__main__":
    actions = {
        "[SUCCESS] Step step_browserM365PasswordSignIn_confirmStaySignedIn_c152093f5_10_3": reopen_app,
        "[SUCCESS] Step step_addAndOpenApp_open_c152093f5_11_1": open_bot_chat,
    }
    for line in sys.stdin:
        print(line, end="", flush=True)
        marker = next((marker for marker in actions if marker in line), None)
        if marker:
            actions.pop(marker)()
