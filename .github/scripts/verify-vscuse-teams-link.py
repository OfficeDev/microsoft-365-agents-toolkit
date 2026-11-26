import json
import os
import subprocess
import sys
import urllib.parse
import urllib.request
import uuid


def deployed_app_id():
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
    for line in file.read_text().splitlines():
        name, _, value = line.partition('=')
        if name == variable[1] and value.strip():
            ids.append(uuid.UUID(value.strip().strip('"').strip("'")))
if len(ids) != 1:
    raise SystemExit('Expected one provisioned Teams app for the remote Chrome profile.')
print(ids[0])
""",
        ],
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        timeout=20,
    )
    return uuid.UUID(result.stdout.strip())


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


def reopen_app():
    app = deployed_app_id()
    tenant = uuid.UUID(os.environ["M365_TENANT_ID"])
    query = urllib.parse.urlencode({"tenantId": str(tenant), "installAppPackage": "true"})
    url = f"https://teams.microsoft.com/l/app/{app}?{query}"
    interact("keyboard_shortcut", {"keys": "ctrl+l"})
    interact("type_text", {"text": url})
    interact("key_press", {"key": "enter"})
    print("Teams verification: submitted the deployed app deep link after sign-in", flush=True)


if __name__ == "__main__":
    marker = "[SUCCESS] Step step_browserM365PasswordSignIn_confirmStaySignedIn_c152093f5_10_3"
    reopened = False
    for line in sys.stdin:
        print(line, end="", flush=True)
        if not reopened and marker in line:
            reopen_app()
            reopened = True
