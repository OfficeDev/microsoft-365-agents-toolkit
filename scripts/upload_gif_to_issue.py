#!/usr/bin/env python3
"""Upload a GIF by committing it to the repo branch via GitHub Contents API.
Returns a raw.githubusercontent.com URL which renders inline in GitHub markdown.

Required env vars: GH_TOKEN, GIF_PATH, ISSUE, REPO, BRANCH
Prints the raw CDN URL on stdout on success.
"""
import os, sys, base64, json

try:
    import requests
except ImportError:
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "requests"], check=True)
    import requests

gif_path = os.environ.get("GIF_PATH", "")
token    = os.environ.get("GH_TOKEN", "")
issue    = os.environ.get("ISSUE", "")
repo     = os.environ.get("REPO", "")
branch   = os.environ.get("BRANCH", "feat/atk-copilot-issue-test-pipeline")

if not gif_path or not os.path.exists(gif_path):
    print(f"GIF not found: {gif_path}", file=sys.stderr)
    sys.exit(0)

file_size = os.path.getsize(gif_path)
print(f"Committing GIF to branch: {gif_path} ({file_size} bytes)", file=sys.stderr)

with open(gif_path, "rb") as f:
    content_b64 = base64.b64encode(f.read()).decode("ascii")

path = f"test-screenshots/issue-{issue}-latest.gif"
api_url = f"https://api.github.com/repos/{repo}/contents/{path}"
headers = {
    "Authorization": f"Bearer {token}",
    "Accept": "application/vnd.github.v3+json",
}

# Check if file already exists (get SHA for update)
existing = requests.get(api_url + f"?ref={branch}", headers=headers, timeout=30)
sha = None
if existing.status_code == 200:
    sha = existing.json().get("sha")
    print(f"File exists, updating (sha={sha[:8]})", file=sys.stderr)
else:
    print("Creating new file", file=sys.stderr)

payload = {
    "message": f"chore: update test screenshot for issue #{issue} [skip ci]",
    "content": content_b64,
    "branch": branch,
}
if sha:
    payload["sha"] = sha

resp = requests.put(api_url, headers=headers, json=payload, timeout=60)
print(f"Status: {resp.status_code}", file=sys.stderr)
if resp.status_code in (200, 201):
    # raw.githubusercontent.com URL renders inline in GitHub markdown
    raw_url = f"https://raw.githubusercontent.com/{repo}/{branch}/{path}"
    print(raw_url)
else:
    print(f"Response: {resp.text[:400]}", file=sys.stderr)
    print("", end="")