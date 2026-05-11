#!/usr/bin/env python3
"""Upload a GIF to a GitHub issue as an embedded asset (CDN URL).
GitHub's uploads.github.com API expects raw binary body, NOT multipart form-data.

Required env vars: GH_TOKEN, GIF_PATH, ISSUE, REPO
Prints the browser_download_url on stdout on success.
"""
import os, sys

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

if not gif_path or not os.path.exists(gif_path):
    print(f"GIF not found: {gif_path}", file=sys.stderr)
    sys.exit(0)

file_size = os.path.getsize(gif_path)
print(f"Uploading GIF: {gif_path} ({file_size} bytes)", file=sys.stderr)

# GitHub uploads API requires raw binary body with Content-Type: image/gif
url = f"https://uploads.github.com/repos/{repo}/issues/{issue}/assets"
with open(gif_path, "rb") as f:
    data = f.read()

resp = requests.post(
    url,
    headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "image/gif",
        "Content-Length": str(file_size),
    },
    data=data,
    timeout=120,
)

print(f"Status: {resp.status_code}", file=sys.stderr)
print(f"Response: {resp.text[:500]}", file=sys.stderr)

if resp.status_code in (200, 201):
    data_json = resp.json()
    gif_url = data_json.get("browser_download_url", data_json.get("url", ""))
    print(gif_url)  # stdout → captured by caller
else:
    print("", end="")  # empty on failure