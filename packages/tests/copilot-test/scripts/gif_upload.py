"""Upload a GIF to a GitHub issue as an asset (multipart/form-data).

Reads env vars:
  GIF_FILE  - path to the GIF file
  ISSUE     - GitHub issue number
  GH_TOKEN  - GitHub personal access token
  REPO      - owner/repo string

Prints the browser_download_url on stdout (or empty string on failure).
Prints HTTP status / error on stderr for debugging.
"""
import os
import sys

try:
    import requests
except ImportError:
    print("requests not installed", file=sys.stderr)
    print("")
    sys.exit(0)

gif_file = os.environ.get("GIF_FILE", "")
issue    = os.environ.get("ISSUE", "")
token    = os.environ.get("GH_TOKEN", "")
repo     = os.environ.get("REPO", "")

if not (gif_file and os.path.exists(gif_file) and issue and token and repo):
    print(f"SKIP: gif_file={gif_file!r} exists={os.path.exists(gif_file) if gif_file else False} issue={issue!r} repo={repo!r}", file=sys.stderr)
    print("")
    sys.exit(0)

url = f"https://uploads.github.com/repos/{repo}/issues/{issue}/assets"
with open(gif_file, "rb") as f:
    resp = requests.post(
        url,
        params={"name": "test-run.gif"},
        headers={
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        },
        files={"data": ("test-run.gif", f, "image/gif")},
    )

print(f"HTTP {resp.status_code}: {resp.text}", file=sys.stderr)
try:
    data = resp.json()
    print(data.get("browser_download_url") or data.get("url") or "")
except Exception as e:
    print(f"JSON parse error: {e}", file=sys.stderr)
    print("")