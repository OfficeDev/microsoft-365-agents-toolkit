"""Upload a test image to a GitHub issue asset to diagnose the Bad Size error.
Tries: (1) first PNG screenshot, (2) GIF without explicit mime type, (3) GIF with image/gif.
"""
import os
import sys
import subprocess

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
test_output_dir = os.environ.get("TEST_OUTPUT_DIR", "/tmp/atk-test-output")

if not (issue and token and repo):
    print(f"SKIP: issue={issue!r} repo={repo!r} token={'set' if token else 'missing'}", file=sys.stderr)
    print("")
    sys.exit(0)

def try_upload(file_path, name, mime_type=None):
    """Try uploading file_path as name with optional mime_type. Returns URL or None."""
    if not (file_path and os.path.exists(file_path)):
        return None
    url = f"https://uploads.github.com/repos/{repo}/issues/{issue}/assets"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    with open(file_path, "rb") as f:
        if mime_type:
            files = {"data": (name, f, mime_type)}
        else:
            files = {"data": (name, f)}
        resp = requests.post(url, params={"name": name}, headers=headers, files=files)
    print(f"  [{name} mime={mime_type}] HTTP {resp.status_code}: {resp.text[:200]}", file=sys.stderr)
    if resp.status_code in (200, 201):
        data = resp.json()
        return data.get("browser_download_url") or data.get("url") or ""
    return None

# Strategy 1: Upload first PNG screenshot (simpler format, no animation)
import glob
png_files = sorted(glob.glob(os.path.join(test_output_dir, "screenshots", "*.png")))
if png_files:
    first_png = png_files[0]
    print(f"Trying PNG upload: {first_png}", file=sys.stderr)
    url = try_upload(first_png, "screenshot-01.png", "image/png")
    if url:
        print(url)
        sys.exit(0)
    # Try without explicit mime type
    url = try_upload(first_png, "screenshot-01.png", None)
    if url:
        print(url)
        sys.exit(0)

# Strategy 2: Upload GIF without explicit mime type
if gif_file and os.path.exists(gif_file):
    print(f"Trying GIF upload (no mime): {gif_file}", file=sys.stderr)
    url = try_upload(gif_file, "test-run.gif", None)
    if url:
        print(url)
        sys.exit(0)
    # Try with image/gif
    print(f"Trying GIF upload (image/gif): {gif_file}", file=sys.stderr)
    url = try_upload(gif_file, "test-run.gif", "image/gif")
    if url:
        print(url)
        sys.exit(0)

print("All upload attempts failed", file=sys.stderr)
print("")