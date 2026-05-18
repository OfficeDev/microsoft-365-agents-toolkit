"""Upload screenshot to atk-test-screenshots branch via Contents API.
Falls back from GIF to first PNG if GIF not available.
Prints raw.githubusercontent.com URL on stdout, debug on stderr.
"""
import os, sys, base64, json, glob, tempfile, subprocess

issue       = os.environ.get("ISSUE", "")
token       = os.environ.get("GH_TOKEN", "")
repo        = os.environ.get("REPO", "")
gif_file    = os.environ.get("GIF_FILE", "")
run_number  = os.environ.get("GITHUB_RUN_NUMBER", "0")
test_dir    = os.environ.get("TEST_OUTPUT_DIR", "/tmp/atk-test-output")
branch      = "atk-test-screenshots"

if not (issue and token and repo):
    print(f"SKIP: issue={issue!r} repo={repo!r}", file=sys.stderr)
    print("")
    sys.exit(0)

# Choose file to upload: prefer GIF, fall back to first PNG
upload_file = gif_file if (gif_file and os.path.exists(gif_file)) else ""
if not upload_file:
    pngs = sorted(glob.glob(os.path.join(test_dir, "screenshots", "*.png")))
    if pngs:
        upload_file = pngs[0]
if not upload_file:
    print("No screenshot file to upload", file=sys.stderr)
    print("")
    sys.exit(0)

ext = os.path.splitext(upload_file)[1]
file_name = f"issue-{issue}-run-{run_number}{ext}"
contents_path = f".github/test-screenshots/{file_name}"

print(f"Uploading {upload_file} ({os.path.getsize(upload_file)//1024}KB) as {file_name}", file=sys.stderr)

with open(upload_file, "rb") as f:
    content_b64 = base64.b64encode(f.read()).decode()

env = {**os.environ, "GH_TOKEN": token}

# Check if file already exists (need its SHA for update)
check = subprocess.run(
    ["gh", "api", f"repos/{repo}/contents/{contents_path}", "--jq", ".sha"],
    capture_output=True, text=True, env=env
)
existing_sha = check.stdout.strip() if check.returncode == 0 else ""
print(f"existing_sha={existing_sha!r}", file=sys.stderr)

# Ensure branch exists - try to create from dev or main
for base_branch in ["dev", "main"]:
    sha_result = subprocess.run(
        ["gh", "api", f"repos/{repo}/git/refs/heads/{base_branch}", "--jq", ".object.sha"],
        capture_output=True, text=True, env=env
    )
    if sha_result.returncode == 0 and sha_result.stdout.strip():
        base_sha = sha_result.stdout.strip()
        subprocess.run(
            ["gh", "api", "--method", "POST", f"repos/{repo}/git/refs",
             "-f", f"ref=refs/heads/{branch}", "-f", f"sha={base_sha}"],
            capture_output=True, env=env
        )
        print(f"branch ensured from {base_branch}", file=sys.stderr)
        break

# Build JSON body in a temp file to avoid arg-length limits on large base64 strings
body = {
    "message": f"test: screenshot issue {issue} run {run_number} [skip ci]",
    "content": content_b64,
    "branch": branch,
}
if existing_sha:
    body["sha"] = existing_sha

tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
json.dump(body, tmp)
tmp.close()

try:
    result = subprocess.run(
        ["gh", "api", "--method", "PUT",
         f"repos/{repo}/contents/{contents_path}",
         "--input", tmp.name, "--jq", ".content.download_url"],
        capture_output=True, text=True, env=env
    )
    print(f"PUT rc={result.returncode} out={result.stdout.strip()[:200]} err={result.stderr.strip()[:200]}", file=sys.stderr)
    if result.returncode == 0:
        raw_url = f"https://raw.githubusercontent.com/{repo}/{branch}/{contents_path}"
        print(raw_url)
    else:
        print("")
finally:
    os.unlink(tmp.name)