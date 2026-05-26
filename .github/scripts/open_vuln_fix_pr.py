#!/usr/bin/env python3
"""
Open a fix PR for the first vulnerability surfaced by the CD pipeline scans.

Reads JSON summaries produced by check_npm_vulnerabilities.py and
check_nuget_vulnerabilities.py (in scan order), picks the first record where
`has_vulnerabilities` is true, takes its first vulnerability, attempts a
mechanical version bump in the relevant manifest, and opens a PR against the
configured base branch.

Falls back to an "empty" PR with a placeholder TODO file when an automatic
version bump is not possible (no fix version known, transitive dep, manifest
in a sibling repo, etc.) so a human can take it from there.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Optional


def safe_print(message: str) -> None:
    try:
        print(message, flush=True)
    except UnicodeEncodeError:
        print(message.encode("ascii", "replace").decode("ascii"), flush=True)


def run(cmd, *, cwd=None, check=True, capture=False, env=None) -> subprocess.CompletedProcess:
    safe_print(f"$ {' '.join(cmd)}")
    result = subprocess.run(
        cmd,
        cwd=cwd,
        text=True,
        capture_output=capture,
        env=env,
    )
    if check and result.returncode != 0:
        if capture:
            safe_print(result.stdout)
            safe_print(result.stderr)
        raise SystemExit(f"Command failed: {' '.join(cmd)} (exit {result.returncode})")
    return result


def load_scan(path: Path) -> Optional[dict]:
    if not path.exists():
        safe_print(f"Scan JSON not found, skipping: {path}")
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        safe_print(f"WARNING: Failed to parse {path}: {e}")
        return None


def pick_first_vuln(scan_jsons, skip_targets=None):
    skip_targets = set(skip_targets or [])
    for path in scan_jsons:
        scan = load_scan(path)
        if not scan:
            continue
        if (scan.get("scan_target") or "") in skip_targets:
            safe_print(f"Skipping scan {path} (scan_target={scan.get('scan_target')!r} excluded)")
            continue
        if not scan.get("has_vulnerabilities"):
            continue
        vulns = scan.get("vulnerabilities") or []
        if not vulns:
            continue
        return scan, vulns[0]
    return None, None


def slugify(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "-", value or "")
    return value.strip("-") or "unknown"


def branch_exists_remote(branch: str, repo: str) -> bool:
    result = run(
        ["gh", "api", f"repos/{repo}/branches/{branch}", "--silent"],
        check=False,
        capture=True,
    )
    return result.returncode == 0


def open_pr_exists(branch: str, repo: str) -> bool:
    result = run(
        ["gh", "pr", "list", "--repo", repo, "--head", branch, "--state", "open", "--json", "number"],
        check=False,
        capture=True,
    )
    if result.returncode != 0:
        return False
    try:
        data = json.loads(result.stdout or "[]")
        return bool(data)
    except json.JSONDecodeError:
        return False


def bump_npm_manifest(file_path: Path, package: str, fixed_version: str) -> bool:
    """Replace a direct dep version in a package.json[.tpl]. Preserve caret/tilde prefix when present."""
    if not file_path.exists():
        return False
    text = file_path.read_text(encoding="utf-8")
    pattern = re.compile(
        r'("' + re.escape(package) + r'"\s*:\s*")([~^]?)([^"\s]+)(")'
    )
    new_text, count = pattern.subn(rf'\g<1>\g<2>{fixed_version}\g<4>', text)
    if count == 0:
        return False
    file_path.write_text(new_text, encoding="utf-8")
    return True


def bump_csproj(file_path: Path, package: str, fixed_version: str) -> bool:
    if not file_path.exists():
        return False
    text = file_path.read_text(encoding="utf-8")
    pattern = re.compile(
        r'(<PackageReference\s+[^>]*Include="' + re.escape(package) + r'"[^>]*Version=")([^"]+)(")',
        re.IGNORECASE,
    )
    new_text, count = pattern.subn(rf'\g<1>{fixed_version}\g<3>', text)
    if count == 0:
        return False
    file_path.write_text(new_text, encoding="utf-8")
    return True


def attempt_file_bump(vuln: dict, ecosystem: str, repo_root: Path) -> bool:
    fixed_version = vuln.get("fixed_version")
    package = vuln.get("package")
    file_rel = vuln.get("file")
    if not (fixed_version and package and file_rel):
        return False
    file_path = (repo_root / file_rel).resolve()
    # Defensive: refuse to write outside the repo.
    try:
        file_path.relative_to(repo_root.resolve())
    except ValueError:
        safe_print(f"Refusing to touch file outside repo root: {file_path}")
        return False
    if ecosystem == "npm":
        return bump_npm_manifest(file_path, package, fixed_version)
    if ecosystem == "nuget":
        return bump_csproj(file_path, package, fixed_version)
    return False


def write_placeholder(repo_root: Path, scan: dict, vuln: dict) -> Path:
    placeholder = repo_root / ".github" / "vuln-fix-todo.md"
    placeholder.parent.mkdir(parents=True, exist_ok=True)
    body_lines = [
        "# Vulnerability fix TODO",
        "",
        "This file was generated automatically because the CD vulnerability scan",
        "found an issue that could not be patched mechanically. Please review and",
        "replace this file with the actual fix, then re-open this PR.",
        "",
        f"- Ecosystem: `{scan.get('ecosystem')}`",
        f"- Scan target: `{scan.get('scan_target')}`",
        f"- File: `{vuln.get('file')}`",
        f"- Package: `{vuln.get('package')}`",
        f"- Current version: `{vuln.get('current_version')}`",
        f"- Fixed version: `{vuln.get('fixed_version')}`",
        f"- Severity: `{vuln.get('severity')}`",
        f"- Advisory: {vuln.get('advisory_url') or 'n/a'}",
        f"- Title: {vuln.get('title') or 'n/a'}",
    ]
    placeholder.write_text("\n".join(body_lines) + "\n", encoding="utf-8")
    return placeholder


def build_pr_body(scan: dict, vuln: dict, automatic_fix: bool) -> str:
    lines = [
        "This PR was opened automatically by the CD vulnerability scan.",
        "",
        f"- **Ecosystem**: {scan.get('ecosystem')}",
        f"- **Scan target**: `{scan.get('scan_target')}`",
        f"- **File**: `{vuln.get('file')}`",
        f"- **Package**: `{vuln.get('package')}`",
        f"- **Current version**: `{vuln.get('current_version') or 'unknown'}`",
        f"- **Fixed version**: `{vuln.get('fixed_version') or 'unknown — manual fix required'}`",
        f"- **Severity**: `{vuln.get('severity') or 'unknown'}`",
    ]
    if vuln.get("advisory_url"):
        lines.append(f"- **Advisory**: {vuln['advisory_url']}")
    if vuln.get("title"):
        lines.append(f"- **Title**: {vuln['title']}")
    lines.append("")
    if automatic_fix:
        lines.append(
            "An automatic version bump was applied. Please verify the change "
            "builds and the dependency is still compatible before merging."
        )
    else:
        lines.append(
            "**No automatic fix was applied** — either the fixed version is "
            "unknown, the package is a transitive dependency, or the manifest "
            "lives outside this repository. Replace the placeholder file with "
            "the real fix and update this PR."
        )
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Open a fix PR for the first vulnerability")
    parser.add_argument("--scan-json", action="append", default=[], help="Path to a scan summary JSON (repeatable, order matters)")
    parser.add_argument("--base-branch", default="dev")
    parser.add_argument("--repo-root", default=".")
    parser.add_argument(
        "--skip-scan-target",
        action="append",
        default=[],
        help="scan_target values to ignore entirely (e.g. samples-repo). Repeatable.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Print actions but do not commit/push/open PR")
    args = parser.parse_args()

    if not args.scan_json:
        safe_print("No --scan-json provided; nothing to do.")
        return 0

    repo_root = Path(args.repo_root).resolve()
    scan_paths = [Path(p) for p in args.scan_json]

    scan, vuln = pick_first_vuln(scan_paths, skip_targets=args.skip_scan_target)
    if not scan or not vuln:
        safe_print("No actionable vulnerability found across scans; nothing to do.")
        return 0

    ecosystem = scan.get("ecosystem", "unknown")
    package = vuln.get("package") or "unknown"
    fixed_version = vuln.get("fixed_version")
    branch = f"auto-fix-vuln/{ecosystem}-{slugify(package)}-{slugify(fixed_version or 'unknown')}"

    repo = os.environ.get("GITHUB_REPOSITORY")
    if not repo:
        safe_print("GITHUB_REPOSITORY env var is not set; cannot check existing PRs.")
        return 1

    if open_pr_exists(branch, repo):
        safe_print(f"Skipped: an open PR already exists for branch {branch}")
        return 0
    if branch_exists_remote(branch, repo):
        # No open PR but branch exists — leftover from a previous failed run.
        # Delete it so this run can recreate cleanly.
        safe_print(f"Remote branch exists without an open PR; deleting stale branch: {branch}")
        if args.dry_run:
            safe_print("[dry-run] would: git push origin --delete " + branch)
        else:
            run(["git", "push", "origin", "--delete", branch], cwd=repo_root, check=False)

    samples_target = (scan.get("scan_target") or "").startswith("samples-repo")
    can_bump_in_this_repo = not samples_target

    if args.dry_run:
        would_bump = bool(can_bump_in_this_repo and fixed_version)
        dry_subject = (
            f"fix(deps): bump {package} to {fixed_version}"
            if would_bump
            else f"chore(security): TODO fix {package} vulnerability"
        )
        safe_print(f"[dry-run] branch={branch} would_attempt_bump={would_bump}")
        safe_print(f"[dry-run] title={dry_subject}")
        safe_print("[dry-run] body:\n" + build_pr_body(scan, vuln, would_bump))
        return 0

    automatic_fix = False
    if can_bump_in_this_repo and fixed_version:
        automatic_fix = attempt_file_bump(vuln, ecosystem, repo_root)

    run(["git", "checkout", "-b", branch], cwd=repo_root)

    if not automatic_fix:
        placeholder = write_placeholder(repo_root, scan, vuln)
        safe_print(f"Wrote placeholder: {placeholder}")

    # Stage everything first so untracked files (e.g. the placeholder) are visible
    # to the diff check below — `git diff --quiet` ignores untracked paths.
    # Exclude samples-repo/: CI checks it out as a sibling repo and `git add -A`
    # would otherwise pick it up as a gitlink/submodule entry.
    run(["git", "add", "-A", "--", ":!samples-repo", ":!samples-repo/**"], cwd=repo_root)

    diff_result = run(["git", "diff", "--cached", "--quiet"], cwd=repo_root, check=False)
    if diff_result.returncode == 0:
        safe_print("No staged changes after add; aborting PR creation.")
        return 0

    commit_subject = (
        f"fix(deps): bump {package} to {fixed_version}"
        if automatic_fix
        else f"chore(security): TODO fix {package} vulnerability"
    )
    run(["git", "commit", "-m", commit_subject], cwd=repo_root)
    run(["git", "push", "origin", branch], cwd=repo_root)

    title = commit_subject
    body = build_pr_body(scan, vuln, automatic_fix)
    # Token selection for `gh pr create`, in priority order:
    #   1. GH_TOKEN_PERSONAL — a user PAT for early validation when org/App
    #      policy hasn't been opened up yet.
    #   2. GH_TOKEN_FOR_PR — typically secrets.GITHUB_TOKEN with pull-requests:
    #      write, so we don't depend on the App token having that scope.
    #   3. Fall back to the existing GH_TOKEN (App token).
    personal_pat = os.environ.get("GH_TOKEN_PERSONAL") or ""
    fallback_pr_token = os.environ.get("GH_TOKEN_FOR_PR") or ""
    pr_token = personal_pat.strip() or fallback_pr_token.strip()
    pr_env = None
    if pr_token:
        pr_env = os.environ.copy()
        pr_env["GH_TOKEN"] = pr_token
        source = "GH_TOKEN_PERSONAL" if personal_pat.strip() else "GH_TOKEN_FOR_PR"
        safe_print(f"Using {source} for gh pr create")
    run(
        [
            "gh", "pr", "create",
            "--repo", repo,
            "--base", args.base_branch,
            "--head", branch,
            "--title", title,
            "--body", body,
        ],
        cwd=repo_root,
        env=pr_env,
    )

    safe_print(f"Opened PR for {package} on branch {branch}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
