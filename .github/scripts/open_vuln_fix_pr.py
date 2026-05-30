#!/usr/bin/env python3
"""
Open a fix PR for the first vulnerability surfaced by the CD pipeline scans.

Reads JSON summaries produced by check_npm_vulnerabilities.py and
check_nuget_vulnerabilities.py (in scan order), picks the first record where
`has_vulnerabilities` is true, takes its first vulnerability, attempts a
mechanical version bump in the relevant manifest, and opens a PR against the
configured base branch.

For npm vulns the bump is verified end-to-end in a temp dir (`npm install` +
`npm audit`) before being committed. Three strategies are tried in order:
  1. direct dependency bump
  2. top-level parent bump (when the vuln is transitive)
  3. `overrides` pin

The first strategy that makes `npm audit` stop reporting the package wins;
if none work the PR is skipped (no placeholder).

NuGet vulns fall back to the historical "write a TODO placeholder" behavior
when an automatic bump isn't possible.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional, Tuple


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


# --------------------------------------------------------------------------- #
# npm helpers: render template, bump, verify
# --------------------------------------------------------------------------- #


MUSTACHE_RE = re.compile(r"\{\{[^}]+\}\}")


def _resolve_executable(name: str) -> str:
    """Return a runnable path for `name` (e.g. 'npm'), handling Windows
    where the real entry point is `npm.cmd` and Python's subprocess does
    not search PATHEXT by default."""
    found = shutil.which(name)
    if found:
        return found
    if os.name == "nt":
        for ext in (".cmd", ".exe", ".bat"):
            alt = shutil.which(name + ext)
            if alt:
                return alt
    return name


def render_manifest_for_install(text: str) -> str:
    """Replace mustache placeholders so a .tpl is consumable by npm install.

    We only need this for the temp-dir verify; the on-disk .tpl keeps its
    placeholders intact. Strategy: replace each `{{...}}` with a safe slug.
    The `name` field then becomes e.g. "placeholder-safeprojectnamelowercase",
    which npm accepts.
    """

    def _sub(match: re.Match) -> str:
        inner = match.group(0).strip("{}").strip()
        return "placeholder-" + slugify(inner).lower()

    return MUSTACHE_RE.sub(_sub, text)


def _npm_install_lock_only(work_dir: Path) -> Tuple[bool, str]:
    proc = subprocess.run(
        [_resolve_executable("npm"), "install", "--package-lock-only", "--no-audit", "--ignore-scripts"],
        cwd=work_dir,
        capture_output=True,
        text=True,
        timeout=300,
    )
    if proc.returncode != 0:
        return False, (proc.stderr or proc.stdout or "")[:500]
    return True, ""


def _npm_audit_json(work_dir: Path) -> dict:
    proc = subprocess.run(
        [_resolve_executable("npm"), "audit", "--json"],
        cwd=work_dir,
        capture_output=True,
        text=True,
        timeout=300,
    )
    try:
        return json.loads(proc.stdout) if proc.stdout else {}
    except json.JSONDecodeError:
        return {}


def verify_npm_no_vuln(manifest_text: str, package_name: str) -> Tuple[bool, str]:
    """Render manifest, install in a temp dir, run npm audit.

    Returns (success, reason). success=True means npm audit no longer lists
    `package_name` among `vulnerabilities`.
    """
    rendered = render_manifest_for_install(manifest_text)
    with tempfile.TemporaryDirectory(prefix="vuln_verify_") as tmp:
        work = Path(tmp)
        (work / "package.json").write_text(rendered, encoding="utf-8")
        ok, err = _npm_install_lock_only(work)
        if not ok:
            return False, f"npm install failed: {err.strip()}"
        audit = _npm_audit_json(work)
        vulns = audit.get("vulnerabilities", {}) or {}
        if package_name in vulns:
            sev = (vulns[package_name].get("severity") or "?").lower()
            return False, f"{package_name} still vulnerable (severity={sev})"
        return True, "clean"


def find_top_level_parent(manifest_text: str, package_name: str) -> Optional[str]:
    """Use `npm ls --all` against the rendered manifest to find the direct dep
    that pulls in `package_name`. Returns the top-level dep name or None.
    """
    rendered = render_manifest_for_install(manifest_text)
    with tempfile.TemporaryDirectory(prefix="vuln_parent_") as tmp:
        work = Path(tmp)
        (work / "package.json").write_text(rendered, encoding="utf-8")
        ok, _ = _npm_install_lock_only(work)
        if not ok:
            return None
        proc = subprocess.run(
            [_resolve_executable("npm"), "ls", package_name, "--all", "--json", "--package-lock-only"],
            cwd=work,
            capture_output=True,
            text=True,
            timeout=180,
        )
        try:
            data = json.loads(proc.stdout) if proc.stdout else {}
        except json.JSONDecodeError:
            return None

        def walk(node: dict, parent_chain):
            deps = node.get("dependencies") or {}
            for name, child in deps.items():
                chain = parent_chain + [name]
                if name == package_name:
                    return chain[0]
                found = walk(child, chain)
                if found:
                    return found
            return None

        return walk(data, [])


def _replace_direct_dep_version(text: str, package: str, fixed_version: str) -> Tuple[str, int]:
    """Replace `"package": "X"` preserving any leading caret/tilde. Returns
    (new_text, count_of_replacements)."""
    pattern = re.compile(
        r'("' + re.escape(package) + r'"\s*:\s*")([~^]?)([^"\s]+)(")'
    )
    return pattern.subn(rf'\g<1>\g<2>{fixed_version}\g<4>', text)


def _latest_npm_version(package: str) -> Optional[str]:
    proc = subprocess.run(
        [_resolve_executable("npm"), "view", package, "version"],
        capture_output=True,
        text=True,
        timeout=60,
    )
    if proc.returncode != 0:
        return None
    out = (proc.stdout or "").strip()
    return out or None


def _inject_overrides(text: str, package: str, fixed_version: str) -> Optional[str]:
    """Add or update an `overrides` block in package.json text.

    Naive JSON manipulation: we parse, mutate, and re-dump with 2-space indent.
    This loses comment-style formatting but package.json doesn't allow comments
    so that's fine. Order of top-level keys is preserved (json.loads keeps
    insertion order in Python 3.7+).
    """
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return None
    overrides = data.get("overrides")
    if not isinstance(overrides, dict):
        overrides = {}
    overrides[package] = fixed_version
    data["overrides"] = overrides
    return json.dumps(data, indent=2) + "\n"


def bump_npm_with_verify(file_path: Path, vuln: dict) -> Tuple[bool, str]:
    """Try direct bump, parent bump, then overrides. Verify each via
    `npm install + npm audit` in a temp dir. Write the winning strategy to
    `file_path` (preserving its original mustache placeholders for direct/parent
    strategies; overrides goes through json round-trip so placeholders inside
    strings survive but formatting is normalized)."""

    package = vuln.get("package") or ""
    fixed_version = vuln.get("fixed_version")
    if not package:
        return False, "no package name in vuln"

    original = file_path.read_text(encoding="utf-8")

    # Strategy A: direct dependency bump.
    if fixed_version:
        candidate, count = _replace_direct_dep_version(original, package, fixed_version)
        if count > 0:
            safe_print(f"[strategy A] direct bump: {package} -> {fixed_version}")
            ok, reason = verify_npm_no_vuln(candidate, package)
            if ok:
                file_path.write_text(candidate, encoding="utf-8")
                return True, f"direct dep bump to {fixed_version}"
            safe_print(f"[strategy A] verify failed: {reason}")

    # Strategy B: bump the top-level dep that pulls in this transitive vuln.
    parent = find_top_level_parent(original, package)
    if parent and parent != package:
        latest = _latest_npm_version(parent)
        if latest:
            candidate, count = _replace_direct_dep_version(original, parent, latest)
            if count > 0:
                safe_print(f"[strategy B] parent bump: {parent} -> {latest} (to evict {package})")
                ok, reason = verify_npm_no_vuln(candidate, package)
                if ok:
                    file_path.write_text(candidate, encoding="utf-8")
                    return True, f"parent {parent} bumped to {latest}"
                safe_print(f"[strategy B] verify failed: {reason}")
        else:
            safe_print(f"[strategy B] could not look up latest version for parent {parent}")
    else:
        safe_print(f"[strategy B] no top-level parent found for {package}")

    # Strategy C: pin via overrides.
    if fixed_version:
        candidate = _inject_overrides(original, package, fixed_version)
        if candidate:
            safe_print(f"[strategy C] overrides pin: {package} -> {fixed_version}")
            ok, reason = verify_npm_no_vuln(candidate, package)
            if ok:
                file_path.write_text(candidate, encoding="utf-8")
                return True, f"overrides pin to {fixed_version}"
            safe_print(f"[strategy C] verify failed: {reason}")

    return False, "all strategies exhausted"


# --------------------------------------------------------------------------- #
# nuget helper (unchanged behavior)
# --------------------------------------------------------------------------- #


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


def build_pr_body(scan: dict, vuln: dict, automatic_fix: bool, strategy: str = "") -> str:
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
            "An automatic fix was applied and verified locally "
            "(`npm install` + `npm audit` no longer flag this package)."
        )
        if strategy:
            lines.append(f"Strategy used: **{strategy}**.")
        lines.append("Please still review the diff before merging.")
    else:
        lines.append(
            "**No automatic fix was applied** — either the fixed version is "
            "unknown, the package is a transitive dependency, or the manifest "
            "lives outside this repository. Replace the placeholder file with "
            "the real fix and update this PR."
        )
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #


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
        safe_print(f"Remote branch exists without an open PR; deleting stale branch: {branch}")
        if args.dry_run:
            safe_print("[dry-run] would: git push origin --delete " + branch)
        else:
            run(["git", "push", "origin", "--delete", branch], cwd=repo_root, check=False)

    samples_target = (scan.get("scan_target") or "").startswith("samples-repo")
    can_bump_in_this_repo = not samples_target

    file_rel = vuln.get("file") or ""
    target_file = (repo_root / file_rel).resolve() if file_rel else None
    if target_file:
        try:
            target_file.relative_to(repo_root.resolve())
        except ValueError:
            safe_print(f"Refusing to touch file outside repo root: {target_file}")
            target_file = None

    automatic_fix = False
    strategy_label = ""

    if args.dry_run:
        safe_print(f"[dry-run] branch={branch}")
        safe_print(f"[dry-run] ecosystem={ecosystem} package={package} fixed={fixed_version}")
        safe_print(f"[dry-run] target_file={target_file}")
        return 0

    if can_bump_in_this_repo and target_file and target_file.exists():
        if ecosystem == "npm":
            ok, label = bump_npm_with_verify(target_file, vuln)
            automatic_fix = ok
            strategy_label = label
            if not ok:
                safe_print(f"npm bump+verify failed for {package}: {label}. Skipping PR.")
                return 0
        elif ecosystem == "nuget" and fixed_version:
            automatic_fix = bump_csproj(target_file, package, fixed_version)
            strategy_label = "csproj version bump" if automatic_fix else ""

    run(["git", "checkout", "-b", branch], cwd=repo_root)

    # NuGet: keep the historical placeholder fallback so the human gets a PR.
    # npm: we returned above if bump+verify failed, so we never land here without a real fix.
    if not automatic_fix and ecosystem != "npm":
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
    body = build_pr_body(scan, vuln, automatic_fix, strategy=strategy_label)
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

    safe_print(f"Opened PR for {package} on branch {branch} (strategy: {strategy_label or 'placeholder'})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
