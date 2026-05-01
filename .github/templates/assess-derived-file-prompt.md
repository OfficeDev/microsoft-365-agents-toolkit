# Assess and Update Derived File

You are checking whether a derived file needs updating after upstream content was synced. If incompatible, regenerate it.

## Target

- **Path:** `{{DERIVED_PATH}}`
- **Purpose:** {{DESCRIPTION}}

## Context

### Files added this sync (NEW — did not exist before)

```
{{ADDED_FILES}}
```

### Files modified this sync

```
{{MODIFIED_FILES}}
```

### Directory listing of all synced content

```
{{DIR_LISTING}}
```

## Instructions

### Phase 1: Assess Compatibility

Read the derived file at `{{DERIVED_PATH}}`. Run ALL of these checks:

#### Check 1 — Reference integrity
Every `[text](path)` link must resolve to a file that exists. Resolve relative paths from the derived file's location. Check against the directory listing. Flag any broken links (targets renamed or deleted).

#### Check 2 — Enumeration completeness (CRITICAL — most common failure)
Find every section in the derived file that lists multiple files from the SAME directory. These are enumeration sections. Examples of patterns to look for:
- A bullet list of links like `- [Name](../docs/foo.md)`, `- [Name](../docs/bar.md)` — this enumerates `docs/`
- A markdown table with file links in one column
- A pipe-separated file inventory line like `file-a.md | file-b.md | file-c.md`
- A domain routing table with links to `index.md` files in subdirectories

**For EACH enumeration section you find:**
1. Resolve the relative links to determine which directory is being enumerated
2. From the directory listing, list ALL `.md` files in that directory (excluding README.md and index.md unless the section already includes them)
3. Compare: every file in the directory should have a corresponding entry
4. Also compare: every ADDED file (from "Files added this sync") whose path falls under that directory MUST appear

**If ANY added file belongs to a directory that has an enumeration section but is missing from that section → INCOMPATIBLE.**

**Example:** If the derived file has a "Platform Comparison Docs" section listing 9 files from `docs/`, and a new `docs/workflows.md` was added in this sync but is NOT listed → INCOMPATIBLE.

#### Check 3 — Description accuracy
Check that numeric counts and factual claims still match. Examples:
- "27 experts" but directory now has 28
- A domain description that contradicts the current index file content

#### Check 4 — Structural coherence
If newly added files introduce an entirely new CATEGORY or DOMAIN (not just new files within an existing domain), check whether the derived file's high-level structure needs a new section.

### Verdict

Output exactly one of:
- `VERDICT: COMPATIBLE` — all checks pass, no changes needed
- `VERDICT: INCOMPATIBLE` — followed by a bullet list of specific failing checks

### Phase 2: Apply fixes (only if INCOMPATIBLE)

Use the `edit_file` tool to make targeted edits to `{{DERIVED_PATH}}`. Do NOT rewrite the entire file.

Rules:
1. **Read before writing** — for each newly added file that needs a description, READ its content (at least the title and first 30 lines) to write an accurate, style-consistent description. Do NOT guess what a file contains.
2. **Targeted edits only** — use the `edit_file` tool to insert or replace only the specific lines that need changing. Do NOT replace the entire file.
3. **Use the SAME relative link style as the original** — if existing entries use `../docs/foo.md`, your additions must too. The derived file is at `{{DERIVED_PATH}}` — all links are relative to its parent directory.
4. **Match formatting exactly** — new entries must use the same bullet style, table format, or list pattern as adjacent existing entries.
5. **Do not remove or rewrite** anything that passed checks.

After making all edits, output exactly:
```
VERDICT: INCOMPATIBLE
EDITS_APPLIED
```
