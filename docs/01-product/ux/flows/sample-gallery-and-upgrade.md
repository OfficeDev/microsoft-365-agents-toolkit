# Sample gallery and project upgrade (v3)

> **v4 design status.** This page is an allowed input to v4 design. It captures the v3 user-facing surfaces around discovering samples and upgrading legacy projects.

Two related but distinct surfaces, grouped here because both are project-creation-adjacent flows that v4 needs to decide whether to preserve.

## Sample gallery

A separate creation flow from templates. **Templates** scaffold an empty starter; **samples** scaffold a complete working app demonstrating a concept.

### CLI

```
atk list samples [--tag <tag>] [--format <table|json>] [--description]
atk new sample <samples-id> [--folder-path <path>] [-i <bool>]
```

Source: [`packages/cli/src/commands/models/listSamples.ts`](../../../packages/cli/src/commands/models/listSamples.ts), [`packages/cli/src/commands/models/createSample.ts`](../../../packages/cli/src/commands/models/createSample.ts).

The sample list is currently a **static array embedded in the source** (`packages/fx-core/src/question/options/CreateSampleProjectOptions.ts`). No external/remote feed. No CDN-based versioning. Each sample has:

- `id` (kebab-case identifier)
- `name` (display name)
- `description`
- `tags` (e.g. `declarative-agent`, `rag`, `weather`, `bot`, `ai-agent`, `ms-teams`)
- `link` (optional GitHub URL for "View source")

### VS Code

Tree view item **"View Samples"** under the Development view → opens a webview gallery (Vite + React + Fluent UI v8) showing samples with filter chips by tag. Click a sample → "Create Project" button → toolkit clones the sample into the chosen folder.

### Discovery UX

| Surface | Filter affordances |
|---------|-------------------|
| CLI `atk list samples` | `--tag` filter; table or JSON output |
| VS Code samples webview | Tag chips, search box, sort options |

### Why samples differ from templates

| Aspect | Templates | Samples |
|--------|-----------|---------|
| Scaffold via | `atk new` (question tree) | `atk new sample <id>` |
| Source files | `templates/vsc/{lang}/{template}/` (in-repo) | Cloned from Microsoft's samples GitHub repo at scaffold time |
| Mustache placeholders | Yes (`appName`, etc.) | No (full working app) |
| Lifecycle YAML | Standard | Per-sample (may differ) |
| Question tree | Question-driven options | One question only: which sample |
| Updates | Per-toolkit-release | Per-sample-repo commit (latest at scaffold time) |

### v4 design implication

Open questions for v4:

- Move sample list to a remote feed (allows update without toolkit release)?
- Unify samples and templates under a single `TemplateRegistry`?
- Surface tag taxonomy as a first-class registry concept?

## Project upgrade (v2 → v3)

For users with legacy v2 (Teams Toolkit pre-v3) projects.

### CLI

```
atk upgrade --force
```

Source: [`packages/cli/src/commands/models/upgrade.ts`](../../../packages/cli/src/commands/models/upgrade.ts). Calls `core.phantomMigrationV3(inputs)` internally.

### VS Code

The Project tree view shows an **"Upgrade"** section when v2 markers are present (`isV3Project = false`):

- View ID: `teamsfx-project-and-check-upgradeV3`.
- Command: `fx-extension.checkProjectUpgrade` (title: "Upgrade Project").
- Shown only when `fx-extension.canUpgradeV3` context is true.

### What the upgrade does

The "phantom migration" transforms:

| v2 → v3 transformation |
|------------------------|
| `.fx/projectSettings.json` + `.fx/configs/*.json` → `m365agents.yml` + `m365agents.local.yml` |
| `.fx/configs/azure.parameters.{env}.json` → `infra/azure.parameters.{env}.json` |
| Manifest schema 1.x partial → 1.x → 1.16+ stable |
| Entry-points wired to v3 driver IDs |

The user sees per-step progress in the output channel. If a transformation fails, the toolkit:

1. Surfaces the failure with `displayMessage` + help link.
2. Suggests the manual edit needed (with code snippet).
3. **Does not proceed** beyond the failed step (transactional).

### Required input

| Input | Notes |
|-------|-------|
| `--force` | Required for CLI; bypasses the confirmation prompt |
| Workspace open in VS Code | Required for the VS Code path |

### Failure modes

- Custom v2 actions not auto-translatable → user-fixable error with a guide link.
- Already on v3 → "Project is already on v3" info message; no-op.
- Incompatible templates (e.g. very old SPFx) → blocked with explanation.

### v4 design implication

Open questions:

- Does v4 ship a `v3 → v4` upgrade path? (Currently the v3↔v4 split is opt-in via `TEAMSFX_V4_CORE`, no migration is needed.)
- Should the "phantom migration" pattern be retired now that v3 has been the default for years?

## See also

- [project-creation-flow.md](project-creation-flow.md) — the standard template-creation flow
- [../surfaces/cli-v3-command-reference.md](../surfaces/cli-v3-command-reference.md) — `atk list samples`, `atk new sample`, `atk upgrade`
- [../surfaces/vscode-extension-commands.md](../surfaces/vscode-extension-commands.md) — `fx-extension.openSamples`, `fx-extension.checkProjectUpgrade`
- [../../01-product/v3-feature-inventory.md](../../01-product/v3-feature-inventory.md) — samples in the capability list
