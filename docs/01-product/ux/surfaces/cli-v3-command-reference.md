# CLI v3 command reference

> **Purpose.** Complete enumeration of `atk` (v3) commands, arguments, options, examples — the user-facing CLI surface as it ships in `packages/cli`.
>
> **v4 design status.** This page **is** an allowed input to v4 design (it describes the user-facing CLI surface, not internals). When designing v4 CLI commands, this is the catalogue to compare against — *not* `packages/cli/src/commands/models/` handler code or the engine implementation under `packages/cli/src/cmds/`.
>
> **Source of truth.** [`packages/cli/src/commands/models/root.ts`](../../../packages/cli/src/commands/models/root.ts) is the single place that wires the root command tree. If this page disagrees with that file, that file wins.

For the higher-level CLI architecture, see [cli.md](cli.md). For per-command implementation conventions (handler shape, telemetry), see [cli.instructions.md](../../../.github/instructions/cli.instructions.md).

## Command tree

All commands below are wired **flat at the root** (no `teamsapp` parent group). Subcommand groups (`auth`, `add`, `regenerate`, `share`, `env`, `collaborator`, `list`, `entra-app`, `set`, `new`) own their own subcommand arrays.

```
atk
├── auth (alias: account)
│   ├── login
│   │   ├── m365
│   │   └── azure
│   ├── logout <service>
│   └── show
├── new
│   └── sample <samples-id>
├── add
│   ├── action
│   ├── auth-config
│   ├── capability
│   └── spfx-web-part
├── regenerate
│   └── action
├── provision
├── deploy
├── share
│   └── remove
├── preview
├── env
│   ├── add <env-name>
│   ├── list
│   └── reset
├── collaborator (alias: permission)
│   ├── grant
│   └── status
├── upgrade
├── init                       (feature flag: GenerateConfigFiles)
├── list
│   ├── samples
│   └── templates
├── update                     (Teams app manifest update)
├── validate                   (Teams app manifest/package validate)
├── package                    (Teams app package zip)
├── publish                    (Teams app publish to tenant catalog)
├── doctor                     (Teams app prerequisite check)
├── entra-app
│   └── update
├── install (alias: sideloading)
├── uninstall (alias: unacquire)
├── launchinfo
├── set                        (feature flag: SensitivityLabelEnabled)
│   └── sensitivitylabel
└── help
```

> **Note on Teams app commands.** `update`, `validate`, `package`, `publish`, `doctor` are exported from `packages/cli/src/commands/models/teamsapp/*.ts` but their `name` fields are bare (`"update"`, `"validate"`, etc.) and they are wired **directly to root**. Users invoke them as `atk validate`, NOT `atk teamsapp validate`. The `teamsapp/` directory is a source-organisation convention only.
>
> Standalone files `packages/cli/src/commands/models/{validate,package,publish}.ts` exist but are **not wired into root** — they are dead code retained for historical reasons. The wired versions live under `teamsapp/`.

## Reserved root options

Defined in [`packages/cli/src/commands/models/root.ts`](../../../packages/cli/src/commands/models/root.ts):

| Option | Short | Type | Default |
|--------|-------|------|---------|
| `--version` | `-v` | boolean | `false` |
| `--help` | `-h` | boolean | `false` |
| `--interactive` | `-i` | boolean | `true` |
| `--debug` | — | boolean | `false` |
| `--verbose` | — | boolean | `false` |
| `--telemetry` | — | boolean | `true` |

## Per-command detail

### `atk auth login m365`

- **Description:** Log in to Microsoft 365.
- **Options:** `--tenant <id>` (default: `""`)

### `atk auth login azure`

- **Description:** Log in to Azure.
- **Options:**
  - `--tenant <id>` (default: `""`)
  - `--service-principal` (default: `false`)
  - `-u, --username <name>` (default: `""`)
  - `-p, --password <secret>` (default: `""`)
  - `--claims-challenge <token>` (default: `""`)
- **Examples:**
  - `atk auth login azure -i false --service-principal -u USERNAME -p SECRET --tenant TENANT_ID`
  - `atk auth login azure -i false --service-principal -u USERNAME -p "C:/Users/mycert.pem" --tenant TENANT_ID`

### `atk auth logout <service>`

- **Argument:** `service` — `azure` or `m365` (required)

### `atk auth show`

- Show current account info. No args / options.

### `atk new`

- **Description:** Create a new project.
- **Options:** `-c, --capability`, `-n, --name`, `--folder-path`, `-l, --programming-language`, `-i, --interactive`, plus 8+ template-driven dynamic options.
- **Examples:**
  - `atk new -c declarative-agent -n myagent -i false`
  - `atk new -c basic-custom-engine-agent -l typescript -n mycea -i false`

### `atk new sample <samples-id>`

- **Argument:** `samples-id` (required)
- **Options:** `--folder-path`, `-i, --interactive`

### `atk add action` / `add auth-config` / `add capability` / `add spfx-web-part`

- Each takes `--folder` plus dynamic question-driven options. Used post-scaffold to extend an existing project.

### `atk regenerate action`

- **Options:** `--folder` plus dynamic options. Re-generates a plugin action (e.g. after spec change).

### `atk provision`

- **Options:**
  - `--env <name>` (required)
  - `--folder <path>` (required)
  - `--resource-group <name>` (hidden)
  - `--region <name>` (hidden)
  - `--ignore-env-file` (default: `false`)

### `atk deploy`

- **Options:**
  - `--env <name>` (required)
  - `--folder <path>` (required)
  - `--ignore-env-file` (default: `false`)
  - `--config-file-path <path>`

### `atk share`

- **Description:** Share a project / app with collaborators.
- **Options:**
  - `--env <name>` (required)
  - `--folder <path>` (required)
  - `--scope <tenant|users>` (required)
  - `--email <list>` (comma-separated)
  - `--ignore-env-file` (default: `false`)
- **Examples:**
  - `atk share` (interactive)
  - `atk share --scope tenant -i false`
  - `atk share --scope users --email 'a@example.com,b@example.com' -i false`

### `atk share remove`

- **Description:** Revoke shared access.
- **Options:**
  - `--env <name>` (required)
  - `--folder <path>` (required)
  - `--users <list>`
  - `--owners <list>`
  - `--ignore-env-file` (default: `false`)

### `atk preview`

- **Description:** Preview app locally.
- **Options:** 15+ including `--env`, `--manifest-file`, `-c, --run-command`, `-p, --running-pattern`, `-o, --open-only`, `-b, --browser`, `-ba, --browser-arg`, `-ep, --exec-path`, etc.

### `atk env add <env-name>`

- **Argument:** `env-name` (required)
- **Options:** `--env-file <path>` (source to copy from), `--folder <path>` (required)

### `atk env list`

- **Options:** `--folder <path>` (required)

### `atk env reset`

- **Options:**
  - `--env <name>`
  - `--env-file <path>`
  - `--ignore-keys <list>`
  - `--folder <path>` (required)

### `atk collaborator grant`

- **Options:**
  - `--env <name>` (required)
  - `--folder <path>` (required)
  - `--email <addr>` (required)
  - `--manifest-file <path>`
  - `--entra-app-manifest-file <path>`
  - `--agent` (grant as agent owner; default: `false`)
  - `--ignore-env-file` (default: `false`)
- **Examples:**
  - `atk collaborator grant -i false --manifest-file ./appPackage/manifest.json --env dev --email other@email.com`
  - `atk collaborator grant -i false --agent true --env dev --email other@email.com`

### `atk collaborator status`

- **Options:** Same as `grant` plus `-a, --all` (list all; default `false`).

### `atk upgrade`

- **Options:** `-f, --force` (default `false`, required)

### `atk validate` (orphaned standalone)

> A separate `validate.ts` exists at `packages/cli/src/commands/models/validate.ts` with a different option set (`--app-package-file`, `--manifest-file`, `--env`, `--folder`). **It is not wired into root** — the wired `atk validate` is the `teamsapp/validate.ts` version above. The standalone file is dead code.

### `atk package` (orphaned standalone)

> A separate `package.ts` exists at `packages/cli/src/commands/models/package.ts`. **It is not wired into root.** Dead code.

### `atk publish` (orphaned standalone)

> A separate `publish.ts` exists at `packages/cli/src/commands/models/publish.ts`. **It is not wired into root.** Dead code.

### `atk update`

Wired from [`packages/cli/src/commands/models/teamsapp/update.ts`](../../../packages/cli/src/commands/models/teamsapp/update.ts). Update Teams app manifest.

- **Options:**
  - `--manifest-file <path>`
  - `--package-file <path>`
  - `-op, --output-package-file <path>`
  - `-of, --output-folder <path>`
  - `--env <name>`
  - `--env-file <path>`
  - `--folder <path>` (required)

### `atk validate`

Wired from [`packages/cli/src/commands/models/teamsapp/validate.ts`](../../../packages/cli/src/commands/models/teamsapp/validate.ts). Validate Teams app manifest or package.

- **Options:**
  - `--manifest-file <path>`
  - `--package-file <path>`
  - `-op, --output-package-file <path>`
  - `-of, --output-folder <path>`
  - `--validation-method <validate-schema|validate-app>` (default: `validate-schema`)
  - `--env <name>`
  - `--env-file <path>`
  - `--folder <path>` (required)
- **Examples:**
  - `atk validate --package-file ./appPackage/build/appPackage.dev.zip --validation-method validate-app --env dev --folder .`
  - `atk validate --manifest-file ./appPackage/manifest.json --env dev --folder .`

### `atk package`

Wired from [`packages/cli/src/commands/models/teamsapp/package.ts`](../../../packages/cli/src/commands/models/teamsapp/package.ts). Zip Teams app package.

- **Options:**
  - `--manifest-file <path>`
  - `-op, --output-package-file <path>`
  - `-of, --output-folder <path>`
  - `--env <name>`
  - `--env-file <path>`
  - `--folder <path>` (required)
- **Example:** `atk package --env dev --folder .`

### `atk publish`

Wired from [`packages/cli/src/commands/models/teamsapp/publish.ts`](../../../packages/cli/src/commands/models/teamsapp/publish.ts). Publish Teams app to tenant catalog.

- **Options:**
  - `--manifest-file <path>`
  - `--package-file <path>`
  - `-op, --output-package-file <path>`
  - `-of, --output-folder <path>`
  - `--env <name>`
  - `--env-file <path>`
  - `--folder <path>` (required)

### `atk doctor`

Wired from [`packages/cli/src/commands/models/teamsapp/doctor.ts`](../../../packages/cli/src/commands/models/teamsapp/doctor.ts). Diagnose Node.js, Azure Functions Core Tools, SSL cert, M365 account. No args / options. See the [DepsChecker section](local-debug-and-prereqs.md#dependency-checker) for the full list of probes.

### `atk entra-app update`

- **Options:**
  - `--entra-app-manifest-file <path>`
  - `--env <name>` (required)
  - `--folder <path>` (required)

### `atk install` (alias: `sideloading`)

- **Options:**
  - `--file-path <path>` (zipped Teams app package)
  - `--xml-path <path>` (Outlook add-in manifest)
  - `--scope <Personal|Shared>`
- **Examples:**
  - `atk install --file-path appPackage.zip`
  - `atk install --file-path appPackage.zip --scope Shared`
  - `atk install --xml-path manifest.xml`

### `atk uninstall` (alias: `unacquire`)

- **Options:**
  - `--mode <title-id|manifest-id|env>` (required)
  - `--title-id <id>` (M365 title ID)
  - `--manifest-id <id>`
  - `--env <name>` (when `--mode env`)
  - `--folder <path>`
  - `--options <list>` — cleanup scope: `m365-app`, `app-registration`, `bot-framework-registration`
- **Examples:**
  - `atk uninstall -i false --mode title-id --title-id U_xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
  - `atk uninstall -i false --mode manifest-id --manifest-id xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx --options 'm365-app,app-registration,bot-framework-registration'`
  - `atk uninstall -i false --mode env --env dev --options 'm365-app,app-registration,bot-framework-registration' --folder ./myapp`

### `atk launchinfo`

- **Options:**
  - `--title-id <id>`
  - `--manifest-id <id>` (at least one required)

### `atk list samples`

- **Options:**
  - `-t, --tag <tag>`
  - `--format <table|json>` (default: `table`)
  - `--description` (show full description; default: `false`)

### `atk list templates`

- **Options:** `--format <table|json>` (default: `table`)

### `atk init` (feature-flagged)

- **Description:** Initialise config files for an existing project.
- **Options:**
  - `--playground` (default: `false`)
  - `--local-debug` (default: `true`)
  - `--remote-deploy` (default: `false`)
  - `--programming-language <typescript|python>`
  - `--manifest-file <path>` (required)
  - `--folder <path>`
- **Examples:**
  - `atk init`
  - `atk init --remote true`
  - `atk init --playground true --local false`

### `atk set sensitivitylabel`

- **Options:**
  - `--sensitivity-label <name|id>` (required)
  - `--env <name>` (required)
  - `--folder <path>`

## Interactive vs non-interactive

The default is **interactive** (`-i true`). The CLI engine discards most options and prompts the user via the question model. Non-interactive mode (`-i false`) requires all needed options on the command line — missing required values surface as `MissingRequiredOptionError`.

### `CI_ENABLED` env var

The `CI_ENABLED` environment variable **does** exist and **does** override interactivity:

- [`packages/cli/src/userInteraction.ts`](../../../packages/cli/src/userInteraction.ts) — `if (process.env.CI_ENABLED === "true") return false`.
- [`packages/cli/src/commands/engine.ts`](../../../packages/cli/src/commands/engine.ts) — sets `globalOptionValues.interactive = false` when `CI_ENABLED === "true"`.

The override sequence is:

1. Per-command `defaultInteractiveOption` (if set).
2. User-supplied `--interactive` / `-i` flag.
3. `CI_ENABLED=true` env var — overrides everything to `false`.

Other CI/CD env vars detected for telemetry (not for interactivity): `GITHUB_ACTIONS`, `JENKINS_URL`, `BUILD_URL`, `BUILD_SOURCEBRANCHNAME`, `AGENT_BUILDDIRECTORY`.

Auth-related env vars consumed by the CLI: `M365_ACCOUNT_NAME`, `M365_ACCOUNT_PASSWORD`, `AZURE_ACCOUNT_NAME`, `AZURE_ACCOUNT_PASSWORD`, `AZURE_SERVICE_PRINCIPAL_ID`, `AZURE_SERVICE_PRINCIPAL_SECRET`, `AZURE_SUBSCRIPTION_ID`, `AZURE_TENANT_ID`, `AZURE_ACCOUNT_OBJECT_ID`, `M365_TENANT_ID`, `TEAMS_APP_TENANT_ID`. Plus `TEAMSFX_CONFIG_FILE_PATH` (consumed by deploy) and `TEAMSFX_CLI_BIN_NAME` (used in help-text examples).

### Commands with `defaultInteractiveOption: false`

These commands default to non-interactive (the inverse of the global default):

- `validate` (teamsapp/validate)
- `package` (teamsapp/package)
- `publish` (teamsapp/publish)
- `update` (teamsapp/update)
- `doctor` (teamsapp/doctor)
- `auth logout`
- `launchinfo`
- `list samples`
- `set sensitivitylabel`
- `env add`
- `entra-app update`
- `init`

All other commands default to interactive (`true`).

`reservedOptionNamesInInteractiveMode` (per command) lists options always parsed even in interactive mode (e.g. `--force`, `--all`).

## Counts

- 24 top-level commands (22 always-on + `init` and `set` behind feature flags `GenerateConfigFiles` and `SensitivityLabelEnabled`).
- 13 subcommand leaves (across `auth`, `add`, `regenerate`, `share`, `env`, `collaborator`, `list`, `entra-app`, `set`, `new`).
- ~37 leaf commands total.
- ~120+ options across all commands (excluding dynamic question-derived options).
- 6 reserved root options.
- 4 alias pairs verified: `auth↔account`, `install↔sideloading`, `uninstall↔unacquire`, `collaborator↔permission`.
- 0 commands with `hidden: true`. (Some individual *options* are hidden — e.g. `provision` has hidden `--resource-group` and `--region`.)

## See also

- [cli.md](cli.md) — architecture, engine phases, login flows
- [vscode-extension-commands.md](vscode-extension-commands.md) — equivalent VS Code surface
- [local-debug-and-prereqs.md](local-debug-and-prereqs.md) — `atk preview`, `atk doctor`, dev tunnels, Test Tool, prerequisites
- [../flows/README.md](../flows/README.md) — end-to-end UX flows that traverse these commands
- [../../01-product/v3-feature-inventory.md](../../01-product/v3-feature-inventory.md) — what these commands let users build
