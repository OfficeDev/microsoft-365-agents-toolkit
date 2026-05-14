# Project creation flow

## VS Code

```
Tree view → "Create a New Agent / App"
  ↓
Webview opens (React / Fluent UI v8)
  ↓
User picks: Type → Capability → Template → Language → App name → Folder
  ↓
Engine: createProject (v3) OR createProjectInteractive (v4)
  ↓
Scaffold: download ZIP → filter → render Mustache → write to folder
  ↓
VS Code opens the folder; tree view refreshes
```

## CLI (interactive)

```
$ atk new
  ↓
Inquirer prompts (same question tree as VS Code)
  ↓
Engine: createProjectInteractive (v4) → scaffoldTemplates
  ↓
Files written under chosen folder
  ↓
PostAction: "open in VS Code? (y/n)"
```

## CLI (non-interactive)

```
$ atk new da basic --app-name MyAgent --folder ./my-agent --language common
  ↓
buildNewCommands resolves the leaf command for da/basic
  ↓
Action: createProjectAction → runOperation(createProjectOp, inputs)
  ↓
Files written; exit code 0
```

## Question tree — TS/JS path (illustrative)

```
projectType
  ├─ "agent" → templateName
  │     ├─ "da/basic" → appName → folder
  │     ├─ "cea/basic" → language → appName → folder
  │     ├─ "ai/chat-bot" → llmProvider → language → appName → folder
  │     └─ ...
  └─ "tab" → ...
```

Per-template extra questions (e.g. `llmProvider`, `foundryEndpoint`, `graphConnectorTenantId`) are declared in the corresponding `TemplateDescriptor`'s `questions?: QuestionSpec[]` and surfaced automatically by `buildQuestionTree`.

## Outputs

| Artifact | Location |
|----------|----------|
| Manifest | `appPackage/manifest.json` (Mustache-rendered) |
| Lifecycle YAML | `m365agents.yml`, `m365agents.local.yml` |
| Bicep | `infra/azure.bicep`, `infra/azure.parameters.json` |
| Source code | `src/` |
| Env files | `env/.env.{envName}` (typically `dev` and `local`) |
| `.vscode/` | `launch.json`, `tasks.json` for F5 |
| Tracking ID | `teamsAppTenantId` etc. in env files (set during provision) |
