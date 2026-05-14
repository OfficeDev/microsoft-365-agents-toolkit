# Information architecture

How users navigate each surface to reach the same engine operations.

## VS Code extension

```
Activity Bar: Microsoft 365 Agents
  └─ Tree views (TreeViewManager)
      ├─ DEVELOPMENT
      │   ├─ Create a New Agent / App
      │   ├─ View Samples
      │   ├─ View How-to Guides
      │   └─ Documentation
      ├─ ENVIRONMENT
      │   ├─ Provision in the Cloud
      │   ├─ Deploy to the Cloud
      │   └─ Publish to Tenant
      ├─ ACCOUNT
      │   ├─ Sign in to Microsoft 365
      │   └─ Sign in to Azure
      └─ FEEDBACK / Help

Command Palette: "Teams: ..." prefix exposes all commands.
Copilot Chat: @teamsapp participant for AI-assisted authoring.
Webviews: project-creation panel, sample gallery (React + Fluent UI v8 + Vite).
```

Source: `packages/vscode-extension/src/treeview/`.

## Visual Studio extension

VS surfaces commands through:

- **Project templates** in the New Project dialog (under "Microsoft 365 Agents Toolkit").
- A **tool window** for environment selection and lifecycle commands.
- **Solution Explorer** context menus for individual project actions.

## CLI

```
atk
├── new                  # interactive scaffold (or atk new <category> <template> --options)
│   ├── da
│   ├── cea
│   ├── ai
│   ├── me
│   ├── bot
│   ├── tab
│   ├── connector
│   └── addin
├── provision
├── deploy
├── publish
├── account
│   ├── login m365
│   └── login azure
├── env
│   ├── list
│   ├── add
│   └── reset
├── teamsapp
│   ├── validate
│   └── package
├── permission
├── add
└── list templates
```

The v4 CLI generates the `new` subtree from `TemplateRegistry` — adding a `TemplateDescriptor` adds a CLI command automatically. See [`packages/cli-next/src/commands/factory.ts`](../../packages/cli-next/src/commands/factory.ts).

## Copilot Chat participant

`@teamsapp` is registered as a chat participant in the VS Code extension. It routes natural-language requests through fx-core operations and provides:

- Project creation guidance
- Capability discovery ("which template should I use for...")
- Troubleshooting assistance (renders `showError` action buttons inline)

See `packages/vscode-extension/src/chat/`.
