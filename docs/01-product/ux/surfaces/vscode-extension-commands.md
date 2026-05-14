# VS Code extension command catalog (v3)

> **Purpose.** Complete enumeration of the user-facing VS Code extension surface — commands, tree views, menus, settings, walkthroughs, Copilot chat participants, and locale coverage — as they ship in `packages/vscode-extension` v3.
>
> **v4 design status.** This page **is** an allowed input to v4 design (it describes the user-visible surface, not internals). When designing v4's VS Code surface, this is the catalogue to compare against — *not* the handler implementations under `packages/vscode-extension/src/handlers/`.
>
> **Source of truth.** [`packages/vscode-extension/package.json`](../../../packages/vscode-extension/package.json) `contributes` and `activationEvents`. If this page disagrees with that file, that file wins.

For the higher-level surface description and architecture layers, see [vscode-extension.md](vscode-extension.md). For per-command implementation conventions (handler shape, telemetry, error display), see [vscode-extension.instructions.md](../../../.github/instructions/vscode-extension.instructions.md).

## Activity bar

One container.

| Container ID | Title |
|--------------|-------|
| `teamsfx` | Microsoft 365 Agents Toolkit |

## Tree views

Source: `packages/vscode-extension/src/treeview/`.

| View ID | Title | Provider | When | Items |
|---------|-------|----------|------|-------|
| `teamsfx-accounts` | Accounts | `accountTreeViewProvider.ts` | `fx-extension.isTeamsFx` | M365AccountNode, AzureAccountNode |
| `teamsfx-environment` | Environment | `environmentTreeViewProvider.ts` | `fx-extension.isTeamsFx && !canUpgradeV3 && isWorkspaceTrusted` | Dynamic environment list |
| `teamsfx-development` | Development | `commandsTreeViewProvider.ts` | `fx-extension.isTeamsFx` | Create Project · View Samples · View Guides · Debug · Get Help (Copilot) |
| `teamsfx-lifecycle` | Lifecycle | `commandsTreeViewProvider.ts` | `fx-extension.isTeamsFx && !canUpgradeV3` | Provision · Deploy · Share (DA only) · Publish |
| `teamsfx-utility` | Utility | `commandsTreeViewProvider.ts` | `fx-extension.isTeamsFx && !canUpgradeV3` | Zip App Package · Validate Application · Publish in Developer Portal |
| `teamsfx-project-and-check-upgradeV3` | Upgrade | `officeDevTreeViewManager.ts` | `fx-extension.isTeamsFx && canUpgradeV3 && isWorkspaceTrusted` | Upgrade-related commands |
| `teamsfx-help-and-feedback` | Help and feedback | `commandsTreeViewProvider.ts` | `fx-extension.isTeamsFx && !canUpgradeV3` | Documentation · Get Started · Report Issues |
| `teamsfx-empty-project-with-chat` | Microsoft 365 Agents Toolkit | `commandsTreeViewProvider.ts` | `!isTeamsFx && !isOfficeAddIn && hideTeamsAgentPreviewTag` | Welcome content (no preview tag) |
| `teamsfx-empty-project-with-chat-preview` | Microsoft 365 Agents Toolkit | `commandsTreeViewProvider.ts` | `!isTeamsFx && !isOfficeAddIn && !hideTeamsAgentPreviewTag` | Welcome content (preview tag) |
| `teamsfx-officedev-development` | Development (Office) | `officeDevTreeViewManager.ts` | `fx-extension.isOfficeAddIn` | Office Add-in development commands |
| `teamsfx-officedev-lifecycle` | Lifecycle (Office) | `officeDevTreeViewManager.ts` | `fx-extension.isOfficeAddIn` | Office Add-in lifecycle commands |
| `teamsfx-officedev-utility` | Utility (Office) | `officeDevTreeViewManager.ts` | `fx-extension.isOfficeAddIn` | Office Add-in utility commands |
| `teamsfx-officedev-help-and-feedback` | Help and feedback (Office) | `officeDevTreeViewManager.ts` | `fx-extension.isOfficeAddIn` | Office Add-in help commands |

## Commands (full catalogue)

80 commands. Grouped by intent for readability; the source-of-truth is `packages/vscode-extension/package.json` `contributes.commands`.

### Project lifecycle (visible in command palette)

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.create` | Create New Agent/App | `!commandLocked` |
| `fx-extension.localdebug` | Debug | `!commandLocked && !isManifestOnlyOfficeAddIn` |
| `fx-extension.provision` | Provision | `isTeamsFx && isWorkspaceTrusted && !commandLocked` |
| `fx-extension.deploy` | Deploy | `isTeamsFx && isWorkspaceTrusted && !commandLocked` |
| `fx-extension.publish` | Publish | `isTeamsFx && isWorkspaceTrusted && !commandLocked` |
| `fx-extension.build` | Zip App Package | `isTeamsFx && isWorkspaceTrusted && !commandLocked` |
| `fx-extension.validateManifest` | Validate Application | `isTeamsFx && isWorkspaceTrusted` |
| `fx-extension.updatePreviewFile` | Update App | `isTeamsFx && isWorkspaceTrusted` |
| `fx-extension.publishInDeveloperPortal` | Publish to Store in Developer Portal | `isTeamsFx && isWorkspaceTrusted && isTDPIntegrationEnabled && !commandLocked` |

### Environment management

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.addEnvironment` | Create New Environment | `isTeamsFx && isWorkspaceTrusted && !commandLocked` |
| `fx-extension.addEnvironmentWithIcon` | Create New Environment (icon) | same as above; in tree view title |
| `fx-extension.refreshEnvironment` | Refresh | tree view title |
| `fx-extension.openResourceGroupInPortal` | Open in Portal (RG) | tree view item context |
| `fx-extension.openSubscriptionInPortal` | Open in Portal (Subscription) | tree view item context |
| `fx-extension.preview` | Preview App | tree view item context (provisioned env) |
| `fx-extension.previewWithIcon` | Preview App (icon) | tree view item context |
| `fx-extension.localdebugWithIcon` | Debug (icon) | tree view item context (local env) |
| `fx-extension.debugInTestToolWithIcon` | Debug in Microsoft 365 Agents Playground | tree view item context (testtool env) |

### Permissions and sharing

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.manageCollaborator` | Manage Microsoft 365 App Collaborators | `!commandLocked` |
| `fx-extension.grantPermission` | Add App Owners | tree view item context |
| `fx-extension.listCollaborator` | List App Owners | command-palette-hidden |
| `fx-extension.removeSharedAccess` | Remove access to the shared app | `isTeamsFx && isWorkspaceTrusted` |

### Account management

| Command ID | Title | Notes |
|------------|-------|-------|
| `fx-extension.cmpAccounts` | Accounts | command palette |
| `fx-extension.createAccount` | Create Account | tree view title |
| `fx-extension.signOut` | Sign Out | tree view item context |
| `fx-extension.m365AccountSettings` | Microsoft 365 Portal | tree view item context (signed-in M365) |
| `fx-extension.azureAccountSettings` | Azure Portal | tree view item context (signed-in Azure) |
| `fx-extension.m365SwitchTenant` | Switch tenant (M365) | tree view item context |
| `fx-extension.azureSwitchTenant` | Switch tenant (Azure) | tree view item context |
| `fx-extension.refreshSideloading` | Refresh (sideloading) | tree view item context |
| `fx-extension.checkSideloading` | Custom App Upload Disabled | tree view item context |
| `fx-extension.refreshCopilot` | Refresh (Copilot) | tree view item context |
| `fx-extension.checkCopilotCallback` | Copilot Access Disabled | tree view item context |
| `fx-extension.checkCopilotAccess` | Check Copilot Access | command-palette-hidden |
| `fx-extension.azureAccountSignOutHelp` | Azure account Sign Out is moved... | tree view item context |

### Declarative Agent capability extension

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.addPlugin` | Add Action | `isTeamsFx && isWorkspaceTrusted && !commandLocked && isDeclarativeCopilotApp` |
| `fx-extension.regeneratePlugin` | Regenerate Action | `isTeamsFx && isWorkspaceTrusted && !commandLocked && isDeclarativeCopilotApp && isKiotaNPMIntegrationEnabled` |
| `fx-extension.addKnowledge` | Add Capability | `isTeamsFx && isWorkspaceTrusted && !commandLocked && isDeclarativeCopilotApp` |
| `fx-extension.addAuthAction` | Add Configurations to Support Actions with Authentication | always |
| `fx-extension.updateActionWithMCP` | Update Action with MCP | `isTeamsFx && isWorkspaceTrusted` |
| `fx-extension.metaOSExtendToDA` | Extend the App to Declarative Agent | `isTeamsFx && isWorkspaceTrusted && !commandLocked && !isDeclarativeCopilotApp && isMetaOSAddinProject` |
| `fx-extension.setSensitivityLabel` | Set Sensitivity Label | `isSensitivityLabelEnabled` |
| `fx-extension.syncManifest` | Sync Manifest | `isSyncManifestEnabled` |

### Entra ID app management

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.updateAadAppManifest` | Update Microsoft Entra App | `isTeamsFx && isAadManifestEnabled && isWorkspaceTrusted && !isSPFx` |
| `fx-extension.openPreviewAadFile` | Preview Microsoft Entra Manifest File | editor title |
| `fx-extension.convertAadToNewSchema` | Upgrade Microsoft Entra Manifest to New Schema | always |
| `fx-extension.openSchema` | Open Manifest Schema | command-palette-hidden |

### Project upgrade

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.checkProjectUpgrade` | Upgrade Project | `initialized && isWorkspaceTrusted` |
| `fx-extension.migrateTeamsTabApp` | Upgrade Teams JS SDK and Code References | `isWorkspaceTrusted` |

### SPFx-specific

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.addWebpart` | Add SPFx web part | `isTeamsFx && isWorkspaceTrusted && !commandLocked && isSPFx` |

### Copilot Chat integration (`@m365agents`)

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.invokeChat` | Get Help from GitHub Copilot | `isChatParticipantUIEntriesEnabled && hideTeamsAgentPreviewTag` |
| `fx-extension.invokeChatWithPreviewTag` | Get Help from GitHub Copilot (Preview) | `isChatParticipantUIEntriesEnabled && !hideTeamsAgentPreviewTag` |
| `fx-extension.teamsAgentTroubleshootSelectedText` | Resolve selected text with @m365agents | `editor.hasSelection && isChatParticipantUIEntriesEnabled && hideTeamsAgentPreviewTag` |
| `fx-extension.teamsAgentTroubleshootSelectedTextWithPreviewTag` | Resolve selected text with @m365agents (preview) | `editor.hasSelection && isChatParticipantUIEntriesEnabled && !hideTeamsAgentPreviewTag` |
| `fx-extension.teamsAgentTroubleshootError` | Resolve with @m365agents | `isChatParticipantUIEntriesEnabled` |
| `fx-extension.findSimilarIssue` | Similar Issues | always |
| `fx-extension.openTeamsAgentWalkthrough` | Teams Agent Walkthrough | `isChatParticipantUIEntriesEnabled` |

### Help and discovery

| Command ID | Title |
|------------|-------|
| `fx-extension.openDocument` | Documentation |
| `fx-extension.openSamples` | View Samples |
| `fx-extension.openWelcome` | Get Started |
| `fx-extension.buildIntelligentAppsWalkthrough` | Build intelligent apps |
| `fx-extension.selectTutorials` | View How-to Guides |
| `fx-extension.openReportIssues` | Report Issues on GitHub |
| `fx-extension.openAccountLink` / `openLifecycleLink` / `openDevelopmentLink` / `openEnvLink` / `openHelpFeedbackLink` / `openDocumentLink` | "Get More Info About …" links per tree view |
| `fx-extension.openOfficeDevLifecycleLink` / `openOfficeDevDevelopmentLink` / `openOfficeDevHelpFeedbackLink` | Office-specific link variants |

### Office Add-in specific

| Command ID | Title | Enablement |
|------------|-------|------------|
| `fx-extension.stopDebugging` | Stop Previewing Your Office Add-in | `isOfficeAddIn && !isManifestOnlyOfficeAddIn` |
| `fx-extension.selectAndDebug` | Select and Start Debugging App | (editor title/run) |

### Other

| Command ID | Title | Notes |
|------------|-------|-------|
| `fx-extension.openAppManagement` | Developer Portal for Teams | `isTeamsFx && !isTDPIntegrationEnabled` |
| `fx-extension.deployManifestFromCtxMenu` | Update App | command-palette-hidden |

## Settings

| Setting key | Type | Default | Purpose |
|-------------|------|---------|---------|
| `M365AgentsToolkit.sovereignCloudEnvironment` | enum (`""`, `"GCC M"`, `"GCC H"`, `"DoD"`) | `""` | Sovereign cloud target. Reload required. For GCC H / DoD, also set `microsoft-sovereign-cloud.environment` to `USGovernment`. |
| `M365AgentsToolkit.logLevel` | enum (`"Info"`, `"Verbose"`, `"Debug"`) | `"Info"` | Log verbosity. `Debug` includes HTTP detail and account info; `Verbose` adds per-action progress; `Info` is summary only. |
| `M365AgentsToolkit.enableDeclarativeAgentMCPSupport` | boolean | `true` | Enable MCP server actions for DA. |
| `M365AgentsToolkit.enableLaunchAgentForTeamsInCopilot` | boolean | `false` | Launch Agent for Teams in Microsoft 365 Copilot. |
| `M365AgentsToolkit.enableDeclarativeAgentInOfficeAddIn` | boolean | `true` | Allow DA in Office Add-in projects. |
| `M365AgentsToolkit.enableCustomFunctionShortcutInOfficeAddIn` | boolean | `true` | Enable Custom Function + Shortcut in Office Add-in projects. |

## Keybindings

None defined.

## Walkthroughs

| Walkthrough ID | Title | Steps | When |
|----------------|-------|-------|------|
| `buildIntelligentApps` | Build intelligent apps | 6 | `!isChatParticipantUIEntriesEnabled` |
| `buildIntelligentAppsWithChat` | Build intelligent apps | 8 | `isChatParticipantUIEntriesEnabled` |
| `teamsToolkitGetStarted` | Teams Toolkit Get Started | 6 | `!isChatParticipantUIEntriesEnabled` |
| `teamsToolkitGetStartedWithChat` | Teams Toolkit Get Started | 9 | `isChatParticipantUIEntriesEnabled` |
| `teamsAgentGetStarted` | Teams Agent Walkthrough | 8 | `isChatParticipantUIEntriesEnabled` |

## Copilot chat participants

| Participant ID | Name | Subcommands |
|----------------|------|-------------|
| `ms-teams-vscode-extension.office` | `office` | `create` · `generatecode` · `nextstep` |

### `@m365agents` — registered by a separate companion extension

Commands like `fx-extension.invokeChat`, `fx-extension.teamsAgentTroubleshootError`, `fx-extension.teamsAgentTroubleshootSelectedText` open Copilot Chat with `@m365agents` queries pre-filled (see [`packages/vscode-extension/src/handlers/copilotChatHandlers.ts`](../../../packages/vscode-extension/src/handlers/copilotChatHandlers.ts)). The `@m365agents` participant itself is **not registered by this extension's `package.json`** — it is provided by a separate companion extension that the user installs alongside (or independently of) this one. From the user's perspective `@m365agents` is the entry point; from this repo's perspective it is an external dependency surfaced through link buttons (see `aka.ms/install-m365agents`) and pre-filled queries.

The single chat participant registered *by this extension* is `office` (Office Add-in helper).

### Chat skills

This extension also contributes a chat skill at `packages/vscode-extension/skills/microsoft-365-agents-toolkit/SKILL.md` via `contributes.chatSkills`. (One skill file.)

## Locale coverage

16 NLS files: base (`en-US`) + `cs`, `de`, `es`, `fr`, `it`, `ja`, `ko`, `pl`, `pt-BR`, `ru`, `tr`, `zh-cn`, `zh-Hans`, `zh-Hant`, `zh-tw`.

## Menu surface (per type)

`contributes.menus` entries by menu type:

| Menu | Entries |
|------|---------|
| `commandPalette` | 45 |
| `view/title` | 11 |
| `view/item/context` | 23 |
| `editor/context` | 2 |
| `editor/title` | 2 |
| `editor/title/run` | 1 |
| **Total** | **84** |

## Other contributes

| Section | Entries |
|---------|---------|
| `icons` | 10 custom icons |
| `viewsWelcome` | 3 welcome-view contents (empty-project / empty-project-preview / project-and-check-upgradeV3) |
| `chatSkills` | 1 (points at `skills/microsoft-365-agents-toolkit/SKILL.md`) |
| `languages` | 1 (`teamsfx-toolkit-output` — syntax highlighting for the toolkit's output channel) |
| `grammars` | 1 (matching `teamsfx-toolkit-output` language) |
| `taskDefinitions` | 1 contributes block (defines task `type` schemas; see [local-debug-and-prereqs.md](local-debug-and-prereqs.md)) |
| `problemMatchers` | 7 — frontend, backend, bot, auth, tunnel, ngrok, local-tunnel |

## Activation events

```
onCommand:fx-extension.create
onCommand:fx-extension.createFromWalkthrough
onCommand:fx-extension.openLifecycleTreeview
onCommand:fx-extension.openDocument
onCommand:fx-extension.openReadMe
onCommand:fx-extension.openSamples
onCommand:fx-extension.openWelcome
onCommand:fx-extension.buildIntelligentAppsWalkthrough
onCommand:fx-extension.selectAndDebug
onCommand:fx-extension.selectTutorials
onCommand:fx-extension.signinM365
onCommand:fx-extension.validate-getStarted-prerequisites
onCommand:workbench.action.tasks.runTask
onCommand:workbench.action.debug.start
workspaceContains:/.fx/**/*
workspaceContains:teamsapp.yml
workspaceContains:m365agents*.yml
workspaceContains:/**/manifest.json
workspaceContains:/manifest*.xml
onView:teamsfx-empty-project-with-chat
onUri
```

Note: `fx-extension.signinM365` is in the activation events but does not appear in `contributes.commands` — it is registered programmatically by the extension at activation time.

## Counts

- **80 commands** in `contributes.commands`
- 13 tree views
- 5 walkthroughs
- 6 settings
- 1 chat participant **registered by this extension** (`office`); `@m365agents` is provided by a separate companion extension
- 84 menu entries across 6 menu types
- 10 custom icons, 3 welcome views, 1 chat skill, 1 language + grammar, 7 problem matchers
- 16 locales
- 21 activation triggers

## See also

- [vscode-extension.md](vscode-extension.md) — architecture layers and singleton patterns
- [cli-v3-command-reference.md](cli-v3-command-reference.md) — equivalent CLI surface
- [local-debug-and-prereqs.md](local-debug-and-prereqs.md) — task types, problem matchers, dev tunnels, Test Tool
- [copilot-chat-participant.md](copilot-chat-participant.md) — the `office` participant and the external `@m365agents` participant
- [../information-architecture.md](../information-architecture.md) — how this surface fits the broader IA
- [../../01-product/v3-feature-inventory.md](../../01-product/v3-feature-inventory.md) — what these commands let users build
