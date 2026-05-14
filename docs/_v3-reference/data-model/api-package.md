# `packages/api` — v3 public exports

> **FORBIDDEN AS v4 DESIGN INPUT.** Archival catalogue only. See [`../README.md`](../README.md).

Source: [`packages/api/src/`](../../../packages/api/src/). Re-exports `neverthrow` (`Result`, `ok`, `err`) and `@microsoft/app-manifest` from `index.ts`.

## Top-level files and folders

```
packages/api/src/
├── cli.ts
├── constants.ts
├── context.ts
├── error.ts
├── generator.ts
├── index.ts        — barrel; re-exports all of the above + neverthrow + app-manifest
├── types.ts
├── qm/             — question model (see qm/index.ts)
├── schemas/        — JSON schemas
└── utils/          — Tools interface + provider re-exports
```

## `error.ts`

| Export | Purpose |
|--------|---------|
| `FxError` | Error interface — `source`, `timestamp`, telemetry properties, recommendation operation |
| `UserError` | User-recoverable error class |
| `SystemError` | Non-recoverable system error class |
| `ErrorOptionBase` | Base error options |
| `UserErrorOptions` | Extends `ErrorOptionBase` with `helpLink` |
| `SystemErrorOptions` | Extends `ErrorOptionBase` with `issueLink` |

## `types.ts`

| Export | Purpose |
|--------|---------|
| `OptionItem` | Display option: `id`, `label`, `description`, `detail`, `buttons[]` |
| `Void` | Empty object type |
| `EnvMeta` | Environment metadata: `name`, `local` flag, `sideloading` flag |
| `Inputs` | `platform`, `projectPath`, `projectId`, `nonInteractive`, `correlationId`, `agent`, `apiAuthData[]` |
| `InputsWithProjectPath` | `Inputs` with required `projectPath` |
| `CreateProjectInputs` | `Inputs` with required `app-name`, `folder` |
| `DeepReadonly` | Recursive readonly utility |
| `MaybePromise<T>` | `T \| Promise<T>` |
| `Settings` | Tooling settings: `version`, `trackingId` |
| `ManifestCapability` | Discriminated union: `staticTab`, `configurableTab`, `Bot`, `MessageExtension`, `WebApplicationInfo` |
| `AuthInfo` | `serverUrl`, `authName`, `authType` |
| `ApiOperation` | `id`, `label`, `groupName`, `data: AuthInfo`, `detail` |
| `Warning` | `type`, `content`, optional `data` |

## `context.ts`

| Export | Purpose |
|--------|---------|
| `Context` | Runtime context: `userInteraction`, `logProvider`, `telemetryReporter`, `expServiceProvider`, `tokenProvider`, `projectPath`, `templateVariables` |

## `generator.ts`

| Export | Purpose |
|--------|---------|
| `IGenerator` | Component interface: `run(context, inputs, destinationPath): Promise<Result<GeneratorResult, FxError>>` |
| `GeneratorResult` | Output: optional `warnings[]` |

## `cli.ts`

| Export | Purpose |
|--------|---------|
| `CLICommand` | Declarative command: `name`, `aliases`, `fullName`, `version`, `description`, `arguments`, `options`, `commands`, `handler`, `telemetry`, `hidden`, `examples` |
| `CLIFoundCommand` | `CLICommand` with required `fullName` |

## `constants.ts`

| Export | Purpose |
|--------|---------|
| `Platform` | Enum: `VSCode`, `CLI`, `VS`, `CLI_HELP` |
| `Stage` | Enum: `create`, `build`, `debug`, `provision`, `deploy`, `package`, `publish`, `share`, `createEnv`, `addFeature`, etc. |
| `ConfigFolderName` | `"fx"` |
| `AppPackageFolderName` | `"appPackage"` |
| Misc string constants | Folder/file names |

## `qm/index.ts`

Re-exports `question.ts`, `ui.ts`, `validation.ts`.

### `qm/question.ts`

| Export | Purpose |
|--------|---------|
| `FunctionRouter` | Namespace + method routing |
| `Func` | Extends `FunctionRouter` with optional params |
| `LocalFunc<T>` | `(inputs: Inputs) => T \| Promise<T>` |
| `OnSelectionChangeFunc` | Selection change callback |
| `StaticOptions` | `string[] \| OptionItem[]` |
| `DynamicOptions` | `LocalFunc<StaticOptions>` |
| `BaseQuestion` | Base shape: `name`, `title`, `value`, `default`, `step`, `totalSteps`, `validation`, `forgetLastValue` |

### `qm/ui.ts`

| Export | Purpose |
|--------|---------|
| `UIConfig<T>` | Base UI config: `name`, `title`, `placeholder`, `prompt`, `step`, `totalSteps`, `default`, `validation`, `buttons` |
| `ConfirmConfig` | Boolean UI with `transformer` |
| `SingleSelectConfig` | Extends `UIConfig` with `options`, `returnObject` |

## `utils/`

`packages/api/src/utils/index.ts` defines the `Tools` interface and re-exports the provider modules:

| Export | Purpose |
|--------|---------|
| `Tools` interface | Aggregates `LogProvider`, `TelemetryReporter`, `TokenProvider`, `UserInteraction`, `ExpServiceProvider` — the v3 `TOOLS` global singleton type |
| `login.ts` | `M365TokenProvider`, `AzureAccountProvider`, `TokenProvider` interfaces |
| `log.ts` | `LogProvider`, `LogLevel` |
| `telemetry.ts` | `TelemetryReporter` |
| `tree.ts` | `TreeItem`, `TreeProvider` |
| `crypto.ts` | `CryptoProvider` |
| `exp.ts` | `ExpServiceProvider` (experimentation) |
| `userInteraction.ts` | `UserInteraction` interface (singleSelect, multiSelect, input, confirm, etc.) |

## `schemas/`

JSON schemas exported for editor IntelliSense / external validation. Currently includes manifest schemas re-exported from `@microsoft/app-manifest`.
