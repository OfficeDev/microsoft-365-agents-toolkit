# Driver system

Source: [`packages/core-next/src/drivers/`](../../../packages/core-next/src/drivers/).

## Layout

```
drivers/
├── types.ts            — DriverDescriptor, DriverConfig, DriverOutput
├── registry.ts         — DriverRegistry class + driverRegistry singleton
├── createDriver.ts     — createDriver() factory
└── builtin/            — 22 built-in implementations
    ├── index.ts        — registerBuiltinDrivers() + builtinDrivers array
    ├── file/           — createOrUpdateEnvironmentFile, createOrUpdateJsonFile
    ├── script/         — run (cross-platform shell with ::set-output)
    ├── cli/            — runNpmCommand, runDotnetCommand
    ├── teamsApp/       — zipAppPackage, validateManifest, validateAppPackage,
    │                     create, configure, update, publishAppPackage, extendToM365
    ├── aadApp/         — create, update
    ├── botAadApp/      — create
    ├── botFramework/   — create
    ├── arm/            — deploy
    ├── azureAppService/ — zipDeploy
    ├── azureFunctions/  — zipDeploy
    ├── oauth/          — register
    └── apiKey/         — register
```

## `createDriver()` factory

```typescript
createDriver({
  id: "myNamespace/myAction",
  name: "Do the thing",
  inputSchema: z.object({ ... }),
  execute: async (input, ctx) => { return ok({ ... }); },
  rollback?: async (input, ctx) => { return ok(); },
});
```

Adds:

1. **Zod pre-validation** — failure → `InvalidDriverInput` with issue path; never enters `execute`.
2. **`validateFn`** — preflight without execution (used by `analyzeSteps`).
3. **Telemetry** — `driver-start`, `driver-end` events with duration.
4. **Error normalisation** — wraps thrown exceptions into `DriverExecutionError`. Recognises plain `AtkError` shapes (with `code`/`message`/`kind`) and returns them unwrapped.

## Built-in driver catalogue (22)

| ID | Purpose |
|----|---------|
| `file/createOrUpdateEnvironmentFile` | Merge KV pairs into `.env` files |
| `file/createOrUpdateJsonFile` | Deep-merge JSON; supports `content` and `appsettings` aliases |
| `script` | Cross-platform shell; parses `::set-output` / `::set-teamsfx-env` |
| `cli/runNpmCommand` | `npm install`, `npm run build` |
| `cli/runDotnetCommand` | `dotnet publish`, `dotnet build` |
| `teamsApp/zipAppPackage` | Bundle manifest + icons |
| `teamsApp/validateManifest` | Schema validation via `TeamsManifestWrapper` |
| `teamsApp/validateAppPackage` | Package validation via TDP API |
| `teamsApp/create` | Create / import in TDP (idempotent via `existingTeamsAppId`) |
| `teamsApp/configure` | Update TDP record |
| `teamsApp/update` | Alias of configure (template references) |
| `teamsApp/publishAppPackage` | Publish to org catalog via Graph `/beta/appCatalogs/teamsApps` |
| `teamsApp/extendToM365` | Sideload via PackageService (V1 classic, V2 DA) |
| `aadApp/create` | Create Entra ID app via MS Graph |
| `aadApp/update` | Update Entra ID app properties; resolves `${{VAR}}` in AAD manifest |
| `botAadApp/create` | Create bot-specific Entra ID app + password |
| `botFramework/create` | Register bot channel via ARM |
| `arm/deploy` | Deploy ARM/Bicep templates |
| `azureAppService/zipDeploy` | Zip deploy to App Service via Kudu |
| `azureFunctions/zipDeploy` | Zip deploy to Functions via Kudu |
| `oauth/register` | Register OAuth config in TDP (Custom + MicrosoftEntra) |
| `apiKey/register` | Register API key secret in TDP |

## Registration

`registerBuiltinDrivers()` in `builtin/index.ts` registers all built-in drivers with the `driverRegistry` singleton. **Idempotent** — safe to call multiple times. Call once at startup before executing any lifecycle. In cli-next this is **deferred to first real command** to keep `--help` cold-start fast.

## Key design decisions vs v3

- **Zod pre-validation before execution** (v3 validated late or not at all).
- **No TypeDI `@Service()` decorators** — plain `createDriver()` factory + `driverRegistry.register()`.
- **Controlled `process.env` sync** — executor temporarily injects envMap into `process.env` before each driver, then cleans up. Drivers themselves return outputs for the executor to write — never write env directly.
- **No external deps for file I/O** — Node.js built-ins instead of `dotenv`, `fs-extra`, `comment-json`.
- **Driver IDs match v3 names** — `m365agents.yml` is portable.

## Tests

`tests/unit/drivers/` — registry tests + per-driver suites (`fileDrivers`, `scriptDriver`, `teamsAppDrivers`, `teamsAppPlatform`, `entraAndBotDrivers`, `azureInfraDrivers`, `authPluginDrivers`).
