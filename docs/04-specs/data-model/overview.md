# Data model overview

```mermaid
erDiagram
    PROJECT ||--|| MANIFEST : "appPackage/manifest.json"
    PROJECT ||--|| LIFECYCLE_YAML : "m365agents.yml"
    PROJECT ||--o| LOCAL_LIFECYCLE_YAML : "m365agents.local.yml"
    PROJECT ||--o{ ENV_FILE : "env/.env.{name}"
    PROJECT ||--o{ INFRA_FILE : "infra/*.bicep"
    PROJECT ||--o| DA_MANIFEST : "appPackage/declarativeAgent.json"
    PROJECT ||--o| API_PLUGIN_MANIFEST : "appPackage/aiPlugin.json"
    PROJECT ||--|| TRACKING_ID : "projectId in env"

    LIFECYCLE_YAML ||--o{ DRIVER_STEP : "actions[]"
    DRIVER_STEP ||--|| DRIVER_DESCRIPTOR : "uses driver id"
    DRIVER_STEP ||--o{ ENV_VAR : "produces outputs"

    TEMPLATE_DESCRIPTOR ||--o{ QUESTION : "questions[]"
    TEMPLATE_REGISTRY ||--o{ TEMPLATE_DESCRIPTOR : "register()"
    DRIVER_REGISTRY ||--o{ DRIVER_DESCRIPTOR : "register()"

    OPERATION ||--|| INPUT_SCHEMA : "Zod schema"
    ATKCONTEXT ||--o| LOGGER : "logger"
    ATKCONTEXT ||--o| TELEMETRY : "telemetry"
    ATKCONTEXT ||--o| UI : "ui"
    ATKCONTEXT ||--o| AUTH : "auth"

    OPERATION }o--|| ATKCONTEXT : "executes within"
```

## Entity catalogue

| Entity | Page | Source |
|--------|------|--------|
| Project + tracking ID | [project-and-tracking-id.md](entities/project-and-tracking-id.md) | env files |
| Teams app manifest | [teamsapp-manifest.md](entities/teamsapp-manifest.md) | `@microsoft/app-manifest` |
| Lifecycle YAML | [m365agents-yml.md](entities/m365agents-yml.md) | `lifecycle/parser.ts` |
| Env file | [env-files.md](entities/env-files.md) | `environment/envManager.ts` |
| Template descriptor | [template-descriptor.md](entities/template-descriptor.md) | `templates/registry.ts` |
| Driver descriptor | [driver-descriptor.md](entities/driver-descriptor.md) | `drivers/createDriver.ts` |
| Operation record | [operation-record.md](entities/operation-record.md) | `core/Operation.ts` |
| `AtkContext` | [atk-context.md](entities/atk-context.md) | `core/AtkContext.ts` |
| `Result` + `FxError` / `AtkError` | [result-and-fxerror.md](entities/result-and-fxerror.md) | `api/error.ts` |
| Question tree | [question-tree.md](entities/question-tree.md) | `questions/treeBuilder.ts` |
| Feature registry | [feature-registry.md](entities/feature-registry.md) | `.dev/features.json` |
