# Feature flags

Boolean gates for safe rollout. Read at runtime from environment variables.

## v3

```ts
import { featureFlagManager, FeatureFlagName } from "../common/featureFlags";

if (featureFlagManager.getBooleanValue(FeatureFlagName.MyFeature)) {
  // new behaviour
} else {
  // old behaviour
}
```

| Concern | v3 implementation |
|---------|-------------------|
| Manager | Singleton `featureFlagManager` |
| Definition | `FeatureFlagName` enum |
| Source | `process.env` |

## v4

```ts
import { createDefaultRegistry } from "../featureFlags";

const registry = createDefaultRegistry();
if (registry.isEnabled("MyFeature")) { ... }
```

| Concern | v4 implementation |
|---------|-------------------|
| Registry | Injectable `FeatureFlagRegistry` |
| Source | `FeatureFlagSource` interface — defaults to `process.env`, swappable for tests |
| Built-in flags | `builtinFlags` array in `core-next/src/featureFlags/` |

## Headline flag

| Flag | Default | Effect |
|------|---------|--------|
| `TEAMSFX_V4_CORE` | `false` | When `true`, consumers route through `core-next` operations instead of `fx-core` |
| `TEAMSFX_MCP_FOR_DA` | per release | Enables MCP actions in DA scaffolds |
| `TEAMSFX_SENSITIVITY_LABEL` | per release | Enables sensitivity-label DA capability |
| `TEAMSFX_DA_METAOS` | per release | Enables MetaOS DA upgrade flow |

## Rules

- **Never hardcode** flag values — always go through the manager / registry.
- **Keep both code paths working** while a flag exists.
- **Remove the flag** once rollout is complete; this is a `refactor` commit, not `feat`.
- **Default off** for new behaviour; flip to default on in a separate PR after telemetry confirms safety.
- **One flag per behaviour** — don't overload a single flag.

## Testing both paths

```ts
describe("with flag", () => {
  it("does new behaviour when flag is on", () => {
    sinon.stub(registry, "isEnabled").returns(true);
    // ...
  });
  it("does old behaviour when flag is off", () => {
    sinon.stub(registry, "isEnabled").returns(false);
    // ...
  });
});
```

The injectable v4 registry is intentionally easy to stub.
