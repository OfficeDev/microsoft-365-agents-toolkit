# Adding a feature flag

Authoritative source: [`fx-core.instructions.md`](../../.github/instructions/fx-core.instructions.md) §"Feature Flags".

Adding a feature flag is a deliberate act — every flag has a maintenance cost. Use one when:

- You're rolling out a behaviour gradually.
- You need to A/B compare two paths in production.
- You need a fast off switch for a risky change.

Don't add a flag when a config option, env var, or runtime check serves better.

## v4 (`core-next`)

### 1. Define the flag

In [`packages/core-next/src/featureFlags/builtinFlags.ts`](../../packages/core-next/src/featureFlags/):

```ts
export const builtinFlags: FeatureFlagDescriptor[] = [
  // ... existing flags ...
  {
    name: "MyFeature",
    envVar: "TEAMSFX_MY_FEATURE",
    defaultValue: false,
    description: "Enables the my-feature behaviour.",
  },
];
```

### 2. Read it at the gate

```ts
import { createDefaultRegistry } from "../featureFlags";

const registry = createDefaultRegistry();
if (registry.isEnabled("MyFeature")) {
  // new behaviour
} else {
  // old behaviour — keep working
}
```

For testability, prefer **injecting the registry** into the function instead of creating one inline:

```ts
async function doThing(input, ctx, registry: FeatureFlagRegistry = createDefaultRegistry()) {
  if (registry.isEnabled("MyFeature")) { ... }
}
```

### 3. Test both paths

```ts
describe("doThing", () => {
  it("does new behaviour when MyFeature is on", async () => {
    const registry = createDefaultRegistry();
    sinon.stub(registry, "isEnabled").returns(true);
    // ...
  });

  it("does old behaviour when MyFeature is off", async () => {
    const registry = createDefaultRegistry();
    sinon.stub(registry, "isEnabled").returns(false);
    // ...
  });
});
```

## v3 (`fx-core`)

### 1. Add the enum entry

In [`packages/fx-core/src/common/featureFlags.ts`](../../packages/fx-core/src/common/):

```ts
export enum FeatureFlagName {
  // ... existing flags ...
  MyFeature = "TEAMSFX_MY_FEATURE",
}
```

### 2. Read it

```ts
import { featureFlagManager, FeatureFlagName } from "../common/featureFlags";

if (featureFlagManager.getBooleanValue(FeatureFlagName.MyFeature)) { ... }
```

## Lifecycle of a flag

1. **Add (default off).** New behaviour gated.
2. **Enable internally.** Set the env var in dev/test environments. Verify telemetry.
3. **Default on.** Flip `defaultValue` to `true` after a stabilisation period.
4. **Remove the flag.** Once telemetry confirms the new path is safe and the old path has no rollback need, **delete the flag and the old code**. This is a `refactor:` commit, not `feat:`.

## Headline flags in flight

| Flag | Status |
|------|--------|
| `TEAMSFX_V4_CORE` | Long-lived; gates the v4 engine |
| `TEAMSFX_MCP_FOR_DA` | Per release |
| `TEAMSFX_SENSITIVITY_LABEL` | Per release |
| `TEAMSFX_DA_METAOS` | Per release |

## See also

- [feature-flag-refactor](../../.github/skills/feature-flag-refactor/SKILL.md) skill — how to refactor existing code behind a new flag.
