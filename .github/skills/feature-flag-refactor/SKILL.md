---
name: feature-flag-refactor
description: "Refactor code behind a feature flag so the original path is preserved and switchable. Use when: refactoring, replacing an implementation, rewriting a module, safe rollout, gradual migration, A/B code paths, keeping old code alongside new code, feature flag gating."
argument-hint: "Describe the refactoring or new implementation to gate behind a flag"
---

# Feature-Flag Refactor

Keep the original implementation intact when refactoring. Gate the new code behind a feature flag so both paths coexist and can be switched at runtime.

## When to Use

- Replacing or rewriting an existing implementation
- Gradual migration from old to new logic
- Risky refactors that need a rollback path
- Any change where "ship new, keep old as fallback" applies

## Procedure

### Step 1 — Define the Feature Flag

Add the flag in `packages/fx-core/src/common/featureFlags.ts`:

```typescript
// In FeatureFlagName class
static readonly MyRefactor = "TEAMSFX_MY_REFACTOR";

// In FeatureFlags class
static readonly MyRefactor = {
  name: FeatureFlagName.MyRefactor,
  defaultValue: "false",   // Off by default — old path runs
};
```

**Naming rules:**
- Environment variable: `TEAMSFX_` prefix, uppercase, underscores
- Flag object: PascalCase, matches the concept (e.g., `BrokerAuth`, `KiotaNPMIntegration`)

### Step 2 — Branch at the Decision Point

Import and branch using `featureFlagManager`:

```typescript
import { featureFlagManager, FeatureFlags } from "../common/featureFlags";

if (featureFlagManager.getBooleanValue(FeatureFlags.MyRefactor)) {
  return await newImplementation(args);  // New path
} else {
  return await originalImplementation(args);  // Old path — untouched
}
```

**Rules:**
- Do NOT modify the original code path — leave it exactly as-is
- The new path should have the same signature and return type
- Keep the branch point as high as possible (function/method entry, not deep in logic)
- One flag per refactor — don't reuse flags across unrelated changes

### Step 3 — Implement the New Path

Write the new implementation alongside the original. Preferred patterns:

**Pattern A — Separate functions (preferred for large changes):**
```typescript
// Original — untouched
async function processTemplate(input: Input): Promise<Result<Output, FxError>> {
  // ... existing code stays exactly as-is
}

// New — added alongside
async function processTemplateV2(input: Input): Promise<Result<Output, FxError>> {
  // ... new implementation
}

// Branch point
async function processTemplateGated(input: Input): Promise<Result<Output, FxError>> {
  if (featureFlagManager.getBooleanValue(FeatureFlags.MyRefactor)) {
    return processTemplateV2(input);
  }
  return processTemplate(input);
}
```

**Pattern B — Inline branch (for small changes):**
```typescript
async function login(scopes: string[]): Promise<string> {
  if (featureFlagManager.getBooleanValue(FeatureFlags.BrokerAuth)) {
    return await this.loginWithBroker(scopes);   // New
  } else {
    return await this.loginWithBrowser(scopes);  // Original
  }
}
```

**Pattern C — Conditional registration (for VS Code commands/providers):**
```typescript
if (featureFlagManager.getBooleanValue(FeatureFlags.SyncManifest)) {
  registerInCommandController(context, "fx-extension.syncManifest", syncManifestHandler);
}
```

### Step 4 — Test Both Paths

Write tests for **both** the old and new paths. Stub the flag:

```typescript
import { featureFlagManager, FeatureFlags } from "../../src/common/featureFlags";

describe("MyRefactor", () => {
  let sandbox: sinon.SinonSandbox;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
  });
  afterEach(() => sandbox.restore());

  it("uses original path when flag is off", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const result = await processTemplateGated(input);
    // assert original behavior
  });

  it("uses new path when flag is on", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").callsFake((flag) => {
      return flag === FeatureFlags.MyRefactor;
    });
    const result = await processTemplateGated(input);
    // assert new behavior
  });
});
```

### Step 5 — (Optional) Wire to VS Code Settings

If the flag should be togglable by users in VS Code settings, add to `packages/vscode-extension/src/config.ts`:

```typescript
// In loadFeatureFlags()
process.env[FeatureFlags.MyRefactor.name] = this.getConfiguration(
  ConfigurationKey.EnableMyRefactor,
  false
).toString();
```

And add the corresponding setting in `packages/vscode-extension/package.json` under `contributes.configuration`.

### Step 6 — Cleanup (After Validation)

Once the new path is validated and the flag is turned on by default:

1. Change `defaultValue` to `"true"` — ship the new path as default
2. After a release cycle with no issues, remove the flag:
   - Delete the branch point — keep only the new path
   - Remove the flag definition from `featureFlags.ts`
   - Remove the old implementation
   - Update tests to only test the new path

## Checklist

- [ ] Flag defined in `featureFlags.ts` with `defaultValue: "false"`
- [ ] Original code path is **completely untouched**
- [ ] Branch point is at the highest reasonable level
- [ ] New and old paths have identical signatures and return types
- [ ] Tests cover both flag=on and flag=off
- [ ] Flag is stubbed in tests (never rely on real env vars)
- [ ] PR description documents: what the flag gates, how to enable it, rollback plan

## Common Pitfalls

- **Modifying the original path** — The whole point is to keep it pristine as a fallback
- **Branching too deep** — Creates tangled flag checks throughout the code; branch at entry points
- **Forgetting to test the old path** — Regressions in the fallback defeat the purpose
- **Reusing flags** — One flag per refactor; shared flags create hidden coupling
- **Never cleaning up** — Stale flags accumulate; plan the cleanup from the start
