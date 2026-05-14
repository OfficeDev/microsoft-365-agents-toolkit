# Adding a driver

Authoritative source: [`fx-core.instructions.md`](../../.github/instructions/fx-core.instructions.md) §"Driver System".

This playbook covers v4. For v3, drivers live in `packages/fx-core/src/component/driver/<area>/` and follow a similar but less structured pattern.

## Steps

### 1. Decide the driver ID

Use a namespaced kebab/path style: `<namespace>/<action>`. Match v3 names where possible for YAML compatibility.

Examples: `arm/deploy`, `teamsApp/publishAppPackage`, `azureFunctions/zipDeploy`.

### 2. Create the file

Under [`packages/core-next/src/drivers/builtin/<namespace>/`](../../packages/core-next/src/drivers/builtin/):

```
src/drivers/builtin/myNamespace/myAction.ts
```

### 3. Author the driver

```ts
import { z } from "zod";
import { ok, err } from "neverthrow";
import { createDriver } from "../../createDriver";
import { AtkError } from "../../../core/AtkError";

const inputSchema = z.object({
  workingDirectory: z.string(),
  threshold: z.number().min(1).max(100),
});

export const myAction = createDriver({
  id: "myNamespace/myAction",
  name: "Do the thing",
  inputSchema,
  execute: async (input, ctx) => {
    // input is fully typed
    // ctx is AtkContext
    try {
      // ... do the work ...
      return ok({
        someOutput: "value",
      });
    } catch (e) {
      return err(new AtkError({
        kind: "system",
        source: "MyAction",
        name: "MyActionFailed",
        message: getDefaultString("teamsfx.myAction.failed"),
        error: e instanceof Error ? e : new Error(String(e)),
      }));
    }
  },
});
```

### 4. Register it

In [`packages/core-next/src/drivers/builtin/index.ts`](../../packages/core-next/src/drivers/builtin/index.ts):

```ts
import { myAction } from "./myNamespace/myAction";

export const builtinDrivers = [
  // ... existing drivers ...
  myAction,
];
```

`registerBuiltinDrivers()` will pick it up automatically.

### 5. Test it

Create `packages/core-next/tests/unit/drivers/builtin/myActionDriver.test.ts`:

```ts
import { expect } from "chai";
import sinon from "sinon";
import { createMockContext } from "../../testHelper";
import { myAction } from "../../../src/drivers/builtin/myNamespace/myAction";

describe("myAction driver", () => {
  afterEach(() => sinon.restore());

  it("should return ok when input is valid", async () => {
    const ctx = createMockContext();
    const result = await myAction.execute({ workingDirectory: ".", threshold: 5 }, ctx);
    expect(result.isOk()).to.be.true;
  });

  it("should return InvalidDriverInput when threshold is out of range", async () => {
    const ctx = createMockContext();
    const result = await myAction.validate({ workingDirectory: ".", threshold: 999 });
    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().name).to.equal("InvalidDriverInput");
  });
});
```

### 6. Surface in YAML

Templates can now reference `uses: myNamespace/myAction` in their `m365agents.yml`. Update the relevant template's YAML if this driver is part of a stage.

### 7. Update docs

- Add a row to the catalogue in [05-engineering/cross-cutting/driver-system.md](../05-engineering/cross-cutting/driver-system.md).
- Mention it in the relevant scenario page under [01-product/scenarios/](../01-product/scenarios/README.md) if it changes the user flow.

## Auth-aware drivers

If your driver needs to call Microsoft 365 / Azure:

- Read tokens via `ctx.auth.m365TokenProvider.getAccessToken(scopes)` or `ctx.auth.azureAccountProvider.getIdentityCredentialAsync()`.
- Add the driver ID to `M365_DRIVERS` or `AZURE_DRIVERS` in `lifecycle/analyze.ts` so prerequisites trigger correctly.

## File-touching drivers

Use **EAFP**:

```ts
try {
  const data = await fs.readFile(path, "utf-8");
} catch (e: any) {
  if (e.code === "ENOENT") return ok(undefined); // or handle missing
  throw e;
}
```

Never `existsSync` + `readFile`.

## Rollback

Provide a `rollback` function for destructive operations (resource creation that can be safely undone). The lifecycle engine does not call rollback automatically yet, but having it positions the driver for future support.
