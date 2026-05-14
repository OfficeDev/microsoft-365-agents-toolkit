# Adding a CLI command

Authoritative source: [`cli.instructions.md`](../../.github/instructions/cli.instructions.md).

## v4 (cli-next)

### 1. Add an action

Under [`packages/cli-next/src/actions/`](../../packages/cli-next/src/actions/):

```ts
import { runOperation } from "@microsoft/teamsfx-core";
import { myOperation } from "@microsoft/teamsfx-core/operations";

export async function myAction(opts: { foo: string }): Promise<Result<void, AtkError>> {
  const ctx = await createCliContext();
  const result = await runOperation(myOperation, { foo: opts.foo }, ctx);
  if (result.isErr()) return result;
  renderPostActions(result.value.postActions);
  return ok(undefined);
}
```

### 2. Wire into a command group

In [`packages/cli-next/src/commands/`](../../packages/cli-next/src/commands/) — pick the right group file (e.g. `project.ts`, `env.ts`, `teamsapp.ts`) or create a new one.

```ts
import { Command } from "commander";
import { wrapHandlerWithContext } from "../wrapHandler";
import { myAction } from "../actions/myAction";

export function buildMyCommands(parent: Command): void {
  parent
    .command("my-cmd")
    .description("Do the thing")
    .option("-f, --foo <value>", "the foo value")
    .action(wrapHandlerWithContext("my-cmd", async (opts) => {
      return myAction(opts);
    }));
}
```

### 3. Register the group

In [`src/commands/factory.ts`](../../packages/cli-next/src/commands/factory.ts) or the relevant command-tree builder:

```ts
import { buildMyCommands } from "./myCommands";
// ...
buildMyCommands(program);
```

### 4. Test

Add an action test:

```ts
import { myAction } from "../../src/actions/myAction";

describe("myAction", () => {
  afterEach(() => sinon.restore());

  it("returns ok when operation succeeds", async () => {
    const result = await myAction({ foo: "bar" });
    expect(result.isOk()).to.be.true;
  });
});
```

For end-to-end coverage, add a Commander integration test in `tests/integration/`.

### 5. Telemetry

`wrapHandler` and `wrapHandlerWithContext` emit start/end telemetry automatically. The first arg is the **stable telemetry name** — never change it across releases.

### 6. Localised strings

Add user-facing strings to the locale bundles via `Localizer.getString()`. English source goes in `package.nls.json`; translations in `package.nls.{locale}.json`.

## v3 (cli)

### 1. Create a `CLICommand` object

Under [`packages/cli/src/commands/models/`](../../packages/cli/src/commands/models/):

```ts
export const myCommand: CLICommand = {
  name: "my-command",
  description: "...",
  options: [...MyCommandOptions],
  handler: async (ctx) => {
    const result = await ctx.core.myOperation(ctx.inputs);
    return result;
  },
  telemetry: { event: TelemetryEvent.MyCommand },
};
```

### 2. Export and register

- Export from `src/commands/models/index.ts`.
- Add to the `commands` array in `src/commands/models/root.ts`.
- Add localised strings to `src/resource/`.

### 3. Test

`packages/cli/tests/unit/commands/<myCommand>.test.ts` — mock `ctx.core` and assert the handler dispatches correctly.

## Adding a `new` subcommand (v4)

Don't. The `atk new <category>` subtree is **auto-generated** from `TemplateRegistry`. To add a new template subcommand, add a `TemplateDescriptor` (see [adding-a-template.md](adding-a-template.md)). Adding bespoke `new` commands defeats the registry-driven design.
