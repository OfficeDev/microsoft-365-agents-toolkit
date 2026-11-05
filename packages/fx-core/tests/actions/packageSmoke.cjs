// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const { mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } = require("node:fs");
const { createRequire } = require("node:module");
const { tmpdir } = require("node:os");
const { join, resolve } = require("node:path");
const { pathToFileURL } = require("node:url");

async function main() {
  const consumer = mkdtempSync(join(tmpdir(), "atk-action-consumer-"));
  try {
    const scope = join(consumer, "node_modules", "@microsoft");
    mkdirSync(scope, { recursive: true });
    const packageRoot = process.argv[2] ? resolve(process.argv[2]) : resolve(__dirname, "..", "..");
    symlinkSync(packageRoot, join(scope, "teamsfx-core"), "junction");
    const consume = createRequire(join(consumer, "consumer.cjs"));
    const actions = consume("@microsoft/teamsfx-core/actions");
    const pureContracts = /[\\/]component[\\/]driver[\\/](m365|workiq|util)[\\/]args\.js$/;
    const loadedDrivers = Object.keys(require.cache).filter(
      (file) =>
        /[\\/]component[\\/](driver[\\/]|m365[\\/]packageService|workiq[\\/]policyPublication\.js$)/.test(
          file
        ) && !pureContracts.test(file)
    );
    assert.deepEqual(loadedDrivers, [], "Metadata import loaded live business modules");
    assert(
      actions.validateActionArgs("teamsApp/extendToM365", { appPackagePath: "app.zip" }).isOk()
    );
    assert(actions.getActionDefinition("teamsApp/uninstallFromM365").isErr());
    assert(
      actions
        .validateActionArgs(actions.ActionNames.PublishPolicy, {
          titleId: "U_title",
          policyPath: "policy.json",
          callerId: "11111111-2222-4333-8444-555555555555",
        })
        .isOk()
    );
    const esm = await import(
      pathToFileURL(consume.resolve("@microsoft/teamsfx-core/actions")).href
    );
    assert.equal(typeof esm.validateActionArgs, "function");
    const { FxCoreClient, ActionNames } = consume("@microsoft/teamsfx-core");
    assert.equal(ActionNames.Acquire, "teamsApp/extendToM365");
    assert.equal(esm.ActionNames.Acquire, ActionNames.Acquire);
    assert.equal(esm.ActionNames.PublishPolicy, "workiq/publishPolicy");
    assert(Object.isFrozen(ActionNames));
    const client = new FxCoreClient({
      tokenProvider: {
        m365TokenProvider: {
          getAccessToken() {
            throw new Error("Unexpected authentication");
          },
        },
      },
      logProvider: {},
      ui: {},
    });
    const result = await client.executeAction(ActionNames.Acquire, {}, { projectPath: consumer });
    assert(result.result.isErr());
    const policyInvalid = await client.executeAction(
      ActionNames.PublishPolicy,
      {},
      { projectPath: consumer }
    );
    assert(policyInvalid.result.isErr());
    assert(consume.resolve("@microsoft/teamsfx-core/build/core/FxCoreClient"));

    const fixture = join(consumer, "consumer.ts");
    writeFileSync(
      fixture,
      [
        'import type { IFxCoreClient, AcquireArgs, ActionResult, PublishPolicyArgs, PolicyPublicationReport } from "@microsoft/teamsfx-core";',
        'import { ActionNames } from "@microsoft/teamsfx-core";',
        'import { validateActionArgs } from "@microsoft/teamsfx-core/actions";',
        "declare const client: IFxCoreClient;",
        'const options = { projectPath: "C:\\\\project" };',
        'const install: AcquireArgs = { appPackagePath: "app.zip", scope: "Personal" };',
        "const execution: Promise<ActionResult<typeof ActionNames.Acquire>> = client.executeAction(ActionNames.Acquire, install, options);",
        'const policy: PublishPolicyArgs = { titleId: "U_title", policyPath: "policy.json", callerId: "11111111-2222-4333-8444-555555555555" };',
        "const publication: Promise<ActionResult<typeof ActionNames.PublishPolicy>> = client.executeAction(ActionNames.PublishPolicy, policy, options);",
        "declare const published: Awaited<typeof publication>;",
        'if ("report" in published) { const report: PolicyPublicationReport | undefined = published.report; void report; }',
        'validateActionArgs("teamsApp/extendToM365", install);',
        "// @ts-expect-error Uninstall is not a standalone action.",
        'void client.executeAction("teamsApp/uninstallFromM365", { titleId: "U_test" }, options);',
        "// @ts-expect-error Unknown actions cannot compile through published declarations.",
        'void client.executeAction("script", {}, options);',
      ].join("\n")
    );
    const check = spawnSync(
      process.execPath,
      [
        require.resolve("typescript/bin/tsc"),
        "--noEmit",
        "--strict",
        "--skipLibCheck",
        "--target",
        "ES2022",
        "--module",
        "Node16",
        "--moduleResolution",
        "Node16",
        fixture,
      ],
      { encoding: "utf8" }
    );
    assert.equal(check.status, 0, `${check.stdout}\n${check.stderr}`);
    console.log("ACTION-10: public CJS/ESM, type declarations and legacy deep imports pass.");
  } finally {
    rmSync(consumer, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
