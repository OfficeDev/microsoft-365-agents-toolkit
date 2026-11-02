// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const ts = require("typescript");

describe("Scheduled resource cleanup", function () {
  it("preserves keep-prefixed resources without changing ordinary cleanup or title retries", async function () {
    const resources = [
      { id: "kept-agent", name: "keep-PineGap" },
      { id: "kept-agent-with-spaces", name: "keep-Study and Learn" },
      { id: "today", name: "fxui1029agent" },
      { id: "old", name: "fxui1028agent" },
      { id: "not-a-prefix", name: "agent-keep-test" },
    ];
    const deleted = {
      aad: [],
      definitions: [],
      registrations: [],
      titles: [],
    };
    const logs = [];
    let acquisitions = resources.map(({ id, name }) => ({
      titleId: id,
      titleDefinition: { name },
    }));
    let acquisitionReads = 0;
    const graph = {
      getUserIdByName: async () => "test-user",
      listTeamsApp: async () => [],
      listAad: async () =>
        resources.map(({ id, name }) => ({ id, displayName: name })),
      listEnterpriseApplications: async () => [],
      listDeletedAad: async () => [],
      listDeletedEnterpriseApplications: async () => [],
      deleteAad: async (id) => deleted.aad.push(id),
    };
    const appStudio = {
      getAppsInAppStudio: async () =>
        resources.map(({ id, name }) => ({
          appDefinitionId: id,
          displayName: name,
        })),
      deleteAppInAppStudio: async (id) => deleted.definitions.push(id),
      getApiKeyRegistration: async () =>
        resources.map(({ id, name }) => ({ id, description: name })),
      deleteApiKeyRegistration: async (id) => deleted.registrations.push(id),
    };
    const titles = {
      listAcquisitions: async () => {
        acquisitionReads++;
        return acquisitions;
      },
      unacquire: async (id) => {
        deleted.titles.push(id);
        acquisitions = acquisitions.filter(
          (acquisition) => acquisition.titleId !== id,
        );
        return true;
      },
    };
    const dependencies = {
      "../utils/constants": { Project: { namePrefix: "fxui" } },
      "../utils/env": { Env: {} },
      "../utils/nameUtil": { getAppNamePrefix: () => "fxui1029" },
      "../utils/retryHandler": { delay: async () => {} },
      "../utils/cleanHelper": {
        GraphApiCleanHelper: { create: async () => graph },
        AppStudioCleanHelper: { create: async () => appStudio },
        M365TitleCleanHelper: { create: async () => titles },
        SharePointApiCleanHelper: {
          create: async () => ({ listApp: async () => [] }),
        },
        DevTunnelCleanHelper: {
          create: async () => ({ deleteAll: async () => {} }),
        },
        filterResourceGroupByName: async () => [],
        deleteResourceGroupByName: async () =>
          assert.fail("Unexpected resource-group deletion"),
      },
    };
    const scriptPath = path.join(__dirname, "../../src/scripts/clean.ts");
    const compiled = ts.transpileModule(fs.readFileSync(scriptPath, "utf8"), {
      compilerOptions: {
        module: ts.ModuleKind.CommonJS,
        target: ts.ScriptTarget.ES2022,
      },
    });
    await vm.runInNewContext(
      compiled.outputText,
      {
        exports: {},
        require: (id) => {
          assert.ok(
            Object.hasOwn(dependencies, id),
            `Unexpected cleanup dependency: ${id}`,
          );
          return dependencies[id];
        },
        console: {
          log: (...args) => logs.push(args.join(" ")),
          error: (error) => {
            throw error;
          },
        },
        process: {
          exit: (code) => {
            assert.fail(`Cleanup exited with code ${code}`);
          },
        },
      },
      { filename: scriptPath, timeout: 1000 },
    );

    assert.deepEqual(deleted, {
      aad: ["old", "not-a-prefix"],
      definitions: ["old", "not-a-prefix"],
      registrations: ["today", "old", "not-a-prefix"],
      titles: ["old", "not-a-prefix"],
    });
    assert.equal(acquisitionReads, 2);
    assert.equal(
      logs.some((line) => /Failed|Get error:/.test(line)),
      false,
      logs.join("\n"),
    );
    assert.equal(logs.at(-1), "Clean Job Done.");
  });
});
