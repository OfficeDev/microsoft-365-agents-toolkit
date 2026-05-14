/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Data-driven deploy pipeline integration tests.
 *
 * For each feature in features.json with "deploy" in its lifecycles,
 * loads the template YAML, stubs external drivers, and runs the
 * deploy lifecycle through the core-next executor.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import * as fs from "node:fs/promises";

import {
  registerBuiltinDrivers,
  parseProjectYaml,
  executeLifecycle,
} from "@microsoft/teamsfx-core-next";
import { getFeaturesForLifecycle, preferredLanguage } from "../helpers/featureRegistry";
import { loadTemplateYaml } from "../helpers/templateYamlLoader";
import { stubExternalDrivers } from "../helpers/driverStubs";
import { createIntegrationContext, makeTmpDir } from "../helpers/mockContext";

describe("Pipeline: Deploy lifecycle (data-driven)", () => {
  let sandbox: sinon.SinonSandbox;
  let tmpDir: string;

  beforeEach(async () => {
    sandbox = sinon.createSandbox();
    registerBuiltinDrivers();
    tmpDir = await makeTmpDir();
  });

  afterEach(async () => {
    sandbox.restore();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  const deployFeatures = getFeaturesForLifecycle("deploy");

  if (deployFeatures.length === 0) {
    it("no features with deploy lifecycle found — update features.json", () => {
      console.log("  ⚠ No features have 'deploy' in lifecycles. Update .dev/features.json.");
    });
    return;
  }

  for (const feature of deployFeatures) {
    const lang = preferredLanguage(feature);

    it(`${feature.name} (${feature.id}, ${lang}) — deploy lifecycle should complete`, async function () {
      const yaml = loadTemplateYaml(feature, lang);
      if (!yaml) {
        this.skip();
        return;
      }

      const parseResult = await parseProjectYaml(yaml);
      if (parseResult.isErr()) {
        console.log(`    ⚠ YAML parse error for ${feature.id}: ${parseResult.error.message}`);
        this.skip();
        return;
      }

      const model = parseResult._unsafeUnwrap();
      const deploySteps = model.deploy;
      if (!deploySteps || deploySteps.length === 0) {
        this.skip();
        return;
      }

      const stubs = stubExternalDrivers(sandbox);
      const ctx = createIntegrationContext(tmpDir);

      // Pre-populate envMap with values that provision would have set
      // (deploy steps reference ARM outputs like BOT_AZURE_APP_SERVICE_RESOURCE_ID)
      const envMap = new Map<string, string>([
        [
          "BOT_AZURE_APP_SERVICE_RESOURCE_ID",
          "/subscriptions/00000000/resourceGroups/rg-test/providers/Microsoft.Web/sites/test-bot",
        ],
        [
          "AZURE_FUNCTION_RESOURCE_ID",
          "/subscriptions/00000000/resourceGroups/rg-test/providers/Microsoft.Web/sites/test-func",
        ],
        ["TEAMS_APP_ID", "mock-teams-app-id-00000"],
        ["BOT_ID", "mock-bot-id-77777"],
      ]);

      const result = await executeLifecycle(ctx, "deploy", deploySteps, envMap);

      expect(result.isOk(), `Deploy failed: ${result.isErr() ? result.error.message : ""}`).to.be
        .true;

      const lifecycleResult = result._unsafeUnwrap();
      expect(lifecycleResult.lifecycle).to.equal("deploy");
      expect(lifecycleResult.steps).to.have.lengthOf(deploySteps.length);

      // Verify expected deploy drivers were called
      for (const step of deploySteps) {
        const stub = stubs.get(step.uses);
        if (stub) {
          expect(stub.called, `${step.uses} should have been called`).to.be.true;
        }
      }
    });
  }
});
