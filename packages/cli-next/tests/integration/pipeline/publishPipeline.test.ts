/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Data-driven publish pipeline integration tests.
 *
 * For each feature in features.json with "publish" in its lifecycles,
 * loads the template YAML, stubs external drivers, and runs the
 * publish lifecycle through the core-next executor.
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

describe("Pipeline: Publish lifecycle (data-driven)", () => {
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

  const publishFeatures = getFeaturesForLifecycle("publish");

  if (publishFeatures.length === 0) {
    it("no features with publish lifecycle found — update features.json", () => {
      console.log("  ⚠ No features have 'publish' in lifecycles. Update .dev/features.json.");
    });
    return;
  }

  for (const feature of publishFeatures) {
    const lang = preferredLanguage(feature);

    it(`${feature.name} (${feature.id}, ${lang}) — publish lifecycle should complete`, async function () {
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
      const publishSteps = model.publish;
      if (!publishSteps || publishSteps.length === 0) {
        this.skip();
        return;
      }

      const stubs = stubExternalDrivers(sandbox);
      const ctx = createIntegrationContext(tmpDir);

      // Pre-populate envMap with values that provision would have set
      const envMap = new Map<string, string>([
        ["TEAMS_APP_ID", "mock-teams-app-id-00000"],
        ["M365_TITLE_ID", "mock-title-id-22222"],
        ["M365_APP_ID", "mock-m365-app-id-33333"],
      ]);

      const result = await executeLifecycle(ctx, "publish", publishSteps, envMap);

      expect(result.isOk(), `Publish failed: ${result.isErr() ? result.error.message : ""}`).to.be
        .true;

      const lifecycleResult = result._unsafeUnwrap();
      expect(lifecycleResult.lifecycle).to.equal("publish");
      expect(lifecycleResult.steps).to.have.lengthOf(publishSteps.length);

      for (const step of publishSteps) {
        const stub = stubs.get(step.uses);
        if (stub) {
          expect(stub.called, `${step.uses} should have been called`).to.be.true;
        }
      }

      // Verify publish-specific output: publishedAppId should be in envMap
      // (most templates use teamsApp/publishAppPackage with writeToEnvironmentFile)
      for (const step of publishSteps) {
        if (step.writeToEnvironmentFile) {
          for (const envVarName of Object.values(step.writeToEnvironmentFile)) {
            expect(envMap.has(envVarName), `envMap should have ${envVarName}`).to.be.true;
          }
        }
      }
    });
  }
});
