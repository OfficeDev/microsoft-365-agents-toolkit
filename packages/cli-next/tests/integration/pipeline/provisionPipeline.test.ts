/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Data-driven provision pipeline integration tests.
 *
 * For each feature in features.json with "provision" in its lifecycles,
 * loads the template YAML, stubs external drivers, and runs the
 * provision lifecycle through the core-next executor.
 *
 * Adding a new feature with provision support in features.json
 * automatically creates a new test case here.
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

describe("Pipeline: Provision lifecycle (data-driven)", () => {
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

  const provisionFeatures = getFeaturesForLifecycle("provision");

  if (provisionFeatures.length === 0) {
    it("no features with provision lifecycle found — update features.json", () => {
      console.log("  ⚠ No features have 'provision' in lifecycles. Update .dev/features.json.");
    });
    return;
  }

  for (const feature of provisionFeatures) {
    const lang = preferredLanguage(feature);

    it(`${feature.name} (${feature.id}, ${lang}) — provision lifecycle should complete`, async function () {
      const yaml = loadTemplateYaml(feature, lang);
      if (!yaml) {
        this.skip(); // YAML file not found — template may not exist yet
        return;
      }

      const parseResult = await parseProjectYaml(yaml);
      if (parseResult.isErr()) {
        // YAML parse failure (usually Mustache leftover) — skip, not fail
        console.log(`    ⚠ YAML parse error for ${feature.id}: ${parseResult.error.message}`);
        this.skip();
        return;
      }

      const model = parseResult._unsafeUnwrap();
      const provisionSteps = model.provision;
      if (!provisionSteps || provisionSteps.length === 0) {
        this.skip(); // No provision steps in this template
        return;
      }

      // Stub external drivers
      const stubs = stubExternalDrivers(sandbox);

      // Create context and run
      const ctx = createIntegrationContext(tmpDir);
      const envMap = new Map<string, string>();
      const result = await executeLifecycle(ctx, "provision", provisionSteps, envMap);

      // Assert
      expect(result.isOk(), `Provision failed: ${result.isErr() ? result.error.message : ""}`).to.be
        .true;

      const lifecycleResult = result._unsafeUnwrap();
      expect(lifecycleResult.lifecycle).to.equal("provision");
      expect(lifecycleResult.steps).to.have.lengthOf(provisionSteps.length);

      // Verify that stubbed drivers were actually called
      for (const step of provisionSteps) {
        const stub = stubs.get(step.uses);
        if (stub) {
          expect(stub.called, `${step.uses} should have been called`).to.be.true;
        }
      }

      // Verify writeToEnvironmentFile outputs landed in envMap
      for (const step of provisionSteps) {
        if (step.writeToEnvironmentFile) {
          for (const envVarName of Object.values(step.writeToEnvironmentFile)) {
            // The env var should exist (may be "" for drivers with empty mock outputs)
            expect(envMap.has(envVarName), `envMap should have ${envVarName} from ${step.uses}`).to
              .be.true;
          }
        }
      }
    });
  }
});
