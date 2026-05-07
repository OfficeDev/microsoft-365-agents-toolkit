// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhiyu You <zhiyou@microsoft.com>
 *
 * Regression tests for issue #15837 (https://github.com/OfficeDev/microsoft-365-agents-toolkit/issues/15837).
 *
 * Original failure: a `declarativeAgent.json` with `capabilities` (or `actions`)
 * provided as an object instead of an array crashed the build with a raw
 * `TypeError: ...filter is not a function`, surfaced as opaque "unknown.wYe".
 *
 * The fix routes manifest reads through the typed converter
 * (`DeclarativeAgentManifestConverter.jsonToManifest`) and adds defensive
 * `Array.isArray` guards in `createAppPackage`. These tests assert the user-
 * visible behavior: a descriptive `UserError` (not a `TypeError`) that names
 * the offending field and the file.
 *
 * Implementation note: each test exercises `teamsapp package`, which invokes
 * the same `teamsApp/zipAppPackage` driver that runs during `teamsapp provision`
 * — i.e. the exact crash site the bug fix targets. Using `package` rather than
 * `provision` keeps the test fast and Azure-credential-free while covering the
 * identical code path.
 *
 * Linked manual test cases (Microsoft Teams Extensibility ADO project,
 * plan #24569079, suite #27971458 "Project - Declarative Agent"):
 *   - 37862958: [VSC][CLI] Provision fails gracefully ... invalid capabilities shape (#15837)
 *   - 37862961: [VSC][CLI] Provision fails gracefully ... invalid actions shape
 *   - 37862962: [VSC][CLI] Zip Teams App Package surfaces schema errors from sub-manifests
 */

import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { it } from "@microsoft/extra-shot-mocha";
import { expect } from "chai";
import * as fs from "fs-extra";
import { describe } from "mocha";
import * as path from "path";

import {
  cleanUpLocalProject,
  getTestFolder,
  getUniqueAppName,
} from "../commonUtils";
import { Executor } from "../../utils/executor";
import { Capability } from "../../utils/constants";

const AUTHOR = "zhiyou@microsoft.com";

interface InvalidShapeAssertionOptions {
  // Substring that MUST appear in the error output (typically the field name).
  mustInclude: string[];
  // Substrings that MUST NOT appear (the regression markers).
  mustNotInclude?: string[];
}

const DEFAULT_FORBIDDEN = ["TypeError", "is not a function", "unknown.wYe"];

function assertInvalidShapeError(
  result: { success: boolean; stdout: string; stderr: string },
  opts: InvalidShapeAssertionOptions,
): void {
  expect(
    result.success,
    `command unexpectedly succeeded. stdout=${result.stdout}`,
  ).to.be.false;

  const combined = `${result.stdout}\n${result.stderr}`;
  for (const needle of opts.mustInclude) {
    expect(combined.toLowerCase()).to.include(
      needle.toLowerCase(),
      `expected error output to mention "${needle}"`,
    );
  }
  for (const needle of opts.mustNotInclude ?? DEFAULT_FORBIDDEN) {
    expect(combined).to.not.include(
      needle,
      `error output should not contain regression marker "${needle}"`,
    );
  }
}

async function createBasicDeclarativeAgentProject(
  appName: string,
  testFolder: string,
): Promise<string> {
  const customized: Record<string, string> = { "with-plugin": "no" };
  const { success } = await Executor.createProject(
    testFolder,
    appName,
    Capability.DeclarativeAgent,
    ProgrammingLanguage.None,
    customized,
  );
  expect(success, "scaffold should succeed").to.be.true;
  return path.resolve(testFolder, appName);
}

async function mutateDeclarativeAgent(
  projectPath: string,
  mutator: (json: Record<string, unknown>) => void,
): Promise<void> {
  const manifestPath = path.join(
    projectPath,
    "appPackage",
    "declarativeAgent.json",
  );
  const json = (await fs.readJSON(manifestPath)) as Record<string, unknown>;
  mutator(json);
  await fs.writeJSON(manifestPath, json, { spaces: 2 });
}

describe("Declarative agent - invalid manifest shape (#15837 regression)", function () {
  this.timeout(5 * 60 * 1000);

  it(
    "package fails with descriptive error when capabilities is not an array (#15837)",
    { testPlanCaseId: 37862958, author: AUTHOR },
    async function () {
      const testFolder = getTestFolder();
      const appName = getUniqueAppName();
      const projectPath = await createBasicDeclarativeAgentProject(
        appName,
        testFolder,
      );

      try {
        // Repro #15837: capabilities supplied as an object rather than an array.
        await mutateDeclarativeAgent(projectPath, (json) => {
          json.capabilities = { name: "CodeInterpreter" };
        });

        const failed = await Executor.package(projectPath);
        assertInvalidShapeError(failed, { mustInclude: ["capabilities"] });

        // Recovery: restoring a valid array shape lets the same command succeed.
        await mutateDeclarativeAgent(projectPath, (json) => {
          json.capabilities = [{ name: "CodeInterpreter" }];
        });
        const recovered = await Executor.package(projectPath);
        expect(
          recovered.success,
          `package should succeed after restoring valid capabilities. stderr=${recovered.stderr}`,
        ).to.be.true;
      } finally {
        await cleanUpLocalProject(projectPath);
      }
    },
  );

  it(
    "package fails with descriptive error when actions is not an array",
    { testPlanCaseId: 37862961, author: AUTHOR },
    async function () {
      const testFolder = getTestFolder();
      const appName = getUniqueAppName();
      const projectPath = await createBasicDeclarativeAgentProject(
        appName,
        testFolder,
      );

      try {
        // Even on a no-plugin template, an `actions` field of the wrong shape
        // must be rejected before the build attempts `actions.map(...)`.
        await mutateDeclarativeAgent(projectPath, (json) => {
          json.actions = { id: "fake-action", file: "ai-plugin.json" };
        });

        const failed = await Executor.package(projectPath);
        assertInvalidShapeError(failed, { mustInclude: ["actions"] });

        // Recovery: removing the invalid field lets the same command succeed.
        await mutateDeclarativeAgent(projectPath, (json) => {
          delete json.actions;
        });
        const recovered = await Executor.package(projectPath);
        expect(
          recovered.success,
          `package should succeed after removing invalid actions. stderr=${recovered.stderr}`,
        ).to.be.true;
      } finally {
        await cleanUpLocalProject(projectPath);
      }
    },
  );

  it(
    "package surfaces sub-manifest schema errors instead of crashing",
    { testPlanCaseId: 37862962, author: AUTHOR },
    async function () {
      const testFolder = getTestFolder();
      const appName = getUniqueAppName();
      const projectPath = await createBasicDeclarativeAgentProject(
        appName,
        testFolder,
      );

      try {
        // A different shape error (string instead of array) must be reported
        // via the same UserError path, not a raw TypeError.
        await mutateDeclarativeAgent(projectPath, (json) => {
          json.capabilities = "CodeInterpreter";
        });

        const failed = await Executor.package(projectPath);
        assertInvalidShapeError(failed, {
          mustInclude: ["capabilities", "declarativeAgent.json"],
        });
      } finally {
        await cleanUpLocalProject(projectPath);
      }
    },
  );
});
