// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author  Renlong Tu <rentu@microsoft.com>
 */

import { expect } from "chai";
import path from "path";
import * as fs from "fs-extra";
import {
  cleanUpLocalProject,
  cleanupSharePointPackage,
  execAsync,
  execAsyncWithRetry,
  getTestFolder,
  getUniqueAppName,
  removeTeamsAppExtendToM365,
} from "../commonUtils";

import { it } from "@microsoft/extra-shot-mocha";

describe("Collaboration", function () {
  const testFolder = getTestFolder();
  let appName = getUniqueAppName();
  let projectPath = path.resolve(testFolder, appName);
  const collaborator = process.env["M365_ACCOUNT_COLLABORATOR"];
  const creator = process.env["M365_ACCOUNT_NAME"];
  let appId: string;

  it(
    "Collaboration: CLI with permission status and permission grant - spfx",
    { testPlanCaseId: 12782618, author: "rentu@microsoft.com" },
    async function () {
      while (await fs.pathExists(projectPath)) {
        appName = getUniqueAppName();
        projectPath = path.resolve(testFolder, appName);
      }

      // new a project
      await execAsync(
        `atk new --interactive false --capability tab-spfx --app-name ${appName}`,
        {
          cwd: testFolder,
          env: process.env,
          timeout: 0,
        }
      );
      console.log(`[Successfully] scaffold to ${projectPath}`);
      const filePath = path.join(projectPath, "teamsapp.yml");
      removeTeamsAppExtendToM365(filePath);

      // provision
      await execAsyncWithRetry(`atk provision`, {
        cwd: projectPath,
        env: process.env,
        timeout: 0,
      });
      console.log("[Successfully] provision");

      const solutionConfig = await fs.readJson(
        `${projectPath}/src/config/package-solution.json`
      );
      appId = solutionConfig["solution"]["id"];

      const checkPermissionResult = await execAsyncWithRetry(
        `atk collaborator status --env dev --interactive false --teams-manifest-file ${projectPath}/appPackage/manifest.json`,
        {
          cwd: projectPath,
          env: process.env,
          timeout: 0,
        }
      );

      expect(checkPermissionResult.stdout).to.contains(
        "Resource Name: Teams App, Permission: Administrator"
      );
      console.log("[Successfully] check permission");

      const grantCollaboratorResult = await execAsyncWithRetry(
        `atk collaborator grant --email ${collaborator} --env dev --teams-manifest-file ${projectPath}/appPackage/manifest.json --interactive false`,
        {
          cwd: projectPath,
          env: process.env,
          timeout: 0,
        }
      );
      // Grant Permission

      expect(grantCollaboratorResult.stdout).to.contains(
        "Administrator Permission granted to Teams App"
      );
      console.log("[Successfully] grant permission");

      const listCollaboratorResult = await execAsync(
        `atk collaborator status --all --env dev --teams-manifest-file ${projectPath}/appPackage/manifest.json --interactive false`,
        {
          cwd: projectPath,
          env: process.env,
          timeout: 0,
        }
      );

      // Check collaborator.
      // When collaborator account is guest account in the tenant. Account name pattern will change.
      // e.g. Guest account "account@example.com" will appear as "account_example#EXT#@exampleTenant.onmicrosoft.com" under tenant "exampleTenant".
      // Thus here will check the account name only.
      expect(listCollaboratorResult.stdout).to.contains(
        `Account used to check: ${creator?.split("@")[0]}`
      );
      expect(listCollaboratorResult.stdout).to.contains(
        `Teams App Owner: ${collaborator?.split("@")[0]}`
      );
      console.log("[Successfully] list collaborator");
    }
  );

  after(async () => {
    // clean up
    if (projectPath) {
      await cleanUpLocalProject(projectPath);
    }
    if (appId) {
      await cleanupSharePointPackage(appId);
    }
  });
});
