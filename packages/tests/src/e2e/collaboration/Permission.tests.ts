// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Bowen Song <bowen.song@microsoft.com>
 */

import { expect } from "chai";
import * as fs from "fs-extra";
import path from "path";
import {
  cleanUp,
  createResourceGroup,
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

  it(
    "Collaboration: CLI with permission status and permission grant",
    { testPlanCaseId: 10753319, author: "bowen.song@microsoft.com" },
    async function () {
      while (await fs.pathExists(projectPath)) {
        appName = getUniqueAppName();
        projectPath = path.resolve(testFolder, appName);
      }

      // new a project
      await execAsync(
        `atk new --interactive false --capability sso-launch-page --app-name ${appName}`,
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
      const resourceGroupName = `${appName}-rg`;
      const result = await createResourceGroup(resourceGroupName, "westus");

      // provision
      await execAsyncWithRetry(`atk provision`, {
        cwd: projectPath,
        env: {
          ...process.env,
          AZURE_RESOURCE_GROUP_NAME: resourceGroupName,
        },
        timeout: 0,
      });
      console.log("[Successfully] provision");

      // Check Permission
      const checkPermissionResult = await execAsyncWithRetry(
        `atk collaborator status --env dev --interactive false --teams-manifest-file ${projectPath}/appPackage/manifest.json --entra-app-manifest-file ${projectPath}/aad.manifest.json`,
        {
          cwd: projectPath,
          env: process.env,
          timeout: 0,
        }
      );

      // TODO: Check for aad admin is skipped since now test account will not be added as admin when create aad app.
      expect(checkPermissionResult.stdout).to.contains(
        "Resource Name: Teams App, Permission: Administrator"
      );
      console.log("[Successfully] check permission");

      // Grant Permission
      const grantCollaboratorResult = await execAsyncWithRetry(
        `atk collaborator grant --email ${collaborator} --env dev --teams-manifest-file ${projectPath}/appPackage/manifest.json --entra-app-manifest-file ${projectPath}/aad.manifest.json --interactive false`,
        {
          cwd: projectPath,
          env: process.env,
          timeout: 0,
        }
      );

      expect(grantCollaboratorResult.stdout).to.contains(
        "Owner Permission granted to Microsoft Entra App"
      );
      expect(grantCollaboratorResult.stdout).to.contains(
        "Administrator Permission granted to Teams App"
      );
      console.log("[Successfully] grant permission");

      const listCollaboratorResult = await execAsync(
        `atk collaborator status --all --env dev --teams-manifest-file ${projectPath}/appPackage/manifest.json --entra-app-manifest-file ${projectPath}/aad.manifest.json --interactive false`,
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
    await cleanUp(appName, projectPath, true, false, false);
  });
});
