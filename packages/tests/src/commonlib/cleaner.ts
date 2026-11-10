// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { M365TokenProvider } from "@microsoft/teamsfx-api";
import { isValidProject } from "@microsoft/teamsfx-core";
import { AadManager } from ".";
import { AzureHelper } from "./azureHelper";
import {
  AADAppIdEnvNames,
  BotIdEnvName,
  M365TitleIdEnvName,
  ResourceGroupEnvName,
  TeamsAppIdEnvName,
} from "../utils/constants";
import { M365TitleHelper } from "./m365TitleHelper";
import { ProjectEnvReader } from "./projectEnvReader";
import { TeamsAppHelper } from "./teamsAppHelper";
import { tryMcpAuthTeardown } from "./mcpAuthTeardown";

/// Clean up the resources created by the test cases for V3 projects.
export class Cleaner {
  static async clean(
    projectPath: string,
    m365TokenProvider?: M365TokenProvider,
  ) {
    if (!isValidProject(projectPath)) {
      return Promise.resolve(true);
    }
    const mcpResult = await tryMcpAuthTeardown(projectPath, {
      deleteTeamsApp: async (id) => {
        const helper = await TeamsAppHelper.init(m365TokenProvider);
        if (!(await helper.deleteTeamsAppById(id))) {
          throw new Error("Failed to delete the test-owned MCP Teams app.");
        }
      },
      unacquireM365Title: async (id) => {
        const helper = await M365TitleHelper.init(
          undefined,
          undefined,
          m365TokenProvider,
        );
        if (!(await helper.unacquire(id))) {
          throw new Error("Failed to unacquire the test-owned MCP title.");
        }
      },
    });
    if (mcpResult !== undefined) {
      console.info(
        `MCP teardown: ${mcpResult.deleted.length} deleted, ` +
          `${mcpResult.skipped.length} protected/skipped, ${mcpResult.retained.length} retained.`,
      );
      return mcpResult;
    }
    const envs = await ProjectEnvReader.readAllEnvFiles(projectPath);
    const azureHelper = AzureHelper.init();
    const aadManager = await AadManager.init(m365TokenProvider);
    const teamsAppHelper = await TeamsAppHelper.init(m365TokenProvider);
    const m365TitleHelper = await M365TitleHelper.init(
      undefined,
      undefined,
      m365TokenProvider,
    );
    return envs.map(async (env) =>
      Promise.all([
        /// clean up resource group
        azureHelper
          .deleteResourceGroup(
            env[ResourceGroupEnvName] ?? process.env[ResourceGroupEnvName],
          )
          .then((result) => console.log(result)),
        /// clean up aad apps
        AADAppIdEnvNames.map((name) =>
          aadManager.deleteAadAppsByClientId(env[name]),
        ),
        /// clean up teams app
        teamsAppHelper.deleteTeamsAppById(env[TeamsAppIdEnvName]),
        /// clean up bot framework app
        teamsAppHelper.deleteBotById(env[BotIdEnvName]),
        /// clean up published teams app
        teamsAppHelper.cancelStagedTeamsAppById(env[TeamsAppIdEnvName]),
        /// clean up m365 app
        m365TitleHelper.unacquire(env[M365TitleIdEnvName]),
      ]),
    );
  }
}

(async () => {
  const projectPath = process.argv[2];
  if (!projectPath) {
    return;
  }
  await Cleaner.clean(projectPath);
})();
