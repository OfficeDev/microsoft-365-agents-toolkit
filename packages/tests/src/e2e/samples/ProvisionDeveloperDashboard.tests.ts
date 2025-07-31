// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import * as fs from "fs-extra";
import * as path from "path";
import { expect } from "chai";
import { removeTeamsAppExtendToM365 } from "../commonUtils";

class AssistDashboardTestCase extends CaseFactory {
  public override async onAfterCreate(projectPath: string): Promise<void> {
    expect(fs.pathExistsSync(path.resolve(projectPath, "infra"))).to.be.true;
    // remove teamsApp/extendToM365 in case it fails
    removeTeamsAppExtendToM365(path.join(projectPath, "m365agents.yml"));

    const envFilePath = path.resolve(projectPath, "env", `.env.dev.user`);
    const envString =
      'PLANNER_GROUP_ID=YOUR_PLANNER_GROUP_ID\nDEVOPS_ORGANIZATION_NAME=msazure\nDEVOPS_PROJECT_NAME="Microsoft Teams Extensibility"\nGITHUB_REPO_NAME=test002\nGITHUB_REPO_OWNER=hellyzh\nPLANNER_PLAN_ID=YOUR_PLAN_ID\nPLANNER_BUCKET_ID=YOUR_BUCKET_ID\nSECRET_DEVOPS_ACCESS_TOKEN=YOUR_DEVOPS_ACCESS_TOKEN\nSECRET_GITHUB_ACCESS_TOKEN=YOUR_GITHUB_ACCESS_TOKEN';
    fs.writeFileSync(envFilePath, envString);
  }
}

new AssistDashboardTestCase(
  TemplateProjectFolder.AssistDashboard,
  24121324,
  "huimiao@microsoft.com",
  ["dashboard"]
).test();
