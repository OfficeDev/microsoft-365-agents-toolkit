// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateDashboardTab } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import * as path from "path";
import * as fs from "fs";

class AssistDashboardTestCase extends CaseFactory {
  override async onValidate(page: Page): Promise<void> {
    return await validateDashboardTab(page);
  }
  public override async onCliValidate(page: Page): Promise<void> {
    return await validateDashboardTab(page);
  }

  override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
  ): Promise<void> {
    const envFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}.user`
    );
    const envString =
      'PLANNER_GROUP_ID=YOUR_PLANNER_GROUP_ID\nDEVOPS_ORGANIZATION_NAME=msazure\nDEVOPS_PROJECT_NAME="Microsoft Teams Extensibility"\nGITHUB_REPO_NAME=test002\nGITHUB_REPO_OWNER=hellyzh\nPLANNER_PLAN_ID=YOUR_PLAN_ID\nPLANNER_BUCKET_ID=YOUR_BUCKET_ID\nSECRET_DEVOPS_ACCESS_TOKEN=YOUR_DEVOPS_ACCESS_TOKEN\nSECRET_GITHUB_ACCESS_TOKEN=YOUR_GITHUB_ACCESS_TOKEN';
    fs.writeFileSync(envFilePath, envString);
  }
}

new AssistDashboardTestCase(
  TemplateProject.AssistDashboard,
  24121324,
  24121439,
  "v-ivanchen@microsoft.com",
  [
    LocalDebugTaskLabel.StartFrontend,
    LocalDebugTaskLabel.WatchBackend,
    LocalDebugTaskLabel.StartBackend,
  ],
  {
    dashboardFlag: true,
    skipInit: true,
  } // [TODO] skipInit browser security block
).test();
