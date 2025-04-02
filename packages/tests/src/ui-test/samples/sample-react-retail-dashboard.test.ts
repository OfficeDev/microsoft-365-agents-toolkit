// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { Page } from "playwright";
import {
  initTeamsPage,
  validateRetailDashboard,
} from "../../utils/playwrightOperation";
import { SampledebugContext } from "./sampledebugContext";
import { Env } from "../../utils/env";

class RetailDashboardTestCase extends CaseFactory {
  public override async onInitPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string,
    options?: {
      teamsAppName: string;
      type: string;
      env: "local" | "dev";
    }
  ): Promise<Page> {
    return await initTeamsPage(
      sampledebugContext.context!,
      teamsAppId,
      Env.username,
      Env.password,
      {
        projectPath: sampledebugContext.projectPath,
        env: options?.env,
        teamsAppName: options?.teamsAppName,
        type: options?.type,
      }
    );
  }

  override async onValidate(
    page: Page,
    options?: { context: SampledebugContext }
  ): Promise<void> {
    return await validateRetailDashboard(page);
  }
}

new RetailDashboardTestCase(
  TemplateProject.RetailDashboard,
  25051148,
  25051150,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.GulpServe],
  {
    teamsAppName: "react-retail-dashboard-",
  }
).test();
