// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import {
  validateDashboardTab,
  reopenPage,
} from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";
import { SampledebugContext } from "./sampledebugContext";

class DashboardTestCase extends CaseFactory {
  override async onValidate(page: Page): Promise<void> {
    return await validateDashboardTab(page);
  }
  public override async onCliValidate(page: Page): Promise<void> {
    return await validateDashboardTab(page);
  }
  public override async onReopenPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string
  ): Promise<Page> {
    return await reopenPage(
      sampledebugContext.context!,
      teamsAppId,
      Env.username,
      Env.password,
      { dashboardFlag: true }
    );
  }
}

new DashboardTestCase(
  TemplateProject.Dashboard,
  17290453,
  "v-ivanchen@microsoft.com",
  "local",
  [
    LocalDebugTaskLabel.StartFrontend,
    LocalDebugTaskLabel.WatchBackend,
    LocalDebugTaskLabel.StartBackend,
  ],
  {
    dashboardFlag: true,
  }
).test();
