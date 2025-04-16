// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import {
  initTeamsPage,
  reopenTeamsPage,
  validateMeeting,
} from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";
import { SampledebugContext } from "./sampledebugContext";

class MyFirstMeetingTestCase extends CaseFactory {
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
  public override async onReopenPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string,
    options?:
      | {
          teamsAppName: string;
          includeFunction: boolean;
          npmName: string;
          dashboardFlag: boolean;
          type: string;
          env: "local" | "dev";
        }
      | undefined
  ): Promise<Page> {
    return await reopenTeamsPage(
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

  override async onCliValidate(page: Page): Promise<void> {
    return await validateMeeting(page, Env.username);
  }

  override async onValidate(page: Page): Promise<void> {
    return await validateMeeting(page, Env.username);
  }
}

new MyFirstMeetingTestCase(
  TemplateProject.MyFirstMeeting,

  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartFrontend],
  {
    teamsAppName: "fxuiMyFirs",
    type: "meeting",
    testPlanCaseId_local: 9958524,
    testPlanCaseId_dev: 14571880,
    //debug: "cli",
  }
).test();
