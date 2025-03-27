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
    }
  ): Promise<Page> {
    return await initTeamsPage(
      sampledebugContext.context!,
      teamsAppId,
      Env.username,
      Env.password,
      {
        projectPath: sampledebugContext.projectPath,
        env: "local",
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
        env: "local",
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
  9958524,
  "v-ivanchen@microsoft.com",
  "local",
  [LocalDebugTaskLabel.StartFrontend],
  {
    teamsAppName: "fxuiMyFirslocal",
    type: "meeting",
    debug: "cli",
  }
).test();
