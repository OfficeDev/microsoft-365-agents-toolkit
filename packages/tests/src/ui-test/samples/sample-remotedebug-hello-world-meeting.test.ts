// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject } from "../../utils/constants";
import {
  initTeamsPage,
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
        env: "dev",
        teamsAppName: options?.teamsAppName,
        type: options?.type,
      }
    );
  }

  override async onValidate(page: Page): Promise<void> {
    return await validateMeeting(page, Env.username);
  }
}

new MyFirstMeetingTestCase(
  TemplateProject.MyFirstMeeting,
  14571880,
  "v-ivanchen@microsoft.com",
  "dev",
  [],
  {
    teamsAppName: "fxuiMyFirsdev",
    type: "meeting",
  }
).test();
