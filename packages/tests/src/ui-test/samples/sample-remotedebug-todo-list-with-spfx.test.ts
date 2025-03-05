// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject } from "../../utils/constants";
import {
  initTeamsPage,
  validateTodoListSpfx,
} from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import { Env } from "../../utils/env";

class TodoListSpfxTestCase extends CaseFactory {
  public override async onAfter(
    sampledebugContext: SampledebugContext
  ): Promise<void> {
    await sampledebugContext.sampleAfter(
      `${sampledebugContext.appName}-dev-rg`
    );
  }
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
        teamsAppName: options?.teamsAppName,
        type: options?.type,
      }
    );
  }
  public override async onValidate(page: Page): Promise<void> {
    return await validateTodoListSpfx(page, "fxuiTodoLidev");
  }
}

new TodoListSpfxTestCase(
  TemplateProject.TodoListSpfx,
  24121511,
  "v-ivanchen@microsoft.com",
  "dev",
  [],
  { teamsAppName: "TodoListSPFx-dev", type: "spfx" }
).test();
