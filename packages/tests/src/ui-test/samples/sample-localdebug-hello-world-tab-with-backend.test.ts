// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateTab, reopenPage } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";
import { SampledebugContext } from "./sampledebugContext";

class HelloWorldTabBackEndTestCase extends CaseFactory {
  override async onValidate(
    page: Page,
    options?: { includeFunction: boolean }
  ): Promise<void> {
    return await validateTab(
      page,
      {
        displayName: Env.displayName,
        includeFunction: options?.includeFunction,
      },
      true
    );
  }
  override async onCliValidate(
    page: Page,
    options?: { includeFunction: boolean }
  ): Promise<void> {
    return await validateTab(page, {
      displayName: Env.displayName,
      includeFunction: options?.includeFunction,
    });
  }
  public override async onReopenPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string
  ): Promise<Page> {
    return await reopenPage(
      sampledebugContext.context!,
      teamsAppId,
      undefined,
      undefined,
      {
        projectPath: sampledebugContext.projectPath,
        env: "local",
      }
    );
  }
}

new HelloWorldTabBackEndTestCase(
  TemplateProject.HelloWorldTabBackEnd,
  12684063,
  "v-ivanchen@microsoft.com",
  "local",
  [
    LocalDebugTaskLabel.StartFrontend,
    LocalDebugTaskLabel.WatchBackend,
    LocalDebugTaskLabel.StartBackend,
  ],
  { debug: "cli" }
).test();
