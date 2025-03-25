// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateContact, reopenPage } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";
import { SampledebugContext } from "./sampledebugContext";

class ContactExporterTestCase extends CaseFactory {
  public override async onValidate(page: Page): Promise<void> {
    return await validateContact(page, { displayName: Env.displayName });
  }
  public override async onCliValidate(page: Page): Promise<void> {
    return await validateContact(page, { displayName: Env.displayName });
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

new ContactExporterTestCase(
  TemplateProject.ContactExporter,
  12599484,
  "v-ivanchen@microsoft.com",
  "local",
  [LocalDebugTaskLabel.StartFrontend],
  { debug: "cli" }
).test();
