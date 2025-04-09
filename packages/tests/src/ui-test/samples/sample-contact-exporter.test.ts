// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateContact } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";

class ContactExporterTestCase extends CaseFactory {
  public override async onValidate(page: Page): Promise<void> {
    return await validateContact(page, { displayName: Env.displayName });
  }
}

new ContactExporterTestCase(
  TemplateProject.ContactExporter,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartFrontend],
  {
    testPlanCaseId_local: 12599484,
    testPlanCaseId_dev: 14571878,
  }
).test();
