// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateAdaptiveCard } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";

class AdaptiveCardTestCase extends CaseFactory {
  override async onValidate(
    page: Page,
    options?: { context: SampledebugContext; env: "local" | "dev" }
  ): Promise<void> {
    return await validateAdaptiveCard(page, {
      context: options?.context,
      env: options?.env,
    });
  }
}

new AdaptiveCardTestCase(
  TemplateProject.AdaptiveCard,
  14524987,
  24121425,
  "v-ivanchen@microsoft.com",
  [
    LocalDebugTaskLabel.StartLocalTunnel,
    LocalDebugTaskLabel.Azurite,
    LocalDebugTaskLabel.Compile,
    LocalDebugTaskLabel.StartBotApp,
  ]
).test();
