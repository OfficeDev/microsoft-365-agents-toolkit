// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class DiceRollerTestCase extends CaseFactory {}

new DiceRollerTestCase(
  TemplateProject.DiceRoller,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartLocalTunnel, LocalDebugTaskLabel.StartWebServer],
  {
    skipInit: true,
    testPlanCaseId_local: 21320394,
    testPlanCaseId_dev: 24121529,
  }
).test();
