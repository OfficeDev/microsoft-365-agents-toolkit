// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yimin Jin <yiminjin@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class BotSSODockerTestCase extends CaseFactory {}

new BotSSODockerTestCase(
  TemplateProjectFolder.BotSSODocker,
  27656551,
  "Ning.Tang@microsoft.com",
  ["aca"]
).test();
