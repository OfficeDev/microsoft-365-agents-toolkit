// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class CopilotConnectorTestCase extends CaseFactory {}

new CopilotConnectorTestCase(
  TemplateProjectFolder.CopilotConnector,
  15277460,
  "junhan@microsoft.com",
  ["tab", "aad"]
).test();
