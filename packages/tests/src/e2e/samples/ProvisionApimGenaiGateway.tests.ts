// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author quke@microsoft.com
 */

// apim-genai-gateway references an *existing* APIM service created in a prior
// deployment step. Full provision requires that pre-existing infrastructure, so
// only project creation and manifest validation are exercised here.

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class ApimGenaiGatewayTestCase extends CaseFactory {}

new ApimGenaiGatewayTestCase(
  TemplateProjectFolder.ApimGenaiGateway,
  0,
  "quke@microsoft.com",
  ["bot"],
  { skipProvision: true }
).test();
