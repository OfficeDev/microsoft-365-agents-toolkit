// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author quke@microsoft.com
 */

// ProxyAgent-NodeJS requires a live Azure AI Foundry Project endpoint, a
// deployed Agent ID, and OAuth App Registration credentials. Provision is
// skipped in automated tests; project creation and manifest validation are
// exercised instead.

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class ProxyAgentNodeJSTestCase extends CaseFactory {}

new ProxyAgentNodeJSTestCase(
  TemplateProjectFolder.ProxyAgentNodeJS,
  0,
  "quke@microsoft.com",
  ["bot"],
  { skipProvision: true }
).test();
