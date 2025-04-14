// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId: 27689383,
  author: "v-ivanchen@microsoft.com",
  lang: "python",
  llm: "llm-service-openai",
  agent: "custom-copilot-agent-new",
});
