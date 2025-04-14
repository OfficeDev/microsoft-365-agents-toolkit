// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId: 28957868,
  author: "v-ivanchen@microsoft.com",
  lang: "python",
  llm: "llm-service-azure-openai",
  agent: "custom-copilot-agent-assistants-api",
});
