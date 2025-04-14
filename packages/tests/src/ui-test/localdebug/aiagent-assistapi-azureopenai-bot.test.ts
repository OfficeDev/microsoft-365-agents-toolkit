// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId: 30570629,
  author: "v-ivanchen@microsoft.com",
  lang: "javascript",
  llm: "llm-service-azure-openai",
  agent: "custom-copilot-agent-assistants-api",
});
