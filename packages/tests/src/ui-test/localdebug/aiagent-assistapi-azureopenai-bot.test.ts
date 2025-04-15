// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId_local: 30570629,
  testPlanCaseId_dev: 30578628,
  author: "v-ivanchen@microsoft.com",
  lang: "JavaScript",
  llm: "llm-service-azure-openai",
  agent: "custom-copilot-agent-assistants-api",
});
