// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId_local: 27042903,
  testPlanCaseId_dev: 27042907,
  author: "v-ivanchen@microsoft.com",
  lang: "TypeScript",
  llm: "llm-service-openai",
  agent: "custom-copilot-agent-assistants-api",
});
