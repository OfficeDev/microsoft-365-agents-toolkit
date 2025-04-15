// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId_local: 28957868,
  testPlanCaseId_dev: 28957869,
  author: "v-ivanchen@microsoft.com",
  lang: "Python",
  llm: "llm-service-azure-openai",
  agent: "custom-copilot-agent-assistants-api",
});
