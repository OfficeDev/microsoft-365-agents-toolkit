// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId_local: 28165244,
  testPlanCaseId_dev: 28165245,
  author: "v-ivanchen@microsoft.com",
  lang: "Python",
  llm: "llm-service-openai",
  agent: "custom-copilot-agent-assistants-api",
});
