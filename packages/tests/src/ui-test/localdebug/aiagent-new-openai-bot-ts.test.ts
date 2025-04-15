// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId_local: 27042863,
  testPlanCaseId_dev: 27042867,
  author: "v-ivanchen@microsoft.com",
  lang: "TypeScript",
  llm: "llm-service-openai",
  agent: "custom-copilot-agent-new",
});
