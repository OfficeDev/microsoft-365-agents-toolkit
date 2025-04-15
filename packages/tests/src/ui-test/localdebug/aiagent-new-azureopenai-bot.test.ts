// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId_local: 27042860,
  testPlanCaseId_dev: 27042864,
  author: "v-ivanchen@microsoft.com",
  lang: "JavaScript",
  llm: "llm-service-azure-openai",
  agent: "custom-copilot-agent-new",
});
