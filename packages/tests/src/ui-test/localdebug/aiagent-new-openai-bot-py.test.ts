// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId_local: 27689383,
  testPlanCaseId_dev: 27689385,
  author: "v-ivanchen@microsoft.com",
  lang: "Python",
  llm: "llm-service-openai",
  agent: "custom-copilot-agent-new",
});
