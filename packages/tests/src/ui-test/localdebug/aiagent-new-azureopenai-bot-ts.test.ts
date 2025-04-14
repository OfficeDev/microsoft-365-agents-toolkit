// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";

happyPathTest({
  testPlanCaseId: 27042862,
  author: "v-ivanchen@microsoft.com",
  lang: "typescript",
  llm: "llm-service-azure-openai",
  agent: "custom-copilot-agent-new",
});
