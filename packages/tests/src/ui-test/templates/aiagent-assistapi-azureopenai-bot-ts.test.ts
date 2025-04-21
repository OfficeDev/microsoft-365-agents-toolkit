// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";
import { Lang } from "../../utils/constants";

happyPathTest({
  testPlanCaseId_local: 30570676,
  testPlanCaseId_dev: 30578611,
  author: "v-ivanchen@microsoft.com",
  lang: Lang.TS,
  llm: "llm-service-azure-openai",
  agent: "custom-copilot-agent-assistants-api",
});
