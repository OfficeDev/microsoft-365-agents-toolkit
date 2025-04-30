// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";
import { Lang } from "../../utils/constants";

happyPathTest({
  testPlanCaseId_local: 27042905,
  testPlanCaseId_dev: 27042909,
  author: "v-ivanchen@microsoft.com",
  lang: Lang.TS,
  llm: "llm-service-openai",
  agent: "custom-copilot-agent-assistants-api",
});
