// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";
import { Lang } from "../../utils/constants";

happyPathTest({
  testPlanCaseId_local: 27042861,
  testPlanCaseId_dev: 27042865,
  author: "v-ivanchen@microsoft.com",
  lang: Lang.JS,
  llm: "llm-service-openai",
  agent: "custom-copilot-agent-new",
});
