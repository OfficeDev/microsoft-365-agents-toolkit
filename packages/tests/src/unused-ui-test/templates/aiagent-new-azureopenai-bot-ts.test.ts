// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";
import { Lang } from "../../utils/constants";

happyPathTest({
  testPlanCaseId_local: 27042862,
  testPlanCaseId_dev: 27042866,
  author: "v-ivanchen@microsoft.com",
  lang: Lang.TS,
  llm: "llm-service-azure-openai",
  agent: "custom-copilot-agent-new",
});
