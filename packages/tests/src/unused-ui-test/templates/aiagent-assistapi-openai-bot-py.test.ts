// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./aiagentBotHappyPathCommon";
import { Lang } from "../../utils/constants";

happyPathTest({
  testPlanCaseId_local: 28165244,
  testPlanCaseId_dev: 28165245,
  author: "v-ivanchen@microsoft.com",
  lang: Lang.PY,
  llm: "llm-service-openai",
  agent: "custom-copilot-agent-assistants-api",
});
