// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./weatherAgentHappyPath";
import { Lang } from "../../../utils/constants";

happyPathTest({
  testPlanCaseId_local: 34648342,
  testPlanCaseId_dev: 34648368,
  author: "v-ivanchen@microsoft.com",
  lang: Lang.TS,
  llm: "llm-service-azure-openai",
});
