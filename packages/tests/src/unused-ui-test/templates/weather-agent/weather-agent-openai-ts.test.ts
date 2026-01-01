// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./weatherAgentHappyPath";
import { Lang } from "../../../utils/constants";

happyPathTest({
  testPlanCaseId_local: 34648300,
  testPlanCaseId_dev: 34648339,
  author: "v-ivanchen@microsoft.com",
  lang: Lang.TS,
  llm: "llm-service-openai",
});
