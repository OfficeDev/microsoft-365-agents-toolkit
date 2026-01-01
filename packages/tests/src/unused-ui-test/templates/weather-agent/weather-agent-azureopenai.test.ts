// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { happyPathTest } from "./weatherAgentHappyPath";
import { Lang } from "../../../utils/constants";

happyPathTest({
  testPlanCaseId_local: 34648323,
  testPlanCaseId_dev: 34648345,
  author: "v-ivanchen@microsoft.com",
  lang: Lang.JS,
  llm: "llm-service-azure-openai",
});
