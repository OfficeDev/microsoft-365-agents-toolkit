// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class DashboardTestCase extends CaseFactory {}

new DashboardTestCase(
  TemplateProjectFolder.Dashboard,
  24132131,
  "Ning.Tang@microsoft.com"
).test();
