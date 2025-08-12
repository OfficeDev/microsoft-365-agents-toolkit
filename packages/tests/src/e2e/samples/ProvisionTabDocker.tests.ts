// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yimin Jin <yiminjin@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class TabDockerTestCase extends CaseFactory {}

new TabDockerTestCase(
  TemplateProjectFolder.TabDocker,
  27676823,
  "Ning.Tang@microsoft.com",
  ["aca"]
).test();
