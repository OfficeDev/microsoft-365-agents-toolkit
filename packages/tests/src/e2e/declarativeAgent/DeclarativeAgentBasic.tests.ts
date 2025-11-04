// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Hui Miao <huimaio@microsoft.com>
 */

import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { Capability } from "../../utils/constants";
import { CaseFactory } from "../caseFactory";

class DeclarativeAgentBasicTestCase extends CaseFactory {
  public async onAfterCreate(projectPath: string): Promise<void> {
    return Promise.resolve();
  }
}

const myRecord: Record<string, string> = {};
myRecord["with-plugin"] = "no";

new DeclarativeAgentBasicTestCase(
  Capability.DeclarativeAgent,
  27971545,
  "huimaio@microsoft.com",
  ["function"],
  ProgrammingLanguage.None,
  { skipValidateForProvision: true, skipDeploy: true },
  myRecord
).test();
