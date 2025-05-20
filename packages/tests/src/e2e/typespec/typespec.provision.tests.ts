// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Bowen Song <bowen.song@microsoft.com>
 */

import { Capability } from "../../utils/constants";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { CaseFactory } from "../caseFactory";

class DeclarativeAgentWithTypeSpec extends CaseFactory {
  public async onAfterCreate(projectPath: string): Promise<void> {
    return Promise.resolve();
  }
}

const myRecord: Record<string, string> = {};
myRecord["with-plugin"] = "type-spec";

new DeclarativeAgentWithTypeSpec(
  Capability.DeclarativeAgent,
  32237977,
  "bowsong@microsoft.com",
  [],
  ProgrammingLanguage.None,
  {
    skipValidateForProvision: true,
    skipDeploy: true,
    skipPackage: true,
    skipErrorMessage: "@microsoft/typespec-m365-copilot",
  },
  myRecord
).test();
