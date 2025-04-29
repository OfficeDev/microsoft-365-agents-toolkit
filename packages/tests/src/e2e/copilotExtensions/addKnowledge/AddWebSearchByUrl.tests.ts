// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yimin Jin <yiminjin@microsoft.com>
 */

import { Capability } from "../../../utils/constants";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { CaseFactory } from "../../caseFactory";
import { Executor } from "../../../utils/executor";
import { expect } from "chai";

class AddWebSearchByUrl extends CaseFactory {
  public async onAfterCreate(projectPath: string): Promise<void> {
    const command = `atk add capability -i false --knowledge-source web-search --search-type url --web-content https://example.com`;
    const { success } = await Executor.execute(command, projectPath);
    expect(success).to.be.true;
    return Promise.resolve();
  }
}

const myRecord: Record<string, string> = {};
myRecord["with-plugin"] = "yes";
myRecord["api-plugin-type"] = "new-api";
myRecord["api-auth"] = "none";

new AddWebSearchByUrl(
  Capability.DeclarativeAgent,
  31721023,
  "yiminjin@microsoft.com",
  ["function"],
  ProgrammingLanguage.JS,
  { skipDeploy: true },
  myRecord
).test();
