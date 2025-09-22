// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Bowen Song <bowen.song@microsoft.com>
 */

import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import * as fs from "fs-extra";
import path from "path";
import { Capability } from "../../../utils/constants";
import { CaseFactory } from "../../caseFactory";

class DeclarativeAgentWithTypeSpec extends CaseFactory {
  public async onAfterCreate(projectPath: string): Promise<void> {
    // Update the project to remove comments
    const mainFilePath = path.join(projectPath, "main.tsp");
    const mainFileContent = await fs.readFile(mainFilePath, "utf-8");
    const updateContent = mainFileContent
      .replace(
        "// Uncomment this part to add a conversation starter to the agent.",
        ""
      )
      .replace(
        "// This will be shown to the user when the agent is first created.",
        ""
      )
      .replace("  // Uncomment this part to add actions to the agent.", "")
      .replace(/\/\/ /g, "");
    await fs.writeFile(mainFilePath, updateContent, "utf-8");
    return Promise.resolve();
  }
}

const myRecord: Record<string, string> = {};
myRecord["with-plugin"] = "type-spec";

new DeclarativeAgentWithTypeSpec(
  Capability.DeclarativeAgent,
  32772441,
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
