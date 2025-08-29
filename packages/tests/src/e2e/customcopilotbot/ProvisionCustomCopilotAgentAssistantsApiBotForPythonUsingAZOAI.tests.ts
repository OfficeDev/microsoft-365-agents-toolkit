// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Frank Qian <frankqian@microsoft.com>
 */

import { Capability } from "../../utils/constants";
import { CaseFactory } from "../caseFactory";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import * as fs from "fs-extra";
import * as path from "path";
import { expect } from "chai";

class AgentAssitantApiAzureOpenAITestCase extends CaseFactory {
  public override async onAfterCreate(projectPath: string): Promise<void> {
    expect(fs.pathExistsSync(path.resolve(projectPath, "infra"))).to.be.true;
    const userFile = path.resolve(projectPath, "env", `.env.dev.user`);
    const AZURE_OPENAI_ASSISTANT_ID = "AZURE_OPENAI_ASSISTANT_ID=fake";
    const KEY = "\n" + AZURE_OPENAI_ASSISTANT_ID;
    fs.appendFileSync(userFile, KEY);
    console.log(`add key ${KEY} to .env.dev.user file`);
  }
}

// OpenAI
const myRecordAzOpenAI: Record<string, string> = {};
myRecordAzOpenAI["custom-copilot-agent"] =
  "custom-copilot-agent-assistants-api";
myRecordAzOpenAI["llm-service"] = "llm-service-azure-openai";
myRecordAzOpenAI["azure-openai-key"] = "fake";
myRecordAzOpenAI["azure-openai-deployment-name"] = "fake";
myRecordAzOpenAI["azure-openai-endpoint"] = "https://test.com";
new AgentAssitantApiAzureOpenAITestCase(
  Capability.Agent,
  28957869,
  "frankqian@microsoft.com",
  ["bot"],
  ProgrammingLanguage.PY,
  {},
  myRecordAzOpenAI
).test();
