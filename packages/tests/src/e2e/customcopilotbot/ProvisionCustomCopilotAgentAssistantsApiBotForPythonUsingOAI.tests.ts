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

class AgentAssitantApiOpenAITestCase extends CaseFactory {
  public override async onAfterCreate(projectPath: string): Promise<void> {
    expect(fs.pathExistsSync(path.resolve(projectPath, "infra"))).to.be.true;
    const userFile = path.resolve(projectPath, "env", `.env.dev.user`);
    const OPENAI_ASSISTANT_ID = "OPENAI_ASSISTANT_ID=fake";
    const KEY = "\n" + OPENAI_ASSISTANT_ID;
    fs.appendFileSync(userFile, KEY);
    console.log(`add key ${KEY} to .env.dev.user file`);
  }
}

// OpenAI
const myRecordOpenAI: Record<string, string> = {};
myRecordOpenAI["custom-copilot-agent"] = "custom-copilot-agent-assistants-api";
myRecordOpenAI["llm-service"] = "llm-service-openai";
myRecordOpenAI["openai-key"] = "fake";
new AgentAssitantApiOpenAITestCase(
  Capability.Agent,
  28165245,
  "frankqian@microsoft.com",
  ["bot"],
  ProgrammingLanguage.PY,
  {},
  myRecordOpenAI
).test();
