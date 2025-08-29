// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Frank Qian <frankqian@microsoft.com>
 */

import { Capability } from "../../utils/constants";
import { CaseFactory } from "../caseFactory";
import * as fs from "fs-extra";
import * as path from "path";
import { expect } from "chai";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";

class AiSearchBotAzureOpenAITestCase extends CaseFactory {
  public override async onAfterCreate(projectPath: string): Promise<void> {
    expect(fs.pathExistsSync(path.resolve(projectPath, "infra"))).to.be.true;
    const userFile = path.resolve(projectPath, "env", `.env.dev.user`);
    const AZURE_OPENAI_EMBEDDING_DEPLOYMENT =
      "AZURE_OPENAI_EMBEDDING_DEPLOYMENT=fake";
    const SECRET_AZURE_SEARCH_KEY = "SECRET_AZURE_SEARCH_KEY=fake";
    const AZURE_SEARCH_ENDPOINT = "AZURE_SEARCH_ENDPOINT=https://test.com";
    const KEY =
      "\n" +
      AZURE_OPENAI_EMBEDDING_DEPLOYMENT +
      "\n" +
      SECRET_AZURE_SEARCH_KEY +
      "\n" +
      AZURE_SEARCH_ENDPOINT;
    fs.appendFileSync(userFile, KEY);
    console.log(`add key ${KEY} to .env.dev.user file`);
  }
}

// Azure OpenAI
const myRecordAzOpenAI: Record<string, string> = {};
myRecordAzOpenAI["custom-copilot-rag"] = "custom-copilot-rag-azureAISearch";
myRecordAzOpenAI["llm-service"] = "llm-service-azure-openai";
myRecordAzOpenAI["azure-openai-key"] = "fake";
myRecordAzOpenAI["azure-openai-deployment-name"] = "fake";
myRecordAzOpenAI["azure-openai-endpoint"] = "https://test.com";
new AiSearchBotAzureOpenAITestCase(
  Capability.RAG,
  27454388,
  "frankqian@microsoft.com",
  ["bot"],
  ProgrammingLanguage.PY,
  {},
  myRecordAzOpenAI
).test();
