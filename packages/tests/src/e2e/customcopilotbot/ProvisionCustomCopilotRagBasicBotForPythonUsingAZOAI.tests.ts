// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Frank Qian <frankqian@microsoft.com>
 */

import { Capability } from "../../utils/constants";
import { CaseFactory } from "../caseFactory";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";

class BasicRAGBotAzureOpenAITestCase extends CaseFactory {}

class BasicRAGBotOpenAITestCase extends CaseFactory {}

// Azure OpenAI
const myRecordAzOpenAI: Record<string, string> = {};
myRecordAzOpenAI["custom-copilot-rag"] = "custom-copilot-rag-customize";
myRecordAzOpenAI["llm-service"] = "llm-service-azure-openai";
myRecordAzOpenAI["azure-openai-key"] = "fake";
myRecordAzOpenAI["azure-openai-deployment-name"] = "fake";
myRecordAzOpenAI["azure-openai-endpoint"] = "https://test.com";
new BasicRAGBotAzureOpenAITestCase(
  Capability.RAG,
  27178092,
  "frankqian@microsoft.com",
  ["bot"],
  ProgrammingLanguage.PY,
  {},
  myRecordAzOpenAI
).test();
