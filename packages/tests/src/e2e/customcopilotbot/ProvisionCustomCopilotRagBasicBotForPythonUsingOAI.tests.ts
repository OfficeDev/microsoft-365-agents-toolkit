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

// OpenAI
const myRecordOpenAI: Record<string, string> = {};
myRecordOpenAI["custom-copilot-rag"] = "custom-copilot-rag-customize";
myRecordOpenAI["llm-service"] = "llm-service-openai";
myRecordOpenAI["openai-key"] = "fake";
new BasicRAGBotOpenAITestCase(
  Capability.RAG,
  27178104,
  "frankqian@microsoft.com",
  ["bot"],
  ProgrammingLanguage.PY,
  {},
  myRecordOpenAI
).test();
