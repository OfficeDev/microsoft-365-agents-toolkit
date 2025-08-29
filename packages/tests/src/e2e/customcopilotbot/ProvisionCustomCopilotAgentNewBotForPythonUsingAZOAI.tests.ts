// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Frank Qian <frankqian@microsoft.com>
 */

import { Capability } from "../../utils/constants";
import { CaseFactory } from "../caseFactory";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";

class AiBotAzureOpenAITestCase extends CaseFactory {}

class AiBotOpenAITestCase extends CaseFactory {}

// Azure OpenAI
const myRecordAzOpenAI: Record<string, string> = {};
myRecordAzOpenAI["custom-copilot-agent"] = "custom-copilot-agent-new";
myRecordAzOpenAI["llm-service"] = "llm-service-azure-openai";
myRecordAzOpenAI["azure-openai-key"] = "fake";
myRecordAzOpenAI["azure-openai-deployment-name"] = "fake";
myRecordAzOpenAI["azure-openai-endpoint"] = "https://test.com";
new AiBotAzureOpenAITestCase(
  Capability.Agent,
  27689384,
  "frankqian@microsoft.com",
  ["bot"],
  ProgrammingLanguage.PY,
  {},
  myRecordAzOpenAI
).test();
