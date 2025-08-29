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

// OpenAI
const myRecordOpenAI: Record<string, string> = {};
myRecordOpenAI["custom-copilot-agent"] = "custom-copilot-agent-new";
myRecordOpenAI["llm-service"] = "llm-service-openai";
myRecordOpenAI["openai-key"] = "fake";
new AiBotOpenAITestCase(
  Capability.Agent,
  27689385,
  "frankqian@microsoft.com",
  ["bot"],
  ProgrammingLanguage.PY,
  {},
  myRecordOpenAI
).test();
