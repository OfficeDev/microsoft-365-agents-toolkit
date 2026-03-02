// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Hui Miao <huimaio@microsoft.com>
 */

import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { Capability } from "../../utils/constants";
import { CaseFactory } from "../caseFactory";

class DeclarativeAgentWithOAutTestCase extends CaseFactory {}

const myRecord: Record<string, string> = {};
myRecord["with-plugin"] = "yes";
myRecord["api-plugin-type"] = "new-api";
myRecord["api-auth"] = "oauth";

new DeclarativeAgentWithOAutTestCase(
  Capability.DeclarativeAgentWithActionFromScratchOAuth,
  30310025,
  "huimaio@microsoft.com",
  ["function"],
  ProgrammingLanguage.JS,
  {},
  myRecord
).test();

new DeclarativeAgentWithOAutTestCase(
  Capability.DeclarativeAgentWithActionFromScratchOAuth,
  30032802,
  "huimaio@microsoft.com",
  ["function"],
  ProgrammingLanguage.TS,
  {},
  myRecord
).test();
