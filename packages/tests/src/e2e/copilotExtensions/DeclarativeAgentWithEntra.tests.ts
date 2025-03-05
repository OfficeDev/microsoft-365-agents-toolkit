// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yimin Jin <yiminjin@microsoft.com>
 */

import { Capability } from "../../utils/constants";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { CaseFactory } from "../caseFactory";

class DeclarativeAgentWithEntra extends CaseFactory {}

const myRecord: Record<string, string> = {};
myRecord["with-plugin"] = "yes";
myRecord["api-plugin-type"] = "new-api";
myRecord["api-auth"] = "microsoft-entra";

new DeclarativeAgentWithEntra(
  Capability.DeclarativeAgent,
  30310142,
  "yiminjin@microsoft.com",
  ["function"],
  ProgrammingLanguage.JS,
  {},
  myRecord
).test();

new DeclarativeAgentWithEntra(
  Capability.DeclarativeAgent,
  30309989,
  "yiminjin@microsoft.com",
  ["function"],
  ProgrammingLanguage.TS,
  {},
  myRecord
).test();
