// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yimin Jin <yiminjin@microsoft.com>
 */

import { Capability } from "../../utils/constants";
import { ProgrammingLanguage } from "@microsoft/teamsfx-core";
import { CaseFactory } from "../caseFactory";

class DeclarativeAgentWithNoneAuth extends CaseFactory {}

const myRecord: Record<string, string> = {};
myRecord["with-plugin"] = "yes";
myRecord["api-plugin-type"] = "new-api";
myRecord["api-auth"] = "none";

new DeclarativeAgentWithNoneAuth(
  Capability.DeclarativeAgent,
  30310138,
  "yiminjin@microsoft.com",
  ["function"],
  ProgrammingLanguage.JS,
  {},
  myRecord
).test();

new DeclarativeAgentWithNoneAuth(
  Capability.DeclarativeAgent,
  30309984,
  "yiminjin@microsoft.com",
  ["function"],
  ProgrammingLanguage.TS,
  {},
  myRecord
).test();
