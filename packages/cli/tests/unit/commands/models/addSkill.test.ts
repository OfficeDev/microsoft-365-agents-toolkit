// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* eslint-disable @typescript-eslint/no-explicit-any */

import { CLIContext, ok } from "@microsoft/teamsfx-api";
import { FxCore } from "@microsoft/teamsfx-core";
import { assert } from "chai";
import "mocha";
import * as sinon from "sinon";
import * as activate from "../../../../src/activate";
import { addSkillCommand } from "../../../../src/commands/models/addSkill";

describe("addSkill command", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("should call FxCore.addSkill with command inputs", async () => {
    const mockCore = new FxCore({} as any);
    const addSkillStub = sandbox.stub(mockCore, "addSkill").resolves(ok(undefined));
    sandbox.stub(activate, "getFxCore").returns(mockCore);

    const ctx: CLIContext = {
      command: { ...addSkillCommand, fullName: "add skill" },
      optionValues: {
        folder: "./",
        skill: "mySkill",
        description: "A test skill",
      },
      globalOptionValues: {},
      argumentValues: [],
      telemetryProperties: {},
    };

    const result = await addSkillCommand.handler!(ctx);

    assert.isTrue(result.isOk());
    assert.isTrue(addSkillStub.calledOnceWith(ctx.optionValues));
  });
});
