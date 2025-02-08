// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

import { err, Inputs, ok, Platform } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import "mocha";
import * as sinon from "sinon";
import { createContext } from "../../../src/common/globalVars";
import { developerPortalScaffoldUtils } from "../../../src/component/developerPortalScaffoldUtils";
import { TdpGenerator } from "../../../src/component/generator/other/tdpGenerator";
import { InputValidationError, UserCancelError } from "../../../src/error";
import { ProgrammingLanguage, QuestionNames } from "../../../src/question";

describe("TdpGenerator", function () {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("activate()", () => {
    it("return true", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: ".",
        teamsAppFromTdp: {},
      };
      const context = createContext();
      const generator = new TdpGenerator();
      const res = generator.activate(context, inputs);
      chai.assert.isTrue(res);
    });
  });
  describe("getTemplateInfos()", () => {
    it("InputValidationError", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: ".",
        teamsAppFromTdp: {},
      };
      const context = createContext();
      const generator = new TdpGenerator();
      const res = await generator.getTemplateInfos(context, inputs, ".");
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof InputValidationError);
      }
    });
    it("happy", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: ".",
        teamsAppFromTdp: {},
        [QuestionNames.TemplateName]: "templateName",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.JS,
      };
      const context = createContext();
      const generator = new TdpGenerator();
      const res = await generator.getTemplateInfos(context, inputs, ".");
      chai.assert.isTrue(res.isOk());
    });
  });
  describe("post()", () => {
    it("update error", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: ".",
        teamsAppFromTdp: {},
      };
      sandbox
        .stub(developerPortalScaffoldUtils, "updateFilesForTdp")
        .resolves(err(new UserCancelError()));
      const context = createContext();
      const generator = new TdpGenerator();
      const res = await generator.post(context, inputs, ".");
      chai.assert.isTrue(res.isErr());
    });
    it("happy", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: ".",
        teamsAppFromTdp: {},
      };
      sandbox.stub(developerPortalScaffoldUtils, "updateFilesForTdp").resolves(ok(undefined));
      const context = createContext();
      const generator = new TdpGenerator();
      const res = await generator.post(context, inputs, ".");
      chai.assert.isTrue(res.isOk());
    });
  });
});
