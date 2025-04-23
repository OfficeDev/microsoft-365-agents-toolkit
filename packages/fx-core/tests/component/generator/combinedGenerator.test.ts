// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author zhaofengxu@microsoft.com
 */

import { Inputs, OptionItem, Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import { RestoreFn } from "mocked-env";
import sinon from "sinon";
import { createContext, setTools } from "../../../src/common/globalVars";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import {
  ApiAuthOptions,
  CapabilityOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../src/question";
import { MockTools } from "../../core/utils";
import { CombinedProjectGenerator } from "../../../src/component/generator/combinedProject/generator";
import { setTemplateNameAndGC } from "../../../src/question/scaffold/vsc/daProjectTypeNode";
import { DACapabilityOptions } from "../../../src/question/scaffold/vsc/CapabilityOptions";

describe("combined generator", async () => {
  setTools(new MockTools());
  let mockedEnvRestore: RestoreFn | undefined;
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });
  describe("activate and get template name", async () => {
    it("declarative-agent-with-graph-connector", async () => {
      const generator = new CombinedProjectGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithGraphConnector,
        [QuestionNames.AppName]: "app",
      };
      let res = await generator.activate(context, inputs);
      let info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, TemplateNames.GraphConnector);

      inputs[QuestionNames.ProgrammingLanguage] = ProgrammingLanguage.CSharp;
      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.microsoftEntra().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.DeclarativeAgentBasic;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isFalse(res);
      assert.equal(info.isOk() && info.value[0].templateName, TemplateNames.DeclarativeAgentBasic);
    });
  });

  describe("post", async () => {
    it("add plugin success", async () => {
      const generator = new CombinedProjectGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithGraphConnector,
        [QuestionNames.AppName]: "app",
      };

      sandbox.stub(fs, "copySync").returns();
      sandbox.stub(fs, "removeSync").returns();

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());
    });
  });

  describe("functions", async () => {
    it("setTemplateNameAndGC", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithGraphConnector,
        [QuestionNames.AppName]: "app",
      };
      const optionItem = {
        id: DACapabilityOptions.withGC().id,
        data: "test-template",
      } as OptionItem;

      setTemplateNameAndGC(optionItem, inputs);

      optionItem.id = DACapabilityOptions.noPlugin().id;
      setTemplateNameAndGC(optionItem, inputs);
    });
  });
});
