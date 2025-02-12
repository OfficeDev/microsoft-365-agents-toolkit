import { Inputs, Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import path from "path";
import sinon, { createSandbox } from "sinon";
import { createContext, setTools } from "../../../src/common/globalVars";
import { Generator } from "../../../src/component/generator/generator";
import { Generators } from "../../../src/component/generator/generatorProvider";
import { DefaultTemplateGenerator } from "../../../src/component/generator/templates/templateGenerator";
import { TemplateInfo } from "../../../src/component/generator/templates/templateInfo";
import {
  TemplateNames,
  inputsToTemplateName,
} from "../../../src/component/generator/templates/templateNames";
import { CapabilityOptions, QuestionNames } from "../../../src/question";
import { ProgrammingLanguage } from "../../../src/question/constants";
import { MockTools, randomAppName } from "../../core/utils";

describe("TemplateGenerator", () => {
  const testInputsToTemplateName = new Map([
    ...inputsToTemplateName,
    [
      {
        [QuestionNames.Capabilities]: CapabilityOptions.tab().id,
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        targetFramework: "net8.0",
      },
      TemplateNames.SsoTabSSR,
    ],
    [
      {
        [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        targetFramework: "net8.0",
      },
      TemplateNames.TabSSR,
    ],
    [
      {
        [QuestionNames.Capabilities]: CapabilityOptions.tab().id,
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        targetFramework: "net9.0",
      },
      TemplateNames.SsoTabSSR,
    ],
    [
      {
        [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        targetFramework: "net9.0",
      },
      TemplateNames.TabSSR,
    ],
  ]);

  setTools(new MockTools());
  const ctx = createContext();
  const destinationPath = path.join(__dirname, "tmp");
  const sandbox = createSandbox();
  let scaffoldingSpy: sinon.SinonSpy;
  let inputs: Inputs;

  beforeEach(() => {
    scaffoldingSpy = sandbox.spy(DefaultTemplateGenerator.prototype, <any>"scaffolding");
    sandbox.stub(Generator, "generate").resolves();
    inputs = {
      platform: Platform.VSCode,
      [QuestionNames.AppName]: randomAppName(),
      [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.JS,
    } as Inputs;
  });

  afterEach(() => {
    sandbox.restore();
  });

  testInputsToTemplateName.forEach(async (templateName, _inputs) => {
    it(`scaffolding ${templateName}`, async () => {
      inputs = { ...inputs, ..._inputs };
      const res = await Generators.find((g) => g.activate(ctx, inputs))?.run(
        ctx,
        inputs,
        destinationPath
      );

      assert.isTrue(res?.isOk());
      assert.isTrue(scaffoldingSpy.calledOnce);
      assert.equal((scaffoldingSpy.args[0][1] as TemplateInfo).templateName, templateName);
      assert.equal(
        (scaffoldingSpy.args[0][1] as TemplateInfo).language,
        inputs?.[QuestionNames.ProgrammingLanguage] || ProgrammingLanguage.JS
      );
    });
  });
});
