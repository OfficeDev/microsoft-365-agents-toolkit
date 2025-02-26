// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  Context,
  FuncValidation,
  Inputs,
  LocalFunc,
  LogProvider,
  Platform,
  Question,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import * as path from "path";
import sinon from "sinon";
import * as utils from "../../src/common/globalVars";
import { getLocalizedString } from "../../src/common/localizeUtils";
import { FileNotFoundError } from "../../src/error";
import { appNameQuestion, folderQuestion, getSolutionName } from "../../src/question";
import { randomAppName } from "../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../plugins/solution/util";

export async function callFuncs(question: Question, inputs: Inputs, answer?: string) {
  try {
    if (question.default && typeof question.default !== "string") {
      await (question.default as LocalFunc<string | undefined>)(inputs);
    }

    if (
      (question.type === "singleSelect" || question.type === "multiSelect") &&
      typeof question.dynamicOptions !== "object" &&
      question.dynamicOptions
    ) {
      await question.dynamicOptions(inputs);
    }
    if (answer && (question as any).validation?.validFunc) {
      await (question as any).validation.validFunc(answer, inputs);
    }

    if ((question as any).placeholder && typeof (question as any).placeholder !== "string") {
      await (question as any).placeholder(inputs);
    }
  } catch (e) {}
}

describe("scaffold question", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("appNameQuestion", () => {
    const question = appNameQuestion();
    const validFunc = (question.validation as FuncValidation<string>).validFunc;
    it("happy path", async () => {
      const inputs: Inputs = { platform: Platform.VSCode, folder: "./" };
      const appName = "1234";
      let validRes = await validFunc(appName, inputs);
      assert.isTrue(validRes === getLocalizedString("core.QuestionAppName.validation.pattern"));
      sandbox.stub<any, any>(fs, "pathExists").resolves(true);
      inputs.appName = randomAppName();
      inputs.folder = "./";
      validRes = await validFunc(inputs.appName, inputs);
      const expected = getLocalizedString(
        "core.QuestionAppName.validation.pathExist",
        path.resolve(inputs.folder, inputs.appName)
      );
      assert.equal(validRes, expected);
      sandbox.restore();
      sandbox.stub<any, any>(fs, "pathExists").resolves(false);
      validRes = await validFunc(inputs.appName, inputs);
      assert.isTrue(validRes === undefined);
    });

    it("app name has 25 length - VSC", async () => {
      const mockedUI = new MockedUserInteraction();
      sandbox.stub(utils, "createContext").returns({
        userInteraction: mockedUI,
      } as Context);
      const showMessageStub = sandbox.stub(mockedUI, "showMessage");

      const input = "abcdefghijklmnopqrstuvwxy";
      await validFunc(input, { platform: Platform.VSCode });

      assert.isTrue(showMessageStub.calledOnce);
    });

    it("app name has 25 length - VS", async () => {
      const mockedLogProvider = new MockedLogProvider();
      sandbox.stub(utils, "createContext").returns({
        logProvider: mockedLogProvider as LogProvider,
      } as Context);
      const warningStub = sandbox.stub(mockedLogProvider, "warning");

      const input = "abcdefghijklmnopqrstuvwxy";
      await validFunc(input, { platform: Platform.VS });

      assert.isTrue(warningStub.calledOnce);

      await validFunc(input);

      assert.isTrue(warningStub.calledTwice);
    });

    it("app name exceed maxlength of 30", async () => {
      const input = "SurveyMonkeyWebhookNotification";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.maxlength"));
    });

    it("app name with only letters", async () => {
      const input = "app";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name starting with digit", async () => {
      const input = "123app";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name count of alphanumerics less than 2", async () => {
      const input = "a..(";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name containing dot", async () => {
      const input = "app.123";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name containing hyphen", async () => {
      const input = "app-123";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name containing multiple special characters", async () => {
      const input = "a..(1";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name containing space", async () => {
      const input = "app 123";
      const result = await validFunc(input);

      assert.isUndefined(result);
    });

    it("app name containing dot at the end - wrong pattern", async () => {
      const input = "app.app.";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name containing space at the end - wrong pattern", async () => {
      const input = "app123 ";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name containing invalid control code", async () => {
      const input = "a\u0001a";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("app name containing invalid character", async () => {
      const input = "app<>123";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });

    it("invalid app name containing &", async () => {
      const input = "app&123";
      const result = await validFunc(input);

      assert.equal(result, getLocalizedString("core.QuestionAppName.validation.pattern"));
    });
  });

  describe("folderQuestion", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("should find taskpane template", () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
      };
      const question = folderQuestion() as any;
      const title = question.title(inputs);
      const defaultV = question.default(inputs);
      assert.equal(title, "Directory where the project folder will be created in");
      assert.equal(defaultV, "./");
    });
  });

  describe("getSolutionName", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("happy path", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJson").resolves({
        "@microsoft/generator-sharepoint": {
          solutionName: "testSolutionName",
        },
      });
      const res = await getSolutionName("");
      assert.equal(res, "testSolutionName");
    });

    it("FileNotFoundError", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      try {
        await getSolutionName(".");
        assert.fail("should throw");
      } catch (e) {
        assert.isTrue(e instanceof FileNotFoundError);
      }
    });

    it("undefined", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJson").resolves({});
      const res = await getSolutionName("");
      assert.isUndefined(res);
    });
  });
});
