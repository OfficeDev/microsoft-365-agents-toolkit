// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  Context,
  FuncValidation,
  Inputs,
  LogProvider,
  OptionItem,
  Platform,
  TokenProvider,
  UserError,
  err,
} from "@microsoft/teamsfx-api";
import axios, { AxiosError, AxiosResponse, InternalAxiosRequestConfig } from "axios";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as path from "path";
import sinon from "sinon";
import * as utils from "../../src/common/globalVars";
import { setTools } from "../../src/common/globalVars";
import { getLocalizedString } from "../../src/common/localizeUtils";
import * as stringUtils from "../../src/common/stringUtils";
import { OneDriveSharePointItemType } from "../../src/component/generator/constant";
import * as generatorHelper from "../../src/component/generator/declarativeAgent/helper";
import * as oneDriveSharePointHandler from "../../src/component/generator/declarativeAgent/oneDriveSharePointHandler";
import { FileNotFoundError } from "../../src/error";
import {
  ActionStartOptions,
  ApiAuthOptions,
  GCNameQuestion,
  MeArchitectureOptions,
  QuestionNames,
  apiAuthQuestion,
  apiOperationQuestion,
  apiPluginStartQuestion,
  appNameQuestion,
  folderQuestion,
  getSolutionName,
  oneDriveSharePointItemQuestion,
  webContentQuestion,
} from "../../src/question";
import { MockTools, MockUserInteraction, randomAppName } from "../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../plugins/solution/util";
import { DACapabilityOptions } from "../../src/question/scaffold/vsc/CapabilityOptions";

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

  describe("api plugin auth question", () => {
    const ui = new MockUserInteraction();
    let mockedEnvRestore: RestoreFn;
    const tools = new MockTools();
    setTools(tools);
    beforeEach(() => {
      mockedEnvRestore = mockedEnv({});
    });

    afterEach(() => {
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });
    it("api message extension", async () => {
      const question = apiAuthQuestion();
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      inputs[QuestionNames.MeArchitectureType] = MeArchitectureOptions.newApi().id;
      assert.isDefined(question.dynamicOptions);
      if (question.dynamicOptions) {
        const options = (await question.dynamicOptions(inputs)) as OptionItem[];
        assert.deepEqual(options, [
          ApiAuthOptions.none(),
          ApiAuthOptions.bearerToken(),
          ApiAuthOptions.microsoftEntra(),
        ]);
      }
    });

    it("api plugin from scratch with auth enabled", async () => {
      const question = apiAuthQuestion();
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      inputs[QuestionNames.ActionType] = ActionStartOptions.newApi().id;
      assert.isDefined(question.dynamicOptions);
      if (question.dynamicOptions) {
        const options = (await question.dynamicOptions(inputs)) as OptionItem[];
        assert.deepEqual(options, [
          ApiAuthOptions.none(),
          ApiAuthOptions.apiKey(),
          ApiAuthOptions.microsoftEntra(),
          ApiAuthOptions.oauth(),
        ]);
      }
    });

    it("api plugin from add action with auth enabled", async () => {
      const question = apiAuthQuestion(true);
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      inputs[QuestionNames.ActionType] = ActionStartOptions.newApi().id;
      assert.isDefined(question.dynamicOptions);
      if (question.dynamicOptions) {
        const options = (await question.dynamicOptions(inputs)) as OptionItem[];
        assert.deepEqual(options, [
          ApiAuthOptions.apiKey(),
          ApiAuthOptions.microsoftEntra(),
          ApiAuthOptions.oauth(),
        ]);
      }
    });
  });
  describe("api plugin auth question (AAD disabled)", () => {
    let mockedEnvRestore: RestoreFn;
    const tools = new MockTools();
    setTools(tools);
    beforeEach(() => {
      mockedEnvRestore = mockedEnv({});
    });

    afterEach(() => {
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });

    it("api plugin from scratch without AAD enabled", async () => {
      const question = apiAuthQuestion();
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      inputs[QuestionNames.ActionType] = ActionStartOptions.newApi().id;
      assert.isDefined(question.dynamicOptions);
      if (question.dynamicOptions) {
        const options = (await question.dynamicOptions(inputs)) as OptionItem[];
        assert.deepEqual(options, [
          ApiAuthOptions.none(),
          ApiAuthOptions.apiKey(),
          ApiAuthOptions.microsoftEntra(),
          ApiAuthOptions.oauth(),
        ]);
      }
    });
  });

  describe("add knowledge", () => {
    const tools = new MockTools();
    setTools(tools);
    afterEach(() => {
      sandbox.restore();
    });

    describe("Web Content", () => {
      it("happy path", async () => {
        const question = webContentQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("https://test.com", inputs);
        assert.isUndefined(res);
      });

      it("happy path", async () => {
        const question = webContentQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("https://test.com", inputs);
        assert.isUndefined(res);
      });

      it("error path: invalid url", async () => {
        const question = webContentQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("fakeUrl", inputs);
        assert.equal(res, "Invalid web content. Please provide a valid URL.");
      });

      it("error path: no inputs", async () => {
        const question = webContentQuestion();

        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        try {
          await validationSchema.validFunc?.("http://fakeUrl.com", undefined);
          assert.fail("Should throw error");
        } catch (err) {
          assert.isNotNull(err);
        }
      });
    });

    describe("OneDrive & SharePoint get ODSP item ", () => {
      it("happy path: site", async () => {
        const question = oneDriveSharePointItemQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const fakeAxiosInstance = axios.create();
        sandbox.stub(axios, "create").returns(fakeAxiosInstance);
        const axiosGetStub = sandbox.stub(fakeAxiosInstance, "get");
        axiosGetStub.onCall(0).resolves({
          status: 200,
          data: {
            id: "fakeId",
            name: "fakeName",
            sharepointIds: {
              webId: "fakeWebId",
              siteId: "fakeSiteId",
            },
          },
        });

        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("https://test.com", inputs);
        assert.deepEqual(inputs.oneDriveSharePointItem, [
          {
            id: "fakeId",
            name: "fakeName",
            siteId: "fakeSiteId",
            webId: "fakeWebId",
          },
        ]);
        assert.isUndefined(res);
      });

      it("happy path: drive", async () => {
        const question = oneDriveSharePointItemQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const fakeAxiosInstance = axios.create();
        sandbox.stub(axios, "create").returns(fakeAxiosInstance);
        const axiosGetStub = sandbox.stub(fakeAxiosInstance, "get");
        axiosGetStub
          .onCall(0)
          .resolves(err(new UserError("fakeError", "fakeError", "fakeError", "fakeError")));
        axiosGetStub.onCall(1).resolves({
          status: 200,
          data: {
            id: "fakeId",
            name: "fakeName",
            sharepointIds: {
              listItemUniqueId: "fakeUniqueId",
              listId: "fakeListId",
              webId: "fakeWebId",
              siteId: "fakeSiteId",
            },
            webUrl: "fakeWebUrl",
            file: "fakeFile",
          },
        });

        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("https://test.com", inputs);
        assert.deepEqual(inputs.oneDriveSharePointItem, [
          {
            id: "fakeId",
            itemType: OneDriveSharePointItemType.File,
            listId: "fakeListId",
            name: "fakeName",
            siteId: "fakeSiteId",
            uniqueId: "fakeUniqueId",
            webId: "fakeWebId",
          },
        ]);
        assert.isUndefined(res);
      });

      it("error path: invalid input url", async () => {
        const question = oneDriveSharePointItemQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };

        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("", inputs);
        assert.equal(res, "Please input a valid URL");
      });

      it("error path: no item url", async () => {
        const question = oneDriveSharePointItemQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        sandbox.stub(stringUtils, "isValidHttpUrl").returns(true);

        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("", inputs);
        assert.isUndefined(res);
      });

      it("error path: graph client result error", async () => {
        const question = oneDriveSharePointItemQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        sandbox
          .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
          .resolves(err(new UserError("fakeError", "fakeError", "fakeError", "fakeError")));

        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("http://fakeUrl.com", inputs);
        assert.isNotNull(res);
      });

      it("error path: no inputs", async () => {
        const question = oneDriveSharePointItemQuestion();

        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        try {
          await validationSchema.validFunc?.("http://fakeUrl.com", undefined);
          assert.fail("Should throw error");
        } catch (err) {
          assert.isNotNull(err);
        }
      });

      it("error path: axios error", async () => {
        const question = oneDriveSharePointItemQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        const config: InternalAxiosRequestConfig = {
          url: "/test",
          method: "get",
          headers: new axios.AxiosHeaders({
            "Content-Type": "application/json",
          }),
          baseURL: "https://test.com",
          timeout: 1000,
        };

        const request = {};
        const response: AxiosResponse = {
          data: { message: "Fake error" },
          status: 500,
          statusText: "Fake error",
          headers: {},
          config: config,
          request: request,
        };
        sandbox
          .stub(oneDriveSharePointHandler, "createGraphClientWithToken")
          .throws(new AxiosError("fake error", "FAKE_ERROR", config, request, response));
        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("https://test.com", inputs);
        assert.isNotNull(res);
      });

      it("error path: non-axios error", async () => {
        const question = oneDriveSharePointItemQuestion();
        const inputs: Inputs = {
          platform: Platform.VSCode,
        };
        sandbox
          .stub(oneDriveSharePointHandler, "createGraphClientWithToken")
          .throws(new UserError("test", "test", "test", "test"));
        const validationSchema = question.additionalValidationOnAccept as FuncValidation<string>;
        const res = await validationSchema.validFunc?.("https://test.com", inputs);
        assert.isNotNull(res);
      });
    });

    describe("Graph Connectors", () => {
      it("happy path", async () => {
        const fakeAxiosInstance = axios.create();
        sandbox.stub(axios, "create").returns(fakeAxiosInstance);
        const axiosGetStub = sandbox.stub(fakeAxiosInstance, "get");
        axiosGetStub.onCall(0).resolves({
          status: 200,
          data: {
            value: [
              {
                id: "fakeId",
                name: "fakeName",
              },
            ],
          },
        });
        const res = await generatorHelper.getGraphConnectors();
        assert.equal(res[0].id, "fakeId");
        assert.equal(res[0].label, "fakeId");
      });

      it("getAccessToken error", async () => {
        sandbox.stub(utils, "createContext").returns({
          tokenProvider: {
            m365TokenProvider: {
              getAccessToken: async () => {
                return Promise.resolve(err(new Error("fakeError")));
              },
            },
          } as unknown as TokenProvider,
        } as Context);
        try {
          await generatorHelper.getGraphConnectors();
          assert.fail("Should throw error");
        } catch (error) {
          assert.isNotNull(error);
        }
      });

      it("api error", async () => {
        const fakeAxiosInstance = axios.create();
        sandbox.stub(axios, "create").returns(fakeAxiosInstance);
        const axiosGetStub = sandbox.stub(fakeAxiosInstance, "get");
        axiosGetStub.onCall(0).rejects({
          status: 404,
          error: "fakeError",
        });
        axiosGetStub.onCall(1).rejects(new Error("fakeError"));
        try {
          await generatorHelper.getGraphConnectors();
          assert.fail("Should throw error");
        } catch (error) {
          assert.isNotNull(error);
        }

        try {
          await generatorHelper.getGraphConnectors();
          assert.fail("Should throw error");
        } catch (error) {
          assert.isNotNull(error);
        }
      });

      it("api 403 error", async () => {
        const fakeAxiosInstance = axios.create();
        sandbox.stub(axios, "create").returns(fakeAxiosInstance);
        const axiosGetStub = sandbox.stub(fakeAxiosInstance, "get");
        axiosGetStub.onCall(0).rejects({
          response: {
            status: 403,
            error: "fakeError",
          },
        });
        try {
          await generatorHelper.getGraphConnectors();
          assert.fail("Should throw error");
        } catch (error) {
          assert.isNotNull(error);
        }
      });
    });
  });

  describe("apiOperationQuestion", () => {
    it("includeExistingAPIs = false", async () => {
      const question = apiOperationQuestion(false);
      if (question.placeholder) {
        const placeholder =
          typeof question.placeholder === "function"
            ? question.placeholder({} as any)
            : question.placeholder;
        assert.equal(
          placeholder,
          getLocalizedString(
            "core.createProjectQuestion.apiSpec.operation.placeholder.skipExisting"
          )
        );
      }
    });
  });
  describe("apiPluginStartQuestion", () => {
    it("Capability === DACapabilityOptions.declarativeAgent().id", async () => {
      const question = apiPluginStartQuestion(false);
      const title =
        typeof question.title === "function"
          ? question.title({
              [QuestionNames.Capabilities]: DACapabilityOptions.declarativeAgent().id,
            } as any)
          : question.title;
      assert.equal(title, getLocalizedString("core.createProjectQuestion.addApiPlugin.title"));
      const placeholder =
        typeof question.placeholder === "function"
          ? question.placeholder({
              [QuestionNames.Capabilities]: DACapabilityOptions.declarativeAgent().id,
            } as any)
          : question.placeholder;
      assert.equal(
        placeholder,
        getLocalizedString("core.createProjectQuestion.addApiPlugin.placeholder")
      );
    });
    it("doesProjectExists = true", async () => {
      const question = apiPluginStartQuestion(true);
      const title =
        typeof question.title === "function" ? question.title({} as any) : question.title;
      assert.equal(title, getLocalizedString("core.createProjectQuestion.addApiPlugin.title"));
    });
    it("doesProjectExists = false", async () => {
      const question = apiPluginStartQuestion(false);
      const title =
        typeof question.title === "function" ? question.title({} as any) : question.title;
      assert.equal(title, getLocalizedString("core.createProjectQuestion.createApiPlugin.title"));
    });
  });

  describe("GCNameQuestion", () => {
    it("happy", async () => {
      const question = GCNameQuestion();
      if ((question.additionalValidationOnAccept as any).validFunc) {
        const res = (question.additionalValidationOnAccept as any).validFunc("test", {} as any);
        assert.isUndefined(res);
      }
      if ((question.validation as any).validFunc) {
        const res = (question.validation as any).validFunc("test", {} as any);
        assert.isUndefined(res);
      }
    });
  });
});
