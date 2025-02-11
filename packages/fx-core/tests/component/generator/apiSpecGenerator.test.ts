// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yuqzho@microsoft.com
 */

import {
  AdaptiveCardGenerator,
  ErrorResult,
  ErrorType,
  ProjectType,
  SpecParser,
  SpecParserError,
  Utils,
  ValidationStatus,
  WarningType,
} from "@microsoft/m365-spec-parser";
import {
  ApiOperation,
  IComposeExtension,
  Inputs,
  Platform,
  ResponseTemplatesFolderName,
  SystemError,
  TeamsAppManifest,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import { assert, expect } from "chai";
import fs from "fs-extra";
import "mocha";
import { OpenAPIV3 } from "openapi-types";
import path from "path";
import * as sinon from "sinon";
import { format } from "util";
import { createContext, setTools } from "../../../src/common/globalVars";
import { getLocalizedString } from "../../../src/common/localizeUtils";
import { manifestUtils } from "../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { PluginManifestUtils } from "../../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import { SpecGenerator } from "../../../src/component/generator/apiSpec/generator";
import * as CopilotPluginHelper from "../../../src/component/generator/apiSpec/helper";
import {
  formatValidationErrors,
  generateScaffoldingSummary,
  listPluginExistingOperations,
  injectAuthAction,
} from "../../../src/component/generator/apiSpec/helper";
import {
  ApiPluginStartOptions,
  CapabilityOptions,
  CustomCopilotRagOptions,
  DeclarativeCopilotTypeOptions,
  MeArchitectureOptions,
  ProgrammingLanguage,
  QuestionNames,
  apiPluginApiSpecOptionId,
} from "../../../src/question";
import { MockTools } from "../../core/utils";
import { copilotGptManifestUtils } from "../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import * as pluginGeneratorHelper from "../../../src/component/generator/apiSpec/helper";
import mockedEnv, { RestoreFn } from "mocked-env";
import { FeatureFlagName } from "../../../src/common/featureFlags";
import * as commonUtils from "../../../src/common/utils";
import * as helper from "../../../src/component/generator/apiSpec/helper";
import { fail } from "assert";
import { ActionInjector } from "../../../src/component/configManager/actionInjector";

const teamsManifest: TeamsAppManifest = {
  name: {
    short: "short name",
    full: "full name",
  },
  description: {
    short: "short description",
    full: "full description",
  },
  developer: {
    name: "developer name",
    websiteUrl: "https://dev.com",
    privacyUrl: "https://dev.com/privacy",
    termsOfUseUrl: "https://dev.com/termsofuse",
  },
  manifestVersion: "1.0.0",
  id: "1",
  version: "1.0.0",
  icons: {
    outline: "outline.png",
    color: "color.png",
  },
  accentColor: "#FFFFFF",
};

describe("generateScaffoldingSummary", async () => {
  const sandbox = sinon.createSandbox();

  afterEach(async () => {
    sandbox.restore();
  });
  it("no warnings", async () => {
    sandbox.stub(fs, "existsSync").returns(true);
    const composeExtension: IComposeExtension = {
      composeExtensionType: "apiBased",
      commands: [
        { id: "command1", type: "query", apiResponseRenderingTemplateFile: "test", title: "" },
        { id: "command1", type: "action", title: "" },
      ],
    };
    const res = await generateScaffoldingSummary(
      [],
      {
        ...teamsManifest,
        composeExtensions: [composeExtension],
      },
      "path",
      undefined,
      ""
    );
    assert.equal(res.length, 0);
  });

  it("warnings about missing property", async () => {
    const res = await generateScaffoldingSummary(
      [],
      {
        ...teamsManifest,
        name: { short: "", full: "" },
        description: { short: "", full: "" },
      },
      "path",
      undefined,
      ""
    );

    assert.isTrue(
      res.includes(
        getLocalizedString(
          "core.copilotPlugin.scaffold.summary.warning.teamsManifest.missingFullDescription"
        )
      )
    );
  });

  it("warnings if exceeding length", async () => {
    const invalidShortName = "a".repeat(65);
    const invalidFullName = "a".repeat(101);
    const invalidShortDescription = "a".repeat(101);
    const invalidFullDescription = "a".repeat(4001);
    const res = await generateScaffoldingSummary(
      [],
      {
        ...teamsManifest,
        name: { short: invalidShortName, full: invalidFullName },
        description: { short: invalidShortDescription, full: invalidFullDescription },
      },
      "path",
      undefined,
      ""
    );
    assert.isTrue(res.includes("name/short"));
  });

  it("no warnings if exceeding length with placeholder in short name", async () => {
    const shortName = "testdebug09051${{APP_NAME_SUFFIX}}";
    const res = await generateScaffoldingSummary(
      [],
      {
        ...teamsManifest,
        name: { short: shortName, full: "full" },
        description: { short: "short", full: "full" },
      },
      "path",
      undefined,
      ""
    );
    assert.equal(res.length, 0);
  });

  it("warnings about API spec", async () => {
    const res = await generateScaffoldingSummary(
      [{ type: WarningType.OperationIdMissing, content: "content" }],
      teamsManifest,
      "path",
      undefined,
      ""
    );

    assert.isTrue(res.includes("content"));
  });

  it("warnings about operationid contains special characters", async () => {
    const res = await generateScaffoldingSummary(
      [
        {
          type: WarningType.OperationIdContainsSpecialCharacters,
          content:
            "Operation id 'user/repo' contained special characters and was renamed to 'user_repo'.",
          data: { operationId: "user/repo" },
        },
        {
          type: WarningType.OperationIdContainsSpecialCharacters,
          content:
            "Operation id 'user/issue' contained special characters and was renamed to 'user_issue'.",
          data: { operationId: "user/issue" },
        },
      ],
      teamsManifest,
      "path",
      undefined,
      ""
    );
    assert.isTrue(res.includes("user_repo"));
    assert.isTrue(res.includes("user_issue"));
  });

  it("warnings about adaptive card template in manifest", async () => {
    const composeExtension: IComposeExtension = {
      composeExtensionType: "apiBased",
      commands: [{ id: "command1", type: "query", title: "" }],
    };
    const res = await generateScaffoldingSummary(
      [],
      {
        ...teamsManifest,
        composeExtensions: [composeExtension],
      },
      "path",
      undefined,
      ""
    );

    assert.isTrue(res.includes("apiResponseRenderingTemplateFile"));
  });

  it("warnings about missing adaptive card template", async () => {
    const composeExtension: IComposeExtension = {
      composeExtensionType: "apiBased",
      commands: [
        { id: "command1", type: "query", apiResponseRenderingTemplateFile: "", title: "" },
      ],
    };
    sandbox.stub(fs, "existsSync").returns(false);
    const res = await generateScaffoldingSummary(
      [{ type: WarningType.GenerateCardFailed, content: "test", data: "command1" }],
      {
        ...teamsManifest,
        composeExtensions: [composeExtension],
      },
      "path",
      undefined,
      ""
    );

    assert.isTrue(res.includes("apiResponseRenderingTemplateFile"));
    assert.isTrue(res.includes("test"));
  });

  it("warnings about command parameters", async () => {
    const composeExtension: IComposeExtension = {
      composeExtensionType: "apiBased",
      apiSpecificationFile: "testApiFile",
      commands: [
        {
          id: "getAll",
          type: "query",
          title: "",
          apiResponseRenderingTemplateFile: "apiResponseRenderingTemplateFile",
          parameters: [
            {
              name: "test",
              title: "test",
            },
          ],
        },
      ],
    };
    const res = await generateScaffoldingSummary(
      [{ type: WarningType.OperationOnlyContainsOptionalParam, content: "", data: "getAll" }],
      {
        ...teamsManifest,
        composeExtensions: [composeExtension],
      },
      "path",
      undefined,
      ""
    );

    assert.isTrue(res.includes("testApiFile"));
  });

  it("warnings about command parameters with some properties missing", async () => {
    const composeExtension: IComposeExtension = {
      composeExtensionType: "apiBased",
      commands: [
        {
          id: "getAll",
          type: "query",
          title: "",
          apiResponseRenderingTemplateFile: "apiResponseRenderingTemplateFile",
          parameters: [],
        },
      ],
    };
    const res = await generateScaffoldingSummary(
      [{ type: WarningType.OperationOnlyContainsOptionalParam, content: "", data: "getAll" }],
      {
        ...teamsManifest,
        composeExtensions: [composeExtension],
      },
      "path",
      undefined,
      ""
    );

    assert.isFalse(res.includes("testApiFile"));
  });

  it("warnings about plugin manifest description", async () => {
    sandbox.stub(PluginManifestUtils.prototype, "readPluginManifestFile").resolves(
      ok({
        functions: [
          { name: "getAll", description: "test" },
          { name: "createNew", description: "" },
        ],
      } as any)
    );
    const res = await generateScaffoldingSummary(
      [{ type: WarningType.FuncDescriptionTooLong, content: "", data: "getAll" }],
      {
        ...teamsManifest,
        copilotExtensions: { plugins: [{ file: "test", id: "1" }] },
      },
      "path",
      "pluginPath",
      ""
    );
    assert.isTrue(res.includes("getAll"));
    assert.isTrue(res.includes("createNew"));
  });

  it("warnings about plugin manifest description: get plugin file error", async () => {
    sandbox
      .stub(PluginManifestUtils.prototype, "readPluginManifestFile")
      .resolves(err(new SystemError("test", "test", "test", "test")));
    const res = await generateScaffoldingSummary(
      [{ type: WarningType.FuncDescriptionTooLong, content: "", data: "getAll" }],
      {
        ...teamsManifest,
        copilotExtensions: { plugins: [{ file: "test", id: "1" }] },
      },
      "path",
      "pluginPath",
      ""
    );

    assert.equal(res.length, 0);
  });
});

describe("isJsonSpecFile", () => {
  afterEach(() => {
    sinon.restore();
  });
  it("should return true for a valid JSON file", async () => {
    const result = await commonUtils.isJsonSpecFile("test.json");
    expect(result).to.be.true;
  });

  it("should return false for an yaml file", async () => {
    const result = await commonUtils.isJsonSpecFile("test.yaml");
    expect(result).to.be.false;
  });

  it("should handle local json files", async () => {
    const readFileStub = sinon.stub(fs, "readFile").resolves('{"name": "test"}' as any);
    const result = await commonUtils.isJsonSpecFile("path/to/localfile");
    expect(result).to.be.true;
  });

  it("should handle remote files", async () => {
    const axiosStub = sinon.stub(axios, "get").resolves({ data: '{"name": "test"}' });
    const result = await commonUtils.isJsonSpecFile("http://example.com/remotefile");
    expect(result).to.be.true;
  });

  it("should return false if it is a yaml file", async () => {
    const readFileStub = sinon.stub(fs, "readFile").resolves("openapi: 3.0.0" as any);
    const result = await commonUtils.isJsonSpecFile("path/to/localfile");
    expect(result).to.be.false;
  });
});

describe("formatValidationErrors", () => {
  it("format validation errors from spec parser", () => {
    const errors: ErrorResult[] = [
      {
        type: ErrorType.SpecNotValid,
        content: "test",
      },
      {
        type: ErrorType.SpecNotValid,
        content: "ResolverError: Error downloading",
      },
      {
        type: ErrorType.RemoteRefNotSupported,
        content: "test",
      },
      {
        type: ErrorType.NoServerInformation,
        content: "test",
      },
      {
        type: ErrorType.UrlProtocolNotSupported,
        content: "protocol",
        data: "http",
      },
      {
        type: ErrorType.RelativeServerUrlNotSupported,
        content: "test",
      },
      {
        type: ErrorType.NoSupportedApi,
        content: "test",
        data: [],
      },
      {
        type: ErrorType.NoSupportedApi,
        content: "test",
        data: [
          {
            api: "GET /api",
            reason: [
              ErrorType.AuthTypeIsNotSupported,
              ErrorType.MissingOperationId,
              ErrorType.PostBodyContainMultipleMediaTypes,
              ErrorType.ResponseContainMultipleMediaTypes,
              ErrorType.ResponseJsonIsEmpty,
              ErrorType.MethodNotAllowed,
              ErrorType.UrlPathNotExist,
            ],
          },
          {
            api: "GET /api2",
            reason: [
              ErrorType.PostBodyContainsRequiredUnsupportedSchema,
              ErrorType.ParamsContainRequiredUnsupportedSchema,
              ErrorType.ExceededRequiredParamsLimit,
              ErrorType.NoParameter,
              ErrorType.NoAPIInfo,
              ErrorType.CircularReferenceNotSupported,
            ],
          },
          { api: "GET /api3", reason: ["unknown"] },
        ],
      },
      {
        type: ErrorType.NoExtraAPICanBeAdded,
        content: "test",
      },
      {
        type: ErrorType.ResolveServerUrlFailed,
        content: "resolveurl",
      },
      {
        type: ErrorType.Cancelled,
        content: "test",
      },
      {
        type: ErrorType.SwaggerNotSupported,
        content: "test",
      },
      {
        type: ErrorType.SpecVersionNotSupported,
        content: "test",
        data: "3.1.0",
      },
      {
        type: ErrorType.Unknown,
        content: "unknown",
      },
      {
        type: ErrorType.AddedAPINotInOriginalSpec,
        content: "test",
      },
    ];

    const res = formatValidationErrors(errors, {
      platform: Platform.VSCode,
      [QuestionNames.ManifestPath]: "testmanifest.json",
    });

    expect(res[0].content).equals("test");
    expect(res[1].content).includes(getLocalizedString("core.common.ErrorFetchApiSpec"));
    expect(res[2].content).equals("test");
    expect(res[3].content).equals(getLocalizedString("core.common.NoServerInformation"));
    expect(res[4].content).equals(
      getLocalizedString("core.common.UrlProtocolNotSupported", "http")
    );
    expect(res[5].content).equals(getLocalizedString("core.common.RelativeServerUrlNotSupported"));
    expect(res[6].content).equals(
      getLocalizedString(
        "core.common.NoSupportedApi",
        getLocalizedString("core.common.invalidReason.NoAPIs")
      )
    );

    const errorMessage1 = [
      getLocalizedString("core.common.invalidReason.AuthTypeIsNotSupported"),
      getLocalizedString("core.common.invalidReason.MissingOperationId"),
      getLocalizedString("core.common.invalidReason.PostBodyContainMultipleMediaTypes"),
      getLocalizedString("core.common.invalidReason.ResponseContainMultipleMediaTypes"),
      getLocalizedString("core.common.invalidReason.ResponseJsonIsEmpty"),
      getLocalizedString("core.common.invalidReason.MethodNotAllowed"),
      getLocalizedString("core.common.invalidReason.UrlPathNotExist"),
    ];
    const errorMessage2 = [
      getLocalizedString("core.common.invalidReason.PostBodyContainsRequiredUnsupportedSchema"),
      getLocalizedString("core.common.invalidReason.ParamsContainRequiredUnsupportedSchema"),
      getLocalizedString("core.common.invalidReason.ExceededRequiredParamsLimit"),
      getLocalizedString("core.common.invalidReason.NoParameter"),
      getLocalizedString("core.common.invalidReason.NoAPIInfo"),
      getLocalizedString("core.common.invalidReason.CircularReference"),
    ];

    expect(res[7].content).equals(
      getLocalizedString(
        "core.common.NoSupportedApi",
        "GET /api: " +
          errorMessage1.join(", ") +
          "\n" +
          "GET /api2: " +
          errorMessage2.join(", ") +
          "\n" +
          "GET /api3: unknown"
      )
    );
    expect(res[8].content).equals(getLocalizedString("error.apime.noExtraAPICanBeAdded"));
    expect(res[9].content).equals("resolveurl");
    expect(res[10].content).equals(getLocalizedString("core.common.CancelledMessage"));
    expect(res[11].content).equals(getLocalizedString("core.common.SwaggerNotSupported"));
    expect(res[12].content).equals(
      format(getLocalizedString("core.common.SpecVersionNotSupported"), res[12].data)
    );
    expect(res[13].content).equals("unknown");
    expect(res[14].content).equals(getLocalizedString("core.common.AddedAPINotInOriginalSpec"));
  });

  it("format validation errors from spec parser: copilot", () => {
    const errors: ErrorResult[] = [
      {
        type: ErrorType.NoSupportedApi,
        content: "test",
        data: [
          {
            api: "GET /api",
            reason: [
              ErrorType.AuthTypeIsNotSupported,
              ErrorType.MissingOperationId,
              ErrorType.PostBodyContainMultipleMediaTypes,
              ErrorType.ResponseContainMultipleMediaTypes,
              ErrorType.ResponseJsonIsEmpty,
              ErrorType.MethodNotAllowed,
              ErrorType.UrlPathNotExist,
            ],
          },
          {
            api: "GET /api2",
            reason: [
              ErrorType.PostBodyContainsRequiredUnsupportedSchema,
              ErrorType.ParamsContainRequiredUnsupportedSchema,
              ErrorType.ExceededRequiredParamsLimit,
              ErrorType.NoParameter,
              ErrorType.NoAPIInfo,
            ],
          },
          { api: "GET /api3", reason: ["unknown"] },
        ],
      },
      {
        type: ErrorType.NoExtraAPICanBeAdded,
        content: "test",
      },
    ];

    const res = formatValidationErrors(errors, {
      platform: Platform.VSCode,
      [QuestionNames.ApiPluginType]: apiPluginApiSpecOptionId,
    });

    const errorMessage1 = [
      getLocalizedString("core.common.invalidReason.AuthTypeIsNotSupported"),
      getLocalizedString("core.common.invalidReason.MissingOperationId"),
      getLocalizedString("core.common.invalidReason.PostBodyContainMultipleMediaTypes"),
      getLocalizedString("core.common.invalidReason.ResponseContainMultipleMediaTypes"),
      getLocalizedString("core.common.invalidReason.ResponseJsonIsEmpty"),
      getLocalizedString("core.common.invalidReason.MethodNotAllowed"),
      getLocalizedString("core.common.invalidReason.UrlPathNotExist"),
    ];
    const errorMessage2 = [
      getLocalizedString("core.common.invalidReason.PostBodyContainsRequiredUnsupportedSchema"),
      getLocalizedString("core.common.invalidReason.ParamsContainRequiredUnsupportedSchema"),
      getLocalizedString("core.common.invalidReason.ExceededRequiredParamsLimit"),
      getLocalizedString("core.common.invalidReason.NoParameter"),
      getLocalizedString("core.common.invalidReason.NoAPIInfo"),
    ];

    expect(res[0].content).equals(
      getLocalizedString(
        "core.common.NoSupportedApiCopilot",
        "GET /api: " +
          errorMessage1.join(", ") +
          "\n" +
          "GET /api2: " +
          errorMessage2.join(", ") +
          "\n" +
          "GET /api3: unknown"
      )
    );
    expect(res[1].content).equals(getLocalizedString("error.copilot.noExtraAPICanBeAdded"));
  });
});

describe("injectAuthAction", async () => {
  const sandbox = sinon.createSandbox();

  afterEach(async () => {
    sandbox.restore();
  });

  it("api key auth", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(Utils, "isBearerTokenAuth").returns(true);
    const injectStub = sandbox.stub(ActionInjector, "injectCreateAPIKeyAction").resolves(undefined);
    const res = await injectAuthAction(
      "oauth",
      "test",
      { scheme: "", type: "http" },
      "test",
      false
    );

    assert.isUndefined(res);
    assert.isTrue(injectStub.calledTwice);
  });

  it("api key auth: no local yaml", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(Utils, "isBearerTokenAuth").returns(true);
    const injectStub = sandbox.stub(ActionInjector, "injectCreateAPIKeyAction").resolves(undefined);
    const res = await injectAuthAction(
      "oauth",
      "test",
      { scheme: "", type: "http" },
      "test",
      false
    );

    assert.isUndefined(res);
    assert.isTrue(injectStub.calledOnce);
  });

  it("oauth auth", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(Utils, "isOAuthWithAuthCodeFlow").returns(true);
    const injectStub = sandbox.stub(ActionInjector, "injectCreateOAuthAction").resolves(undefined);
    const res = await injectAuthAction(
      "oauth",
      "test",
      { scheme: "", type: "http" },
      "test",
      false
    );

    assert.isUndefined(res);
    assert.isTrue(injectStub.calledTwice);
  });

  it("oauth auth: no local yaml", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(Utils, "isOAuthWithAuthCodeFlow").returns(true);
    const injectStub = sandbox.stub(ActionInjector, "injectCreateOAuthAction").resolves(undefined);
    const res = await injectAuthAction(
      "oauth",
      "test",
      { scheme: "", type: "http" },
      "test",
      false
    );

    assert.isUndefined(res);
    assert.isTrue(injectStub.calledOnce);
  });

  it("api key auth from authType", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    // sandbox.stub(Utils, "isBearerTokenAuth").returns(true);
    const injectStub = sandbox.stub(ActionInjector, "injectCreateAPIKeyAction").resolves(undefined);
    const res = await injectAuthAction(
      "oauth",
      "test",
      undefined,
      "test",
      false,
      "ApiKeyPluginVault"
    );

    assert.isUndefined(res);
    assert.isTrue(injectStub.calledTwice);
  });

  it("oauth auth from authType", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    const injectStub = sandbox.stub(ActionInjector, "injectCreateOAuthAction").resolves(undefined);
    const res = await injectAuthAction(
      "oauth",
      "test",
      undefined,
      "test",
      false,
      "OAuthPluginVault"
    );

    assert.isUndefined(res);
    assert.isTrue(injectStub.calledTwice);
  });
});

describe("listPluginExistingOperations", () => {
  const teamsManifestWithPlugin: TeamsAppManifest = {
    ...teamsManifest,
    copilotExtensions: {
      plugins: [
        {
          file: "resources/plugin.json",
          id: "plugin1",
        },
      ],
    },
  };

  const sandbox = sinon.createSandbox();
  afterEach(async () => {
    sandbox.restore();
  });

  it("success", async () => {
    sandbox
      .stub(PluginManifestUtils.prototype, "getApiSpecFilePathFromTeamsManifest")
      .resolves(ok(["openapi.yaml"]));

    sandbox
      .stub(SpecParser.prototype, "validate")
      .resolves({ status: ValidationStatus.Valid, warnings: [], errors: [] });
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api1",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    const res = await listPluginExistingOperations(
      teamsManifestWithPlugin,
      "manifestPath",
      "openapi.yaml"
    );
    expect(res).to.be.deep.equal(["api1"]);
  });

  it("get api spec error", async () => {
    sandbox
      .stub(PluginManifestUtils.prototype, "getApiSpecFilePathFromTeamsManifest")
      .resolves(err(new SystemError("getApiSpecFilePathFromTeamsManifest", "name", "", "")));

    let hasException = false;

    try {
      await listPluginExistingOperations(teamsManifestWithPlugin, "manifestPath", "openapi.yaml");
    } catch (e) {
      hasException = true;
      expect(e.source).equal("getApiSpecFilePathFromTeamsManifest");
    }
    expect(hasException).to.be.true;
  });

  it("openapi is not referenced for plugin", async () => {
    sandbox
      .stub(PluginManifestUtils.prototype, "getApiSpecFilePathFromTeamsManifest")
      .resolves(ok(["openapi.yaml"]));
    let hasException = false;

    try {
      await listPluginExistingOperations(teamsManifestWithPlugin, "manifestPath", "notexist.yaml");
    } catch (e) {
      hasException = true;
      expect(e.source).equal("listPluginExistingOperations");
      expect(e.name).equal("api-spec-not-used-in-plugin");
    }
    expect(hasException).to.be.true;
  });
});

describe("updateForCustomApi", async () => {
  const sandbox = sinon.createSandbox();
  const spec = {
    openapi: "3.0.0",
    info: {
      title: "My API",
      version: "1.0.0",
    },
    description: "test",
    paths: {
      "/hello": {
        get: {
          operationId: "getHello",
          summary: "Returns a greeting",
          parameters: [
            {
              name: "query",
              in: "query",
              schema: { type: "string" },
            },
          ],
          responses: {
            "200": {
              description: "A greeting message",
              content: {
                "application/json": {
                  schema: {
                    type: "string",
                  },
                },
              },
            },
          },
        },
        post: {
          operationId: "createPet",
          summary: "Create a pet",
          description: "Create a new pet in the store",
          requestBody: {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  required: ["name"],
                  properties: {
                    name: {
                      type: "string",
                      description: "Name of the pet",
                    },
                  },
                },
              },
            },
          },
        },
      },
    },
  } as OpenAPIV3.Document;

  const manifest: TeamsAppManifest = {
    manifestVersion: "version",
    id: "mock-app-id",
    name: { short: "short-name" },
    description: { short: "", full: "" },
    version: "version",
    icons: { outline: "outline.png", color: "color.png" },
    accentColor: "#ffffff",
    developer: {
      privacyUrl: "",
      websiteUrl: "",
      termsOfUseUrl: "",
      name: "developer-name",
    },
    bots: [
      {
        botId: "${{BOT_ID}}",
        scopes: ["personal", "team", "groupChat"],
        supportsFiles: false,
        isNotificationOnly: false,
      },
    ],
    validDomains: ["valid-domain"],
  };

  afterEach(async () => {
    sandbox.restore();
  });

  it("happy path: ts", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));

    sandbox
      .stub(manifestUtils, "_writeAppManifest")
      .callsFake(async (updatedManifest, manifestPath) => {
        expect(manifestPath.replace(/\\/g, "/")).to.be.equal("path/appPackage/manifest.json");
        expect(updatedManifest.bots![0].commandLists![0].commands[0].title).to.be.equal(
          "Hello, how can you help me?"
        );
        expect(updatedManifest.bots![0].commandLists![0].commands[1].title).to.be.equal(
          "How to build apps with TTK?"
        );
        expect(updatedManifest.bots![0].commandLists![0].commands[2].title).to.be.equal(
          "Returns a greeting"
        );
        expect(updatedManifest.bots![0].commandLists![0].commands[3].title).to.be.equal(
          "Create a pet"
        );
        return ok(undefined);
      });
    await CopilotPluginHelper.updateForCustomApi(spec, "typescript", "path", "openapi.yaml");
  });

  it("read manifest failed", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox
      .stub(manifestUtils, "_readAppManifest")
      .resolves(err(new SystemError("test", "", "", "")));
    try {
      await CopilotPluginHelper.updateForCustomApi(spec, "typescript", "path", "openapi.yaml");
      assert.fail("should throw error");
    } catch (e) {
      expect(e.source).to.be.equal("test");
    }
  });

  it("happy path: should contain warning if generate adaptive card failed", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        assert.fail("should not generate adaptive card");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox
      .stub(AdaptiveCardGenerator, "generateAdaptiveCard")
      .throws(new Error("generate adaptive card failed"));

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));

    const result = await CopilotPluginHelper.updateForCustomApi(
      spec,
      "typescript",
      "path",
      "openapi.yaml"
    );

    expect(result).to.be.deep.equal([
      {
        type: WarningType.GenerateCardFailed,
        content:
          "Failed to create the adaptive card for API 'getHello': generate adaptive card failed. Mitigation: Not required but you can manually add it to the adaptiveCards folder.",
        data: "getHello",
      },
      {
        type: WarningType.GenerateCardFailed,
        content:
          "Failed to create the adaptive card for API 'createPet': generate adaptive card failed. Mitigation: Not required but you can manually add it to the adaptiveCards folder.",
        data: "createPet",
      },
    ]);
  });

  it("happy path: js", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    await CopilotPluginHelper.updateForCustomApi(spec, "javascript", "path", "openapi.yaml");
  });

  it("happy path: should contain warning if generate adaptive card data failed", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        assert.fail("should not generate adaptive card");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      } else if (file == path.join("path", "src", "adaptiveCard", "hello.data.json")) {
        expect(data).to.deep.equal({});
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox.stub(AdaptiveCardGenerator, "generateAdaptiveCard").returns([
      {
        type: "AdaptiveCard",
        $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.5",
        body: [
          {
            type: "TextBlock",
            text: "name: ${if(name, name, 'N/A')}",
            wrap: true,
          },
        ],
      },
      "$",
      {},
      [{ type: WarningType.GenerateJsonDataFailed, content: "generate json data failed" }],
    ]);

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));

    const result = await CopilotPluginHelper.updateForCustomApi(
      spec,
      "typescript",
      "path",
      "openapi.yaml"
    );

    expect(result).to.be.deep.equal([
      {
        type: WarningType.GenerateJsonDataFailed,
        content:
          "Failed to create the adaptive card mock data for API 'getHello': generate json data failed. Mitigation: Not required but you can manually add it to the adaptiveCards folder.",
        data: "getHello",
      },
      {
        type: WarningType.GenerateJsonDataFailed,
        content:
          "Failed to create the adaptive card mock data for API 'createPet': generate json data failed. Mitigation: Not required but you can manually add it to the adaptiveCards folder.",
        data: "createPet",
      },
    ]);
  });

  it("happy path: js", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    await CopilotPluginHelper.updateForCustomApi(spec, "javascript", "path", "openapi.yaml");
  });

  it("happy path: python", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file == path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file == path.join("path", "src", "bot.py")) {
        expect(data).to.contains(`@bot_app.ai.action("getHello")`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("# Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code # Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    await CopilotPluginHelper.updateForCustomApi(spec, "python", "path", "openapi.yaml");
  });

  it("happy path: csharp", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file == path.join("path", "APIActions.cs")) {
        expect(data).to.contains(`[Action("getHello")]`);
        expect(data).to.contains(`public async Task<string> GetHelloAsync`);
        expect(data).to.contains("openapi.yaml");
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("# Replace with action code");
      }

      if (file.toString().endsWith("actions.json")) {
        expect(file == path.join("path", "prompts", "Chat", "actions.json")).to.be.true;
      }

      if (file.toString().endsWith("skprompt.txt")) {
        expect(file == path.join("path", "prompts", "Chat", "skprompt.txt")).to.be.true;
      }

      if (file.toString().endsWith("getHello.json")) {
        expect(file == path.join("path", "adaptiveCards", "getHello.json")).to.be.true;
      }
    });

    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    //sandbox fs.readdir(destinationPath)
    sandbox.stub(fs, "readdir").resolves(["MyApp.csproj"] as any);
    await CopilotPluginHelper.updateForCustomApi(spec, "csharp", "path", "openapi.yaml");
  });

  it("unknown language: unknown", async () => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file == path.join("path", "APIActions.cs")) {
        fail("actions.json should not be created for unknown language");
      }

      if (file.toString().endsWith("actions.json")) {
        fail("actions.json should not be created for unknown language");
      }

      if (file.toString().endsWith("skprompt.txt")) {
        fail("actions.json should not be created for unknown language");
      }

      if (file.toString().endsWith("getHello.json")) {
        fail("actions.json should not be created for unknown language");
      }
    });

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    await CopilotPluginHelper.updateForCustomApi(spec, "unknown", "path", "openapi.yaml");
  });

  it("happy path with spec without path", async () => {
    const limitedSpec = {
      openapi: "3.0.0",
      info: {
        title: "My API",
        version: "1.0.0",
      },
    } as OpenAPIV3.Document;
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.equals("[]");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    await CopilotPluginHelper.updateForCustomApi(limitedSpec, "javascript", "path", "openapi.yaml");
  });

  it("happy path with spec without pathItem", async () => {
    const limitedSpec = {
      openapi: "3.0.0",
      info: {
        title: "My API",
        version: "1.0.0",
      },
      paths: {},
    } as OpenAPIV3.Document;
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.equals("[]");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    await CopilotPluginHelper.updateForCustomApi(limitedSpec, "javascript", "path", "openapi.yaml");
  });

  it("happy path with spec with patch", async () => {
    const limitedSpec = {
      openapi: "3.0.0",
      info: {
        title: "My API",
        version: "1.0.0",
      },
      paths: {
        patch: {
          operationId: "createPet",
          summary: "Create a pet",
          description: "Create a new pet in the store",
          requestBody: {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  required: ["name"],
                  properties: {
                    name: {
                      type: "string",
                      description: "Name of the pet",
                    },
                  },
                },
              },
            },
          },
        },
      },
    } as OpenAPIV3.Document;
    sandbox.stub(fs, "ensureDir").resolves();
    const mockWriteFile = sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.equals("[]");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.equals("[]");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    await CopilotPluginHelper.updateForCustomApi(limitedSpec, "javascript", "path", "openapi.yaml");
    expect(mockWriteFile.calledThrice).to.be.true;
  });

  it("happy path with spec with required and multiple parameter", async () => {
    const newSpec = {
      openapi: "3.0.0",
      info: {
        title: "My API",
        version: "1.0.0",
      },
      description: "test",
      paths: {
        "/hello": {
          get: {
            operationId: "getHello",
            summary: "Returns a greeting",
            parameters: [
              {
                name: "query",
                in: "query",
                schema: { type: "string" },
                required: true,
              },
              {
                name: "query2",
                in: "query",
                schema: { type: "string" },
                requried: false,
              },
              {
                name: "query3",
                in: "query",
                schema: { type: "string" },
                requried: true,
                description: "test",
              },
            ],
            responses: {
              "200": {
                description: "",
                content: {
                  "application/json": {
                    schema: {
                      type: "string",
                    },
                  },
                },
              },
            },
          },
          post: {
            operationId: "createPet",
            summary: "Create a pet",
            description: "",
            requestBody: {
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    required: ["name"],
                    properties: {
                      name: {
                        type: "string",
                        description: "",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      },
    } as OpenAPIV3.Document;
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    await CopilotPluginHelper.updateForCustomApi(newSpec, "typescript", "path", "openapi.yaml");
  });

  it("happy path with spec request body and schema contains format", async () => {
    const newSpec = {
      openapi: "3.0.0",
      info: {
        title: "My API",
        version: "1.0.0",
      },
      description: "test",
      paths: {
        "/hello": {
          get: {
            operationId: "getHello",
            summary: "Returns a greeting",
            parameters: [
              {
                name: "query",
                in: "query",
                schema: { type: "string" },
                required: true,
              },
              {
                name: "query2",
                in: "query",
                schema: { type: "string" },
                requried: false,
              },
              {
                name: "query3",
                in: "query",
                schema: { type: "string" },
                requried: true,
                description: "test",
              },
              {
                name: "query4",
                in: "query",
                schema: {
                  type: "array",
                  items: {
                    type: "string",
                    format: "test",
                  },
                },
              },
            ],
            responses: {
              "200": {
                description: "",
                content: {
                  "application/json": {
                    schema: {
                      type: "string",
                    },
                  },
                },
              },
            },
          },
          post: {
            operationId: "createPet",
            summary: "Create a pet",
            description: "",
            requestBody: {
              required: true,
              description: "request body description",
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    required: ["date"],
                    properties: {
                      date: {
                        type: "string",
                        description: "",
                        format: "date-time",
                      },
                      array: {
                        type: "array",
                        items: {
                          type: "string",
                          format: "test",
                        },
                      },
                      object: {
                        type: "object",
                        properties: {
                          nestedObjProperty: {
                            type: "string",
                            description: "",
                            format: "test",
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        },
      },
    } as OpenAPIV3.Document;
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
        expect(data).to.contains("body");
        expect(data).to.not.contains("format");
        expect(data).to.contains("nestedObjProperty");
        expect(data).to.contains("array");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    await CopilotPluginHelper.updateForCustomApi(newSpec, "typescript", "path", "openapi.yaml");
  });

  it("happy path with spec with auth", async () => {
    const authSpec = {
      openapi: "3.0.0",
      info: {
        title: "My API",
        version: "1.0.0",
      },
      description: "test",
      paths: {
        "/hello": {
          get: {
            operationId: "getHello",
            summary: "Returns a greeting",
            parameters: [
              {
                name: "query",
                in: "query",
                schema: { type: "string" },
              },
            ],
            responses: {
              "200": {
                description: "A greeting message",
                content: {
                  "application/json": {
                    schema: {
                      type: "string",
                    },
                  },
                },
              },
            },
            security: [
              {
                api_key: [],
              },
            ],
          },
          post: {
            operationId: "createPet",
            summary: "Create a pet",
            description: "Create a new pet in the store",
            requestBody: {
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    required: ["name"],
                    properties: {
                      name: {
                        type: "string",
                        description: "Name of the pet",
                      },
                    },
                  },
                },
              },
            },
          },
        },
      },
      components: {
        securitySchemes: {
          api_key: {
            type: "apiKey",
            name: "api_key",
            in: "header",
          },
        },
      },
    } as OpenAPIV3.Document;
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    await CopilotPluginHelper.updateForCustomApi(authSpec, "typescript", "path", "openapi.yaml");
  });

  it("happy path with spec with jsonPath", async () => {
    const specWithJsonPath = {
      openapi: "3.0.0",
      info: {
        title: "My API",
        version: "1.0.0",
      },
      description: "test",
      paths: {
        "/hello": {
          get: {
            operationId: "getHello",
            summary: "Returns a greeting",
            parameters: [
              {
                name: "query",
                in: "query",
                schema: { type: "string" },
              },
            ],
            responses: {
              "200": {
                description: "A greeting message",
                content: {
                  "application/json": {
                    schema: {
                      type: "object",
                      properties: {
                        results: {
                          type: "array",
                          items: {
                            type: "object",
                            properties: {
                              id: {
                                type: "string",
                                description: "id",
                              },
                            },
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        },
      },
    } as OpenAPIV3.Document;
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").callsFake((file, data) => {
      if (file === path.join("path", "src", "prompts", "chat", "skprompt.txt")) {
        expect(data).to.contains("The following is a conversation with an AI assistant.");
      } else if (file === path.join("path", "src", "adaptiveCard", "hello.json")) {
        expect(data).to.contains("${results}");
      } else if (file === path.join("path", "src", "prompts", "chat", "actions.json")) {
        expect(data).to.contains("getHello");
      } else if (file === path.join("path", "src", "app", "app.ts")) {
        expect(data).to.contains(`app.ai.action("getHello"`);
        expect(data).not.to.contains("{{");
        expect(data).not.to.contains("// Replace with action code");
      }
    });
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox.stub(manifestUtils, "_writeAppManifest").resolves(ok(undefined));
    sandbox
      .stub(fs, "readFile")
      .resolves(Buffer.from("test code // Replace with action code {{OPENAPI_SPEC_PATH}}"));
    await CopilotPluginHelper.updateForCustomApi(
      specWithJsonPath,
      "typescript",
      "path",
      "openapi.yaml"
    );
  });
});

describe("listOperations", async () => {
  const context = createContext();
  const sandbox = sinon.createSandbox();
  const spec = {
    openapi: "3.0.0",
    info: {
      title: "My API",
      version: "1.0.0",
    },
    description: "test",
    paths: {
      "/hello": {
        get: {
          operationId: "getHello",
          summary: "Returns a greeting",
          parameters: [
            {
              name: "query",
              in: "query",
              schema: { type: "string" },
            },
          ],
          responses: {
            "200": {
              description: "A greeting message",
              content: {
                "application/json": {
                  schema: {
                    type: "string",
                  },
                },
              },
            },
          },
          security: [
            {
              api_key: [],
            },
          ],
        },
        post: {
          operationId: "createPet",
          summary: "Create a pet",
          description: "Create a new pet in the store",
          requestBody: {
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  required: ["name"],
                  properties: {
                    name: {
                      type: "string",
                      description: "Name of the pet",
                    },
                  },
                },
              },
            },
          },
        },
      },
    },
    components: {
      securitySchemes: {
        api_key: {
          type: "apiKey",
          name: "api_key",
          in: "header",
        },
      },
    },
  } as OpenAPIV3.Document;

  afterEach(async () => {
    sandbox.restore();
  });

  it("allow auth for teams ai project", async () => {
    const inputs = {
      "custom-copilot-rag": "custom-copilot-rag-customApi",
      platform: Platform.VSCode,
    };
    sandbox.stub(CopilotPluginHelper, "formatValidationErrors").resolves([]);
    sandbox.stub(CopilotPluginHelper, "logValidationResults").resolves();
    sandbox.stub(SpecParser.prototype, "validate").resolves({
      status: ValidationStatus.Valid,
      warnings: [],
      errors: [],
      specHash: "xxx",
    });
    sandbox
      .stub(SpecParser.prototype, "list")
      .resolves({ APIs: [], allAPICount: 1, validAPICount: 0 });

    const res = await CopilotPluginHelper.listOperations(context, "", inputs, true, false, "");
    expect(res.isOk()).to.be.true;
  });

  it("will show invalid api reasons", async () => {
    const inputs = {
      "custom-copilot-rag": "custom-copilot-rag-customApi",
      platform: Platform.VSCode,
    };
    sandbox.stub(CopilotPluginHelper, "formatValidationErrors").resolves([]);
    sandbox.stub(CopilotPluginHelper, "logValidationResults").resolves();
    sandbox.stub(SpecParser.prototype, "validate").resolves({
      status: ValidationStatus.Valid,
      warnings: [],
      errors: [],
      specHash: "",
    });
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "1",
          server: "https://test",
          operationId: "id1",
          isValid: false,
          reason: [ErrorType.NoParameter],
        },
        {
          api: "2",
          server: "https://test",
          operationId: "id2",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 2,
      validAPICount: 1,
    });
    const warningSpy = sandbox.spy(context.logProvider, "warning");

    const res = await CopilotPluginHelper.listOperations(context, "", inputs, true, false, "");
    expect(res.isOk()).to.be.true;
    expect(warningSpy.calledOnce).to.be.true;
  });

  it("should throw error if list api not from original OpenAPI spec", async () => {
    const inputs = {
      platform: Platform.VSCode,
      "manifest-path": "fake-path",
    };
    sandbox.stub(CopilotPluginHelper, "formatValidationErrors").resolves([]);
    sandbox.stub(CopilotPluginHelper, "logValidationResults").resolves();
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({} as any));
    sandbox.stub(manifestUtils, "getOperationIds").returns(["getHello"]);
    sandbox.stub(CopilotPluginHelper, "listPluginExistingOperations").resolves(["getHello"]);
    sandbox.stub(SpecParser.prototype, "validate").resolves({
      status: ValidationStatus.Valid,
      warnings: [],
      errors: [],
      specHash: "",
    });
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "GET /api",
          server: "https://test",
          operationId: "getApi",
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 0,
    });

    const res = await CopilotPluginHelper.listOperations(context, "", inputs, false, false, "");
    expect(res.isErr()).to.be.true;
    if (res.isErr()) {
      expect(res.error.length).to.be.equal(1);
      expect(res.error[0].type).to.be.equal(ErrorType.AddedAPINotInOriginalSpec);
    }
  });

  it("should not allow auth for VS project", async () => {
    const inputs = {
      platform: Platform.VS,
    };
    sandbox.stub(CopilotPluginHelper, "formatValidationErrors").resolves([]);
    sandbox.stub(CopilotPluginHelper, "logValidationResults").resolves();
    sandbox.stub(SpecParser.prototype, "validate").resolves({
      status: ValidationStatus.Valid,
      warnings: [],
      errors: [],
      specHash: "xxx",
    });
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "1",
          server: "https://test",
          operationId: "id1",
          isValid: false,
          reason: [ErrorType.AuthTypeIsNotSupported],
        },
      ],
      allAPICount: 1,
      validAPICount: 0,
    });

    const res = await CopilotPluginHelper.listOperations(context, "", inputs, true, false, "");
    expect(res.isOk()).to.be.true;
  });

  it("should not allow auth for VS copilot project", async () => {
    const inputs = {
      platform: Platform.VS,
      "api-plugin-type": "api-spec",
    };
    sandbox.stub(CopilotPluginHelper, "formatValidationErrors").resolves([]);
    sandbox.stub(CopilotPluginHelper, "logValidationResults").resolves();
    sandbox.stub(SpecParser.prototype, "validate").resolves({
      status: ValidationStatus.Valid,
      warnings: [],
      errors: [],
      specHash: "xxx",
    });
    sandbox.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "1",
          server: "https://test",
          operationId: "id1",
          isValid: false,
          reason: [ErrorType.AuthTypeIsNotSupported],
        },
      ],
      allAPICount: 1,
      validAPICount: 0,
    });

    const res = await CopilotPluginHelper.listOperations(context, "", inputs, true, false, "");
    expect(res.isOk()).to.be.true;
  });
});

describe("SpecGenerator", async () => {
  describe("activate", async () => {
    it("should activate and get correct template name", async () => {
      const generator = new SpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      };
      let res = await generator.activate(context, inputs);
      let templateName = generator.getTemplateName(inputs);
      assert.isTrue(res);
      assert.equal(templateName, "api-plugin-existing-api");

      delete inputs[QuestionNames.Capabilities];
      inputs[QuestionNames.MeArchitectureType] = MeArchitectureOptions.apiSpec().id;
      res = generator.activate(context, inputs);
      templateName = generator.getTemplateName(inputs);
      assert.isTrue(res);
      assert.equal(templateName, "copilot-plugin-existing-api");

      delete inputs[QuestionNames.MeArchitectureType];
      inputs[QuestionNames.Capabilities] = CapabilityOptions.customCopilotRag().id;
      inputs[QuestionNames.CustomCopilotRag] = CustomCopilotRagOptions.customApi().id;
      res = generator.activate(context, inputs);
      templateName = generator.getTemplateName(inputs);
      assert.isTrue(res);
      assert.equal(templateName, "custom-copilot-rag-custom-api");
    });
  });

  describe("getTempalteInfos", async () => {
    const sandbox = sinon.createSandbox();
    let mockedEnvRestore: RestoreFn | undefined;
    afterEach(async () => {
      sandbox.restore();
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });
    it("happy path", async () => {
      const generator = new SpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: "testapp",
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      inputs.apiAuthData = [
        { serverUrl: "https://test.com", authName: "test", authType: "apiKey" },
      ];
      let res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "api-plugin-existing-api");
        assert.equal(res.value[0].replaceMap!["DeclarativeCopilot"], "");

        let filterResult = res.value[0].filterFn!("declarativeAgent.json.tpl");
        assert.isFalse(filterResult);
        filterResult = res.value[0].filterFn!("test.json");
        assert.isTrue(filterResult);
        filterResult = res.value[0].filterFn!("instruction.txt");
        assert.isFalse(filterResult);
      }

      inputs[QuestionNames.Capabilities] = CapabilityOptions.declarativeAgent().id;
      res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "api-plugin-existing-api");
        assert.equal(res.value[0].replaceMap!["DeclarativeCopilot"], "true");

        let filterResult = res.value[0].filterFn!("declarativeAgent.json.tpl");
        assert.isTrue(filterResult);
        filterResult = res.value[0].filterFn!("instruction.txt");
        assert.isTrue(filterResult);
      }

      delete inputs[QuestionNames.Capabilities];
      delete inputs.apiAuthData;
      inputs[QuestionNames.MeArchitectureType] = MeArchitectureOptions.apiSpec().id;
      res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "copilot-plugin-existing-api");
      }

      delete inputs[QuestionNames.MeArchitectureType];
      inputs[QuestionNames.Capabilities] = CapabilityOptions.customCopilotRag().id;
      inputs[QuestionNames.CustomCopilotRag] = CustomCopilotRagOptions.customApi().id;
      res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "custom-copilot-rag-custom-api");
      }
    });

    it("succeed even get yaml file failed", async () => {
      const generator = new SpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: "testapp",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      inputs.apiAuthData = [
        { serverUrl: "https://test.com", authName: "test", authType: "apiKey" },
      ];
      sandbox.stub(commonUtils, "isJsonSpecFile").throws();
      const res = await generator.getTemplateInfos(context, inputs, ".", { telemetryProps: {} });
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "api-plugin-existing-api");
        assert.equal(res.value[0].language, ProgrammingLanguage.CSharp);
      }
    });

    it("happy path for kiota integration with auth", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      const generator = new SpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: "testapp",
        [QuestionNames.ApiPluginManifestPath]: "ai-plugin.json",
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      sandbox.stub(helper, "listOperations").resolves(
        ok([
          {
            id: "operation1",
            label: "operation1",
            groupName: "1",
            data: {
              serverUrl: "https://server1",
              authName: "auth",
              authType: "apiKey",
            },
          },
        ])
      );
      const res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "api-plugin-existing-api");
        assert.equal(res.value[0].replaceMap!["DeclarativeCopilot"], "");

        let filterResult = res.value[0].filterFn!("declarativeAgent.json.tpl");
        assert.isFalse(filterResult);
        filterResult = res.value[0].filterFn!("test.json");
        assert.isTrue(filterResult);
        filterResult = res.value[0].filterFn!("instruction.txt");
        assert.isFalse(filterResult);
      }
    });

    it("happy path for kiota integration without auth", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      const generator = new SpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: "testapp",
        [QuestionNames.ApiPluginManifestPath]: "ai-plugin.json",
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      sandbox.stub(helper, "listOperations").resolves(
        ok([
          {
            id: "operation1",
            label: "operation1",
            groupName: "1",
            data: {
              serverUrl: "https://server1",
            },
          },
        ])
      );
      const res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "api-plugin-existing-api");
        assert.equal(res.value[0].replaceMap!["DeclarativeCopilot"], "");

        let filterResult = res.value[0].filterFn!("declarativeAgent.json.tpl");
        assert.isFalse(filterResult);
        filterResult = res.value[0].filterFn!("test.json");
        assert.isTrue(filterResult);
        filterResult = res.value[0].filterFn!("instruction.txt");
        assert.isFalse(filterResult);
      }
    });

    it("should not throw error parse failed for kiota integration", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      const generator = new SpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: "testapp",
        [QuestionNames.ApiPluginManifestPath]: "ai-plugin.json",
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      sandbox.stub(helper, "listOperations").resolves(
        err([
          {
            type: ErrorType.SpecNotValid,
            content: "test",
          },
        ])
      );
      const res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
    });
  });

  describe("SpecGenerator: post", function () {
    const tools = new MockTools();
    setTools(tools);
    const sandbox = sinon.createSandbox();
    let mockedEnvRestore: RestoreFn | undefined;

    const apiOperations: ApiOperation[] = [
      {
        id: "operation1",
        label: "operation1",
        groupName: "1",
        data: {
          serverUrl: "https://server1",
        },
      },
      {
        id: "operation2",
        label: "operation2",
        groupName: "1",
        data: {
          serverUrl: "https://server1",
          authName: "auth",
        },
      },
    ];

    afterEach(async () => {
      sandbox.restore();
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });

    it("API ME success", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath", {
        telemetryProps: {
          "project-id": "test",
        },
      });

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("API ME success with API Key authentication", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ApiSpecLocation]: "test.json",
        [QuestionNames.ApiOperation]: ["operation2"],
        supportedApisFromApiSpec: apiOperations,
        apiAuthData: [
          {
            authType: "apiKey",
            serverUrl: "",
          },
        ],
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("API plugin success", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generateForCopilot")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("success with api spec warning and generate warnings", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.ApiSpecLocation]: "test.json",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").resolves({
        status: ValidationStatus.Warning,
        errors: [],
        warnings: [
          {
            type: WarningType.OperationIdMissing,
            content: "warning",
            data: ["operation1", " operation2"],
          },
          {
            type: WarningType.ConvertSwaggerToOpenAPI,
            content: "",
          },
        ],
        specHash: "xxx",
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({ ...teamsManifest }));
      const generateParser = sandbox.stub(SpecParser.prototype, "generate").resolves({
        allSuccess: true,
        warnings: [
          { type: WarningType.GenerateCardFailed, content: "test", data: "getPets" },
          {
            type: WarningType.OperationOnlyContainsOptionalParam,
            content: "test",
            data: "getPets",
          },
        ],
      });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("warning message");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      if (result.isOk()) {
        assert.isTrue(result.value.warnings!.length === 4);
        assert.isFalse(result.value.warnings![0].content.includes("operation2"));
        assert.isUndefined(result.value.warnings![0].data);
        assert.equal(result.value.warnings![1].type, WarningType.ConvertSwaggerToOpenAPI);
        assert.equal(result.value.warnings![2].type, WarningType.GenerateCardFailed);
        assert.equal(
          result.value.warnings![3].type,
          WarningType.OperationOnlyContainsOptionalParam
        );
        assert.equal(result.value.warnings![3].content, "");
        assert.isTrue(generateParser.args[0][3]?.includes(ResponseTemplatesFolderName));
      }
    });

    it("success without api spec warning after filtering", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").resolves({
        status: ValidationStatus.Warning,
        errors: [],
        warnings: [
          { type: WarningType.OperationIdMissing, content: "warning", data: ["operation2"] },
        ],
        specHash: "xxx",
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({ ...teamsManifest }));
      sandbox.stub(SpecParser.prototype, "generate").resolves({ allSuccess: true, warnings: [] });

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      if (result.isOk()) {
        assert.isTrue(result.value.warnings!.length === 0);
      }
    });

    it("success with warnings when CSharp", async function () {
      const inputs: Inputs = {
        platform: Platform.VS,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").resolves({
        status: ValidationStatus.Warning,
        errors: [],
        warnings: [{ type: WarningType.OperationIdMissing, content: "warning" }],
        specHash: "xxx",
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox
        .stub(manifestUtils, "_readAppManifest")
        .resolves(ok({ ...teamsManifest, name: { short: "", full: "" } }));
      sandbox.stub(SpecParser.prototype, "generate").resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("warn message");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
    });

    it("invalid API spec", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").resolves({
        status: ValidationStatus.Error,
        errors: [{ type: ErrorType.NoServerInformation, content: "" }],
        warnings: [],
        specHash: "xxx",
      });

      sandbox.stub(SpecParser.prototype, "generate").resolves();

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.isTrue(result.error.name === "invalid-api-spec");
      }
    });

    it("read manifest error", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox
        .stub(manifestUtils, "_readAppManifest")
        .resolves(err(new SystemError("readManifest", "name", "", "")));
      sandbox.stub(SpecParser.prototype, "generate").resolves({ allSuccess: true, warnings: [] });

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.source, "readManifest");
      }
    });

    it("throws exception", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").throws(new Error("test"));

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
    });

    it("throws specParser error", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox
        .stub(SpecParser.prototype, "generate")
        .throws(new SpecParserError("test", ErrorType.Unknown));
      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.message, "test");
      }
    });

    it("generateCustomCopilot: success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        getTemplateInfosState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(SpecParser.prototype, "getFilteredSpecs").resolves([
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
      ]);
      sandbox.stub(CopilotPluginHelper, "updateForCustomApi").resolves([]);
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("generateCustomCopilot for csharp: success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        getTemplateInfosState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(SpecParser.prototype, "getFilteredSpecs").resolves([
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
      ]);
      sandbox.stub(CopilotPluginHelper, "updateForCustomApi").resolves([]);
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("generateCustomCopilot: CLI with warning", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        getTemplateInfosState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(SpecParser.prototype, "getFilteredSpecs").resolves([
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
      ]);
      sandbox.stub(CopilotPluginHelper, "updateForCustomApi").resolves([]);
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("warning message");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("generateCustomCopilot: error", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        getTemplateInfosState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(SpecParser.prototype, "getFilteredSpecs").resolves([
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
      ]);
      sandbox.stub(CopilotPluginHelper, "updateForCustomApi").throws(new Error("test"));
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(SpecParser.prototype, "generate").resolves({ allSuccess: true, warnings: [] });

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr() && result.error.message === "test");
    });

    it("generate for oauth: success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        supportedApisFromApiSpec: [
          {
            id: "operation1",
            label: "operation1",
            groupName: "1",
            data: {
              serverUrl: "https://server1",
              authName: "auth",
              authType: "oauth2",
            },
          },
        ] as ApiOperation[],
        getTemplateInfosState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();

      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generateForCopilot")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("declarative copilot with plugin success", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.Capabilities]: CapabilityOptions.declarativeAgent().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.WithPlugin]: DeclarativeCopilotTypeOptions.withPlugin().id,
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const addAction = sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generateForCopilot")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
      assert.isTrue(addAction.calledOnce);
    });

    it("declarative copilot with plugin error", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.Capabilities]: CapabilityOptions.declarativeAgent().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.WithPlugin]: DeclarativeCopilotTypeOptions.withPlugin().id,
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        getTemplateInfosState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const addAction = sandbox
        .stub(copilotGptManifestUtils, "addAction")
        .resolves(err(new SystemError("test", "test", "test", "test")));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generateForCopilot")
        .resolves({ allSuccess: true, warnings: [] });

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr() && result.error.name === "test");
      assert.isTrue(generateBasedOnSpec.calledOnce);
      assert.isTrue(addAction.calledOnce);
    });

    it("generate for kiota", async function () {
      mockedEnvRestore = mockedEnv({ [FeatureFlagName.KiotaIntegration]: "true" });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiPluginManifestPath]: "test.json",
        [QuestionNames.ProjectType]: "copilot-agent-type",
        getTemplateInfosState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "api1",
            server: "https://test",
            operationId: "get",
            auth: {
              name: "test",
              authScheme: {
                type: "http",
                scheme: "bearer",
              },
            },
            isValid: true,
            reason: [],
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(false);
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "readJSON").resolves({});
      sandbox.stub(SpecParser.prototype, "generateAdaptiveCardInPlugin").resolves();
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

      const copyKiotaFolder = sandbox.stub(fs, "copy").callsFake((src, dest, options) => {
        assert.isTrue(src.endsWith(".kiota"));
        assert.isTrue(dest.endsWith(".kiota"));
      });
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(copyKiotaFolder.calledOnce);
    });

    it("generate for kiota without .kiota folder", async function () {
      mockedEnvRestore = mockedEnv({ [FeatureFlagName.KiotaIntegration]: "true" });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiPluginManifestPath]: "test.json",
        [QuestionNames.ProjectType]: "copilot-agent-type",
        getTemplateInfosState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "api1",
            server: "https://test",
            operationId: "get",
            auth: {
              name: "test",
              authScheme: {
                type: "http",
                scheme: "bearer",
              },
            },
            isValid: true,
            reason: [],
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(false).onSecondCall().resolves(false);
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "readJSON").resolves({});
      sandbox.stub(SpecParser.prototype, "generateAdaptiveCardInPlugin").resolves();
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

      const copyKiotaFolder = sandbox.stub(fs, "copy").resolves();
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(copyKiotaFolder.notCalled);
    });

    it("generate for kiota with .kiota folder already exists", async function () {
      mockedEnvRestore = mockedEnv({ [FeatureFlagName.KiotaIntegration]: "true" });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiPluginManifestPath]: "test.json",
        [QuestionNames.ProjectType]: "copilot-agent-type",
        getTemplateInfosState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "api1",
            server: "https://test",
            operationId: "get",
            auth: {
              name: "test",
              authScheme: {
                type: "http",
                scheme: "bearer",
              },
            },
            isValid: true,
            reason: [],
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(true);
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "readJSON").resolves({});
      sandbox.stub(SpecParser.prototype, "generateAdaptiveCardInPlugin").resolves();
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

      const copyKiotaFolder = sandbox.stub(fs, "copy").resolves();
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(copyKiotaFolder.notCalled);
    });

    it("generate for kiota without adaptive card", async function () {
      mockedEnvRestore = mockedEnv({ [FeatureFlagName.KiotaIntegration]: "true" });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiPluginManifestPath]: "test.json",
        [QuestionNames.ProjectType]: "copilot-agent-type",
        getTemplateInfosState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "api1",
            server: "https://test",
            operationId: "get",
            auth: {
              name: "test",
              authScheme: {
                type: "http",
                scheme: "bearer",
              },
            },
            isValid: true,
            reason: [],
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(false);
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "readJSON").resolves({});
      sandbox
        .stub(SpecParser.prototype, "generateAdaptiveCardInPlugin")
        .throws(new Error("generate ac failed"));
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

      const copyKiotaFolder = sandbox.stub(fs, "copy").resolves();
      sandbox.stub(pluginGeneratorHelper, "generateScaffoldingSummary").resolves("");

      const generator = new SpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(copyKiotaFolder.calledOnce);
    });
  });
});
