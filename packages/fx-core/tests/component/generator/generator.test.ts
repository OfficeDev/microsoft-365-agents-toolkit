// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Platform } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import axios, { AxiosError, AxiosHeaders, AxiosResponse } from "axios";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import Mustache from "mustache";
import path from "path";
import { createSandbox } from "sinon";
import * as folderUtils from "../../../../fx-core/src/folder";
import { createContext, setTools } from "../../../src/common/globalVars";
import { getLocalizedString } from "../../../src/common/localizeUtils";
import * as requestUtils from "../../../src/common/requestUtils";
import { sendRequestWithRetry, sendRequestWithTimeout } from "../../../src/common/requestUtils";
import { SampleConfig, SampleUrlInfo, sampleProvider } from "../../../src/common/samples";
import templateConfig from "../../../src/common/templates-config.json";
import {
  commonTemplateName,
  placeholderDelimiters,
} from "../../../src/component/generator/constant";
import { DefaultTemplateGenerator } from "../../../src/component/generator/defaultGenerator";
import {
  DownloadSampleApiLimitError,
  DownloadSampleNetworkError,
  FetchSampleInfoError,
} from "../../../src/component/generator/error";
import { Generator } from "../../../src/component/generator/generator";
import {
  GeneratorContext,
  ScaffoldLocalTemplateAction,
  ScaffoldRemoteTemplateAction,
  TemplateActionSeq,
  fetchSampleInfoAction,
} from "../../../src/component/generator/generatorAction";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import { getTemplateReplaceMap } from "../../../src/component/generator/templates/templateReplaceMap";
import * as generatorUtils from "../../../src/component/generator/utils";
import {
  downloadDirectory,
  getSampleInfoFromName,
  isApiLimitError,
  renderTemplateFileData,
  renderTemplateFileName,
  runWithLimitedConcurrency,
  simplifyAxiosError,
} from "../../../src/component/generator/utils";
import { ActionContext } from "../../../src/component/middleware/actionExecutionMW";
import { CapabilityOptions, ProgrammingLanguage, QuestionNames } from "../../../src/question";
import sampleConfigV3 from "../../common/samples-config-v3.json";
import { MockTools, randomAppName } from "../../core/utils";

const mockedSampleInfo: SampleConfig = {
  id: "test-id",
  onboardDate: new Date(),
  title: "test-title",
  shortDescription: "test-sd",
  fullDescription: "test-fd",
  types: [],
  tags: [],
  time: "",
  configuration: "test-configuration",
  suggested: false,
  thumbnailPath: "",
  gifUrl: "",
  downloadUrlInfo: {
    owner: "OfficeDev",
    repository: "TeamsFx-Samples",
    ref: "dev",
    dir: "test",
  },
};

// The sample prefix is present in the downloadurl of the external sample
const mockedExternalSampleConfig = {
  samples: [
    {
      id: "test",
      onboardDate: new Date(),
      title: "test-title",
      shortDescription: "test-sd",
      fullDescription: "test-fd",
      types: [],
      tags: [],
      time: "",
      configuration: "test-configuration",
      suggested: false,
      gifUrl: "",
      downloadUrlInfo: {
        owner: "Org",
        repository: "Repo",
        ref: "main",
        dir: "sample/test",
      },
    },
  ],
};

describe("Generator utils", () => {
  const tmpDir = path.join(__dirname, "tmp");
  const sandbox = createSandbox();
  let mockedEnvRestore: RestoreFn = () => {};

  afterEach(async () => {
    sandbox.restore();
    if (await fs.pathExists(tmpDir)) {
      await fs.remove(tmpDir);
    }
    mockedEnvRestore();
  });

  it("return rc if set env rc", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_TEMPLATE_PRERELEASE: "rc",
    });
    const tagList = "1.0.0\n 2.0.0\n 2.1.0\n 3.0.0\n 0.0.0-rc";
    sandbox.replace(templateConfig, "useLocalTemplate", false);
    sandbox.stub(axios, "get").resolves({ data: tagList, status: 200 } as AxiosResponse);
    const url = await generatorUtils.getTemplateUrl(
      "templateName",
      generatorUtils.getTemplateLatestVersion
    );
    assert.isTrue(url?.includes("0.0.0-rc"));
  });

  it("return correct version", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_TEMPLATE_PRERELEASE: "",
    });
    const tagList = "1.0.0\n 2.0.0\n 2.1.0\n 3.0.0";
    const tag = "2.1.0";
    sandbox.replace(templateConfig, "useLocalTemplate", false);
    sandbox.stub(axios, "get").resolves({ data: tagList, status: 200 } as AxiosResponse);
    sandbox.stub(templateConfig, "version").value("^2.0.0");
    sandbox.replace(templateConfig, "tagPrefix", "templates@");
    const templateName = "templateName";
    const selectedTag = await generatorUtils.getTemplateLatestVersion();
    const url = generatorUtils.getTemplateZipUrlByVersion(templateName, selectedTag);
    assert.isTrue(url.includes(tag));
  });

  it("return error if version pattern cannot match tag list", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_TEMPLATE_PRERELEASE: "",
    });
    const tagList = "1.0.0\n 2.0.0\n 2.1.0\n 3.0.0";
    sandbox.stub(axios, "get").resolves({ data: tagList, status: 200 } as AxiosResponse);
    sandbox.stub(templateConfig, "version").value("^4.0.0");
    sandbox.replace(templateConfig, "tagPrefix", "templates@");
    try {
      await generatorUtils.getTemplateLatestVersion();
    } catch (e) {
      assert.exists(e);
      return;
    }
    assert.fail("Should not reach here.");
  });

  it("sendRequestWithRetry throw error if requestFn returns error status code", async () => {
    const requestFn = async () => {
      return { status: 400 } as AxiosResponse;
    };
    try {
      await sendRequestWithRetry(requestFn, 1);
    } catch (e) {
      assert.exists(e);
      return;
    }
    assert.fail("Should not reach here.");
  });

  it("sendRequestWithRetry throw error if requestFn throw error", async () => {
    const requestFn = async () => {
      throw new Error("test");
    };
    try {
      await sendRequestWithRetry(requestFn, 1);
    } catch (e) {
      assert.exists(e);
      return;
    }
    assert.fail("Should not reach here.");
  });

  it("sendRequestWithTimeout throw error if requestFn throw error", async () => {
    const requestFn = async () => {
      throw new Error("test");
    };
    try {
      await sendRequestWithTimeout(requestFn, 1000, 1);
    } catch (e) {
      assert.exists(e);
      return;
    }
    assert.fail("Should not reach here.");
  });

  it("sendRequestWithTimeout throw request timeout if requestFn throw error", async () => {
    const requestFn = async () => {
      throw new Error("test");
    };
    sandbox.stub(axios, "isCancel").returns(true);
    try {
      await sendRequestWithTimeout(requestFn, 1000, 2);
    } catch (e) {
      assert.exists(e);
      return;
    }
    assert.fail("Should not reach here.");
  });

  it("fetch zip from url", async () => {
    sandbox.stub(axios, "get").resolves({ status: 200, data: new AdmZip().toBuffer() });
    const url = "ut";
    const zip = await generatorUtils.fetchZipFromUrl(url);
    assert.equal(zip.getEntries().length, 0);
  });

  it("unzip", async () => {
    const inputDir = path.join(tmpDir, "input");
    const outputDir = path.join(tmpDir, "output");
    await fs.ensureDir(inputDir);
    const fileData = "{{appName}}";
    await fs.writeFile(path.join(inputDir, "test.txt.tpl"), fileData);
    const zip = new AdmZip();
    zip.addLocalFolder(inputDir);
    zip.writeZip(path.join(tmpDir, "test.zip"));
    await generatorUtils.unzip(
      new AdmZip(path.join(tmpDir, "test.zip")),
      outputDir,
      (fileName: string, fileData: Buffer) => renderTemplateFileName(fileName, fileData, {}),
      (fileName: string, fileData: Buffer) =>
        renderTemplateFileData(fileName, fileData, { appName: "test" })
    );
    const content = await fs.readFile(path.join(outputDir, "test.txt"), "utf8");
    assert.equal(content, "test");
  });

  it("unzip with no render function", async () => {
    const inputDir = path.join(tmpDir, "input");
    const outputDir = path.join(tmpDir, "output");
    await fs.ensureDir(inputDir);
    const fileData = "{{appName}}";
    await fs.writeFile(path.join(inputDir, "test.txt"), fileData);
    const zip = new AdmZip();
    zip.addLocalFolder(inputDir);
    zip.writeZip(path.join(tmpDir, "test.zip"));
    await generatorUtils.unzip(new AdmZip(path.join(tmpDir, "test.zip")), outputDir);
    const content = await fs.readFile(path.join(outputDir, "test.txt"), "utf8");
    assert.equal(content, fileData);
  });

  it("unzip with relative path", async () => {
    const inputDir = path.join(tmpDir, "input");
    const outputDir = path.join(tmpDir, "output");
    await fs.ensureDir(inputDir);
    const fileData = "{{appName}}";
    await fs.writeFile(path.join(inputDir, "test.txt.tpl"), fileData);
    const zip = new AdmZip();
    zip.addLocalFolder(inputDir);
    zip.writeZip(path.join(tmpDir, "test.zip"));
    await generatorUtils.unzip(
      new AdmZip(path.join(tmpDir, "test.zip")),
      outputDir,
      (fileName: string, fileData: Buffer) => renderTemplateFileName(fileName, fileData, {}),
      (fileName: string, fileData: Buffer) =>
        renderTemplateFileData(fileName, fileData, { appName: "test" }),
      (fileName: string) => fileName.startsWith("test1")
    );
    assert.isFalse(await fs.pathExists(path.join(outputDir, "test.txt")));
  });

  it("get sample info from name", async () => {
    const sampleName = "test";
    try {
      getSampleInfoFromName(sampleName);
    } catch (e) {
      assert.equal(e.message, "Invalid inputs: sample 'test' not found");
    }
  });

  it("not render if file doensn't end with .tpl", async () => {
    const res = renderTemplateFileData("fileName", Buffer.from("appName", "utf-8"), {
      appName: "test",
    });
    assert.equal(res.toString(), "appName");
  });

  it("zip folder", async () => {
    const inputDir = path.join(tmpDir, "input");
    await fs.ensureDir(inputDir);
    const fileData = "test";
    await fs.writeFile(path.join(inputDir, "test.txt"), fileData);
    const zip = generatorUtils.zipFolder(inputDir);
    zip.getEntry("test.txt")!.getData().toString();
    zip.getEntries().forEach((entry) => {
      assert.equal(entry.getData().toString(), "test");
      assert.equal(zip.getEntries().length, 1);
    });
  });

  it("download directory get file info error", async () => {
    const axiosStub = sandbox.stub(axios, "get");
    const error = new Error("Network error");
    (error as any).isAxiosError = true;
    (error as any).response = {
      status: 403,
      headers: {
        "x-ratelimit-remaining": "0",
      },
    };
    axiosStub.onFirstCall().rejects(error);
    try {
      await downloadDirectory(
        {
          owner: "OfficeDev",
          repository: "TeamsFx-Samples",
          ref: "dev",
          dir: "test",
        },
        tmpDir
      );
    } catch (e) {
      assert.exists(e);
      return;
    }
    assert.fail("Should not reach here.");
  });

  it("download directory happy path", async () => {
    const axiosStub = sandbox.stub(axios, "get");
    const sampleName = "test";
    const mockFileName = "test.txt";
    const mockFileData = "test data";
    const fileInfo = [{ type: "file", path: `${sampleName}/${mockFileName}` }];
    axiosStub.onFirstCall().resolves({ status: 200, data: { tree: fileInfo } });
    axiosStub.onSecondCall().resolves({ status: 200, data: mockFileData });
    await fs.ensureDir(tmpDir);
    await downloadDirectory(
      {
        owner: "OfficeDev",
        repository: "TeamsFx-Samples",
        ref: "dev",
        dir: "test",
      },
      tmpDir
    );
    const data = await fs.readFile(path.join(tmpDir, mockFileName), "utf8");
    assert.equal(data, mockFileData);
  });

  it("download directory with LFS files", async () => {
    const axiosStub = sandbox.stub(axios, "get");
    const sampleName = "test";
    const mockFileName = "test.txt";
    const mockFileData = "test data";
    const lfsData =
      "version https://git-lfs.github.com/spec/v1\noid sha256:548c1fe07b6b278da680ccd84483be06262521f2e3\nsize 100";
    const fileInfo = [{ type: "file", path: `${sampleName}/${mockFileName}` }];
    axiosStub.onFirstCall().resolves({ status: 200, data: { tree: fileInfo } });
    axiosStub.onSecondCall().resolves({ status: 200, data: lfsData });
    axiosStub.onThirdCall().resolves({ status: 200, data: mockFileData });
    await fs.ensureDir(tmpDir);
    await downloadDirectory(
      {
        owner: "OfficeDev",
        repository: "TeamsFx-Samples",
        ref: "dev",
        dir: "test",
      },
      tmpDir
    );
    const data = await fs.readFile(path.join(tmpDir, mockFileName), "utf8");
    assert.equal(data, mockFileData);
  });

  it("limit concurrency", async () => {
    const data = [1, 10, 2, 3];
    let res: number[] = [];
    const callback = async (num: number) => {
      await new Promise((resolve) => setTimeout(resolve, num * 10));
      res.push(num);
    };
    await runWithLimitedConcurrency(data, callback, 2);
    assert.deepEqual(res, [1, 2, 3, 10]);
    res = [];
    await runWithLimitedConcurrency(data, callback, 1);
    assert.deepEqual(res, [1, 10, 2, 3]);
  });

  it("convert sample info to url", async () => {
    const sampleInfo: SampleUrlInfo = {
      owner: "OfficeDev",
      repository: "TeamsFx-Samples",
      ref: "dev",
      dir: "test",
    };
    const url = generatorUtils.convertToUrl(sampleInfo);
    assert.equal(url, "https://github.com/OfficeDev/TeamsFx-Samples/tree/dev/test");
  });

  it("should simplify an AxiosError", () => {
    const mockError: AxiosError = {
      message: "API rate limit exceeded",
      name: "AxiosError",
      code: "403",
      stack: "Error stack",
      response: {
        config: {
          headers: new AxiosHeaders(),
        },
        status: 403,
        statusText: "Forbidden",
        headers: {
          "x-ratelimit-remaining": "0",
        },
        data: "Error data",
      },
      isAxiosError: true,
      toJSON: () => ({}),
    };
    const simplifiedError = simplifyAxiosError(mockError);
    const expectedError = {
      message: "API rate limit exceeded",
      name: "AxiosError",
      code: "403",
      config: undefined,
      stack: "Error stack",
      status: 403,
      statusText: "Forbidden",
      headers: {
        "x-ratelimit-remaining": "0",
      },
      data: "Error data",
    };

    assert.deepEqual(simplifiedError, expectedError);
  });
  it("should simplify an AxiosError with no response", () => {
    const mockError: AxiosError = {
      message: "API rate limit exceeded",
      name: "AxiosError",
      code: "403",
      stack: "Error stack",
      isAxiosError: true,
      toJSON: () => ({}),
    };
    const simplifiedError = simplifyAxiosError(mockError);
    const expectedError = {
      message: "API rate limit exceeded",
      name: "AxiosError",
      code: "403",
      data: undefined,
      headers: undefined,
      status: undefined,
      statusText: undefined,
      config: undefined,
      stack: "Error stack",
    };

    assert.deepEqual(simplifiedError, expectedError);
  });

  it("should return true for an API limit error", () => {
    const mockError: AxiosError = {
      message: "API rate limit exceeded",
      name: "AxiosError",
      code: "403",
      stack: "Error stack",
      response: {
        config: {
          headers: new AxiosHeaders(),
        },
        status: 403,
        statusText: "Forbidden",
        headers: {
          "x-ratelimit-remaining": "0",
        },
        data: "Error data",
      },
      isAxiosError: true,
      toJSON: () => ({}),
    };

    assert.isTrue(isApiLimitError(mockError));
  });

  it("should return false for a non-API limit error", () => {
    const mockError: AxiosError = {
      message: "Not Found",
      name: "AxiosError",
      code: "404",
      stack: "Error stack",
      isAxiosError: true,
      toJSON: () => ({}),
    };
    assert.isFalse(isApiLimitError(mockError));
  });

  it("convertToLangKey for none", () => {
    const key = generatorUtils.convertToLangKey(ProgrammingLanguage.None);
    assert.equal(key, "common");
  });
});

describe("Generator error", async () => {
  const tools = new MockTools();
  setTools(tools);
  const ctx = createContext();
  const inputs = {
    platform: Platform.VSCode,
    [QuestionNames.AppName]: randomAppName(),
    [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.JS,
    [QuestionNames.Capabilities]: CapabilityOptions.basicBot().id,
  } as Inputs;
  const sandbox = createSandbox();
  const tmpDir = path.join(__dirname, "tmp");

  afterEach(async () => {
    if (await fs.pathExists(tmpDir)) {
      await fs.remove(tmpDir);
    }
    sandbox.restore();
  });

  [false, true].forEach((newGeneratorFlag) => {
    it("template fallback error", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_NEW_GENERATOR: `${newGeneratorFlag}` });
      sandbox.stub(ScaffoldRemoteTemplateAction, "run").resolves();
      sandbox.stub(folderUtils, "getTemplatesFolder").resolves("foobar");
      const result = newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(ctx, inputs, tmpDir)
        : await Generator.generateTemplate(ctx, tmpDir, "bot", "ts");
      if (result.isErr()) {
        assert.equal(result.error.name, "ScaffoldLocalTemplateError");
      } else {
        assert.fail("template fallback error should be thrown.");
      }
    });

    it("template not found error", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_NEW_GENERATOR: `${newGeneratorFlag}` });
      sandbox.stub(ScaffoldRemoteTemplateAction, "run").resolves();
      sandbox.stub(generatorUtils, "unzip").resolves();
      const result = newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(ctx, inputs, tmpDir)
        : await Generator.generateTemplate(ctx, tmpDir, "bot", "ts");
      if (result.isErr()) {
        assert.equal(result.error.innerError.name, "TemplateNotFoundError");
      } else {
        assert.fail("template not found error should be thrown.");
      }
    });
  });

  it("fetch sample info fail", async () => {
    sandbox.stub(fetchSampleInfoAction, "run").throws(new Error("test"));
    const result = await Generator.generateSample(ctx, tmpDir, "test");
    if (result.isErr()) {
      assert.equal(result.error.name, "FetchSampleInfoError");
    } else {
      assert.fail("fetch sample info error should be thrown.");
    }
  });

  it("sample not found error", async () => {
    sandbox.stub(generatorUtils, "getSampleInfoFromName").resolves(mockedSampleInfo);
    sandbox.stub(generatorUtils, "downloadDirectory").resolves([] as string[]);
    sandbox
      .stub(requestUtils, "sendRequestWithTimeout")
      .resolves({ data: sampleConfigV3 } as AxiosResponse);

    const result = await Generator.generateSample(ctx, tmpDir, "test");
    if (result.isErr()) {
      assert.equal(result.error.name, "SampleNotFoundError");
    } else {
      assert.fail("Sample not found error should be thrown.");
    }
  });
  it("create download sample network error with correct inner error", async () => {
    const url = "http://example.com";
    const mockError: AxiosError = {
      message: "Test error",
      name: "AxiosError",
      config: {
        headers: new AxiosHeaders(),
      },
      code: "500",
      stack: "Error stack",
      response: {
        config: {
          headers: new AxiosHeaders(),
        },
        status: 500,
        statusText: "Internal Server Error",
        headers: {},
        data: "Error data",
      },
      isAxiosError: true,
      toJSON: () => ({}),
    };
    const error = new DownloadSampleNetworkError(url, mockError);
    assert.deepEqual(error.innerError, simplifyAxiosError(mockError));
  });
  it("create fetch sample info error with correct inner error", async () => {
    const mockError: AxiosError = {
      message: "Test error",
      name: "AxiosError",
      config: {
        headers: new AxiosHeaders(),
      },
      code: "500",
      stack: "Error stack",
      response: {
        config: {
          headers: new AxiosHeaders(),
        },
        status: 500,
        statusText: "Internal Server Error",
        headers: {},
        data: "Error data",
      },
      isAxiosError: true,
      toJSON: () => ({}),
    };
    const error = new FetchSampleInfoError(mockError);
    assert.deepEqual(error.innerError, simplifyAxiosError(mockError));
  });
  it("create download sample api limit error with correct inner error", async () => {
    const url = "http://example.com";
    const mockError: AxiosError = {
      message: "API rate limit exceeded",
      name: "AxiosError",
      code: "403",
      stack: "Error stack",
      response: {
        config: {
          headers: new AxiosHeaders(),
        },
        status: 403,
        statusText: "Forbidden",
        headers: {
          "x-ratelimit-remaining": "0",
        },
        data: "Error data",
      },
      isAxiosError: true,
      toJSON: () => ({}),
    };
    const error = new DownloadSampleApiLimitError(url, mockError);
    assert.deepEqual(error.innerError, simplifyAxiosError(mockError));
  });

  it("scaffold remote, miss key error: language", async () => {
    try {
      const ctx = { name: "bot", destination: tmpDir } as GeneratorContext;
      await ScaffoldRemoteTemplateAction.run(ctx);
    } catch (err: any) {
      assert.equal(err?.name, "MissKeyError");
      assert.include(err?.message, "language");
    }
  });
  it("scaffold local, missing key error: language", async () => {
    try {
      const ctx = { name: "bot", destination: tmpDir } as GeneratorContext;
      await ScaffoldLocalTemplateAction.run(ctx);
    } catch (err: any) {
      assert.equal(err?.name, "MissKeyError");
      assert.include(err?.message, "language");
    }
  });
});

describe("render template", () => {
  it("escape undefined or variable", () => {
    [{ variable: "test" }, { variable: "test", app: null }].forEach((variables) => {
      // arrange
      const filename = "test.tpl";
      const fileData = Buffer.from("{{variable}}{{app}}");
      const expectedResult = "test{{app}}";

      // execute
      const result = renderTemplateFileData(filename, fileData, variables as any);

      assert.equal(result, expectedResult);
    });
  });

  it("escape nested undefined variables", () => {
    // arrange
    const filename = "test.tpl";
    const fileData = Buffer.from("{{#parent}}test{{child}}{{/parent}}");
    const variables1 = { parent: "true", child: null };
    const variables2 = { parent: "true", child: "hello" };
    const expectedResult1 = "test{{child}}";
    const expectedResult2 = "testhello";

    // execute
    const result1 = renderTemplateFileData(filename, fileData, variables1 as any);
    const result2 = renderTemplateFileData(filename, fileData, variables2 as any);

    // assert
    assert.equal(result1, expectedResult1);
    assert.equal(result2, expectedResult2);
  });

  it("should render array element", () => {
    const filename = "test.tpl";
    const fileData = Buffer.from("{{#person}}{{name}},{{email}}{{/person}}");
    const variables1 = {
      person: [
        { name: "name1", email: "email1" },
        { name: "name2", email: "email2" },
      ],
    };
    const expectedResult1 = "name1,email1name2,email2";

    // execute
    const result1 = renderTemplateFileData(filename, fileData, variables1 as any);

    // assert
    assert.equal(result1, expectedResult1);
  });

  it("should render object element", () => {
    const filename = "test.tpl";
    const fileData = Buffer.from("{{#person}}{{name}},{{email}}{{/person}}");
    const variables1 = {
      person: { name: "name1", email: "email1" },
    };
    const expectedResult1 = "name1,email1";

    // execute
    const result1 = renderTemplateFileData(filename, fileData, variables1 as any);

    // assert
    assert.equal(result1, expectedResult1);
  });

  it("do not escape empty string variable", () => {
    // arrange
    const filename = "test.tpl";
    const fileData = Buffer.from("{{variable}}{{app}}");
    const variables = { variable: "test", app: "" };

    // execute
    const result = renderTemplateFileData(filename, fileData, variables);
    const expectedResult = Mustache.render(
      fileData.toString(),
      variables,
      {},
      placeholderDelimiters
    );

    assert.equal(result, expectedResult);
  });

  it("skip non template file", () => {
    // arrange
    const filename = "test.txt";
    const fileData = Buffer.from("{{variable}}{{app}}");
    const variables = { variable: "test", app: "" };
    const expectedResult = fileData;
    // execute
    const result = renderTemplateFileData(filename, fileData, variables);

    assert.deepEqual(result, expectedResult);
  });

  it("no variables", () => {
    // arrange
    const filename = "test.tpl";
    const fileData = Buffer.from("{{variable}}{{app}}");
    const expectedResult = fileData.toString();
    // execute
    const result = renderTemplateFileData(filename, fileData);

    assert.deepEqual(result, expectedResult);
  });
});

[false, true].forEach((newGeneratorFlag) => {
  describe(`Generator happy path with new generator enabled=${newGeneratorFlag}`, async () => {
    const tools = new MockTools();
    setTools(tools);
    const context = createContext();
    let inputs: Inputs;
    const sandbox = createSandbox();
    const tmpDir = path.join(__dirname, "tmp");
    const templateName = TemplateNames.DefaultBot;
    const language = "ts";
    let mockedEnvRestore: RestoreFn = () => {};

    async function buildFakeTemplateZip(templateName: string, mockFileName: string) {
      const mockFileData = "test data";
      const fallbackDir = path.join(tmpDir, "fallback");
      await fs.ensureDir(fallbackDir);
      const templateZip = new AdmZip();
      templateZip.addFile(path.join(templateName, mockFileName), Buffer.from(mockFileData));
      templateZip.writeZip(path.join(fallbackDir, "ts.zip"));
      return templateZip;
    }

    beforeEach(() => {
      inputs = {
        platform: Platform.VSCode,
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.Capabilities]: CapabilityOptions.basicBot().id,
        [QuestionNames.TemplateName]: TemplateNames.DefaultBot,
      } as Inputs;
      sandbox.stub(process, "env").value({ TEAMSFX_NEW_GENERATOR: "true" });
    });

    afterEach(async () => {
      sandbox.restore();
      if (await fs.pathExists(tmpDir)) {
        await fs.remove(tmpDir);
      }
      mockedEnvRestore();
    });

    it("external sample", async () => {
      const axiosStub = sandbox.stub(axios, "get");
      sandbox
        .stub(sampleProvider, "SampleCollection")
        .value(Promise.resolve(mockedExternalSampleConfig));
      const sampleName = "test";
      const mockFileName = "test.txt";
      const mockFileData = "test data";
      const foobarName = "foobar";
      const foobarFileName = "foobar.txt";
      const fileInfo = [
        { type: "file", path: `sample/${sampleName}/${mockFileName}` },
        { type: "file", path: `sample/${foobarName}/${foobarFileName}` },
      ];
      axiosStub.onFirstCall().resolves({ status: 200, data: { tree: fileInfo } });
      axiosStub.onSecondCall().resolves({ status: 200, data: mockFileData });
      const result = await Generator.generateSample(context, tmpDir, sampleName);
      assert.isTrue(result.isOk());
      if (!fs.existsSync(path.join(tmpDir, mockFileName))) {
        assert.fail("file creation failure");
      }
      if (fs.existsSync(path.join(tmpDir, foobarFileName))) {
        assert.fail("file should not be created");
      }
    });

    it("template", async () => {
      const inputDir = path.join(tmpDir, "input");
      await fs.ensureDir(path.join(inputDir, templateName));
      const fileData = "{{appName}}";
      await fs.writeFile(path.join(inputDir, templateName, "test.txt.tpl"), fileData);
      const zip = new AdmZip();
      zip.addLocalFolder(inputDir);
      zip.writeZip(path.join(tmpDir, "test.zip"));
      sandbox.stub(generatorUtils, "getTemplateZipUrlByVersion").resolves("test.zip");
      sandbox
        .stub(generatorUtils, "fetchZipFromUrl")
        .resolves(new AdmZip(path.join(tmpDir, "test.zip")));
      context.templateVariables = Generator.getDefaultVariables("test");
      const result = newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(context, inputs, tmpDir)
        : await Generator.generateTemplate(context, tmpDir, templateName, language);
      assert.isTrue(result.isOk());
    });

    it("template variables when test tool enabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_TEST_TOOL: "true" });
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test");
      assert.equal(vars.enableTestToolByDefault, "true");
    });

    it("template variables when test tool disabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_TEST_TOOL: "false" });
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test");
      assert.equal(vars.enableTestToolByDefault, "");
    });

    it("template variables when ME test tool enabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_ME_TEST_TOOL: "true" });
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test");
      assert.equal(vars.enableMETestToolByDefault, "true");
    });

    it("template variables when ME test tool disabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_ME_TEST_TOOL: "false" });
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test");
      assert.equal(vars.enableMETestToolByDefault, "");
    });

    it("template variables when new project enabled", async () => {
      sandbox.stub(process, "env").value({
        TEAMSFX_NEW_PROJECT_TYPE: "true",
        TEAMSFX_NEW_PROJECT_TYPE_NAME: "M365",
        TEAMSFX_NEW_PROJECT_TYPE_EXTENSION: "maproj",
      });
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test");
      assert.equal(vars.isNewProjectTypeEnabled, "true");
    });

    it("template variables when test tool disabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_NEW_PROJECT_TYPE: "false" });
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test");
      assert.equal(vars.isNewProjectTypeEnabled, "");
    });

    it("template variables when set placeProjectFileInSolutionDir to true", async () => {
      inputs.placeProjectFileInSolutionDir = "true";
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test", undefined, undefined, true);
      assert.equal(vars.PlaceProjectFileInSolutionDir, "true");
    });

    it("template variables with custom copilot - OpenAI", async () => {
      inputs.projectId = "test-id";
      inputs[QuestionNames.LLMService] = "llm-service-openai";
      inputs[QuestionNames.OpenAIKey] = "test-key";
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test", "test", undefined, false, undefined, {
            llmService: "llm-service-openai",
            openAIKey: "test-key",
          });
      assert.equal(vars.useOpenAI, "true");
      assert.equal(vars.useAzureOpenAI, "");
      if (newGeneratorFlag) {
        assert.isTrue(vars.openAIKey.startsWith("crypto_"));
      } else {
        assert.equal(vars.openAIKey, "test-key");
      }
      assert.equal(vars.azureOpenAIKey, "");
      assert.equal(vars.azureOpenAIEndpoint, "");
    });

    it("template variables with custom copilot - Azure OpenAI", async () => {
      inputs.projectId = "test-id";
      inputs[QuestionNames.LLMService] = "llm-service-azure-openai";
      inputs[QuestionNames.AzureOpenAIKey] = "test-key";
      inputs[QuestionNames.AzureOpenAIEndpoint] = "test-endpoint";
      inputs[QuestionNames.AzureOpenAIDeploymentName] = "test-deployment";
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test", "test", undefined, false, undefined, {
            llmService: "llm-service-azure-openai",
            azureOpenAIKey: "test-key",
            azureOpenAIEndpoint: "test-endpoint",
            azureOpenAIDeploymentName: "test-deployment",
          });
      assert.equal(vars.useOpenAI, "");
      assert.equal(vars.useAzureOpenAI, "true");
      assert.equal(vars.openAIKey, "");
      if (newGeneratorFlag) {
        assert.isTrue(vars.azureOpenAIKey.startsWith("crypto_"));
      } else {
        assert.equal(vars.azureOpenAIKey, "test-key");
      }
      assert.equal(vars.azureOpenAIEndpoint, "test-endpoint");
      assert.equal(vars.azureOpenAIDeploymentName, "test-deployment");
    });

    it("template variables with custom copilot - AI Search for csharp", async () => {
      inputs.projectId = "test-id";
      inputs[QuestionNames.AzureOpenAIKey] = "test-key";
      inputs[QuestionNames.AzureAISearchApiKey] = "test-search-key";
      inputs[QuestionNames.AzureAISearchEndpoint] = "test-search-endpoint";
      inputs[QuestionNames.OpenAIEmbeddingModel] = "test-openai-embedding-model";
      inputs[QuestionNames.AzureOpenAIEmbeddingDeploymentName] = "test-azure-embedding-deployment";
      const vars = getTemplateReplaceMap(inputs);
      assert.isTrue(vars.azureAISearchApiKey.startsWith("crypto_"));
      assert.equal(vars.azureAISearchEndpoint, "test-search-endpoint");
      assert.equal(vars.openAIEmbeddingModel, "test-openai-embedding-model");
      assert.equal(vars.azureOpenAIEmbeddingDeploymentName, "test-azure-embedding-deployment");
    });

    it("template variables when contains auth", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_TEST_TOOL: "false" });
      const vars = Generator.getDefaultVariables("Test", "Test", "net6", false, [
        {
          authName: "authName",
          openapiSpecPath: "path/to/spec.yaml",
          registrationIdEnvName: "AUTHNAME_REGISTRATION_ID",
          authType: "apiKey",
        },
      ]);
      assert.equal(vars.enableTestToolByDefault, "");
      assert.equal(vars.appName, "Test");
      assert.equal(vars.ApiKey[0].ApiSpecAuthName, "authName");
      assert.equal(vars.ApiKey[0].ApiSpecPath, "path/to/spec.yaml");
      assert.equal(vars.ApiKey[0].ApiSpecAuthRegistrationIdEnvName, "AUTHNAME_REGISTRATION_ID");
      assert.equal(vars.SafeProjectName, "Test");
      assert.equal(vars.SafeProjectNameLowerCase, "test");
    });

    it("template variables when contains auth with special characters", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_TEST_TOOL: "false" });
      const vars = Generator.getDefaultVariables("Test", "Test", "net6", false, [
        {
          authName: "authName",
          openapiSpecPath: "path/to/spec.yaml",
          registrationIdEnvName: "AUTH-NAME_REGISTRATION*ID",
          authType: "oauth2",
        },
        {
          authName: "authName2",
          openapiSpecPath: "path/to/spec.yaml",
          registrationIdEnvName: "AUTH-NAME2_REGISTRATION*ID",
          authType: "apiKey",
        },
      ]);
      assert.equal(vars.enableTestToolByDefault, "");
      assert.equal(vars.appName, "Test");
      assert.equal(vars.OAuth[0].ApiSpecAuthName, "authName");
      assert.equal(vars.OAuth[0].ApiSpecPath, "path/to/spec.yaml");
      assert.equal(vars.OAuth[0].ApiSpecAuthRegistrationIdEnvName, "AUTH_NAME_REGISTRATION_ID");
      assert.equal(vars.ApiKey[0].ApiSpecAuthName, "authName2");
      assert.equal(vars.ApiKey[0].ApiSpecPath, "path/to/spec.yaml");
      assert.equal(vars.ApiKey[0].ApiSpecAuthRegistrationIdEnvName, "AUTH_NAME2_REGISTRATION_ID");
      assert.equal(vars.SafeProjectName, "Test");
      assert.equal(vars.SafeProjectNameLowerCase, "test");
    });

    it("template variables when contains auth with name not start with [A-Z]", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_TEST_TOOL: "false" });
      const vars = Generator.getDefaultVariables("Test", "Test", undefined, false, [
        {
          authName: "authName",
          openapiSpecPath: "path/to/spec.yaml",
          registrationIdEnvName: "*AUTH-NAME_REGISTRATION*ID",
          authType: "apiKey",
        },
      ]);
      assert.equal(vars.enableTestToolByDefault, "");
      assert.equal(vars.appName, "Test");
      assert.equal(vars.ApiKey[0].ApiSpecAuthName, "authName");
      assert.equal(vars.ApiKey[0].ApiSpecPath, "path/to/spec.yaml");
      assert.equal(
        vars.ApiKey[0].ApiSpecAuthRegistrationIdEnvName,
        "PREFIX__AUTH_NAME_REGISTRATION_ID"
      );
      assert.equal(vars.SafeProjectName, "Test");
      assert.equal(vars.SafeProjectNameLowerCase, "test");
    });

    it("generate templates from local when remote download processing fails", async () => {
      const mockFileName = "test.txt";
      const actionContext: ActionContext = {
        telemetryProps: {},
      };
      await buildFakeTemplateZip(templateName, mockFileName);

      sandbox.replace(templateConfig, "useLocalTemplate", true);
      sandbox.stub(folderUtils, "getTemplatesFolder").returns(tmpDir);
      sandbox.stub(ScaffoldRemoteTemplateAction, "run").throws(new Error("test"));

      const result = newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(context, inputs, tmpDir, actionContext)
        : await Generator.generateTemplate(context, tmpDir, templateName, language, actionContext);

      const isFallback = actionContext.telemetryProps?.fallback === "true";
      if (isFallback === false) {
        assert.fail("template should be generated by fallback");
      }

      if (!fs.existsSync(path.join(tmpDir, mockFileName))) {
        assert.fail("template creation failure");
      }
      assert.isTrue(result.isOk());
    });

    it("template from local when using local template tag", async () => {
      const mockFileName = "test.txt";
      const actionContext: ActionContext = {
        telemetryProps: {},
      };
      await buildFakeTemplateZip(templateName, mockFileName);

      sandbox.replace(templateConfig, "useLocalTemplate", true);
      sandbox.stub(folderUtils, "getTemplatesFolder").returns(tmpDir);

      const result = newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(context, inputs, tmpDir, actionContext)
        : await Generator.generateTemplate(context, tmpDir, templateName, language, actionContext);

      const isFallback = actionContext.telemetryProps?.fallback === "true";
      if (isFallback === true) {
        assert.fail("template should not be generated from remote to local");
      }

      if (!fs.existsSync(path.join(tmpDir, mockFileName))) {
        assert.fail("local template creation failure");
      }
      assert.isTrue(result.isOk());
    });

    it("template from local when local version is higher than git tag version", async () => {
      const mockFileName = "test.txt";
      const actionContext: ActionContext = {
        telemetryProps: {},
      };
      await buildFakeTemplateZip(templateName, mockFileName);

      sandbox.replace(templateConfig, "useLocalTemplate", false);
      sandbox.replace(templateConfig, "localVersion", "9.9.9");
      sandbox.replace(templateConfig, "version", "~3.0.0");
      const tagList = "1.0.0\n 2.0.0\n 2.1.0\n 3.0.0";
      sandbox.stub(axios, "get").resolves({ data: tagList, status: 200 } as AxiosResponse);
      sandbox.stub(folderUtils, "getTemplatesFolder").returns(tmpDir);
      sandbox
        .stub(generatorUtils, "getTemplateZipUrlByVersion")
        .resolves("fooUrl/templates@0.1.0/test.zip");

      const result = newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(context, inputs, tmpDir, actionContext)
        : await Generator.generateTemplate(context, tmpDir, templateName, language, actionContext);

      const isFallback = actionContext.telemetryProps?.fallback === "true";
      if (isFallback === true) {
        assert.fail("template should not be generated from remote to local");
      }

      if (!fs.existsSync(path.join(tmpDir, mockFileName))) {
        assert.fail("local template creation failure");
      }
      assert.isTrue(result.isOk());
    });

    it("template from downloading when local version is not higher than online version", async () => {
      const mockFileName = "test.txt";
      const zip = await buildFakeTemplateZip(templateName, mockFileName);
      const actionContext: ActionContext = {
        telemetryProps: {},
      };

      sandbox.replace(templateConfig, "useLocalTemplate", false);
      sandbox.replace(templateConfig, "localVersion", "0.1.0");
      sandbox.stub(folderUtils, "getTemplatesFolder").returns(tmpDir);
      sandbox.stub(generatorUtils, "getTemplateLatestVersion").resolves("0.1.1");
      sandbox.stub(generatorUtils, "fetchZipFromUrl").resolves(zip);

      const result = newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(context, inputs, tmpDir, actionContext)
        : await Generator.generateTemplate(context, tmpDir, templateName, language, actionContext);

      const isFallback = actionContext.telemetryProps?.fallback === "true";
      if (isFallback === true) {
        assert.fail("template should not be generated from remote to local");
      }

      if (!fs.existsSync(path.join(tmpDir, mockFileName))) {
        assert.fail("local template creation failure");
      }
      assert.isTrue(result.isOk());
    });

    it("template from downloading when TEAMSFX_TEMPLATE_PRERELEASE feature flag is set", async () => {
      const mockFileName = "test.txt";
      const zip = await buildFakeTemplateZip(templateName, mockFileName);
      const actionContext: ActionContext = {
        telemetryProps: {},
      };

      mockedEnvRestore = mockedEnv({
        TEAMSFX_TEMPLATE_PRERELEASE: "rc",
      });
      sandbox.replace(templateConfig, "useLocalTemplate", false);
      sandbox.replace(templateConfig, "localVersion", "0.1.0");
      sandbox.stub(folderUtils, "getTemplatesFolder").returns(tmpDir);
      sandbox.stub(generatorUtils, "getTemplateLatestVersion").resolves("0.1.1");
      sandbox.stub(generatorUtils, "fetchZipFromUrl").resolves(zip);

      const result = newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(context, inputs, tmpDir, actionContext)
        : await Generator.generateTemplate(context, tmpDir, templateName, language, actionContext);

      const isFallback = actionContext.telemetryProps?.fallback === "true";
      if (isFallback === true) {
        assert.fail("template should not be generated from remote to local");
      }

      if (!fs.existsSync(path.join(tmpDir, mockFileName))) {
        assert.fail("local template creation failure");
      }
      assert.isTrue(result.isOk());
    });

    it("telemetry contains correct template name", async () => {
      const actionContext: ActionContext = {
        telemetryProps: {},
      };

      sandbox.replace(TemplateActionSeq, "values", () => [] as any);
      newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(context, inputs, tmpDir, actionContext)
        : await Generator.generateTemplate(context, tmpDir, templateName, language, actionContext);

      assert.equal(actionContext.telemetryProps?.["template-name"], `${templateName}-${language}`);
    });

    it("telemetry contains correct template name when language undefined", async () => {
      const actionContext: ActionContext = {
        telemetryProps: {},
      };
      inputs[QuestionNames.ProgrammingLanguage] = undefined;

      sandbox.replace(TemplateActionSeq, "values", () => [] as any);
      newGeneratorFlag
        ? await new DefaultTemplateGenerator().run(context, inputs, tmpDir, actionContext)
        : await Generator.generateTemplate(context, tmpDir, templateName, undefined, actionContext);

      assert.equal(
        actionContext.telemetryProps?.["template-name"],
        `${templateName}-${commonTemplateName}`
      );
    });

    it("template variables when CEA enabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_CEA_ENABLED: "true" });
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test");
      assert.equal(vars.CEAEnabled, "true");
    });

    it("template variables when CEA disabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_CEA_ENABLED: "false" });
      const vars = newGeneratorFlag
        ? getTemplateReplaceMap(inputs)
        : Generator.getDefaultVariables("test");
      assert.equal(vars.CEAEnabled, "");
    });

    it("CEA works in M365 tag shows when CEA enabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_CEA_ENABLED: "true" });
      const descriptionAnswer = getLocalizedString(
        "core.createProjectQuestion.capability.customEngineAgent.description"
      );
      assert.equal(CapabilityOptions.customCopilotBasic().description, descriptionAnswer);
      assert.equal(CapabilityOptions.customCopilotRag().description, descriptionAnswer);
      assert.equal(CapabilityOptions.customCopilotAssistant().description, descriptionAnswer);
    });

    it("CEA works in M365 tag doesn't show when CEA disabled", async () => {
      sandbox.stub(process, "env").value({ TEAMSFX_CEA_ENABLED: "false" });
      assert.equal(CapabilityOptions.customCopilotBasic().description, undefined);
      assert.equal(CapabilityOptions.customCopilotRag().description, undefined);
      assert.equal(CapabilityOptions.customCopilotAssistant().description, undefined);
    });
  });
});

describe("Generate sample using download directory", () => {
  const tmpDir = path.join(__dirname, "tmp");
  const sandbox = createSandbox();
  let mockedEnvRestore = mockedEnv({});
  const tools = new MockTools();
  setTools(tools);
  const ctx = createContext();
  beforeEach(async () => {
    mockedEnvRestore = mockedEnv({
      DOWNLOAD_DIRECTORY: "true",
    });
    sandbox.stub(generatorUtils, "getSampleInfoFromName").resolves(mockedSampleInfo);
  });

  afterEach(async () => {
    sandbox.restore();
    mockedEnvRestore();
    if (await fs.pathExists(tmpDir)) {
      await fs.remove(tmpDir);
    }
  });

  it("generate sample using download directory", async () => {
    const axiosStub = sandbox.stub(axios, "get");
    const sampleName = "test";
    const mockFileName = "test.txt";
    const mockFileData = "test data";
    const fileInfo = [{ type: "file", path: `${sampleName}/${mockFileName}` }];
    axiosStub.onFirstCall().resolves({ status: 200, data: { tree: fileInfo } });
    axiosStub.onSecondCall().resolves({ status: 200, data: mockFileData });
    const result = await Generator.generateSample(ctx, tmpDir, "test");
    assert.isTrue(result.isOk());
  });

  it("download directory throw api limit error", async () => {
    const axiosStub = sandbox.stub(axios, "get");
    const error = new Error("Network error");
    (error as any).isAxiosError = true;
    (error as any).response = {
      status: 403,
      headers: {
        "x-ratelimit-remaining": "0",
      },
    };
    axiosStub.onSecondCall().rejects(error);
    const result = await Generator.generateSample(ctx, tmpDir, "test");
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "DownloadSampleApiLimitError");
    }
  });

  it("download directory throw network error", async () => {
    const axiosStub = sandbox.stub(axios, "get");
    axiosStub.onFirstCall().resolves({ status: 502 });
    axiosStub.onSecondCall().resolves({ status: 502 });
    const result = await Generator.generateSample(ctx, tmpDir, "test");
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "DownloadSampleNetworkError");
    }
  });

  it("throw error if one file download failed", async () => {
    const axiosStub = sandbox.stub(axios, "get");
    const sampleName = "test";
    const mockFileName = "test.txt";
    const mockFileData = "test data";
    const fileInfo = [
      { type: "file", path: `${sampleName}/${mockFileName}` },
      { type: "file", path: `${sampleName}/${mockFileName}_1` },
      { type: "file", path: `${sampleName}/${mockFileName}_2` },
    ];
    axiosStub.onCall(0).resolves({ status: 200, data: { tree: fileInfo } });
    axiosStub.onCall(1).resolves({ status: 200, data: mockFileData });
    axiosStub.onCall(2).resolves({ status: 200, data: mockFileData });
    axiosStub.onCall(3).resolves({ status: 502 });
    axiosStub.onCall(4).resolves({ status: 502 });
    const result = await Generator.generateSample(ctx, tmpDir, "test");
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "DownloadSampleNetworkError");
    }
    assert.isFalse(await fs.pathExists(path.join(tmpDir, sampleName)));
  });

  it("clean up if downloading failed", async () => {
    const rmStub = sandbox.stub(fs, "remove").resolves();
    const existsStub = sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(generatorUtils, "downloadDirectory").rejects();
    const result = await Generator.generateSample(ctx, tmpDir, "test");
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "DownloadSampleNetworkError");
    }
    assert.isTrue(rmStub.calledOnce);
    assert.isTrue(existsStub.calledOnce);
  });
});
