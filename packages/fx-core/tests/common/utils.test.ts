import chai from "chai";
import fs from "fs-extra";
import "mocha";
import sinon from "sinon";
import { jsonUtils } from "../../src/common/jsonUtils";
import { convertToAlphanumericOnly } from "../../src/common/stringUtils";
import {
  isYamlFileName,
  isYamlFileNameV3,
  isYamlFileNameV4,
} from "../../src/common/versionMetadata";
import {
  FileNotFoundError,
  JSONSyntaxError,
  ReadFileError,
  WriteFileError,
} from "../../src/error/common";

describe("convert to valid AppName in ProjectSetting", () => {
  it("convert app name", () => {
    const appName = "app.123";
    const expected = "app123";
    const projectSettingsName = convertToAlphanumericOnly(appName);

    chai.assert.equal(projectSettingsName, expected);
  });

  it("convert app name", () => {
    const appName = "app.1@@2！3";
    const expected = "app123";
    const projectSettingsName = convertToAlphanumericOnly(appName);

    chai.assert.equal(projectSettingsName, expected);
  });
});

describe("JSONUtils", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("readJSONFileSync JSONSyntaxError", () => {
    sandbox.stub(fs, "readJSONSync").throws(new SyntaxError());
    const res = jsonUtils.readJSONFileSync(".");
    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.isTrue(res.error instanceof JSONSyntaxError);
    }
  });
  it("readJSONFileSync ReadFileError", () => {
    sandbox.stub(fs, "readJSONSync").throws(new Error());
    const res = jsonUtils.readJSONFileSync(".");
    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.isTrue(res.error instanceof ReadFileError);
    }
  });
  it("readJSONFileSync FileNotFoundError", () => {
    sandbox.stub(fs, "readJSONSync").throws(new Error("no such file or directory"));
    const res = jsonUtils.readJSONFileSync(".");
    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.isTrue(res.error instanceof FileNotFoundError);
    }
  });
});

describe("Errors", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("WriteFileError", () => {
    const error = new WriteFileError(new Error("write file error"), "common");
    chai.assert(error.name === "WriteFileError");
  });
  it("WriteFileError", () => {
    const error = new WriteFileError(new Error(""), "common");
    chai.assert(error.name === "WriteFileError");
  });
});

describe("versionMetadata", () => {
  it("isYamlFileName - true", () => {
    const res = isYamlFileName("m365agents.local.yml");
    chai.assert.isTrue(res);
  });
  it("isYamlFileName - false", () => {
    const res = isYamlFileName("abc.local.yml");
    chai.assert.isFalse(res);
  });
  it("isYamlFileNameV3 - true", () => {
    const res = isYamlFileNameV3("teamsapp.local.yml");
    chai.assert.isTrue(res);
  });
  it("isYamlFileNameV3 - false", () => {
    const res = isYamlFileNameV3("m365agents.local.yml");
    chai.assert.isFalse(res);
  });
  it("isYamlFileNameV4 - true", () => {
    const res = isYamlFileNameV4("m365agents.local.yml");
    chai.assert.isTrue(res);
  });
  it("isYamlFileNameV4 - false", () => {
    const res = isYamlFileNameV4("teamsapp.yml");
    chai.assert.isFalse(res);
  });
});
