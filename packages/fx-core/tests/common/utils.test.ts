import "mocha";
import chai from "chai";
import { convertToAlphanumericOnly } from "../../src/common/stringUtils";
import { jsonUtils } from "../../src/common/jsonUtils";
import {
  FileNotFoundError,
  JSONSyntaxError,
  ReadFileError,
  WriteFileError,
} from "../../src/error/common";
import sinon from "sinon";
import fs from "fs-extra";
import { isYamlFileName } from "../../src/common/versionMetadata";

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

describe("versionMedadata", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("isYamlFileName - true", () => {
    const res = isYamlFileName("m365agents.local.yml");
    chai.assert.isTrue(res);
  });
  it("isYamlFileName - false", () => {
    const res = isYamlFileName("abc.local.yml");
    chai.assert.isTrue(res);
  });
  it("isYamlFileNameV3 - true", () => {
    const res = isYamlFileName("teamsapp.local.yml");
    chai.assert.isFalse(res);
  });
  it("isYamlFileNameV3 - false", () => {
    const res = isYamlFileName("m365agents.local.yml");
    chai.assert.isFalse(res);
  });
  it("isYamlFileNameV4 - true", () => {
    const res = isYamlFileName("m365agents.local.yml");
    chai.assert.isFalse(res);
  });
  it("isYamlFileNameV4 - false", () => {
    const res = isYamlFileName("teamsapp.yml");
    chai.assert.isFalse(res);
  });
});
