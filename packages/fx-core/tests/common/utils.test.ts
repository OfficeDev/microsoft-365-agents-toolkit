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
