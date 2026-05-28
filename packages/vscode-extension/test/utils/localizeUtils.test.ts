import * as chai from "chai";
import fs from "fs-extra";
import sinon from "ts-sinon";
import VsCodeLogInstance from "../../src/commonlib/log";
import * as globalVariables from "../../src/globalVariables";
import {
  _resetCollections,
  loadLocalizedStrings,
  parseLocale,
} from "../../src/utils/localizeUtils";

afterEach(() => {
  sinon.restore();
});

describe("localizeUtils", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    _resetCollections();
    sandbox.restore();
  });

  describe("loadLocalizedStrings", () => {
    it("should log error if no default string collection", () => {
      sandbox.stub(fs, "pathExistsSync").callsFake((directory: string) => {
        if (directory.includes("package.nls.json")) {
          return false;
        }
        return true;
      });
      sandbox.stub(fs, "readJsonSync").returns({});
      sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
      const vscodeLogStub = sandbox.stub(VsCodeLogInstance, "error");
      _resetCollections();

      loadLocalizedStrings();

      chai.expect(vscodeLogStub.calledOnce).to.be.true;
    });

    it("should log error if no string file found for current locale", () => {
      sandbox.stub(process, "env").value({ VSCODE_NLS_CONFIG: '{ "locale": "zh-cn" }' });
      sandbox.stub(fs, "pathExistsSync").callsFake((directory: string) => {
        if (directory.includes("package.nls.json")) {
          return true;
        }
        return false;
      });
      sandbox.stub(fs, "readJsonSync").returns({});
      sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
      const vscodeLogStub = sandbox.stub(VsCodeLogInstance, "error");
      _resetCollections();

      loadLocalizedStrings();

      chai.expect(vscodeLogStub.calledOnce).to.be.true;
    });
  });

  describe("parseLocale", () => {
    it("should return current locale", () => {
      sandbox.stub(process, "env").value({ VSCODE_NLS_CONFIG: '{ "locale": "zh-cn" }' });

      const locale = parseLocale();

      chai.expect(locale).to.equal("zh-cn");
    });
  });
});
