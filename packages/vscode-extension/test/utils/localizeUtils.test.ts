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

    it("should log info if no string file found for current non-English locale", () => {
      sandbox.stub(navigator, "language").value(undefined);
      sandbox.stub(process, "env").value({ VSCODE_NLS_CONFIG: '{ "locale": "zh-cn" }' });
      sandbox.stub(fs, "pathExistsSync").callsFake((directory: string) => {
        if (directory.includes("package.nls.json")) {
          return true;
        }
        return false;
      });
      sandbox.stub(fs, "readJsonSync").returns({});
      sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
      const vscodeLogStub = sandbox.stub(VsCodeLogInstance, "info");
      _resetCollections();

      loadLocalizedStrings();

      chai.expect(vscodeLogStub.called).to.be.true;
    });

    it("should not log for English variant locales (en-gb, en-au, etc.)", () => {
      sandbox.stub(navigator, "language").value(undefined);
      sandbox.stub(process, "env").value({ VSCODE_NLS_CONFIG: '{ "locale": "en-gb" }' });
      sandbox.stub(fs, "pathExistsSync").callsFake((directory: string) => {
        if (directory.includes("package.nls.json")) {
          return true;
        }
        return false;
      });
      sandbox.stub(fs, "readJsonSync").returns({});
      sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
      const errorStub = sandbox.stub(VsCodeLogInstance, "error");
      const infoStub = sandbox.stub(VsCodeLogInstance, "info");
      _resetCollections();

      loadLocalizedStrings();

      // Should not log error or info about missing locale file for English variants
      chai.expect(errorStub.called).to.be.false;
      // info is called once for "Current VS Code locale is: en-gb" but not for missing file
      chai
        .expect(
          infoStub.calledWithMatch(/No localized strings file found for locale/)
        )
        .to.be.false;
    });
  });

  describe("parseLocale", () => {
    it("should return current locale", () => {
      sandbox.stub(navigator, "language").value(undefined);
      sandbox.stub(process, "env").value({ VSCODE_NLS_CONFIG: '{ "locale": "zh-cn" }' });
      const locale = parseLocale();

      chai.expect(locale).to.equal("zh-cn");
    });
  });
});
