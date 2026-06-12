import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as systemEnvUtils from "../../src/utils/systemEnvUtils";
import { Inputs, Platform, VsCodeEnv } from "@microsoft/teamsfx-api";

describe("SystemEnvUtils", () => {
  describe("detectVsCodeEnv()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("locally run", () => {
      const expectedResult = {
        extensionKind: vscode.ExtensionKind.UI,
        id: "",
        extensionUri: vscode.Uri.file(""),
        extensionPath: "",
        isActive: true,
        packageJSON: {},
        exports: undefined,
        activate: sandbox.spy(),
      };
      const getExtension = sandbox
        .stub(vscode.extensions, "getExtension")
        .callsFake((name: string) => {
          return expectedResult;
        });

      chai.expect(systemEnvUtils.detectVsCodeEnv()).equals(VsCodeEnv.local);
      getExtension.restore();
    });

    it("Remotely run", () => {
      const expectedResult = {
        extensionKind: vscode.ExtensionKind.Workspace,
        id: "",
        extensionUri: vscode.Uri.file(""),
        extensionPath: "",
        isActive: true,
        packageJSON: {},
        exports: undefined,
        activate: sandbox.spy(),
      };
      const getExtension = sandbox
        .stub(vscode.extensions, "getExtension")
        .callsFake((name: string) => {
          return expectedResult;
        });

      chai
        .expect(systemEnvUtils.detectVsCodeEnv())
        .oneOf([VsCodeEnv.remote, VsCodeEnv.codespaceVsCode, VsCodeEnv.codespaceBrowser]);
      getExtension.restore();
    });
  });

  describe("getSystemInputs()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("getSystemInputs()", () => {
      const input: Inputs = systemEnvUtils.getSystemInputs();

      chai.expect(input.platform).equals(Platform.VSCode);
    });
  });
});
