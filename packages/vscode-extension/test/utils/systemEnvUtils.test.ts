import * as chai from "chai";
import * as vscode from "vscode";
import * as systemEnvUtils from "../../src/utils/systemEnvUtils";
import { Inputs, Platform, VsCodeEnv } from "@microsoft/teamsfx-api";
import { vi } from "vitest";

describe("SystemEnvUtils", () => {
  describe("detectVsCodeEnv()", () => {
    it("locally run", () => {
      const expectedResult = {
        extensionKind: vscode.ExtensionKind.UI,
        id: "",
        extensionUri: vscode.Uri.file(""),
        extensionPath: "",
        isActive: true,
        packageJSON: {},
        exports: undefined,
        activate: vi.fn(),
      };
      const getExtension = vi
        .spyOn(vscode.extensions, "getExtension")
        .mockImplementation((name: string) => {
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
        activate: vi.fn(),
      };
      const getExtension = vi
        .spyOn(vscode.extensions, "getExtension")
        .mockImplementation((name: string) => {
          return expectedResult;
        });

      chai
        .expect(systemEnvUtils.detectVsCodeEnv())
        .oneOf([VsCodeEnv.remote, VsCodeEnv.codespaceVsCode, VsCodeEnv.codespaceBrowser]);
      getExtension.restore();
    });
  });

  describe("getSystemInputs()", () => {
    it("getSystemInputs()", () => {
      const input: Inputs = systemEnvUtils.getSystemInputs();

      chai.expect(input.platform).equals(Platform.VSCode);
    });
  });
});
