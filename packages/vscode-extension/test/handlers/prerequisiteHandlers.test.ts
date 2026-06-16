/**
 * @author HuihuiWu-Microsoft <73154171+HuihuiWu-Microsoft@users.noreply.github.com>
 */
import { SystemError, err } from "@microsoft/teamsfx-api";
import { DepsManager, DepsType } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import path from "path";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as getStartedChecker from "../../src/debug/depsChecker/getStartedChecker";
import * as globalVariables from "../../src/globalVariables";
import {
  getDotnetPathHandler,
  getPathDelimiterHandler,
  installAdaptiveCardExt,
  validateGetStartedPrerequisitesHandler,
} from "../../src/handlers/prerequisiteHandlers";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";

describe("prerequisiteHandlers", () => {
  describe("getDotnetPathHandler", async () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("dotnet is installed", async () => {
      sandbox.stub(DepsManager.prototype, "getStatus").resolves([
        {
          name: ".NET Core SDK",
          type: DepsType.Dotnet,
          isInstalled: true,
          command: "",
          details: {
            isLinuxSupported: false,
            installVersion: "",
            supportedVersions: [],
            binFolders: ["dotnet-bin-folder/dotnet"],
          },
        },
      ]);

      const dotnetPath = await getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}dotnet-bin-folder${path.delimiter}`);
    });

    it("dotnet is not installed", async () => {
      sandbox.stub(DepsManager.prototype, "getStatus").resolves([
        {
          name: ".NET Core SDK",
          type: DepsType.Dotnet,
          isInstalled: false,
          command: "",
          details: {
            isLinuxSupported: false,
            installVersion: "",
            supportedVersions: [],
            binFolders: undefined,
          },
        },
      ]);

      const dotnetPath = await getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}`);
    });

    it("failed to get dotnet path", async () => {
      sandbox.stub(DepsManager.prototype, "getStatus").rejects(new Error("failed to get status"));
      const dotnetPath = await getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}`);
    });
  });

  describe("getPathDelimiterHandler", () => {
    it("happy path", async () => {
      const actualPath = await getPathDelimiterHandler();
      chai.assert.equal(actualPath, path.delimiter);
    });
  });

  describe("validateGetStartedPrerequisitesHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("error", async () => {
      const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox
        .stub(getStartedChecker, "checkPrerequisitesForGetStarted")
        .resolves(err(new SystemError("test", "test", "test")));

      const result = await validateGetStartedPrerequisitesHandler();

      chai.assert.isTrue(sendTelemetryStub.called);
      chai.assert.isTrue(result.isErr());
    });
  });

  describe("installAdaptiveCardExt", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Happy path()", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

      sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .resolves("Install" as unknown as vscode.MessageItem);

      await installAdaptiveCardExt();

      chai.assert.isTrue(executeCommandStub.calledOnce);
    });
  });
});
