import * as vscode from "vscode";
import * as sinon from "sinon";
import fs from "fs-extra";
import * as chai from "chai";
import { featureFlagManager, FeatureFlags } from "@microsoft/teamsfx-core";
import * as globalVariables from "../../src/globalVariables";
import * as extTelemetryEvents from "../../src/telemetry/extTelemetryEvents";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { PanelType } from "../../src/controls/PanelType";
import { TreatmentVariableValue } from "../../src/exp/treatmentVariables";
import { WebviewPanel } from "../../src/controls/webviewPanel";
import {
  openReadMeHandler,
  openSampleReadmeHandler,
  openWorkspaceMCPConfigHandler,
} from "../../src/handlers/readmeHandlers";

describe("readmeHandlers", () => {
  describe("openReadMeHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Happy Path", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "isTeamsFxProject").value(true);
      const executeCommands = sandbox.stub(vscode.commands, "executeCommand");
      sandbox
        .stub(vscode.workspace, "workspaceFolders")
        .value([{ uri: { fsPath: "readmeTestFolder" } }]);
      sandbox.stub(fs, "pathExists").resolves(true);
      const openTextDocumentStub = sandbox
        .stub(vscode.workspace, "openTextDocument")
        .resolves({} as any as vscode.TextDocument);

      await openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

      chai.assert.isTrue(openTextDocumentStub.calledOnce);
      chai.assert.isTrue(executeCommands.calledOnce);
    });

    it("Create Project", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "isTeamsFxProject").value(false);
      sandbox.stub(globalVariables, "core").value(undefined);
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .callsFake(
          (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
            return Promise.resolve({
              title: "Yes",
              run: (options as any).run,
            } as vscode.MessageItem);
          }
        );
      await openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

      chai.assert.isTrue(showMessageStub.calledOnce);
    });

    it("Open Folder", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "isTeamsFxProject").value(false);
      sandbox.stub(globalVariables, "core").value(undefined);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .callsFake(
          (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
            return Promise.resolve({
              title: "Yes",
              run: (items[0] as any).run,
            } as vscode.MessageItem);
          }
        );
      await openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

      chai.assert.isTrue(executeCommandStub.calledOnce);
    });

    it("Function Notification Bot Template", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "isTeamsFxProject").value(true);
      sandbox
        .stub(vscode.workspace, "workspaceFolders")
        .value([{ uri: { fsPath: "readmeTestFolder" } }]);
      sandbox.stub(TreatmentVariableValue, "inProductDoc").value(true);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox
        .stub(fs, "readFile")
        .resolves(Buffer.from("## Get Started with the Notification bot"));
      const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");

      await openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

      sandbox.assert.calledOnceWithExactly(
        createOrShow,
        PanelType.FunctionBasedNotificationBotReadme
      );
    });

    it("Express Notification Bot Template", async () => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "isTeamsFxProject").value(true);
      sandbox
        .stub(vscode.workspace, "workspaceFolders")
        .value([{ uri: { fsPath: "readmeTestFolder" } }]);
      sandbox.stub(TreatmentVariableValue, "inProductDoc").value(true);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox
        .stub(fs, "readFile")
        .resolves(Buffer.from("## Get Started with the Notification bot express"));
      const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");

      await openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

      sandbox.assert.calledOnceWithExactly(
        createOrShow,
        PanelType.ExpressServerNotificationBotReadme
      );
    });
  });

  describe("openSampleReadmeHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Trigger from Walkthrough", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox.stub(vscode.workspace, "openTextDocument");
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

      await openSampleReadmeHandler(["WalkThrough"]);

      chai.assert.isTrue(executeCommandStub.calledOnce);
    });
  });

  describe("openWorkspaceMCPConfigHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("no workspace folder - returns ok without opening", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value(undefined);
      const openTextDocumentStub = sandbox.stub(vscode.workspace, "openTextDocument");

      const res = await openWorkspaceMCPConfigHandler();

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(openTextDocumentStub.notCalled);
    });

    it("config file missing - returns ok without opening", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox.stub(fs, "pathExists").resolves(false);
      const openTextDocumentStub = sandbox.stub(vscode.workspace, "openTextDocument");

      const res = await openWorkspaceMCPConfigHandler();

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(openTextDocumentStub.notCalled);
    });

    it("DT flag off - shows Fetch Action notification and opens mcp.json", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const openTextDocumentStub = sandbox
        .stub(vscode.workspace, "openTextDocument")
        .resolves({} as vscode.TextDocument);
      sandbox.stub(vscode.window, "showTextDocument").resolves();
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .resolves("Fetch Action" as any);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

      const res = await openWorkspaceMCPConfigHandler();

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(showMessageStub.calledOnce);
      chai.assert.isTrue(openTextDocumentStub.calledOnce);
      // Fetch Action selection wires the updateActionWithMCP command.
      await Promise.resolve();
      chai.assert.isTrue(executeCommandStub.calledWith("fx-extension.updateActionWithMCP"));
    });

    it("DT flag on - opens ai-plugin.json without Fetch Action notification", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox
        .stub(featureFlagManager, "getBooleanValue")
        .callsFake((flag) => flag === FeatureFlags.MCPForDADT);
      const openTextDocumentStub = sandbox
        .stub(vscode.workspace, "openTextDocument")
        .resolves({} as vscode.TextDocument);
      sandbox.stub(vscode.window, "showTextDocument").resolves();
      const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage");

      const res = await openWorkspaceMCPConfigHandler();

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(showMessageStub.notCalled);
      chai.assert.isTrue(openTextDocumentStub.calledOnce);
      const openedUri = openTextDocumentStub.firstCall.args[0] as vscode.Uri;
      chai.assert.isTrue(openedUri.fsPath.endsWith("ai-plugin.json"));
    });

    it("DT flag on but ai-plugin.json missing - falls back to mcp.json", async () => {
      sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
      sandbox
        .stub(fs, "pathExists")
        .callsFake((p: string) => Promise.resolve(p.endsWith("mcp.json")) as any);
      sandbox
        .stub(featureFlagManager, "getBooleanValue")
        .callsFake((flag) => flag === FeatureFlags.MCPForDADT);
      const openTextDocumentStub = sandbox
        .stub(vscode.workspace, "openTextDocument")
        .resolves({} as vscode.TextDocument);
      sandbox.stub(vscode.window, "showTextDocument").resolves();

      const res = await openWorkspaceMCPConfigHandler();

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(openTextDocumentStub.calledOnce);
      const openedUri = openTextDocumentStub.firstCall.args[0] as vscode.Uri;
      chai.assert.isTrue(openedUri.fsPath.endsWith("mcp.json"));
    });
  });
});
