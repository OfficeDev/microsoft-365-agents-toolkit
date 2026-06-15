import * as vscode from "vscode";
import fs from "fs-extra";
import * as chai from "chai";
import * as globalVariables from "../../src/globalVariables";
import * as extTelemetryEvents from "../../src/telemetry/extTelemetryEvents";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { PanelType } from "../../src/controls/PanelType";
import { TreatmentVariableValue } from "../../src/exp/treatmentVariables";
import { WebviewPanel } from "../../src/controls/webviewPanel";
import { openReadMeHandler, openSampleReadmeHandler } from "../../src/handlers/readmeHandlers";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";

describe("readmeHandlers", () => {
  describe("openReadMeHandler", () => {
    it("Happy Path", async () => {
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "isTeamsFxProject", true);
      const executeCommands = vi.spyOn(vscode.commands, "executeCommand");
      vi.spyOn(vscode.workspace, "workspaceFolders").value([
        { uri: { fsPath: "readmeTestFolder" } },
      ]);
      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      const openTextDocumentStub = vi
        .spyOn(vscode.workspace, "openTextDocument")
        .mockResolvedValue({} as any as vscode.TextDocument);

      await openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

      chai.assert.isTrue(openTextDocumentStub.calledOnce);
      chai.assert.isTrue(executeCommands.calledOnce);
    });

    it("Create Project", async () => {
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "isTeamsFxProject", false);
      mockValue(globalVariables, "core", undefined);
      const showMessageStub = vi
        .spyOn(vscode.window, "showInformationMessage")
        .mockImplementation(
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
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "isTeamsFxProject", false);
      mockValue(globalVariables, "core", undefined);
      const executeCommandStub = vi.spyOn(vscode.commands, "executeCommand");
      const showMessageStub = vi
        .spyOn(vscode.window, "showInformationMessage")
        .mockImplementation(
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
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "isTeamsFxProject", true);
      vi.spyOn(vscode.workspace, "workspaceFolders").value([
        { uri: { fsPath: "readmeTestFolder" } },
      ]);
      mockValue(TreatmentVariableValue, "inProductDoc", true);
      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockResolvedValue(
        Buffer.from("## Get Started with the Notification bot")
      );
      const createOrShow = vi
        .spyOn(WebviewPanel, "createOrShow")
        .mockImplementation(() => undefined);

      await openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

      expect(createOrShow).toHaveBeenCalledTimes(1);
      expect(createOrShow).toHaveBeenCalledWith(PanelType.FunctionBasedNotificationBotReadme);
    });

    it("Express Notification Bot Template", async () => {
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "isTeamsFxProject", true);
      vi.spyOn(vscode.workspace, "workspaceFolders").value([
        { uri: { fsPath: "readmeTestFolder" } },
      ]);
      mockValue(TreatmentVariableValue, "inProductDoc", true);
      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      vi.spyOn(fs, "readFile").mockResolvedValue(
        Buffer.from("## Get Started with the Notification bot express")
      );
      const createOrShow = vi
        .spyOn(WebviewPanel, "createOrShow")
        .mockImplementation(() => undefined);

      await openReadMeHandler([extTelemetryEvents.TelemetryTriggerFrom.Auto]);

      expect(createOrShow).toHaveBeenCalledTimes(1);
      expect(createOrShow).toHaveBeenCalledWith(PanelType.ExpressServerNotificationBotReadme);
    });
  });

  describe("openSampleReadmeHandler", () => {
    it("Trigger from Walkthrough", async () => {
      mockValue(vscode.workspace, "workspaceFolders", [{ uri: vscode.Uri.file("test") }]);
      vi.spyOn(vscode.workspace, "openTextDocument");
      const executeCommandStub = vi.spyOn(vscode.commands, "executeCommand");

      await openSampleReadmeHandler(["WalkThrough"]);

      chai.assert.isTrue(executeCommandStub.calledOnce);
    });
  });
});
