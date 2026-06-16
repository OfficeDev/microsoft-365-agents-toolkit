import { ok, TeamsAppManifest } from "@microsoft/teamsfx-api";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import * as chai from "chai";
import fs from "fs-extra";
import * as path from "path";
import * as vscode from "vscode";
import { PanelType } from "../../src/controls/PanelType";
import { WebviewPanel } from "../../src/controls/webviewPanel";
import * as globalVariables from "../../src/globalVariables";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";
import {
  controlHandlersOps,
  openFolderHandler,
  openLifecycleTreeview,
  openSamplesHandler,
  openWelcomeHandler,
  saveTextDocumentHandler,
  selectWalkthrough,
} from "../../src/handlers/controlHandlers";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryUpdateAppReason,
} from "../../src/telemetry/extTelemetryEvents";
import { getDefaultString } from "../../src/utils/localizeUtils";
import * as teamsfxCore from "@microsoft/teamsfx-core";

const controlHandlersDeps = controlHandlersOps;

describe("Control Handlers", () => {
  describe("openWelcomeHandler", () => {
    it("opens intelligent app walkthrough for API plugin apps", async () => {
      vi.spyOn(teamsfxCore.featureFlagManager, "getBooleanValue").mockReturnValue(false);
      vi.spyOn(teamsfxCore.manifestUtils, "readAppManifest").mockResolvedValue(
        ok({} as TeamsAppManifest)
      );
      vi.spyOn(teamsfxCore.manifestUtils, "getCapabilities").mockReturnValue(["copilotGpt"]);
      mockValue(globalVariables, "workspaceUri", { fsPath: "/test" });
      const executeCommands = vi.spyOn(vscode.commands, "executeCommand");
      const sendTelemetryEvent = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      await openWelcomeHandler("invalidArgs");

      expect(executeCommands).toHaveBeenCalledTimes(1);
      expect(executeCommands).toHaveBeenCalledWith(
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
      );
    });

    it("opens intelligent app walkthrough with chat for API plugin apps", async () => {
      vi.spyOn(teamsfxCore.featureFlagManager, "getBooleanValue").mockReturnValue(true);
      vi.spyOn(teamsfxCore.manifestUtils, "readAppManifest").mockResolvedValue(
        ok({} as TeamsAppManifest)
      );
      vi.spyOn(teamsfxCore.manifestUtils, "getCapabilities").mockReturnValue(["copilotGpt"]);
      mockValue(globalVariables, "workspaceUri", { fsPath: "/test" });
      const executeCommands = vi.spyOn(vscode.commands, "executeCommand");
      const sendTelemetryEvent = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      await openWelcomeHandler("invalidArgs");

      expect(executeCommands).toHaveBeenCalledTimes(1);
      expect(executeCommands).toHaveBeenCalledWith(
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
      );
    });

    it("opens intelligent app walkthrough for JS/TS custom engine copilot apps", async () => {
      vi.spyOn(teamsfxCore.featureFlagManager, "getBooleanValue").mockReturnValue(false);
      vi.spyOn(teamsfxCore.manifestUtils, "readAppManifest").mockResolvedValue(
        ok({} as TeamsAppManifest)
      );
      vi.spyOn(teamsfxCore.manifestUtils, "getCapabilities").mockReturnValue(["bot"]);
      mockValue(globalVariables, "workspaceUri", { fsPath: "/test" });
      vi.spyOn(fs, "pathExists").mockImplementation(async (path: string) => {
        return path.includes("package.json");
      });
      vi.spyOn(fs, "readFile").mockResolvedValue(Buffer.from('"@microsoft/teams-ai"'));
      const executeCommands = vi.spyOn(vscode.commands, "executeCommand");
      const sendTelemetryEvent = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      await openWelcomeHandler();

      expect(executeCommands).toHaveBeenCalledTimes(1);
      expect(executeCommands).toHaveBeenCalledWith(
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
      );
    });

    it("opens intelligent app walkthrough for python custom engine copilot apps", async () => {
      vi.spyOn(teamsfxCore.featureFlagManager, "getBooleanValue").mockReturnValue(false);
      vi.spyOn(teamsfxCore.manifestUtils, "readAppManifest").mockResolvedValue(
        ok({} as TeamsAppManifest)
      );
      vi.spyOn(teamsfxCore.manifestUtils, "getCapabilities").mockReturnValue(["bot"]);
      mockValue(globalVariables, "workspaceUri", { fsPath: "/test" });
      vi.spyOn(fs, "pathExists").mockImplementation(async (path: string) => {
        return path.includes("requirements.txt");
      });
      vi.spyOn(fs, "readFile").mockResolvedValue(Buffer.from("teams-ai"));
      const executeCommands = vi.spyOn(vscode.commands, "executeCommand");
      const sendTelemetryEvent = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      await openWelcomeHandler();

      expect(executeCommands).toHaveBeenCalledTimes(1);
      expect(executeCommands).toHaveBeenCalledWith(
        "workbench.action.openWalkthrough",
        "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
      );
    });
  });

  describe("openSamplesHandler", () => {
    it("openSamplesHandler", async () => {
      const createOrShow = vi
        .spyOn(WebviewPanel, "createOrShow")
        .mockImplementation(() => undefined);
      const sendTelemetryEvent = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      await openSamplesHandler();

      expect(createOrShow).toHaveBeenCalledTimes(1);
      expect(createOrShow).toHaveBeenCalledWith(PanelType.SampleGallery, []);
    });
  });

  describe("openFolderHandler", () => {
    it("empty args", async () => {
      const sendTelemetryStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      const result = await openFolderHandler();

      chai.assert.isTrue(sendTelemetryStub.called);
      chai.assert.isTrue(result.isOk());
    });

    it("happy path", async () => {
      const sendTelemetryStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      const openFolderInExplorerStub = vi
        .spyOn(controlHandlersDeps, "openFolderInExplorer")
        .mockImplementation(() => {
          return;
        });

      const result = await openFolderHandler("file://path/to/folder");

      const expectedPath = "/path/to/folder".split("/").join(path.sep);
      chai.assert.isTrue(sendTelemetryStub.called);
      chai.assert.isTrue(openFolderInExplorerStub.calledOnceWith(expectedPath));
      chai.assert.isTrue(result.isOk());
    });
  });

  describe("saveTextDocumentHandler", () => {
    it("non valid project", () => {
      const isValidProjectStub = vi
        .spyOn(controlHandlersDeps, "isValidProject")
        .mockReturnValue(false);
      mockValue(globalVariables, "workspaceUri", { fsPath: "/path/to/workspace" });

      saveTextDocumentHandler({ document: {} } as any);

      chai.assert.isTrue(isValidProjectStub.calledOnceWith("/path/to/workspace"));
    });

    it("manual save reason", () => {
      const isValidProjectStub = vi
        .spyOn(controlHandlersDeps, "isValidProject")
        .mockReturnValue(true);
      const sendTelemetryEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "workspaceUri", { fsPath: "/path/to/workspace" });

      saveTextDocumentHandler({
        document: { fileName: "/dirname/fileName" },
        reason: vscode.TextDocumentSaveReason.Manual,
      } as vscode.TextDocumentWillSaveEvent);

      chai.assert.isTrue(isValidProjectStub.calledTwice);
      chai.assert.equal(isValidProjectStub.getCall(0).args[0], "/path/to/workspace");
      chai.assert.equal(isValidProjectStub.getCall(1).args[0], "/dirname");
      chai.assert.equal(sendTelemetryEventStub.getCall(0).args[0], TelemetryEvent.UpdateTeamsApp);
      chai.assert.equal(
        sendTelemetryEventStub.getCall(0).args[1][TelemetryProperty.UpdateTeamsAppReason],
        TelemetryUpdateAppReason.Manual
      );
    });

    it("after delay save reason", () => {
      const isValidProjectStub = vi
        .spyOn(controlHandlersDeps, "isValidProject")
        .mockReturnValue(true);
      const sendTelemetryEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "workspaceUri", { fsPath: "/path/to/workspace" });

      saveTextDocumentHandler({
        document: { fileName: "/dirname/fileName" },
        reason: vscode.TextDocumentSaveReason.AfterDelay,
      } as vscode.TextDocumentWillSaveEvent);

      chai.assert.isTrue(isValidProjectStub.calledTwice);
      chai.assert.equal(isValidProjectStub.getCall(0).args[0], "/path/to/workspace");
      chai.assert.equal(isValidProjectStub.getCall(1).args[0], "/dirname");
      chai.assert.equal(sendTelemetryEventStub.getCall(0).args[0], TelemetryEvent.UpdateTeamsApp);
      chai.assert.equal(
        sendTelemetryEventStub.getCall(0).args[1][TelemetryProperty.UpdateTeamsAppReason],
        TelemetryUpdateAppReason.AfterDelay
      );
    });

    it("focus out save reason", () => {
      const dirname = "/dirname";
      const parentDir = path.join(dirname, "..");
      const isValidProjectStub = vi
        .spyOn(controlHandlersDeps, "isValidProject")
        .mockImplementation((p: string | undefined) => {
          return p !== dirname;
        });
      const sendTelemetryEventStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "workspaceUri", { fsPath: "/path/to/workspace" });

      saveTextDocumentHandler({
        document: { fileName: "/dirname/fileName" },
        reason: vscode.TextDocumentSaveReason.FocusOut,
      } as vscode.TextDocumentWillSaveEvent);

      chai.assert.isTrue(isValidProjectStub.calledThrice);
      chai.assert.equal(isValidProjectStub.getCall(0).args[0], "/path/to/workspace");
      chai.assert.equal(isValidProjectStub.getCall(1).args[0], dirname);
      chai.assert.equal(isValidProjectStub.getCall(2).args[0], parentDir);
      chai.assert.equal(sendTelemetryEventStub.getCall(0).args[0], TelemetryEvent.UpdateTeamsApp);
      chai.assert.equal(
        sendTelemetryEventStub.getCall(0).args[1][TelemetryProperty.UpdateTeamsAppReason],
        TelemetryUpdateAppReason.FocusOut
      );
    });
  });

  describe("openLifecycleTreeview", () => {
    it("TeamsFx Project", async () => {
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "isTeamsFxProject", true);
      const executeCommandStub = vi.spyOn(vscode.commands, "executeCommand");

      await openLifecycleTreeview();

      chai.assert.isTrue(executeCommandStub.calledWith("teamsfx-lifecycle.focus"));
    });

    it("non-TeamsFx Project", async () => {
      vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
      mockValue(globalVariables, "isTeamsFxProject", false);
      const executeCommandStub = vi.spyOn(vscode.commands, "executeCommand");

      await openLifecycleTreeview();

      chai.assert.isTrue(executeCommandStub.calledWith("workbench.view.extension.teamsfx"));
    });
  });

  describe("selectWalkthrough", () => {
    let quickPickStub: ReturnType<typeof vi.spyOn>;
    let executeCommandStub: ReturnType<typeof vi.spyOn>;

    beforeEach(() => {
      // Stubbing VS Code APIs
      quickPickStub = vi.spyOn(vscode.window, "showQuickPick");
      executeCommandStub = vi.spyOn(vscode.commands, "executeCommand");
    });

    it("should select the declarative agent walkthrough", async () => {
      quickPickStub.mockResolvedValue({
        label: getDefaultString("teamstoolkit.walkthroughs.buildIntelligentApps.title"),
        detail: "Some description",
      });

      executeCommandStub.mockImplementation((command: string, ...args: any[]) => {
        chai.assert(command, "workbench.action.openWalkthrough");
        chai.assert(args[0], "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps");
        return "Success";
      });

      const result = await selectWalkthrough();

      chai.assert.isTrue(quickPickStub.calledOnce);
      chai.assert.isTrue(executeCommandStub.calledOnce);
      chai.assert.isTrue(result.isOk());
    });
  });
});
