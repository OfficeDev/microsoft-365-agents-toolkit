import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import path from "path";
import * as globalVariables from "../../src/globalVariables";
import * as vsc_ui from "../../src/qm/vsc_ui";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import VsCodeLogInstance from "../../src/commonlib/log";
import { ok, ManifestUtil, err, UserError, SystemError } from "@microsoft/teamsfx-api";
import { manifestUtils, pluginManifestUtils } from "@microsoft/teamsfx-core";
import { GlobalKey } from "../../src/constants";
import { VsCodeUI } from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import { autoOpenProjectHandler } from "../../src/handlers/autoOpenProjectHandler";
import * as pluginGeneratorHelper from "@microsoft/teamsfx-core/build/component/generator/apiSpec/helper";

describe("autoOpenProjectHandler", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("opens walk through", async () => {
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openWalkThrough") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandFunc = sandbox.stub(vscode.commands, "executeCommand");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.notCalled);
    chai.assert.isTrue(executeCommandFunc.notCalled);
  });

  it("opens walk through if workspace Uri exists", async () => {
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openWalkThrough") {
        return true;
      } else {
        return false;
      }
    });
    const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.parse("test"));
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandFunc = sandbox.stub(vscode.commands, "executeCommand");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.notCalled);
    chai.assert.isTrue(executeCommandFunc.notCalled);
    chai.assert.isTrue(globalStateUpdateStub.calledTwice);
  });

  it("opens README", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else {
        return "";
      }
    });
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({} as any));
    sandbox.stub(ManifestUtil, "parseCommonProperties").resolves({ isCopilotPlugin: false });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.calledOnce);
  });

  it("opens sample README", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox.stub(vscode.window, "showInformationMessage");
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openSampleReadMe") {
        return true;
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    await autoOpenProjectHandler();

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  it("opens README and show APIE ME warnings successfully", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return JSON.stringify([{ type: "type", content: "content" }]);
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(
      ok({
        name: { short: "short", full: "full" },
        description: { short: "short", full: "" },
        composeExtensions: [{ apiSpecificationFile: "test.json", commands: [{ id: "command1" }] }],
      } as any)
    );
    const parseRes = {
      id: "",
      version: "",
      capabilities: [""],
      manifestVersion: "",
      isApiME: true,
      isSPFx: false,
      isApiMeAAD: false,
    };
    const parseManifestStub = sandbox.stub(ManifestUtil, "parseCommonProperties").returns(parseRes);
    VsCodeLogInstance.outputChannel = {
      show: () => {},
      info: () => {},
    } as unknown as vscode.OutputChannel;
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const generateWarningStub = sandbox
      .stub(pluginGeneratorHelper, "generateScaffoldingSummary")
      .resolves("warning message");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.calledTwice);
    chai.assert.isTrue(parseManifestStub.called);
    chai.assert.isTrue(generateWarningStub.called);
  });

  it("opens README and show warnings", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return JSON.stringify([{ type: "type", content: "content" }]);
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(
      ok({
        name: { short: "short", full: "full" },
        description: { short: "short", full: "" },
        composeExtensions: [{ commands: [{ id: "command1" }] }],
      } as any)
    );
    const parseRes = {
      id: "",
      version: "",
      capabilities: [""],
      manifestVersion: "",
      isApiME: true,
      isSPFx: false,
      isApiMeAAD: false,
    };
    const parseManifestStub = sandbox.stub(ManifestUtil, "parseCommonProperties").returns(parseRes);
    VsCodeLogInstance.outputChannel = {
      show: () => {},
      info: () => {},
    } as unknown as vscode.OutputChannel;
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const generateWarningStub = sandbox
      .stub(pluginGeneratorHelper, "generateScaffoldingSummary")
      .resolves("warning message");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.calledTwice);
    chai.assert.isTrue(parseManifestStub.called);
    chai.assert.isFalse(generateWarningStub.called);
  });

  it("opens README and show copilot plugin warnings successfully", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    sandbox.stub(vscode.window, "showInformationMessage").resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return JSON.stringify([{ type: "type", content: "content" }]);
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(path, "relative").returns("test");

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(
      ok({
        name: { short: "short", full: "full" },
        description: { short: "short", full: "" },
        copilotExtensions: { plugins: [{ file: "ai-plugin.json", id: "plugin1" }] },
      } as any)
    );
    const parseRes = {
      id: "",
      version: "",
      capabilities: ["plugin"],
      manifestVersion: "",
      isApiME: false,
      isSPFx: false,
      isApiMeAAD: false,
    };
    const parseManifestStub = sandbox.stub(ManifestUtil, "parseCommonProperties").returns(parseRes);
    const getApiSpecStub = sandbox
      .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
      .resolves(ok(["test"]));
    VsCodeLogInstance.outputChannel = {
      show: () => {},
      info: () => {},
    } as unknown as vscode.OutputChannel;
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const generateWarningStub = sandbox
      .stub(pluginGeneratorHelper, "generateScaffoldingSummary")
      .resolves("warning message");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendTelemetryStub.calledTwice);
    chai.assert.isTrue(parseManifestStub.called);
    chai.assert.isTrue(getApiSpecStub.called);
    chai.assert.isTrue(generateWarningStub.called);
  });
  it("skip show warnings if parsing error", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return "string";
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const sendErrorTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendErrorTelemetryStub.called);
  });

  it("skip show warnings if cannot get manifest", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return "string";
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox
      .stub(manifestUtils, "_readAppManifest")
      .resolves(err(new UserError("source", "name", "", "")));

    const sendErrorTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendErrorTelemetryStub.called);
  });

  it("skip show warnings if get plugin api spec error", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else if (key === GlobalKey.CreateWarnings) {
        return JSON.stringify([{ type: "type", content: "content" }]);
      } else {
        return "";
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(
      ok({
        name: { short: "short", full: "full" },
        description: { short: "short", full: "" },
        copilotExtensions: { plugins: [{ file: "ai-plugin.json", id: "plugin1" }] },
      } as any)
    );
    const parseRes = {
      id: "",
      version: "",
      capabilities: ["plugin"],
      manifestVersion: "",
      isApiME: false,
      isSPFx: false,
      isApiBasedMe: true,
      isApiMeAAD: false,
    };
    sandbox.stub(ManifestUtil, "parseCommonProperties").returns(parseRes);
    const getApiSpecStub = sandbox
      .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
      .resolves(err(new SystemError("test", "test", "", "")));
    VsCodeLogInstance.outputChannel = {
      show: () => {},
      info: () => {},
    } as unknown as vscode.OutputChannel;
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const sendErrorTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");

    await autoOpenProjectHandler();

    chai.assert.isTrue(sendErrorTelemetryStub.called);
    chai.assert.equal(
      sendErrorTelemetryStub.args[0][0],
      TelemetryEvent.ShowScaffoldingWarningSummaryError
    );
    chai.assert.isTrue(getApiSpecStub.called);
  });

  it("auto install dependency", async () => {
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "teamsToolkit:autoInstallDependency") {
        return true;
      } else {
        return false;
      }
    });
    const globalStateStub = sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(vsc_ui, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
    const runCommandStub = sandbox.stub(vsc_ui.VS_CODE_UI, "runCommand");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await autoOpenProjectHandler();

    chai.assert.isTrue(globalStateStub.calledWith("teamsToolkit:autoInstallDependency", false));
    chai.assert.isTrue(runCommandStub.calledOnce);
  });
});
