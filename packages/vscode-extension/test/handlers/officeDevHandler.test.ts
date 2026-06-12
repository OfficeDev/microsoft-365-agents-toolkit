import { FxError, Result, ok } from "@microsoft/teamsfx-api";
import * as globalState from "@microsoft/teamsfx-core";
import * as chai from "chai";
import * as mockfs from "mock-fs";
import * as sinon from "sinon";
import * as vscode from "vscode";
import { OfficeDevTerminal, TriggerCmdType } from "../../src/debug/taskTerminal/officeDevTerminal";
import * as globalVariables from "../../src/globalVariables";
import * as officeDevHandlers from "../../src/handlers/officeDevHandlers";
import { generateManifestGUID, stopOfficeAddInDebug } from "../../src/handlers/officeDevHandlers";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { openOfficeDevFolder } from "../../src/utils/workspaceUtils";
import * as autoOpenHelper from "../../src/utils/autoOpenHelper";
import * as readmeHandlers from "../../src/handlers/readmeHandlers";
import * as telemetryUtils from "../../src/utils/telemetryUtils";

describe("officeDevHandler", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
    mockfs.restore();
  });

  async function testOpenUrlHandler(
    openLinkFunc: (args?: any[]) => Promise<Result<boolean, FxError>>,
    urlPath: string
  ) {
    sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    const openUrl = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
    const res = await openLinkFunc(undefined);
    chai.assert.isTrue(openUrl.calledOnce);
    chai.assert.isTrue(res.isOk());
    chai.assert.equal(openUrl.args[0][0], urlPath);
  }

  it("openOfficePartnerCenterHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openOfficePartnerCenterHandler,
      "https://aka.ms/WXPAddinPublish"
    );
  });

  it("openGetStartedLinkHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openGetStartedLinkHandler,
      "https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins"
    );
  });

  it("openOfficeDevDeployHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openOfficeDevDeployHandler,
      "https://aka.ms/WXPAddinDeploy"
    );
  });

  it("publishToAppSourceHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.publishToAppSourceHandler,
      "https://learn.microsoft.com/partner-center/marketplace/submit-to-appsource-via-partner-center"
    );
  });

  it("openDebugLinkHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openDebugLinkHandler,
      "https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-overview"
    );
  });

  it("openDocumentHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openDocumentHandler,
      "https://learn.microsoft.com/office/dev/add-ins/"
    );
  });

  it("openDevelopmentLinkHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openDevelopmentLinkHandler,
      "https://learn.microsoft.com/office/dev/add-ins/develop/develop-overview"
    );
  });

  it("openLifecycleLinkHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openLifecycleLinkHandler,
      "https://learn.microsoft.com/office/dev/add-ins/overview/core-concepts-office-add-ins"
    );
  });

  it("openHelpFeedbackLinkHandler", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openHelpFeedbackLinkHandler,
      "https://learn.microsoft.com/answers/tags/9/m365"
    );
  });

  it("openReportIssues", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openReportIssues,
      "https://github.com/OfficeDev/office-js/issues"
    );
  });

  it("openScriptLabLink", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openScriptLabLink,
      "https://learn.microsoft.com/office/dev/add-ins/overview/explore-with-script-lab"
    );
  });

  it("openPromptLibraryLink", async () => {
    testOpenUrlHandler(
      officeDevHandlers.openPromptLibraryLink,
      "https://aka.ms/OfficeAddinsPromptLibrary"
    );
  });
});

describe("autoOpenOfficeDevProjectHandler", () => {
  const sandbox = sinon.createSandbox();
  let inMemoryGlobalState: Map<string, any>;

  beforeEach(async () => {
    inMemoryGlobalState = new Map<string, any>();
    sandbox
      .stub(globalState, "globalStateGet")
      .callsFake(async (key: string, defaultValue?: any) => {
        return inMemoryGlobalState.has(key) ? inMemoryGlobalState.get(key) : defaultValue;
      });
    sandbox.stub(globalState, "globalStateUpdate").callsFake(async (key: string, value: any) => {
      inMemoryGlobalState.set(key, value);
    });

    await globalState.globalStateUpdate("fx-extension.openWalkThrough", false);
    await globalState.globalStateUpdate("fx-extension.openReadMe", "");
    await globalState.globalStateUpdate("fx-extension.openSampleReadMe", false);
    await globalState.globalStateUpdate("CreateWarnings", "");
  });

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(autoOpenHelper, "ShowScaffoldingWarningSummary").resolves();
  });

  afterEach(() => {
    sandbox.restore();
    mockfs.restore();
  });

  it("opens walk through", async () => {
    await globalState.globalStateUpdate("fx-extension.openWalkThrough", true);

    await officeDevHandlers.autoOpenOfficeDevProjectHandler();

    const isOpenWalkThrough = (await globalState.globalStateGet(
      "fx-extension.openWalkThrough",
      true
    )) as boolean;
    chai.assert.isFalse(isOpenWalkThrough);
  });

  it("opens README", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(readmeHandlers, "openReadMeHandler").resolves();
    const readmePath = vscode.Uri.file("test").fsPath;
    await globalState.globalStateUpdate("fx-extension.openReadMe", readmePath);

    await officeDevHandlers.autoOpenOfficeDevProjectHandler();

    const currentReadMe = (await globalState.globalStateGet(
      "fx-extension.openReadMe",
      ""
    )) as string;
    chai.assert.equal(currentReadMe, "");
  });

  it("opens sample README", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    sandbox.stub(globalVariables, "isOfficeAddInProject").resolves(false);
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .resolves(undefined);
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    sandbox.stub(autoOpenHelper, "showLocalDebugMessage").resolves();
    sandbox.stub(readmeHandlers, "openSampleReadmeHandler").resolves();
    await globalState.globalStateUpdate("fx-extension.openSampleReadMe", true);

    await officeDevHandlers.autoOpenOfficeDevProjectHandler();

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  it("openOfficeDevFolder", async () => {
    await globalState.globalStateUpdate("ShowLocalDebugMessage", false);
    await globalState.globalStateUpdate("fx-extension.openReadMe", "");
    await globalState.globalStateUpdate("fx-extension.openWalkThrough", true);
    const folderPath = vscode.Uri.file("/test");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    sandbox.stub(telemetryUtils, "isTriggerFromWalkThrough").returns(false);

    await openOfficeDevFolder(folderPath, true, [{ type: "warnning", content: "test" }]);

    const openWalkThrough = await globalState.globalStateGet("fx-extension.openWalkThrough", true);
    const openReadMe = await globalState.globalStateGet("fx-extension.openReadMe", "");
    const showLocalDebugMessage = await globalState.globalStateGet("ShowLocalDebugMessage", false);
    chai.assert.equal(openWalkThrough, false);
    chai.assert.equal(openReadMe, folderPath.fsPath);
    chai.assert.equal(showLocalDebugMessage, true);
    chai.assert(executeCommandStub.calledWithExactly("vscode.openFolder", folderPath, true));
  });
});

describe("OfficeDevTerminal", () => {
  const sandbox = sinon.createSandbox();
  let getInstanceStub: any, showStub: any, sendTextStub: any;

  beforeEach(() => {
    getInstanceStub = sandbox.stub(OfficeDevTerminal, "getInstance");
    showStub = sandbox.stub();
    sendTextStub = sandbox.stub();
    getInstanceStub.returns({ show: showStub, sendText: sendTextStub });
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
  });

  afterEach(() => {
    getInstanceStub.restore();
    sandbox.restore();
  });

  it("should validate Office AddIn Manifest", async () => {
    const result = await officeDevHandlers.validateOfficeAddInManifest();
    chai.expect(result.isOk()).to.be.true;
    sinon.assert.calledOnce(showStub);
    sinon.assert.calledWith(sendTextStub, TriggerCmdType.triggerValidate); // replace triggerValidate with actual value
  });

  it("should install Office AddIn Dependencies", async () => {
    const result = await officeDevHandlers.installOfficeAddInDependencies();
    chai.expect(result.isOk()).to.be.true;
    sinon.assert.calledOnce(showStub);
    sinon.assert.calledWith(sendTextStub, TriggerCmdType.triggerInstall); // replace triggerInstall with actual value
  });
});

class TerminalStub implements vscode.Terminal {
  shellIntegration: vscode.TerminalShellIntegration | undefined;
  name!: string;
  processId!: Thenable<number | undefined>;
  creationOptions!: Readonly<vscode.TerminalOptions | vscode.ExtensionTerminalOptions>;
  exitStatus: vscode.TerminalExitStatus | undefined;
  state!: vscode.TerminalState;
  hide(): void {
    throw new Error("Method not implemented.");
  }
  dispose(): void {
    throw new Error("Method not implemented.");
  }
  // Implement all methods from the Terminal interface
  // ...

  sendText(text: string, addNewLine?: boolean): void {
    // This is a stubbed method
  }

  show(preserveFocus?: boolean): void {
    // This is a stubbed method
  }
}

describe("stopOfficeAddInDebug", () => {
  let getInstanceStub: sinon.SinonStub;
  let showStub: sinon.SinonStub;
  let sendTextStub: sinon.SinonStub;
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should call getInstance, show and sendText", async () => {
    const terminalStub = new TerminalStub();
    getInstanceStub = sandbox.stub(OfficeDevTerminal, "getInstance").returns(terminalStub);
    showStub = sandbox.stub(terminalStub, "show");
    sendTextStub = sandbox.stub(terminalStub, "sendText");
    await stopOfficeAddInDebug();

    sinon.assert.calledOnce(getInstanceStub);
    sinon.assert.calledOnce(showStub);
    sinon.assert.calledOnce(sendTextStub);
  });
});

describe("generateManifestGUID", () => {
  let getInstanceStub: sinon.SinonStub;
  let showStub: sinon.SinonStub;
  let sendTextStub: sinon.SinonStub;
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should call getInstance, show and sendText with correct arguments", async () => {
    const terminalStub = new TerminalStub();
    getInstanceStub = sandbox.stub(OfficeDevTerminal, "getInstance").returns(terminalStub);
    showStub = sandbox.stub(terminalStub, "show");
    sendTextStub = sandbox.stub(terminalStub, "sendText");

    await generateManifestGUID();

    sinon.assert.calledOnce(getInstanceStub);
    sinon.assert.calledOnce(showStub);
    sinon.assert.calledOnce(sendTextStub);
    sinon.assert.calledWithExactly(sendTextStub, TriggerCmdType.triggerGenerateGUID);
  });
});
