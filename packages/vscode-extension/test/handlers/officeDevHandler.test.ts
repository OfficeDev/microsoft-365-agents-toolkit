import { FxError, Result, ok } from "@microsoft/teamsfx-api";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
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

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
  });

  afterEach(() => {
    sandbox.restore();
    mockfs.restore();
  });

  it("opens walk through", async () => {
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openWalkThrough") {
        return true;
      } else {
        return false;
      }
    });
    const stateUpdate = sandbox.stub(globalState, "globalStateUpdate");

    await officeDevHandlers.autoOpenOfficeDevProjectHandler();

    chai.assert.isTrue(stateUpdate.calledOnce);
  });

  it("opens README", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "fx-extension.openReadMe") {
        return vscode.Uri.file("test").fsPath;
      } else {
        return "";
      }
    });

    const openReadMeHandlerStub = sandbox.stub(readmeHandlers, "openReadMeHandler");
    const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate");
    const ShowScaffoldingWarningSummaryStub = sandbox.stub(
      autoOpenHelper,
      "ShowScaffoldingWarningSummary"
    );

    await officeDevHandlers.autoOpenOfficeDevProjectHandler();

    chai.assert.isTrue(openReadMeHandlerStub.calledOnce);
    chai.assert.isTrue(globalStateUpdateStub.calledTwice);
    chai.assert.isTrue(ShowScaffoldingWarningSummaryStub.calledOnce);
  });

  it("opens sample README", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    sandbox.stub(globalVariables, "isTeamsFxProject").resolves(false);
    sandbox.stub(globalVariables, "isOfficeAddInProject").resolves(false);
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

    await officeDevHandlers.autoOpenOfficeDevProjectHandler();

    chai.assert.isTrue(executeCommandStub.calledOnce);
  });

  it("openOfficeDevFolder", async () => {
    const folderPath = vscode.Uri.file("/test");
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
    const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate");

    await openOfficeDevFolder(folderPath, true, [{ type: "warnning", content: "test" }]);

    chai.assert(globalStateUpdateStub.callCount == 5);
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
