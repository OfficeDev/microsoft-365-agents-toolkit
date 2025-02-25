import { err, ok, Platform, SystemError, UserError } from "@microsoft/teamsfx-api";
import {
  AppDefinition,
  FeatureFlagName,
  teamsDevPortalClient,
  UnhandledError,
  UserCancelError,
} from "@microsoft/teamsfx-core";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import { ProgressHandler } from "@microsoft/vscode-ui";
import { assert } from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import * as copilotHandler from "../../src/handlers/copilotChatHandlers";
import {
  addAuthActionHandler,
  addPluginHandler,
  addWebpartHandler,
  copilotPluginAddAPIHandler,
  createNewProjectHandler,
  deployHandler,
  provisionHandler,
  publishHandler,
  scaffoldFromDeveloperPortalHandler,
  addKnowledgeHandler,
} from "../../src/handlers/lifecycleHandlers";
import * as shared from "../../src/handlers/sharedOpts";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { TelemetryEvent } from "../../src/telemetry/extTelemetryEvents";
import envTreeProviderInstance from "../../src/treeview/environmentTreeViewProvider";
import * as telemetryUtils from "../../src/utils/telemetryUtils";
import * as workspaceUtils from "../../src/utils/workspaceUtils";
import M365TokenInstance from "../../src/commonlib/m365Login";
import { MockCore } from "../mocks/mockCore";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import mockedEnv, { RestoreFn } from "mocked-env";
import VsCodeLogInstance from "../../src/commonlib/log";

describe("Lifecycle handlers", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("provision handlers", () => {
    it("error", async () => {
      sandbox.stub(shared, "runCommand").resolves(err(new UserCancelError()));
      const res = await provisionHandler();
      assert.isTrue(res.isErr());
    });
  });

  describe("createNewProjectHandler", function () {
    const sandbox = sinon.createSandbox();
    let mockedEnvRestore: RestoreFn;

    afterEach(() => {
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
      sandbox.restore();
    });

    it("invokeTeamsAgent", async () => {
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "abc",
          shouldInvokeTeamsAgent: true,
          projectId: "mockId",
        })
      );
      sandbox.stub(copilotHandler, "invokeTeamsAgent").resolves();
      const res = await createNewProjectHandler();
      assert.isTrue(res.isOk());
    });

    it("triggered in office agent", async () => {
      sandbox.stub(projectSettingsHelper, "isValidOfficeAddInProject").returns(true);
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "abc",
          shouldInvokeTeamsAgent: false,
          projectId: "mockId",
        })
      );
      sandbox.stub(copilotHandler, "invokeTeamsAgent").resolves();
      const res = await createNewProjectHandler("", { agent: "office" });
      assert.isTrue(res.isOk());
    });

    it("office add-in", async () => {
      sandbox.stub(projectSettingsHelper, "isValidOfficeAddInProject").returns(true);
      const openOfficeDevFolder = sandbox.stub(workspaceUtils, "openOfficeDevFolder").resolves();
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "abc",
          shouldInvokeTeamsAgent: false,
          projectId: "mockId",
        })
      );
      const res = await createNewProjectHandler();
      assert.isTrue(res.isOk());
      assert.isTrue(openOfficeDevFolder.calledOnce);
    });

    it("none office add-in", async () => {
      sandbox.stub(projectSettingsHelper, "isValidOfficeAddInProject").returns(false);
      const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "abc",
          shouldInvokeTeamsAgent: false,
          projectId: "mockId",
        })
      );
      const res = await createNewProjectHandler({ teamsAppFromTdp: true }, {});
      assert.isTrue(res.isOk());
      assert.isTrue(openFolder.calledOnce);
    });

    it("kiota integration: kiota installed release version", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "",
          lastCommand: "command",
        })
      );
      sandbox.stub(vscode.extensions, "getExtension").returns({
        id: "mockedId",
        extensionUri: vscode.Uri.parse("file://mockedUri"),
        isActive: true,
        extensionPath: "mockedPath",
        extensionKind: vscode.ExtensionKind.UI,
        exports: {},
        packageJSON: {
          version: "1.18.100000002",
        },
        activate: () => Promise.resolve(),
      });
      const executeCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const logError = sandbox.stub(VsCodeLogInstance, "error").resolves();
      const res = await createNewProjectHandler();
      assert.isTrue(res.isOk());
      assert.isTrue(executeCommand.calledOnce);
      assert.isTrue(logError.notCalled);
    });

    it("kiota integration: kiota installed pre-release version", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "",
          lastCommand: "command",
        })
      );
      sandbox.stub(vscode.extensions, "getExtension").returns({
        id: "mockedId",
        extensionUri: vscode.Uri.parse("file://mockedUri"),
        isActive: true,
        extensionPath: "mockedPath",
        extensionKind: vscode.ExtensionKind.UI,
        exports: {},
        packageJSON: {
          version: "1.19.24090901",
        },
        activate: () => Promise.resolve(),
      });
      const executeCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const logError = sandbox.stub(VsCodeLogInstance, "error").resolves();
      const res = await createNewProjectHandler();
      assert.isTrue(res.isOk());
      assert.isTrue(executeCommand.calledOnce);
      assert.isTrue(logError.notCalled);
    });

    it("kiota integration: kiota not installed and click install", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "",
          lastCommand: "command",
        })
      );
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .callsFake((title: string, ...items: any[]) => {
          return Promise.resolve(items[0]);
        });
      const executeCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const logError = sandbox.stub(VsCodeLogInstance, "error").resolves();
      const res = await createNewProjectHandler();
      assert.isTrue(res.isOk());
      assert.isTrue(showMessageStub.calledOnce);
      assert.isTrue(executeCommand.calledOnce);
      assert.isTrue(logError.calledOnce);
    });

    it("kiota integration: kiota version not match and click install", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "",
          lastCommand: "command",
        })
      );
      sandbox.stub(vscode.extensions, "getExtension").returns({
        id: "mockedId",
        extensionUri: vscode.Uri.parse("file://mockedUri"),
        isActive: true,
        extensionPath: "mockedPath",
        extensionKind: vscode.ExtensionKind.UI,
        exports: {},
        packageJSON: {
          version: "1.18.100000001",
        },
        activate: () => Promise.resolve(),
      });
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .callsFake((title: string, ...items: any[]) => {
          return Promise.resolve(items[0]);
        });
      const executeCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const logError = sandbox.stub(VsCodeLogInstance, "error").resolves();
      const res = await createNewProjectHandler();
      assert.isTrue(res.isOk());
      assert.isTrue(showMessageStub.calledOnce);
      assert.isTrue(executeCommand.calledOnce);
      assert.isTrue(logError.calledOnce);
    });

    it("kiota integration: no kiota version and click install", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "",
          lastCommand: "command",
        })
      );
      sandbox.stub(vscode.extensions, "getExtension").returns({
        id: "mockedId",
        extensionUri: vscode.Uri.parse("file://mockedUri"),
        isActive: true,
        extensionPath: "mockedPath",
        extensionKind: vscode.ExtensionKind.UI,
        exports: {},
        packageJSON: {},
        activate: () => Promise.resolve(),
      });
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .callsFake((title: string, ...items: any[]) => {
          return Promise.resolve(items[0]);
        });
      const executeCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const logError = sandbox.stub(VsCodeLogInstance, "error").resolves();
      const res = await createNewProjectHandler();
      assert.isTrue(res.isOk());
      assert.isTrue(showMessageStub.calledOnce);
      assert.isTrue(executeCommand.calledOnce);
      assert.isTrue(logError.calledOnce);
    });

    it("kiota integration: kiota not installed and click cancel", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "",
          lastCommand: "command",
        })
      );
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .callsFake((title: string, ...items: any[]) => {
          return Promise.resolve(items[1]);
        });
      const executeCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const logError = sandbox.stub(VsCodeLogInstance, "error").resolves();
      const res = await createNewProjectHandler();
      assert.isTrue(res.isOk());
      assert.isTrue(showMessageStub.calledOnce);
      assert.isTrue(executeCommand.notCalled);
      assert.isTrue(logError.calledOnce);
    });
  });

  describe("provisionHandler", function () {
    it("happy", async () => {
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      sandbox.stub(envTreeProviderInstance, "reloadEnvironments");
      const res = await provisionHandler();
      assert.isTrue(res.isOk());
    });
  });

  describe("deployHandler", function () {
    it("happy", async () => {
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      const res = await deployHandler();
      assert.isTrue(res.isOk());
    });
  });

  describe("publishHandler", function () {
    it("happy()", async () => {
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      const res = await publishHandler();
      assert.isTrue(res.isOk());
    });
  });

  describe("addWebpartHandler", function () {
    it("happy()", async () => {
      sandbox.stub(shared, "runCommand").resolves(ok(undefined));
      const res = await addWebpartHandler();
      assert.isTrue(res.isOk());
    });
  });

  describe("scaffoldFromDeveloperPortalHandler", async () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(globalVariables, "checkIsSPFx").returns(false);
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("missing args", async () => {
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      const createProgressBar = sandbox
        .stub(vsc_ui.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);

      const res = await scaffoldFromDeveloperPortalHandler();

      assert.equal(res.isOk(), true);
      assert.equal(createProgressBar.notCalled, true);
    });

    it("incorrect number of args", async () => {
      const progressHandler = new ProgressHandler("title", 1);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      const createProgressBar = sandbox
        .stub(vsc_ui.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);

      const res = await scaffoldFromDeveloperPortalHandler();

      assert.equal(res.isOk(), true);
      assert.equal(createProgressBar.notCalled, true);
    });

    it("general error when signing in M365", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sandbox.stub(progressHandler, "start").resolves();
      const endProgress = sandbox.stub(progressHandler, "end").resolves();
      sandbox.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").throws("error1");
      const createProgressBar = sandbox
        .stub(vsc_ui.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sandbox.stub(vscode.window, "showErrorMessage");

      const res = await scaffoldFromDeveloperPortalHandler(["appId"]);
      assert.isTrue(res.isErr());
      assert.isTrue(createProgressBar.calledOnce);
      assert.isTrue(startProgress.calledOnce);
      assert.isTrue(endProgress.calledOnceWithExactly(false));
      assert.isTrue(showErrorMessage.calledOnce);
      if (res.isErr()) {
        assert.isTrue(res.error instanceof UnhandledError);
      }
    });

    it("error when signing M365", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sandbox.stub(progressHandler, "start").resolves();
      const endProgress = sandbox.stub(progressHandler, "end").resolves();
      sandbox
        .stub(M365TokenInstance, "signInWhenInitiatedFromTdp")
        .resolves(err(new UserError("source", "name", "message", "displayMessage")));
      const createProgressBar = sandbox
        .stub(vsc_ui.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sandbox.stub(vscode.window, "showErrorMessage");

      const res = await scaffoldFromDeveloperPortalHandler(["appId"]);

      assert.equal(res.isErr(), true);
      assert.equal(createProgressBar.calledOnce, true);
      assert.equal(startProgress.calledOnce, true);
      assert.equal(endProgress.calledOnceWithExactly(false), true);
      assert.equal(showErrorMessage.calledOnce, true);
    });

    it("error when signing in M365 but missing display message", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sandbox.stub(progressHandler, "start").resolves();
      const endProgress = sandbox.stub(progressHandler, "end").resolves();
      sandbox
        .stub(M365TokenInstance, "signInWhenInitiatedFromTdp")
        .resolves(err(new UserError("source", "name", "", "")));
      const createProgressBar = sandbox
        .stub(vsc_ui.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sandbox.stub(vscode.window, "showErrorMessage");

      const res = await scaffoldFromDeveloperPortalHandler(["appId"]);

      assert.equal(res.isErr(), true);
      assert.equal(createProgressBar.calledOnce, true);
      assert.equal(startProgress.calledOnce, true);
      assert.equal(endProgress.calledOnceWithExactly(false), true);
      assert.equal(showErrorMessage.calledOnce, true);
    });

    it("failed to get teams app", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sandbox.stub(progressHandler, "start").resolves();
      const endProgress = sandbox.stub(progressHandler, "end").resolves();
      sandbox.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").resolves(ok("token"));
      sandbox
        .stub(M365TokenInstance, "getAccessToken")
        .resolves(err(new SystemError("source", "name", "", "")));
      const createProgressBar = sandbox
        .stub(vsc_ui.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      sandbox.stub(globalVariables, "core").value(new MockCore());
      sandbox.stub(vscode.commands, "executeCommand");
      sandbox.stub(globalState, "globalStateUpdate");
      const getApp = sandbox.stub(teamsDevPortalClient, "getApp").throws("error");

      const res = await scaffoldFromDeveloperPortalHandler(["appId"]);

      assert.isTrue(res.isErr());
      assert.isTrue(getApp.calledOnce);
      assert.isTrue(createProgressBar.calledOnce);
      assert.isTrue(startProgress.calledOnce);
      assert.isTrue(endProgress.calledOnceWithExactly(true));
    });

    it("happy path", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sandbox.stub(progressHandler, "start").resolves();
      const endProgress = sandbox.stub(progressHandler, "end").resolves();
      sandbox.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").resolves(ok("token"));
      sandbox.stub(M365TokenInstance, "getAccessToken").resolves(ok("authSvcToken"));
      sandbox.stub(teamsDevPortalClient, "setRegionEndpointByToken").resolves();
      const createProgressBar = sandbox
        .stub(vsc_ui.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const createProject = sandbox.spy(globalVariables.core, "createProjectFromTdp");
      sandbox.stub(vscode.commands, "executeCommand");
      sandbox.stub(globalState, "globalStateUpdate");
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
      };
      sandbox.stub(teamsDevPortalClient, "getApp").resolves(appDefinition);

      const res = await scaffoldFromDeveloperPortalHandler("appId", "testuser");

      assert.equal(createProject.args[0][0].teamsAppFromTdp.teamsAppId, "mock-id");
      assert.isTrue(res.isOk());
      assert.isTrue(createProgressBar.calledOnce);
      assert.isTrue(startProgress.calledOnce);
      assert.isTrue(endProgress.calledOnceWithExactly(true));
    });
  });

  describe("copilotPluginAddAPIHandler", async () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("API ME:", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const addAPIHanlder = sandbox.spy(globalVariables.core, "copilotPluginAddAPI");
      const args = [
        {
          fsPath: "manifest.json",
        },
      ];

      await copilotPluginAddAPIHandler(args);

      sinon.assert.calledOnce(addAPIHanlder);
    });

    it("API Plugin", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const addAPIHanlder = sandbox.spy(globalVariables.core, "copilotPluginAddAPI");
      const args = [
        {
          fsPath: "openapi.yaml",
          isFromApiPlugin: true,
          manifestPath: "manifest.json",
        },
      ];

      await copilotPluginAddAPIHandler(args);

      sinon.assert.calledOnce(addAPIHanlder);
    });
  });

  describe("AddPluginHandler", async () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("success:", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const addPluginHanlder = sandbox.spy(globalVariables.core, "addPlugin");

      await addPluginHandler();

      sinon.assert.calledOnce(addPluginHanlder);
    });

    it("success: success call kiota", async () => {
      const mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(globalVariables, "core").value(new MockCore());
      sandbox
        .stub(globalVariables.core, "addPlugin")
        .resolves(ok({ lastCommand: "addPlugin", manifestPath: "manifest.json" }));
      sandbox.stub(shared, "runCommand").resolves(
        ok({
          projectPath: "",
          lastCommand: "command",
        })
      );
      sandbox.stub(vscode.extensions, "getExtension").returns({
        id: "mockedId",
        extensionUri: vscode.Uri.parse("file://mockedUri"),
        isActive: true,
        extensionPath: "mockedPath",
        extensionKind: vscode.ExtensionKind.UI,
        exports: {},
        packageJSON: {
          version: "1.18.100000002",
        },
        activate: () => Promise.resolve(),
      });
      const executeCommand = sandbox.stub(vscode.commands, "executeCommand").resolves();
      const logError = sandbox.stub(VsCodeLogInstance, "error").resolves();

      const result = await addPluginHandler();

      assert.isTrue(result.isOk());
      assert.isTrue(executeCommand.calledOnce);
      assert.isTrue(logError.notCalled);
      mockedEnvRestore();
    });
  });

  describe("AddAuthActionHandler", async () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .callsFake((title: string, ...items: any[]) => {
          return Promise.resolve(items[0]);
        });
      const addAuthAction = sandbox.spy(globalVariables.core, "addAuthAction");
      const provisionction = sandbox.spy(globalVariables.core, "provisionResources");
      await addAuthActionHandler();
      sandbox.assert.calledOnce(addAuthAction);
      sandbox.assert.calledOnce(provisionction);
    });
  });

  describe("addKnowledgeHandler", async () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path", async () => {
      sandbox.stub(globalVariables, "core").value(new MockCore());
      const addKnowledge = sandbox.spy(globalVariables.core, "addKnowledge");

      await addKnowledgeHandler();

      sinon.assert.calledOnce(addKnowledge);
    });
  });
});
