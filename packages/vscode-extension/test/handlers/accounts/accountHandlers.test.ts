import * as vscode from "vscode";
import * as sinon from "sinon";
import * as chai from "chai";
import M365TokenInstance from "../../../src/commonlib/m365Login";
import { err, ok } from "@microsoft/teamsfx-api";
import { AzureAccountManager } from "../../../src/commonlib/azureLogin";
import * as vsc_ui from "../../../src/qm/vsc_ui";
import {
  azureAccountSignOutHelpHandler,
  cmpAccountsHandler,
  createAccountHandler,
} from "../../../src/handlers/accounts/accountHandlers";
import { ExtTelemetry } from "../../../src/telemetry/extTelemetry";
import * as localizeUtils from "../../../src/utils/localizeUtils";

describe("AccountHandlers", () => {
  describe("createAccountHandler", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    beforeEach(() => {
      sandbox.stub(localizeUtils, "localize").returns("test");
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    });

    it("create M365 account", async () => {
      const selectOptionStub = sandbox
        .stub(vsc_ui.VS_CODE_UI, "selectOption")
        .resolves(ok({ result: "createAccountM365" } as any));
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl");
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      await createAccountHandler([]);

      chai.expect(selectOptionStub.calledOnce).to.be.true;
      chai.expect(
        openUrlStub.calledOnceWith("https://developer.microsoft.com/microsoft-365/dev-program")
      ).to.be.true;
      chai.expect(sendTelemetryEventStub.args[1][1]).to.deep.equal({
        "account-type": "m365",
        "trigger-from": "CommandPalette",
      });
    });

    it("create Azure account", async () => {
      const selectOptionStub = sandbox
        .stub(vsc_ui.VS_CODE_UI, "selectOption")
        .resolves(ok({ result: "createAccountAzure" } as any));
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl");
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      await createAccountHandler([]);

      chai.expect(selectOptionStub.calledOnce).to.be.true;
      chai.expect(openUrlStub.calledOnceWith("https://azure.microsoft.com/en-us/free/")).to.be.true;
      chai.expect(sendTelemetryEventStub.args[1][1]).to.deep.equal({
        "account-type": "azure",
        "trigger-from": "CommandPalette",
      });
    });

    it("create account error", async () => {
      const selectOptionStub = sandbox
        .stub(vsc_ui.VS_CODE_UI, "selectOption")
        .resolves(err("error") as any);
      const sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      await createAccountHandler([]);

      chai.expect(selectOptionStub.calledOnce).to.be.true;
      chai.expect(sendTelemetryEventStub.calledOnce).to.be.true;
      chai.expect(sendTelemetryErrorEventStub.calledOnce).to.be.true;
    });
  });

  describe("cmpAccountsHandler", () => {
    const sandbox = sinon.createSandbox();
    let changeSelectionCallback: (e: readonly vscode.QuickPickItem[]) => any;
    let stubQuickPick: any;

    afterEach(() => {
      sandbox.restore();
    });

    beforeEach(() => {
      changeSelectionCallback = () => {};
      stubQuickPick = {
        items: [],
        onDidChangeSelection: (
          _changeSelectionCallback: (e: readonly vscode.QuickPickItem[]) => any
        ) => {
          changeSelectionCallback = _changeSelectionCallback;
          return {
            dispose: () => {},
          };
        },
        onDidHide: () => {
          return {
            dispose: () => {},
          };
        },
        show: () => {},
        hide: () => {},
        onDidAccept: () => {},
      };
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(vscode.window, "createQuickPick").returns(stubQuickPick as any);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      sandbox.stub(vsc_ui.VS_CODE_UI, "selectOption").resolves(ok({ result: "unknown" } as any));
    });

    it("Sign out happy path", async () => {
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .resolves(undefined);
      const M365SignOutStub = sandbox.stub(M365TokenInstance, "signout");
      sandbox
        .stub(M365TokenInstance, "getStatus")
        .resolves(ok({ status: "SignedIn", accountInfo: { upn: "test.email.com" } }));
      sandbox
        .stub(AzureAccountManager.prototype, "getStatus")
        .resolves({ status: "SignedIn", accountInfo: { upn: "test.email.com" } });
      const hideStub = sandbox.stub(stubQuickPick, "hide");

      await cmpAccountsHandler([]);
      changeSelectionCallback([stubQuickPick.items[1]]);

      for (const i of stubQuickPick.items) {
        await (i as any).function();
      }

      chai.assert.isTrue(showMessageStub.calledTwice);
      chai.assert.isTrue(M365SignOutStub.calledOnce);
      chai.assert.isTrue(hideStub.calledOnce);
    });

    it("Sign in happy path", async () => {
      const showMessageStub = sandbox
        .stub(vscode.window, "showInformationMessage")
        .resolves(undefined);
      const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");
      sandbox
        .stub(M365TokenInstance, "getStatus")
        .resolves(ok({ status: "SignedOut", accountInfo: { upn: "test.email.com" } }));
      sandbox
        .stub(AzureAccountManager.prototype, "getStatus")
        .resolves({ status: "SignedOut", accountInfo: { upn: "test.email.com" } });
      const hideStub = sandbox.stub(stubQuickPick, "hide");

      await cmpAccountsHandler([]);
      changeSelectionCallback([stubQuickPick.items[1]]);

      for (const i of stubQuickPick.items) {
        await (i as any).function();
      }

      chai.assert.isTrue(showMessageStub.notCalled);
      chai.assert.isTrue(executeCommandStub.calledThrice);
      chai.expect(executeCommandStub.args[0][0]).to.be.equal("fx-extension.signinAzure");
      chai.expect(executeCommandStub.args[1][0]).to.be.equal("fx-extension.signinM365");
      chai.expect(executeCommandStub.args[2][0]).to.be.equal("fx-extension.signinAzure");
      chai.assert.isTrue(hideStub.calledOnce);
    });

    it("Sign out happy path - unique_name", async () => {
      sandbox.stub(vscode.window, "showInformationMessage").resolves(undefined);
      sandbox.stub(M365TokenInstance, "signout");
      sandbox
        .stub(M365TokenInstance, "getStatus")
        .resolves(ok({ status: "SignedIn", accountInfo: { unique_name: "test.email.com" } }));
      sandbox
        .stub(AzureAccountManager.prototype, "getStatus")
        .resolves({ status: "SignedIn", accountInfo: { upn: "test.email.com" } });
      sandbox.stub(stubQuickPick, "hide");

      await cmpAccountsHandler([]);

      chai.assert.isTrue((stubQuickPick.items[0].label as string).includes("test.email.com"));
    });

    it("Sign out happy path - undefined", async () => {
      sandbox.stub(vscode.window, "showInformationMessage").resolves(undefined);
      sandbox.stub(M365TokenInstance, "signout");
      sandbox
        .stub(M365TokenInstance, "getStatus")
        .resolves(ok({ status: "SignedIn", accountInfo: {} }));
      sandbox
        .stub(AzureAccountManager.prototype, "getStatus")
        .resolves({ status: "SignedIn", accountInfo: { upn: "test.email.com" } });
      sandbox.stub(stubQuickPick, "hide");

      await cmpAccountsHandler([]);

      chai.assert.equal(
        stubQuickPick.items[0].label as string,
        localizeUtils.localize("teamstoolkit.handlers.signOutOfM365")
      );
    });
  });

  describe("azureAccountSignOutHelpHandler", () => {
    it("happy path", async () => {
      try {
        azureAccountSignOutHelpHandler();
      } catch (e) {
        chai.assert.isTrue(e instanceof Error);
      }
    });
  });
});
