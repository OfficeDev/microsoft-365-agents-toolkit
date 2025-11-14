import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import { err, ok, SystemError } from "@microsoft/teamsfx-api";
import { NetworkError, UserCancelError } from "@microsoft/teamsfx-core";
import * as tool from "@microsoft/teamsfx-core/build/src/common/tools";
import azureAccountManager from "../../../src/commonlib/azureLogin";
import { LoginFailureError } from "../../../src/commonlib/codeFlowLogin";
import M365TokenInstance from "../../../src/commonlib/m365Login";
import {
  onSwitchAzureTenant,
  onSwitchM365Tenant,
} from "../../../src/handlers/accounts/switchTenantHandler";
import * as vsc_ui from "../../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../../src/telemetry/extTelemetry";
import { TelemetryTriggerFrom } from "../../../src/telemetry/extTelemetryEvents";

describe("onSwitchM365Tenant", () => {
  const sandbox = sinon.createSandbox();
  let sendTelemetryEventStub: sinon.SinonStub;
  let sendTelemetryErrorEventStub: sinon.SinonStub;
  let selectOptionStub: sinon.SinonStub;

  beforeEach(() => {
    sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI({} as vscode.ExtensionContext));
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("Failed to retrieve access token", async () => {
    sandbox
      .stub(M365TokenInstance, "getAccessToken")
      .resolves(err(new NetworkError("extension", "")));

    await onSwitchM365Tenant(TelemetryTriggerFrom.SideBar);

    chai.assert.isTrue(sendTelemetryEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.args[0][1] instanceof NetworkError);
  });

  it("Failed to select tenant in UI", async () => {
    sandbox.stub(M365TokenInstance, "getAccessToken").resolves(ok("faked token"));
    sandbox.stub(M365TokenInstance, "switchTenant").resolves(ok("faked token"));
    sandbox.stub(tool, "listAllTenants").resolves([
      {
        tenantId: "0022fd51-06f5-4557-8a34-69be98de6e20",
        displayName: "MSFT",
        defaultDomain: "t815h.onmicrosoft.com",
      },
      {
        tenantId: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
        displayName: "Cisco",
        defaultDomain: "Cisco561.onmicrosoft.com",
      },
    ]);
    selectOptionStub = sandbox
      .stub(vsc_ui.VS_CODE_UI, "selectOption")
      .resolves(err(new UserCancelError()));

    await onSwitchM365Tenant(TelemetryTriggerFrom.SideBar);

    chai.assert.isTrue(sendTelemetryEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.args[0][1] instanceof UserCancelError);
  });

  it("Failed to switch tenant", async () => {
    sandbox.stub(M365TokenInstance, "getAccessToken").resolves(ok("faked token"));
    sandbox
      .stub(M365TokenInstance, "switchTenant")
      .resolves(err(new NetworkError("extension", "")));
    sandbox.stub(tool, "listAllTenants").resolves([
      {
        tenantId: "0022fd51-06f5-4557-8a34-69be98de6e20",
        displayName: "MSFT",
        defaultDomain: "t815h.onmicrosoft.com",
      },
      {
        tenantId: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
        displayName: "Cisco",
        defaultDomain: "Cisco561.onmicrosoft.com",
      },
    ]);
    selectOptionStub = sandbox
      .stub(vsc_ui.VS_CODE_UI, "selectOption")
      .resolves(ok({ type: "success" }));

    await onSwitchM365Tenant(TelemetryTriggerFrom.SideBar);

    chai.assert.isTrue(sendTelemetryEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.args[0][1] instanceof NetworkError);
  });

  it("Succeed to switch tenant", async () => {
    sandbox.stub(M365TokenInstance, "getAccessToken").resolves(ok("faked token"));
    sandbox.stub(M365TokenInstance, "switchTenant").resolves(ok("faked token"));
    sandbox.stub(tool, "listAllTenants").resolves([
      {
        tenantId: "0022fd51-06f5-4557-8a34-69be98de6e20",
        displayName: "MSFT",
        defaultDomain: "t815h.onmicrosoft.com",
      },
      {
        tenantId: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
        displayName: "Cisco",
        defaultDomain: "Cisco561.onmicrosoft.com",
      },
    ]);
    selectOptionStub = sandbox
      .stub(vsc_ui.VS_CODE_UI, "selectOption")
      .resolves(ok({ type: "success" }));

    await onSwitchM365Tenant(TelemetryTriggerFrom.SideBar);

    chai.assert.isTrue(sendTelemetryEventStub.calledTwice);
    chai.assert.isTrue(sendTelemetryErrorEventStub.notCalled);
    const items = await selectOptionStub.args[0][0].options();
    chai.assert.deepEqual(items, [
      {
        id: "0022fd51-06f5-4557-8a34-69be98de6e20",
        label: "MSFT",
        description: "t815h.onmicrosoft.com",
      },
      {
        id: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
        label: "Cisco",
        description: "Cisco561.onmicrosoft.com",
      },
    ]);
  });
});

describe("onSwitchAzureTenant", () => {
  const sandbox = sinon.createSandbox();
  let sendTelemetryEventStub: sinon.SinonStub;
  let sendTelemetryErrorEventStub: sinon.SinonStub;
  let selectOptionStub: sinon.SinonStub;

  beforeEach(() => {
    sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sendTelemetryErrorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI({} as vscode.ExtensionContext));
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("Failed to retrieve access token", async () => {
    sandbox.stub(azureAccountManager, "getIdentityCredentialAsync").resolves({
      getToken: () => {
        return Promise.resolve(null);
      },
    });
    selectOptionStub = sandbox.stub(vsc_ui.VS_CODE_UI, "selectOption").resolves(
      err({
        name: "switchTenantFailed",
        source: "extension",
        timestamp: new Date(),
        message: "failed",
      })
    );

    await onSwitchAzureTenant(TelemetryTriggerFrom.SideBar);

    chai.assert.isTrue(sendTelemetryEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    try {
      await selectOptionStub.args[0][0].options();
    } catch (e) {
      chai.assert.isTrue(e instanceof SystemError);
    }
  });

  it("User cancelled", async () => {
    sandbox.stub(azureAccountManager, "getIdentityCredentialAsync").resolves({
      getToken: () => {
        return Promise.resolve(null);
      },
    });
    selectOptionStub = sandbox
      .stub(vsc_ui.VS_CODE_UI, "selectOption")
      .resolves(err(new UserCancelError()));

    await onSwitchAzureTenant(TelemetryTriggerFrom.SideBar);

    chai.assert.isTrue(sendTelemetryEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
  });

  it("Failed to switch tenant", async () => {
    sandbox.stub(azureAccountManager, "getIdentityCredentialAsync").resolves({
      getToken: () => {
        return Promise.resolve({ token: "faked token", expiresOnTimestamp: 0 });
      },
    });
    sandbox.stub(tool, "listAllTenants").resolves([
      {
        tenantId: "0022fd51-06f5-4557-8a34-69be98de6e20",
        displayName: "MSFT",
        defaultDomain: "t815h.onmicrosoft.com",
      },
      {
        tenantId: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
        displayName: "Cisco",
        defaultDomain: "Cisco561.onmicrosoft.com",
      },
    ]);
    selectOptionStub = sandbox
      .stub(vsc_ui.VS_CODE_UI, "selectOption")
      .resolves(ok({ type: "success" }));
    const switchTenantStub = sandbox
      .stub(azureAccountManager, "switchTenant")
      .resolves(err(LoginFailureError()));

    await onSwitchAzureTenant(TelemetryTriggerFrom.SideBar);

    chai.assert.isTrue(sendTelemetryEventStub.calledOnce);
    chai.assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    chai.assert.isTrue(selectOptionStub.calledOnce);
    chai.assert.isTrue(switchTenantStub.calledOnce);
  });

  it("Succeed to switch tenant", async () => {
    sandbox.stub(azureAccountManager, "getIdentityCredentialAsync").resolves({
      getToken: () => {
        return Promise.resolve({ token: "faked token", expiresOnTimestamp: 0 });
      },
    });
    sandbox.stub(tool, "listAllTenants").resolves([
      {
        tenantId: "0022fd51-06f5-4557-8a34-69be98de6e20",
        displayName: "MSFT",
        defaultDomain: "t815h.onmicrosoft.com",
      },
      {
        tenantId: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
        displayName: "Cisco",
        defaultDomain: "Cisco561.onmicrosoft.com",
      },
    ]);
    selectOptionStub = sandbox
      .stub(vsc_ui.VS_CODE_UI, "selectOption")
      .resolves(ok({ type: "success" }));
    const switchTenantStub = sandbox.stub(azureAccountManager, "switchTenant").resolves(
      ok({
        getToken: () => {
          return Promise.resolve(null);
        },
      })
    );

    await onSwitchAzureTenant(TelemetryTriggerFrom.SideBar);

    chai.assert.isTrue(sendTelemetryEventStub.calledTwice);
    chai.assert.isTrue(sendTelemetryErrorEventStub.notCalled);
    chai.assert.isTrue(selectOptionStub.calledOnce);
    chai.assert.isTrue(switchTenantStub.calledOnce);
    const items = await selectOptionStub.args[0][0].options();
    chai.assert.deepEqual(items, [
      {
        id: "0022fd51-06f5-4557-8a34-69be98de6e20",
        label: "MSFT",
        description: "t815h.onmicrosoft.com",
      },
      {
        id: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
        label: "Cisco",
        description: "Cisco561.onmicrosoft.com",
      },
    ]);
  });
});
