import { ok } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import { PanelType } from "../../../src/controls/PanelType";
import { WebviewPanel } from "../../../src/controls/webviewPanel";
import {
  checkCopilotCallback,
  checkSideloadingCallback,
} from "../../../src/handlers/accounts/checkAccessCallback";
import * as vsc_ui from "../../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../../src/telemetry/extTelemetry";
import * as localizeUtils from "../../../src/utils/localizeUtils";

describe("checkAccessCallback", () => {
  describe("checkCopilotCallback", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    beforeEach(() => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    });

    it("checkCopilotCallback() and open url", async () => {
      sandbox.stub(localizeUtils, "localize").returns("Enroll");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const showMessageStub = sandbox.stub(vsc_ui.VS_CODE_UI, "showMessage").resolves(ok("Enroll"));
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl");

      await checkCopilotCallback();

      chai.expect(showMessageStub.callCount).to.be.equal(1);
      chai.expect(openUrlStub.callCount).to.be.equal(1);
    });

    it("checkCopilotCallback() and fail to open url", async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const showMessageStub = sandbox.stub(vsc_ui.VS_CODE_UI, "showMessage").resolves(ok("Enroll"));
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl");

      await checkCopilotCallback();

      chai.expect(showMessageStub.callCount).to.be.equal(1);
      chai.expect(openUrlStub.callCount).to.be.equal(0);
    });

    it("checkCopilotCallback() and fail to show message", async () => {
      const localizeStub = sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const showMessageStub = sandbox
        .stub(vsc_ui.VS_CODE_UI, "showMessage")
        .rejects(new Error("error"));

      await checkCopilotCallback();

      chai.expect(showMessageStub.callCount).to.be.equal(1);
      chai.expect(localizeStub.callCount).to.be.equal(2);
    });
  });

  describe("CheckSideloading", () => {
    const sandbox = sinon.createSandbox();
    let clock: sinon.SinonFakeTimers;

    afterEach(() => {
      if (clock) {
        clock.restore();
      }
      clock.restore();
      sandbox.restore();
    });

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    });

    it("checkSideloadingCallback() - click enable custom app upload button", async () => {
      const showMessageStub = sandbox
        .stub(vsc_ui.VS_CODE_UI, "showMessage")
        .resolves(ok("Enable Custom App Upload"));
      const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl");

      clock = sandbox.useFakeTimers();
      await checkSideloadingCallback();
      await clock.tickAsync(5000);

      sinon.assert.calledOnce(showMessageStub);
      sinon.assert.calledOnceWithExactly(
        openUrlStub,
        "https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/tools-prerequisites#enable-custom-app-upload-using-admin-center"
      );
    });

    it("checkSideloadingCallback() - click use test tenant button", async () => {
      const showMessageStub = sandbox
        .stub(vsc_ui.VS_CODE_UI, "showMessage")
        .resolves(ok("Use Test Tenant"));
      const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");

      clock = sandbox.useFakeTimers();
      await checkSideloadingCallback();
      await clock.tickAsync(5000);

      sinon.assert.calledOnce(showMessageStub);
      sinon.assert.calledOnceWithExactly(createOrShow, PanelType.AccountHelp);
    });
  });
});
