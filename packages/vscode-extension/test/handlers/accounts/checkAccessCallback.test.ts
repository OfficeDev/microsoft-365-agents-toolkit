import * as sinon from "sinon";
import * as chai from "chai";
import * as localizeUtils from "../../../src/utils/localizeUtils";
import * as vsc_ui from "../../../src/qm/vsc_ui";
import * as vscode from "vscode";
import {
  checkCopilotCallback,
  checkSideloadingCallback,
} from "../../../src/handlers/accounts/checkAccessCallback";
import { ok } from "@microsoft/teamsfx-api";
import { ExtTelemetry } from "../../../src/telemetry/extTelemetry";
import { WebviewPanel } from "../../../src/controls/webviewPanel";
import { PanelType } from "../../../src/controls/PanelType";

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

    afterEach(() => {
      sandbox.restore();
    });

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    });

    it("checkSideloadingCallback()", async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      let showMessageCalledCount = 0;
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        showMessage: async () => {
          showMessageCalledCount += 1;
          return Promise.resolve(ok("Get More Info"));
        },
      });
      const createOrShow = sandbox.stub(WebviewPanel, "createOrShow");

      checkSideloadingCallback();

      chai.expect(showMessageCalledCount).to.be.equal(1);
      sinon.assert.calledOnceWithExactly(createOrShow, PanelType.AccountHelp);
    });
  });
});
