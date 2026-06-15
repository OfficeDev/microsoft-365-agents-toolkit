import { ok } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as vscode from "vscode";
import { PanelType } from "../../../src/controls/PanelType";
import { WebviewPanel } from "../../../src/controls/webviewPanel";
import { vi } from "vitest";
import { mockValue } from "../../mocks/vitestMockUtils";
import {
  checkCopilotCallback,
  checkSandboxCallback,
  checkSideloadingCallback,
} from "../../../src/handlers/accounts/checkAccessCallback";
import { checkAccessCallbackDeps } from "../../../src/handlers/accounts/checkAccessCallback";
import * as vsc_ui from "../../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../../src/telemetry/extTelemetry";
import * as localizeUtils from "../../../src/utils/localizeUtils";

describe("checkAccessCallback", () => {
  describe("checkCopilotCallback", () => {
    beforeEach(() => {
      mockValue(vsc_ui, "VS_CODE_UI", new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    });

    it("checkCopilotCallback() and open url", async () => {
      vi.spyOn(checkAccessCallbackDeps, "localize").mockReturnValue("Enroll");
      vi.spyOn(checkAccessCallbackDeps, "sendTelemetryEvent");
      const showMessageStub = vi
        .spyOn(checkAccessCallbackDeps, "showMessage")
        .mockResolvedValue(ok("Enroll"));
      const openUrlStub = vi.spyOn(checkAccessCallbackDeps, "openUrl");

      await checkCopilotCallback();

      chai.expect(showMessageStub.callCount).to.be.equal(1);
      chai.expect(openUrlStub.callCount).to.be.equal(1);
    });

    it("checkCopilotCallback() and fail to open url", async () => {
      vi.spyOn(checkAccessCallbackDeps, "localize").mockReturnValue("");
      vi.spyOn(checkAccessCallbackDeps, "sendTelemetryEvent");
      const showMessageStub = vi
        .spyOn(checkAccessCallbackDeps, "showMessage")
        .mockResolvedValue(ok("Enroll"));
      const openUrlStub = vi.spyOn(checkAccessCallbackDeps, "openUrl");

      await checkCopilotCallback();

      chai.expect(showMessageStub.callCount).to.be.equal(1);
      chai.expect(openUrlStub.callCount).to.be.equal(0);
    });

    it("checkCopilotCallback() and fail to show message", async () => {
      const localizeStub = vi.spyOn(checkAccessCallbackDeps, "localize").mockReturnValue("");
      vi.spyOn(checkAccessCallbackDeps, "sendTelemetryEvent");
      const showMessageStub = vi
        .spyOn(checkAccessCallbackDeps, "showMessage")
        .mockRejectedValue(new Error("error"));

      await checkCopilotCallback();

      chai.expect(showMessageStub.callCount).to.be.equal(1);
      chai.expect(localizeStub.callCount).to.be.equal(2);
    });
  });

  describe("CheckSideloading", () => {
    let clock: ReturnType<typeof vi.useFakeTimers>;

    afterEach(() => {
      if (clock) {
        clock.restore();
      }
      clock.restore();
      vi.restoreAllMocks();
    });

    beforeEach(() => {
      vi.spyOn(checkAccessCallbackDeps, "sendTelemetryEvent");
      mockValue(vsc_ui, "VS_CODE_UI", new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    });

    it("checkSideloadingCallback() - click enable custom app upload button", async () => {
      const showMessageStub = vi
        .spyOn(checkAccessCallbackDeps, "showMessage")
        .mockResolvedValue(ok("Enable Custom App Upload"));
      const openUrlStub = vi.spyOn(checkAccessCallbackDeps, "openUrl");

      clock = vi.useFakeTimers();
      await checkSideloadingCallback();
      await clock.tickAsync(5000);

      expect(showMessageStub).toHaveBeenCalledTimes(1);
      expect(openUrlStub).toHaveBeenCalledTimes(1);
      expect(openUrlStub).toHaveBeenCalledWith(
        "https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/tools-prerequisites#enable-custom-app-upload-using-admin-center"
      );
    });

    it("checkSideloadingCallback() - click use test tenant button", async () => {
      const showMessageStub = vi
        .spyOn(checkAccessCallbackDeps, "showMessage")
        .mockResolvedValue(ok("Use Test Tenant"));
      const createOrShow = vi.spyOn(checkAccessCallbackDeps, "createOrShow");

      clock = vi.useFakeTimers();
      await checkSideloadingCallback();
      await clock.tickAsync(5000);

      expect(showMessageStub).toHaveBeenCalledTimes(1);
      expect(createOrShow).toHaveBeenCalledTimes(1);
      expect(createOrShow).toHaveBeenCalledWith(PanelType.AccountHelp);
    });
  });

  describe("checkSandboxCallback", () => {
    beforeEach(() => {
      mockValue(vsc_ui, "VS_CODE_UI", new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    });

    it("executes quick open command when user confirms", async () => {
      vi.spyOn(checkAccessCallbackDeps, "localize").mockImplementation((key: string) => {
        if (key === "teamstoolkit.accountTree.sandboxedTeam.button") {
          return "Debug in Sandbox";
        }
        return key;
      });
      vi.spyOn(checkAccessCallbackDeps, "showMessage").mockResolvedValue(ok("Debug in Sandbox"));
      const executeCommandStub = vi
        .spyOn(checkAccessCallbackDeps, "executeCommand")
        .mockResolvedValue(undefined as any);

      await checkSandboxCallback();

      chai.expect(executeCommandStub.calledOnce).to.be.true;
    });

    it("does not execute command when user skips", async () => {
      vi.spyOn(checkAccessCallbackDeps, "localize").mockReturnValue("Debug in Sandbox");
      vi.spyOn(checkAccessCallbackDeps, "showMessage").mockResolvedValue(ok("Skip"));
      const executeCommandStub = vi
        .spyOn(checkAccessCallbackDeps, "executeCommand")
        .mockResolvedValue(undefined as any);

      await checkSandboxCallback();

      chai.expect(executeCommandStub.called).to.be.false;
    });
  });

  describe("checkAccessCallbackDeps delegation", () => {
    beforeEach(() => {
      mockValue(vsc_ui, "VS_CODE_UI", new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
    });

    it("showMessage and openUrl delegate to VS_CODE_UI", async () => {
      const showMessageStub = vi.spyOn(vsc_ui.VS_CODE_UI, "showMessage").mockResolvedValue(ok(""));
      const openUrlStub = vi.spyOn(vsc_ui.VS_CODE_UI, "openUrl").mockResolvedValue(undefined);

      await checkAccessCallbackDeps.showMessage("info", "msg", false, "ok");
      await checkAccessCallbackDeps.openUrl("https://example.com");

      chai.expect(showMessageStub.calledOnce).to.be.true;
      chai.expect(openUrlStub.calledOnce).to.be.true;
    });

    it("localize and sendTelemetryEvent delegate", () => {
      const localizeStub = vi.spyOn(localizeUtils, "localize").mockReturnValue("x");
      const telemetryStub = vi.spyOn(ExtTelemetry, "sendTelemetryEvent");

      const localized = checkAccessCallbackDeps.localize("key");
      checkAccessCallbackDeps.sendTelemetryEvent("evt");

      chai.expect(localized).to.equal("x");
      chai.expect(localizeStub.calledOnce).to.be.true;
      chai.expect(telemetryStub.calledOnce).to.be.true;
    });

    it("createOrShow delegates to WebviewPanel", () => {
      const createOrShowStub = vi.spyOn(WebviewPanel, "createOrShow").mockImplementation(() => {});

      checkAccessCallbackDeps.createOrShow(PanelType.AccountHelp);

      chai.expect(createOrShowStub.calledOnce).to.be.true;
    });
  });
});
