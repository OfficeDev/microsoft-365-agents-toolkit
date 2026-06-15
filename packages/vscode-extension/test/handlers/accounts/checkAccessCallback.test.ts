import { ok } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as vscode from "vscode";
import { PanelType } from "../../../src/controls/PanelType";
import { WebviewPanel } from "../../../src/controls/webviewPanel";
import { vi } from "vitest";
import { mockValue } from "../../mocks/vitestMockUtils";
import {
  checkCopilotCallback,
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
});
