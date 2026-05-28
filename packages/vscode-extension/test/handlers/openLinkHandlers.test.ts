import { ok } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import M365TokenInstance from "../../src/commonlib/m365Login";
import { signedIn, signedOut } from "../../src/commonlib/common/constant";
import { DeveloperPortalHomeLink } from "../../src/constants";
import {
  openAccountLinkHandler,
  openAppManagement,
  openAzureAccountHandler,
  openBotManagement,
  openDevelopmentLinkHandler,
  openDocumentHandler,
  openDocumentLinkHandler,
  openEnvLinkHandler,
  openExternalHandler,
  openHelpFeedbackLinkHandler,
  openLifecycleLinkHandler,
  openM365AccountHandler,
  openReportIssues,
  openResourceGroupInPortal,
  openSubscriptionInPortal,
} from "../../src/handlers/openLinkHandlers";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as envTreeUtils from "../../src/utils/envTreeUtils";
import * as localizeUtils from "../../src/utils/localizeUtils";
import { MockCore } from "../mocks/mockCore";
import { TelemetryTriggerFrom } from "../../src/telemetry/extTelemetryEvents";

describe("Open link handlers", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent").resolves();
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent").resolves();
    sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("openAppManagement", async () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("open link with loginHint", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      sandbox.stub(globalVariables, "core").value(new MockCore());
      sandbox.stub(M365TokenInstance, "getStatus").resolves(
        ok({
          status: signedIn,
          token: undefined,
          accountInfo: { upn: "test" },
        })
      );
      const openUrl = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));

      const res = await openAppManagement();

      chai.assert.isTrue(openUrl.calledOnce);
      chai.assert.isTrue(res.isOk());
      chai.assert.equal(openUrl.args[0][0], `${DeveloperPortalHomeLink}?login_hint=test`);
    });

    it("open link without loginHint", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      sandbox.stub(M365TokenInstance, "getStatus").resolves(
        ok({
          status: signedOut,
          token: undefined,
          accountInfo: { upn: "test" },
        })
      );
      const openUrl = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));

      const res = await openAppManagement();

      chai.assert.isTrue(openUrl.calledOnce);
      chai.assert.isTrue(res.isOk());
      chai.assert.equal(openUrl.args[0][0], DeveloperPortalHomeLink);
    });
  });

  describe("openEnvLinkHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openEnvLinkHandler([]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openDevelopmentLinkHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDevelopmentLinkHandler([]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openDocumentHandler", () => {
    it("opens upgrade guide when clicked from sidebar", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
      const openUrl = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));

      await openDocumentHandler(TelemetryTriggerFrom.SideBar, "learnmore");

      chai.assert.isTrue(openUrl.calledOnceWith("https://aka.ms/teams-toolkit-5.0-upgrade"));
    });
  });

  describe("openLifecycleLinkHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openLifecycleLinkHandler([]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openHelpFeedbackLinkHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openHelpFeedbackLinkHandler([]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openM365AccountHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openM365AccountHandler();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openAzureAccountHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openAzureAccountHandler();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openBotManagement", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openBotManagement();
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openAccountLinkHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openAccountLinkHandler([]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openReportIssues", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openReportIssues([]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openExternalHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openExternalHandler([{ url: "abc" }]);
      chai.assert.isTrue(res.isOk());
    });
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openExternalHandler([]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openDocumentHandler", () => {
    it("happy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentHandler(["", ""]);
      chai.assert.isTrue(res.isOk());
    });
    it("happy learnmore", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentHandler(["", "learnmore"]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openDocumentLinkHandler", () => {
    it("signinAzure", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([{ contextValue: "signinAzure" }]);
      chai.assert.isTrue(res.isOk());
    });
    it("fx-extension.create", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([{ contextValue: "fx-extension.create" }]);
      chai.assert.isTrue(res.isOk());
    });
    it("fx-extension.provision", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([{ contextValue: "fx-extension.provision" }]);
      chai.assert.isTrue(res.isOk());
    });
    it("fx-extension.build", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([{ contextValue: "fx-extension.build" }]);
      chai.assert.isTrue(res.isOk());
    });
    it("fx-extension.deploy", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([{ contextValue: "fx-extension.deploy" }]);
      chai.assert.isTrue(res.isOk());
    });
    it("fx-extension.publish", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([{ contextValue: "fx-extension.publish" }]);
      chai.assert.isTrue(res.isOk());
    });
    it("fx-extension.publishInDeveloperPortal", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([
        { contextValue: "fx-extension.publishInDeveloperPortal" },
      ]);
      chai.assert.isTrue(res.isOk());
    });
    it("empty", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([]);
      chai.assert.isTrue(res.isOk());
    });
    it("none", async () => {
      sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl").resolves(ok(true));
      const res = await openDocumentLinkHandler([{ contextValue: "" }]);
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("openSubscriptionInPortal", () => {
    it("subscriptionInfo not found", async () => {
      sandbox.stub(envTreeUtils, "getSubscriptionInfoFromEnv");
      const res = await openSubscriptionInPortal("local");
      chai.assert.equal(res.isErr() ? res.error.name : "Not Error", "EnvResourceInfoNotFoundError");
    });

    it("happy path", async () => {
      sandbox.stub(envTreeUtils, "getSubscriptionInfoFromEnv").returns({
        subscriptionName: "subscriptionName",
        subscriptionId: "subscriptionId",
        tenantId: "tenantId",
      } as any);
      const openExternalStub = sandbox.stub(vscode.env, "openExternal");
      await openSubscriptionInPortal("local");
      chai.assert.equal(openExternalStub.callCount, 1);
      chai.assert.deepEqual(
        openExternalStub.args[0][0],
        vscode.Uri.parse(
          `https://portal.azure.com/#@tenantId/resource/subscriptions/subscriptionId`
        )
      );
    });
  });

  describe("openResourceGroupInPortal", () => {
    it("subscriptionInfo not found", async () => {
      sandbox.stub(localizeUtils, "localize").returns("Unable to load %s info for environment %s.");
      sandbox.stub(envTreeUtils, "getSubscriptionInfoFromEnv");
      sandbox.stub(envTreeUtils, "getResourceGroupNameFromEnv").returns("resourceGroupName" as any);
      const res = await openResourceGroupInPortal("local");
      chai.assert.equal(
        res.isErr() ? res.error.message : "Not Error",
        "Unable to load Subscription info for environment local."
      );
    });

    it("resourceGroupName not found", async () => {
      sandbox.stub(localizeUtils, "localize").returns("Unable to load %s info for environment %s.");
      sandbox.stub(envTreeUtils, "getSubscriptionInfoFromEnv").returns({
        subscriptionName: "subscriptionName",
        subscriptionId: "subscriptionId",
        tenantId: "tenantId",
      } as any);
      sandbox.stub(envTreeUtils, "getResourceGroupNameFromEnv");
      const res = await openResourceGroupInPortal("local");
      chai.assert.equal(
        res.isErr() ? res.error.message : "Not Error",
        "Unable to load Resource Group info for environment local."
      );
    });

    it("subscriptionInfo and resourceGroupName not found", async () => {
      sandbox.stub(envTreeUtils, "getSubscriptionInfoFromEnv");
      sandbox.stub(envTreeUtils, "getResourceGroupNameFromEnv");
      const res = await openResourceGroupInPortal("local");
      chai.assert.equal(
        res.isErr() ? res.error.message : "Not Error",
        "Unable to load Subscription and Resource Group info for environment local."
      );
    });

    it("happy path", async () => {
      sandbox.stub(envTreeUtils, "getSubscriptionInfoFromEnv").returns({
        subscriptionName: "subscriptionName",
        subscriptionId: "subscriptionId",
        tenantId: "tenantId",
      } as any);
      sandbox.stub(envTreeUtils, "getResourceGroupNameFromEnv").returns("resourceGroupName" as any);
      const openExternalStub = sandbox.stub(vscode.env, "openExternal");
      await openResourceGroupInPortal("local");
      chai.assert.equal(openExternalStub.callCount, 1);
      chai.assert.deepEqual(
        openExternalStub.args[0][0],
        vscode.Uri.parse(
          `https://portal.azure.com/#@tenantId/resource/subscriptions/subscriptionId/resourceGroups/resourceGroupName`
        )
      );
    });
  });
});
