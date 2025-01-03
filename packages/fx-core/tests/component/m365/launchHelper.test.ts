// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok, ManifestProperties } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import "mocha";
import sinon from "sinon";
import { NotExtendedToM365Error } from "../../../src/component/m365/errors";
import { LaunchHelper } from "../../../src/component/m365/launchHelper";
import { PackageService } from "../../../src/component/m365/packageService";
import { HubTypes } from "../../../src/question";
import { outlookCopilotAppId } from "../../../src/component/m365/constants";
import { MockedM365Provider } from "../../core/utils";
import mockedEnv, { RestoreFn } from "mocked-env";
import { FeatureFlagName } from "../../../src";

describe("LaunchHelper", () => {
  const m365TokenProvider = new MockedM365Provider();
  const launchHelper = new LaunchHelper(m365TokenProvider);

  afterEach(() => {
    sinon.restore();
  });

  describe("getLaunchUrl", () => {
    it("getLaunchUrl: Teams, signed in", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      const properties: ManifestProperties = {
        capabilities: ["staticTab"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.teams, "test-id", properties);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://teams.microsoft.com/l/app/test-id?installAppPackage=true&webjoin=true&tenantId=test-tid&appTenantId=test-tid&login_hint=test-upn"
      );
    });

    it("getLaunchUrl: Teams, signed in, copilot plugin", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      const properties: ManifestProperties = {
        capabilities: ["plugin"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.teams, "test-id", properties, true);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://teams.microsoft.com/?tenantId=test-tid&appTenantId=test-tid&login_hint=test-upn"
      );
    });

    it("getLaunchUrl: Teams, signed in, copilot plugin + staticTab", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      const properties: ManifestProperties = {
        capabilities: ["MessageExtension", "staticTab"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.teams, "test-id", properties, true);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://teams.microsoft.com/l/app/test-id?installAppPackage=true&webjoin=true&tenantId=test-tid&appTenantId=test-tid&login_hint=test-upn"
      );
    });

    it("getLaunchUrl: Teams, signed in, copilot plugin + configurableTab", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      const properties: ManifestProperties = {
        capabilities: ["MessageExtension", "configurableTab"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.teams, "test-id", properties, true);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://teams.microsoft.com/l/app/test-id?installAppPackage=true&webjoin=true&tenantId=test-tid&appTenantId=test-tid&login_hint=test-upn"
      );
    });

    it("getLaunchUrl: Teams, signed in, copilot plugin + bot", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      const properties: ManifestProperties = {
        capabilities: ["MessageExtension", "Bot", "plugin"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.teams, "test-id", properties, true);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://teams.microsoft.com/l/app/test-id?installAppPackage=true&webjoin=true&tenantId=test-tid&appTenantId=test-tid&login_hint=test-upn"
      );
    });

    it("Teams, signed out", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
        })
      );
      const properties: ManifestProperties = {
        capabilities: ["staticTab"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.teams, "test-id", properties);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://teams.microsoft.com/l/app/test-id?installAppPackage=true&webjoin=true&login_hint=login_your_m365_account"
      );
    });

    it("Outlook, staticTab, acquired, signed in", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      const properties: ManifestProperties = {
        capabilities: ["staticTab"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      sinon.stub(LaunchHelper.prototype, <any>"getM365AppId").resolves(ok("test-app-id"));
      const result = await launchHelper.getLaunchUrl(HubTypes.outlook, "test-id", properties);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://outlook.office.com/host/test-app-id?login_hint=test-upn"
      );
    });

    it("Outlook, staticTab, unacquired, signed in", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      sinon.stub(LaunchHelper.prototype, <any>"getM365AppId").resolves(err({ foo: "bar" }));
      const properties: ManifestProperties = {
        capabilities: ["staticTab"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.outlook, "test-id", properties);
      chai.assert(result.isErr());
      chai.assert.deepEqual((result as any).error, { foo: "bar" });
    });

    it("Outlook, Bot, signed in", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      sinon.stub(LaunchHelper.prototype, <any>"getM365AppId").resolves(ok("test-app-id"));
      const properties: ManifestProperties = {
        capabilities: ["Bot"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.outlook, "test-id", properties);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://outlook.office.com/mail?login_hint=test-upn"
      );
    });

    it("Outlook, signed in", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      sinon.stub(LaunchHelper.prototype, <any>"getM365AppId").resolves(ok("test-app-id"));
      const properties: ManifestProperties = {
        capabilities: ["Bot"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.office, "test-id", properties);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://www.office.com/m365apps/test-app-id?auth=2&login_hint=test-upn"
      );
    });

    it("Outlook, copilot extension", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      const properties: ManifestProperties = {
        capabilities: ["plugin"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      sinon.stub(LaunchHelper.prototype, <any>"getM365AppId").resolves(ok("test-app-id"));
      const result = await launchHelper.getLaunchUrl(HubTypes.outlook, "test-id", properties);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        `https://outlook.office.com/host/${outlookCopilotAppId}?login_hint=test-upn`
      );
    });

    it("Office, copilot extension", async () => {
      sinon.stub(m365TokenProvider, "getStatus").resolves(
        ok({
          status: "",
          accountInfo: {
            tid: "test-tid",
            upn: "test-upn",
          },
        })
      );
      sinon.stub(LaunchHelper.prototype, <any>"getM365AppId").resolves(ok("test-app-id"));
      const properties: ManifestProperties = {
        capabilities: ["copilotGpt"],
        id: "test-id",
        version: "1.0.0",
        manifestVersion: "1.16",
        isApiME: false,
        isSPFx: false,
        isApiMeAAD: false,
      };
      const result = await launchHelper.getLaunchUrl(HubTypes.office, "test-id", properties);
      chai.assert(result.isOk());
      chai.assert.equal(
        (result as any).value,
        "https://www.office.com/chat?auth=2&login_hint=test-upn"
      );
    });
  });

  describe("getM365AppId", () => {
    it("getAccessToken error", async () => {
      sinon.stub(m365TokenProvider, "getAccessToken").resolves(err({ foo: "bar" } as any));
      const result = await launchHelper.getM365AppId("test-id");
      chai.assert(result.isErr());
      chai.assert.deepEqual((result as any).error, { foo: "bar" });
    });

    it("retrieveAppId 404", async () => {
      sinon.stub(m365TokenProvider, "getAccessToken").resolves(ok(""));
      sinon
        .stub(PackageService.prototype, "retrieveAppId")
        .rejects(new NotExtendedToM365Error("test"));
      const result = await launchHelper.getM365AppId("test-id");
      chai.assert(result.isErr());
      chai.assert.deepEqual((result as any).error.name, "NotExtendedToM365Error");
    });

    it("retrieveAppId undefined", async () => {
      sinon.stub(m365TokenProvider, "getAccessToken").resolves(ok(""));
      sinon.stub(PackageService.prototype, "retrieveAppId").resolves(undefined);
      const result = await launchHelper.getM365AppId("test-id");
      chai.assert(result.isErr());
      chai.assert.deepEqual((result as any).error.name, "NotExtendedToM365Error");
    });

    it("happy path", async () => {
      sinon.stub(m365TokenProvider, "getAccessToken").resolves(ok(""));
      sinon.stub(PackageService.prototype, "retrieveAppId").resolves("test-app-id");
      const result = await launchHelper.getM365AppId("test-id");
      chai.assert(result.isOk());
      chai.assert.deepEqual((result as any).value, "test-app-id");
    });
  });
});
