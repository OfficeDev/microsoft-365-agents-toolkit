// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import axios, { AxiosResponse } from "axios";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import fs from "fs-extra";
import "mocha";
import mockFs from "mock-fs";
import * as path from "path";
import Sinon, * as sinon from "sinon";
import { getProjectMetadata } from "../../src/common/projectSettingsHelper";
import * as telemetry from "../../src/common/telemetry";
import {
  getSPFxToken,
  getSideloadingStatus,
  listAllTenants,
  listDevTunnels,
  isTestToolEnabledProject,
} from "../../src/common/tools";
import { PackageService } from "../../src/component/m365/packageService";
import { isVideoFilterProject } from "../../src/core/middleware/videoFilterAppBlocker";
import { isUserCancelError } from "../../src/error/common";
import { MockTools } from "../core/utils";

chai.use(chaiAsPromised);

describe("tools", () => {
  describe("getSideloadingStatus()", () => {
    let mockGet: () => AxiosResponse;
    let events: number;
    let errors: number;

    beforeEach(() => {
      sinon.restore();

      const mockInstance = axios.create();
      sinon.stub(mockInstance, "get").callsFake(async () => mockGet());
      sinon.stub(axios, "create").returns(mockInstance);

      events = 0;
      sinon.stub(telemetry, "sendTelemetryEvent").callsFake(() => {
        ++events;
      });

      errors = 0;
      sinon.stub(telemetry, "sendTelemetryErrorEvent").callsFake(() => {
        ++errors;
      });
    });

    afterEach(() => {
      sinon.restore();
    });

    it("sideloading enabled", async () => {
      mockGet = () => {
        return {
          status: 200,
          data: {
            value: {
              isSideloadingAllowed: true,
            },
          },
        } as AxiosResponse;
      };

      const result = await getSideloadingStatus("fake-token");

      chai.assert.isDefined(result);
      chai.assert.isTrue(result);
      chai.assert.equal(events, 1);
      chai.assert.equal(errors, 0);
    });

    it("sideloading not enabled", async () => {
      mockGet = () => {
        return {
          status: 200,
          data: {
            value: {
              isSideloadingAllowed: false,
            },
          },
        } as AxiosResponse;
      };

      const result = await getSideloadingStatus("fake-token");

      chai.assert.isDefined(result);
      chai.assert.isFalse(result);
      chai.assert.equal(events, 1);
      chai.assert.equal(errors, 0);
    });

    it("sideloading unknown", async () => {
      mockGet = () => {
        return {
          status: 200,
          data: {
            value: {
              foo: "bar",
            },
          },
        } as AxiosResponse;
      };

      const result = await getSideloadingStatus("fake-token");

      chai.assert.isUndefined(result);
      chai.assert.equal(events, 0);
      chai.assert.equal(errors, 1);
    });

    it("error and retry", async () => {
      mockGet = () => {
        throw new Error("test");
      };
      const clock = sinon.useFakeTimers();

      const resultPromise = getSideloadingStatus("fake-token");
      await clock.tickAsync(100000);
      const result = await resultPromise;
      clock.restore();

      chai.assert.isUndefined(result);
      chai.assert.equal(events, 0);
      chai.assert.equal(errors, 1);
    });
  });

  describe("listAllTenants", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("returns empty for invalid token", async () => {
      const tenants = await listAllTenants("");

      chai.assert.equal(tenants.length, 0);
    });

    it("returns empty when API call failure", async () => {
      sandbox.stub(axios, "get").throws({ name: 404, message: "failed" });

      const tenants = await listAllTenants("faked token");

      chai.assert.equal(tenants.length, 0);
    });

    it("returns tenant list", async () => {
      const fakedTenants = {
        data: {
          value: [
            {
              tenantId: "0022fd51-06f5-4557-8a34-69be98de6e20",
              countryCode: "SG",
              displayName: "MSFT",
            },
            {
              tenantId: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
              countryCode: "CN",
              displayName: "Cisco",
            },
          ],
        },
      };
      sandbox.stub(axios, "get").resolves(fakedTenants);

      const tenants = await listAllTenants("faked token");

      chai.assert.equal(tenants, fakedTenants.data.value);
    });
  });

  describe("getCopilotStatus", () => {
    let mockGet: () => AxiosResponse;
    let errors: number;
    beforeEach(() => {
      sinon.restore();

      const mockInstance = axios.create();
      sinon.stub(mockInstance, "get").callsFake(async () => mockGet());
      sinon.stub(axios, "create").returns(mockInstance);

      errors = 0;
      sinon.stub(telemetry, "sendTelemetryErrorEvent").callsFake(() => {
        ++errors;
      });
    });

    it("copilot status unknown", async () => {
      mockGet = () => {
        return {
          status: 200,
          data: {
            value: {
              foo: "bar",
            },
          },
        } as AxiosResponse;
      };

      const result = await PackageService.GetSharedInstance().getCopilotStatus("fake-token");

      chai.assert.isUndefined(result);
      chai.assert.equal(errors, 1);
    });
  });

  describe("getProjectMetadata", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("happy path V3", async () => {
      try {
        sandbox.stub<any, any>(fs, "readFileSync").callsFake((file: string) => {
          return `version: 1.0.0
projectId: 00000000-0000-0000-0000-000000000000`;
        });
        sandbox.stub<any, any>(fs, "pathExistsSync").callsFake((file: string) => {
          return true;
        });
        const result = getProjectMetadata("root-path");
        chai.assert.isNotEmpty(result);
        chai.assert.equal(result!.projectId, "00000000-0000-0000-0000-000000000000");
      } finally {
      }
    });

    it("project settings not exists", async () => {
      sandbox.stub<any, any>(fs, "pathExistsSync").callsFake((file: string) => {
        return false;
      });
      const result = getProjectMetadata("root-path");
      chai.assert.isUndefined(result);
    });

    it("throw error", async () => {
      sandbox.stub<any, any>(fs, "pathExistsSync").callsFake((file: string) => {
        throw new Error("new error");
      });
      const result = getProjectMetadata("root-path");
      chai.assert.isUndefined(result);
    });

    it("empty root path", async () => {
      const result = getProjectMetadata("");
      chai.assert.isUndefined(result);
    });
  });

  describe("isVideoFilterProject", async () => {
    let sandbox: Sinon.SinonSandbox;
    const mockProjectRoot = "video-filter";
    beforeEach(() => {
      sandbox = sinon.createSandbox();
    });
    afterEach(() => {
      sandbox.restore();
      mockFs.restore();
    });

    it("Can recognize normal video filter project", async () => {
      // Arrange
      const manifest = {
        meetingExtensionDefinition: {
          videoFiltersConfigurationUrl: "https://a.b.c/",
        },
      };
      mockFs({
        [path.join(mockProjectRoot, "appPackage", "manifest.json")]: JSON.stringify(manifest),
      });

      // Act
      const result = await isVideoFilterProject(mockProjectRoot);

      // Assert
      chai.expect(result.isOk()).to.be.true;
      chai.expect(result._unsafeUnwrap()).to.be.true;
    });

    it("Should not recognize tab project as video filter", async () => {
      // Arrange
      const manifest = {
        $schema:
          "https://developer.microsoft.com/en-us/json-schemas/teams/v1.14/MicrosoftTeams.schema.json",
        manifestVersion: "1.14",
        version: "1.0.0",
        id: "{{state.fx-resource-appstudio.teamsAppId}}",
        packageName: "com.microsoft.teams.extension",
        developer: {
          name: "Teams App, Inc.",
          websiteUrl: "https://www.example.com",
          privacyUrl: "https://www.example.com/termofuse",
          termsOfUseUrl: "https://www.example.com/privacy",
        },
        icons: {
          color: "{{config.manifest.icons.color}}",
          outline: "{{config.manifest.icons.outline}}",
        },
        name: {
          short: "{{config.manifest.appName.short}}",
          full: "{{config.manifest.appName.full}}",
        },
        description: {
          short: "{{config.manifest.description.short}}",
          full: "{{config.manifest.description.full}}",
        },
        accentColor: "#FFFFFF",
        bots: [],
        composeExtensions: [],
        configurableTabs: [
          {
            configurationUrl:
              "{{{state.fx-resource-frontend-hosting.endpoint}}}{{{state.fx-resource-frontend-hosting.indexPath}}}/config",
            canUpdateConfiguration: true,
            scopes: ["team", "groupchat"],
          },
        ],
        staticTabs: [
          {
            entityId: "index0",
            name: "Personal Tab",
            contentUrl:
              "{{{state.fx-resource-frontend-hosting.endpoint}}}{{{state.fx-resource-frontend-hosting.indexPath}}}/tab",
            websiteUrl:
              "{{{state.fx-resource-frontend-hosting.endpoint}}}{{{state.fx-resource-frontend-hosting.indexPath}}}/tab",
            scopes: ["personal"],
          },
        ],
        permissions: ["identity", "messageTeamMembers"],
        validDomains: ["{{state.fx-resource-frontend-hosting.domain}}"],
      };
      mockFs({
        [path.join(mockProjectRoot, "appPackage", "manifest.json")]: JSON.stringify(manifest),
      });

      // Act
      const result = await isVideoFilterProject(mockProjectRoot);

      // Assert
      chai.expect(result.isOk()).to.be.true;
      chai.expect(result._unsafeUnwrap()).to.be.false;
    });
  });

  describe("getSPFxToken", async () => {
    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      const mockTools = new MockTools();
      sinon.stub(mockTools.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("xxx"));
      sinon.stub(axios, "get").resolves({ data: { webUrl: "122" } });
      const res = await getSPFxToken(mockTools.tokenProvider.m365TokenProvider);
    });
  });

  describe("listDevTunnels", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });

    it("should return an error when the API call fails", async () => {
      const token = "test-token";

      const result = await listDevTunnels(token);
      chai.assert.isTrue(result.isErr());
    });
  });

  describe("listDevTunnels using github token", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });

    it("should return an error when the API call fails", async () => {
      const token = "test-token";

      const result = await listDevTunnels(token, true);
      chai.assert.isTrue(result.isErr());
    });
  });

  describe("isUserCancelError()", () => {
    it("should return true if error is UserCancelError", () => {
      const error = new Error();
      error.name = "UserCancelError";
      chai.expect(isUserCancelError(error)).is.true;
    });
  });

  describe("isTestToolEnabledProject", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("should return true if test tool YAML file exists", () => {
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const result = isTestToolEnabledProject("test-project-path");
      chai.expect(result).to.be.true;
    });

    it("should return false if test tool YAML file does not exist", () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);
      const result = isTestToolEnabledProject("test-project-path");
      chai.expect(result).to.be.false;
    });
  });
});
