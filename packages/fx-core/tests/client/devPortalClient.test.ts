// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios from "axios";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { devPortalClient } from "../../src/client/devPortalClient";
import { legacyTeamsDevPortalClient } from "../../src/client/teamsDevPortalClient";
import { teamsDevPortalClient } from "../../src/client/teamsDevPortalClientProvider";
import { FeatureFlagName } from "../../src/common/featureFlags";
import { setTools } from "../../src/common/globalVars";
import { RetryHandler } from "../../src/common/retryHandler";
import { AppUser } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appUser";
import { Constants } from "../../src/component/driver/teamsApp/constants";
import { MockTools } from "../core/utils";
import { AppResponse } from "../../src/client/devPortalModels";

const resourceId = "11111111-1111-4111-8111-111111111111";
const manifestId = "22222222-2222-4222-8222-222222222222";
const timestamp = "2026-10-21T01:23:45.000Z";

describe("native Developer Portal response contract", () => {
  beforeEach(() => {
    setTools(new MockTools());
    RetryHandler.RETRIES = 1;
    process.env[FeatureFlagName.NewDeveloperPortalApis] = "true";
  });

  afterEach(() => {
    vi.restoreAllMocks();
    RetryHandler.RETRIES = 6;
    delete process.env[FeatureFlagName.NewDeveloperPortalApis];
  });

  it("AC-05: create/get/update retain the same nested native response without legacy fields", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const app = {
      appId: resourceId,
      appProfile: {
        appAccessControl: null,
        appMetadata: { ownerAadId: null },
        appDetails: {
          applicationManifest: {
            id: manifestId,
            name: null,
            developer: null,
            bots: null,
            composeExtensions: null,
          },
        },
      },
    };
    vi.spyOn(requester, "post").mockResolvedValue({ data: app });
    vi.spyOn(requester, "get").mockResolvedValue({ data: app });
    vi.spyOn(requester, "put").mockResolvedValue({ data: app });

    expect(await teamsDevPortalClient.createApp("token", "app")).toBe(app);
    expect(await teamsDevPortalClient.getApp("token", resourceId)).toBe(app);
    expect(await teamsDevPortalClient.updateApp("token", resourceId, Buffer.from(""))).toBe(app);
    expect(app).not.toHaveProperty("teamsAppId");
  });

  it.each([null, {}, { appDetails: null }, { appDetails: { applicationManifest: null } }])(
    "AC-05: supports nullable or absent optional native profile data (%j)",
    async (appProfile) => {
      const requester = axios.create();
      vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
      const app = { appId: resourceId, appProfile };
      vi.spyOn(requester, "get").mockResolvedValue({ data: app });
      expect(await teamsDevPortalClient.getApp("token", resourceId)).toBe(app);
    }
  );

  it.each([
    {},
    { appId: null },
    { appId: "" },
    { appId: "invalid" },
    { appId: resourceId },
    { appId: manifestId, appProfile: null },
    { appId: resourceId, appProfile: undefined },
    { appId: resourceId, appProfile: [] },
    { appId: resourceId, appProfile: "invalid" },
  ])("AC-05: rejects invalid or mismatched response identity (%j)", async (app) => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "get").mockResolvedValue({ data: app });
    await expect(teamsDevPortalClient.getApp("token", resourceId)).rejects.toBeDefined();
  });

  it("AC-05/AC-16: list retains native nullable summaries and continuation pagination", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const first = { appId: resourceId, appExternalId: manifestId, appName: null, appVersion: null };
    const second = { appId: manifestId, appName: "Other app", appVersion: null, updatedAt: null };
    const get = vi
      .spyOn(requester, "get")
      .mockResolvedValueOnce({ data: { items: [first], continuationToken: "next" } })
      .mockResolvedValueOnce({ data: { items: [second], continuationToken: null } });
    expect(await teamsDevPortalClient.listApps("token")).toEqual([first, second]);
    expect(get.mock.calls[1]).toEqual([
      "/v1.0/apps",
      { params: { pageSize: 100 }, headers: { "x-ms-continuation": "next" } },
    ]);
  });

  it.each(["true", "false"])(
    "AC-06: equal-ID get uses the saved resource directly (flag %s)",
    async (flag) => {
      process.env[FeatureFlagName.NewDeveloperPortalApis] = flag;
      const active = flag === "true" ? devPortalClient : legacyTeamsDevPortalClient;
      const requester = axios.create();
      vi.spyOn(active, "createRequesterWithToken").mockReturnValue(requester);
      const app =
        flag === "true"
          ? {
              appId: resourceId,
              appProfile: { appDetails: { applicationManifest: { id: resourceId } } },
            }
          : { teamsAppId: resourceId, appId: resourceId };
      const get = vi.spyOn(requester, "get").mockResolvedValue({ data: app });
      expect(await teamsDevPortalClient.getApp("token", resourceId)).toBe(app);
      expect(get).toHaveBeenCalledExactlyOnceWith(
        flag === "true" ? `/v1.0/apps/${resourceId}` : `/api/appdefinitions/${resourceId}`
      );
    }
  );

  it.each(["getApp", "getAppPackage", "deleteApp"] as const)(
    "AC-16: %s does not scan list or retry with a manifest ID after a 404",
    async (method) => {
      const requester = axios.create();
      vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
      const get = vi.spyOn(requester, "get").mockRejectedValue({ response: { status: 404 } });
      const remove = vi.spyOn(requester, "delete").mockRejectedValue({ response: { status: 404 } });
      await expect(teamsDevPortalClient[method]("token", resourceId)).rejects.toBeDefined();
      expect([...get.mock.calls, ...remove.mock.calls]).toHaveLength(1);
    }
  );

  it("AC-10: freshness uses only the summary with the exact resource appId", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const get = vi.spyOn(requester, "get").mockResolvedValue({
      data: {
        items: [
          { appId: manifestId, appExternalId: resourceId, updatedAt: "2000-01-01T00:00:00Z" },
          { appId: resourceId, appExternalId: manifestId, updatedAt: timestamp },
        ],
      },
    });
    expect(await teamsDevPortalClient.getAppLastModified("token", resourceId)).toBe(timestamp);
    expect(get.mock.calls.map((call) => call[0])).toEqual(["/v1.0/apps"]);
  });

  it.each([undefined, null, "", "invalid"])(
    "AC-10: missing or invalid summary timestamp fails (%j)",
    async (updatedAt) => {
      const requester = axios.create();
      vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
      vi.spyOn(requester, "get").mockResolvedValue({
        data: { items: [{ appId: resourceId, updatedAt }] },
      });
      await expect(
        teamsDevPortalClient.getAppLastModified("token", resourceId)
      ).rejects.toBeDefined();
    }
  );

  it("AC-10: an external-ID-only match cannot provide freshness", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "get").mockResolvedValue({
      data: { items: [{ appId: manifestId, appExternalId: resourceId, updatedAt: timestamp }] },
    });
    await expect(
      teamsDevPortalClient.getAppLastModified("token", resourceId)
    ).rejects.toBeDefined();
  });

  it("AC-10: legacy freshness retains the detail updatedAt source", async () => {
    process.env[FeatureFlagName.NewDeveloperPortalApis] = "false";
    vi.spyOn(legacyTeamsDevPortalClient, "getApp").mockResolvedValue({
      teamsAppId: resourceId,
      updatedAt: timestamp,
    });
    expect(await teamsDevPortalClient.getAppLastModified("token", resourceId)).toBe(timestamp);
  });

  it("AC-05: permissions consume native appAccessControl.appUsers and retain ETag", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const owner = { aadId: "owner", tenantId: "tenant", role: "Owner" };
    const user: AppUser = {
      aadId: "member",
      tenantId: "tenant",
      isAdministrator: false,
      displayName: "",
      userPrincipalName: "",
    };
    vi.spyOn(requester, "get").mockResolvedValue({
      data: { appId: resourceId, appProfile: { appAccessControl: { appUsers: [owner] } } },
      headers: { etag: "etag" },
    });
    const post = vi.spyOn(requester, "post").mockResolvedValue({ status: 200 });
    await teamsDevPortalClient.grantPermission("token", resourceId, user);
    expect(post).toHaveBeenCalledWith(
      `/v1.0/apps/${resourceId}/owners`,
      { appUsers: [owner, { aadId: "member", tenantId: "tenant", role: "Member" }] },
      { headers: { "If-Match": "etag" } }
    );
  });

  it("AC-05: native user listing preserves actual fields without fabricating names", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const users = [{ aadId: "owner", tenantId: "tenant", role: "Owner" }];
    vi.spyOn(requester, "get").mockResolvedValue({
      data: { appId: resourceId, appProfile: { appAccessControl: { appUsers: users } } },
    });
    expect(await teamsDevPortalClient.getUserList("token", resourceId)).toBe(users);
    expect(await teamsDevPortalClient.checkPermission("token", resourceId, "owner")).toBe(
      Constants.PERMISSIONS.admin
    );
    expect(users[0]).not.toHaveProperty("displayName");
    expect(users[0]).not.toHaveProperty("userPrincipalName");
    expect(users[0]).not.toHaveProperty("isAdministrator");
  });

  it("AC-05: populated native draft capabilities and metadata round-trip without mutation", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const app = {
      appId: resourceId,
      appProfile: {
        appAccessControl: { appUsers: null },
        appMetadata: {
          ownerAadId: "owner",
          updatedByAadId: "editor",
          source: "DeveloperPortal",
          appType: 1,
          appPublishInfos: [],
        },
        appDetails: {
          applicationManifest: {
            id: manifestId,
            manifestVersion: "1.24",
            version: "1.0.0",
            name: { short: "App", full: null },
            description: { short: "Description", full: null },
            developer: {
              name: "Developer",
              websiteUrl: "https://example.com",
              privacyUrl: null,
              termsOfUseUrl: null,
              mpnId: null,
            },
            icons: { color: "color.png", outline: null },
            webApplicationInfo: { id: resourceId, resource: null },
            bots: [{ botId: manifestId, scopes: ["personal"] }],
            staticTabs: [
              {
                entityId: "tab",
                name: "Tab",
                contentUrl: null,
                websiteUrl: null,
                scopes: ["personal"],
                context: ["personalTab"],
              },
            ],
            configurableTabs: [
              {
                configurationUrl: null,
                scopes: ["team"],
                context: ["channelTab"],
                canUpdateConfiguration: true,
                supportedSharePointHosts: ["sharePointFullPage"],
              },
            ],
            composeExtensions: [{ botId: manifestId, composeExtensionType: "botBased" }],
            connectors: [{ connectorId: manifestId, scopes: ["team"] }],
            meetingExtensionDefinition: {
              scenes: [{ id: "scene" }],
              supportsStreaming: false,
              supportsAnonymousGuestUsers: null,
            },
            activities: {
              activityTypes: [
                { type: "activity", description: "Activity", templateText: "Template" },
              ],
            },
            supportsChannelFeatures: "tier1",
            copilotAgents: {
              declarativeAgents: [{ id: "declarative", file: "declarativeAgent.json" }],
              customEngineAgents: [{ id: manifestId, type: "bot" }],
            },
          },
        },
      },
    };
    vi.spyOn(requester, "get").mockResolvedValue({ data: app });
    expect(await teamsDevPortalClient.getApp("token", resourceId)).toBe(app);
  });

  it.each([
    null,
    {},
    { appDetails: null, appAccessControl: null },
    { appDetails: {}, appAccessControl: {} },
    { appDetails: { applicationManifest: null }, appAccessControl: { appUsers: null } },
    { appDetails: { applicationManifest: {} }, appAccessControl: { appUsers: [] } },
    { appDetails: { applicationManifest: { bots: [] } } },
    { appDetails: { applicationManifest: { bots: [{ botId: null }] } } },
  ] satisfies AppResponse["appProfile"][])(
    "AC-05: missing draft bots and users are treated as absent (%j)",
    async (appProfile) => {
      const requester = axios.create();
      vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
      vi.spyOn(requester, "get").mockResolvedValue({ data: { appId: resourceId, appProfile } });
      expect(await teamsDevPortalClient.getBotId("token", resourceId)).toBeUndefined();
      expect(await teamsDevPortalClient.checkPermission("token", resourceId, "missing")).toBe(
        Constants.PERMISSIONS.noPermission
      );
      expect(
        teamsDevPortalClient.checkUser(
          { appId: resourceId, appProfile },
          {
            aadId: "missing",
            tenantId: "tenant",
            isAdministrator: false,
            displayName: "Missing",
            userPrincipalName: "missing@example.com",
          }
        )
      ).toBe(false);
    }
  );

  it("AC-05: membership checks work for both response shapes independently of the selected API", async () => {
    const user: AppUser = {
      aadId: "member",
      tenantId: "tenant",
      isAdministrator: false,
      displayName: "Member",
      userPrincipalName: "member@example.com",
    };
    const app: AppResponse = {
      appId: resourceId,
      appProfile: { appAccessControl: { appUsers: [{ ...user, role: "Member" }] } },
    };
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "get").mockResolvedValue({ data: app });
    expect(await teamsDevPortalClient.checkPermission("token", resourceId, "member")).toBe(
      Constants.PERMISSIONS.operative
    );
    for (const flag of ["true", "false"]) {
      process.env[FeatureFlagName.NewDeveloperPortalApis] = flag;
      expect(teamsDevPortalClient.checkUser(app, user)).toBe(true);
      expect(
        teamsDevPortalClient.checkUser({ teamsAppId: resourceId, userList: [user] }, user)
      ).toBe(true);
    }
  });

  it("AC-05: permission lookup failure never grants access", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "get").mockRejectedValue(new Error("lookup failed"));
    expect(await teamsDevPortalClient.checkPermission("token", resourceId, "owner")).toBe(
      Constants.PERMISSIONS.noPermission
    );
  });

  it.each([null, {}, { appAccessControl: null }, { appAccessControl: { appUsers: null } }])(
    "AC-05: ownership updates handle an absent owner list (%j)",
    async (appProfile) => {
      const requester = axios.create();
      vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
      vi.spyOn(requester, "get").mockResolvedValue({ data: { appId: resourceId, appProfile } });
      const post = vi.spyOn(requester, "post").mockResolvedValue({ status: 200 });
      const user: AppUser = {
        aadId: "owner",
        tenantId: "tenant",
        isAdministrator: true,
        displayName: "Owner",
        userPrincipalName: "owner@example.com",
      };
      await teamsDevPortalClient.grantPermission("token", resourceId, user);
      expect(post.mock.calls[0][1]).toEqual({
        appUsers: [{ aadId: "owner", tenantId: "tenant", role: "Owner" }],
      });
      await teamsDevPortalClient.removePermission("token", resourceId, user);
      expect(post).toHaveBeenCalledTimes(1);
    }
  );

  it("AC-05: bot lookup reads the nested draft manifest", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "get").mockResolvedValue({
      data: {
        appId: resourceId,
        appProfile: { appDetails: { applicationManifest: { bots: [{ botId: manifestId }] } } },
      },
    });
    expect(await teamsDevPortalClient.getBotId("token", resourceId)).toBe(manifestId);
  });

  it.each([
    { appId: resourceId, appProfile: { appAccessControl: { appUsers: [{ aadId: "owner" }] } } },
    { appId: resourceId, appProfile: { appMetadata: { ownerAadId: 42 } } },
    {
      appId: resourceId,
      appProfile: { appDetails: { applicationManifest: { name: { short: 42 } } } },
    },
    {
      appId: resourceId,
      appProfile: { appDetails: { applicationManifest: { bots: [{ botId: 42 }] } } },
    },
    {
      appId: resourceId,
      appProfile: { appDetails: { applicationManifest: { staticTabs: [null] } } },
    },
    { appId: resourceId, appProfile: { appDetails: { applicationManifest: { id: "invalid" } } } },
    { appId: resourceId, appProfile: null, teamsAppId: manifestId },
  ])("AC-05: create/get/update do not validate unrelated profile fields (%j)", async (app) => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "post").mockResolvedValue({ data: app });
    vi.spyOn(requester, "get").mockResolvedValue({ data: app });
    vi.spyOn(requester, "put").mockResolvedValue({ data: app });

    expect(await teamsDevPortalClient.createApp("token", "app")).toBe(app);
    expect(await teamsDevPortalClient.getApp("token", resourceId)).toBe(app);
    expect(await teamsDevPortalClient.updateApp("token", resourceId, Buffer.from(""))).toBe(app);
  });

  it.each([
    null,
    {},
    { appId: "invalid", appProfile: null },
    { appId: resourceId },
    { appId: resourceId, appProfile: undefined },
    { appId: resourceId, appProfile: [] },
    { appId: resourceId, appProfile: "invalid" },
  ])("AC-05: create and update reject malformed response identity (%j)", async (app) => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "post").mockResolvedValue({ data: app });
    vi.spyOn(requester, "put").mockResolvedValue({ data: app });
    await expect(teamsDevPortalClient.createApp("token", "app")).rejects.toBeDefined();
    await expect(
      teamsDevPortalClient.updateApp("token", resourceId, Buffer.from(""))
    ).rejects.toBeDefined();
  });

  it("AC-05: update rejects a response for a different resource", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "put").mockResolvedValue({
      data: { appId: manifestId, appProfile: null },
    });
    await expect(
      teamsDevPortalClient.updateApp("token", resourceId, Buffer.from(""))
    ).rejects.toBeDefined();
  });

  it.each([
    { items: [{ appId: "invalid" }] },
    { items: [{ appId: resourceId, appName: 42 }] },
    { items: [], continuationToken: 42 },
    { items: null },
  ])("AC-05: list rejects malformed response identities and fields (%j)", async (data) => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "get").mockResolvedValue({ data });
    await expect(teamsDevPortalClient.listApps("token")).rejects.toBeDefined();
  });

  it.each([undefined, "invalid"])(
    "AC-10: invalid legacy freshness also fails (%j)",
    async (updatedAt) => {
      process.env[FeatureFlagName.NewDeveloperPortalApis] = "false";
      vi.spyOn(legacyTeamsDevPortalClient, "getApp").mockResolvedValue({
        teamsAppId: resourceId,
        updatedAt,
      });
      await expect(
        teamsDevPortalClient.getAppLastModified("token", resourceId)
      ).rejects.toBeDefined();
    }
  );
});
