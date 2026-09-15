// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios from "axios";
import AdmZip from "adm-zip";
import { afterEach, beforeEach, chai, describe, expect, it, vi } from "vitest";
import { devPortalClient } from "../../src/client/devPortalClient";
import {
  getActiveTeamsDevPortalClient,
  teamsDevPortalClient,
} from "../../src/client/teamsDevPortalClientProvider";
import { legacyTeamsDevPortalClient } from "../../src/client/teamsDevPortalClient";
import { FeatureFlagName } from "../../src/common/featureFlags";
import { setTools } from "../../src/common/globalVars";
import { RetryHandler } from "../../src/common/retryHandler";
import { WrappedAxiosClient } from "../../src/common/wrappedAxiosClient";
import type { AppUser } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appUser";
import { MockTools } from "../core/utils";

describe("teamsDevPortalClientProvider", () => {
  beforeEach(() => {
    setTools(new MockTools());
    RetryHandler.RETRIES = 1;
  });

  afterEach(() => {
    vi.restoreAllMocks();
    RetryHandler.RETRIES = 6;
    delete process.env[FeatureFlagName.NewDeveloperPortalApis];
    legacyTeamsDevPortalClient.regionEndpoint = undefined;
    devPortalClient.regionEndpoint = undefined;
  });

  it("selects the client using the new Developer Portal API feature flag", () => {
    chai.assert.strictEqual(getActiveTeamsDevPortalClient(), legacyTeamsDevPortalClient);

    process.env[FeatureFlagName.NewDeveloperPortalApis] = "true";

    chai.assert.strictEqual(getActiveTeamsDevPortalClient(), devPortalClient);
  });

  it("rejects new-only operations when the feature flag is disabled", () => {
    expect(() => teamsDevPortalClient.createApp("token", "app")).toThrow(
      "The operation requires the new Developer Portal APIs."
    );
  });

  it("supports assigning an own proxy property", () => {
    const property = Symbol("provider-test");
    const facade = teamsDevPortalClient as unknown as Record<PropertyKey, unknown>;
    Object.defineProperty(teamsDevPortalClient, property, {
      configurable: true,
      writable: true,
      value: "initial",
    });

    try {
      facade[property] = "updated";

      chai.assert.equal(facade[property], "updated");
    } finally {
      delete facade[property];
    }
  });

  it("restores concrete client properties when a proxy property is deleted", () => {
    legacyTeamsDevPortalClient.regionEndpoint = "https://legacy.example.com";
    devPortalClient.regionEndpoint = "https://new.example.com";

    Object.defineProperty(teamsDevPortalClient, "regionEndpoint", {
      configurable: true,
      writable: true,
      value: "https://temporary.example.com",
    });
    delete teamsDevPortalClient.regionEndpoint;

    chai.assert.equal(legacyTeamsDevPortalClient.regionEndpoint, "https://legacy.example.com");
    chai.assert.equal(devPortalClient.regionEndpoint, "https://new.example.com");
  });

  it("sets the legacy regional endpoint returned by AuthSvc", async () => {
    const requester = axios.create();
    vi.spyOn(WrappedAxiosClient, "create").mockReturnValue(requester);
    vi.spyOn(RetryHandler, "Retry").mockResolvedValue({
      data: { regionGtms: { teamsDevPortal: "https://legacy.example.com" } },
    });

    await legacyTeamsDevPortalClient.setRegionEndpointByToken("token");

    chai.assert.equal(legacyTeamsDevPortalClient.regionEndpoint, "https://legacy.example.com");
  });

  it("imports an app through the legacy endpoint", async () => {
    const requester = axios.create();
    const app = { teamsAppId: "app-id" };
    vi.spyOn(legacyTeamsDevPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const post = vi.spyOn(requester, "post").mockResolvedValue({ data: app });

    const result = await legacyTeamsDevPortalClient.importApp("token", Buffer.from("package"));

    chai.assert.strictEqual(result, app);
    expect(post).toHaveBeenCalledWith(
      "/api/appdefinitions/v2/import",
      expect.any(Buffer),
      expect.objectContaining({ params: { overwriteIfAppAlreadyExists: false } })
    );
  });

  it("wraps an empty legacy import response", async () => {
    const requester = axios.create();
    vi.spyOn(legacyTeamsDevPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "post").mockResolvedValue(undefined);

    await expect(
      legacyTeamsDevPortalClient.importApp("token", Buffer.from("package"))
    ).rejects.toBeDefined();
  });

  it("reports an invalid manifest app ID returned by the legacy import API", async () => {
    const requester = axios.create();
    const zip = new AdmZip();
    zip.addFile("manifest.json", Buffer.from(JSON.stringify({ id: "not-a-guid" })));
    vi.spyOn(legacyTeamsDevPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "post").mockRejectedValue({
      response: { status: 400, data: "App Id must be a GUID" },
    });

    await expect(
      legacyTeamsDevPortalClient.importApp("token", zip.toBuffer())
    ).rejects.toBeDefined();
  });

  it("uses legacy app CRUD endpoints", async () => {
    const requester = axios.create();
    const app = { teamsAppId: "app-id" };
    legacyTeamsDevPortalClient.regionEndpoint = "https://legacy.example.com";
    vi.spyOn(legacyTeamsDevPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const get = vi
      .spyOn(requester, "get")
      .mockResolvedValueOnce({ data: [app] })
      .mockResolvedValueOnce({ data: app });
    const remove = vi.spyOn(requester, "delete").mockResolvedValue({ data: true });

    chai.assert.deepEqual(await legacyTeamsDevPortalClient.listApps("token"), [app]);
    chai.assert.strictEqual(await legacyTeamsDevPortalClient.getApp("token", "app-id"), app);
    chai.assert.isTrue(await legacyTeamsDevPortalClient.deleteApp("token", "app-id"));
    expect(get.mock.calls.map((call) => call[0])).toEqual([
      "/api/appdefinitions",
      "/api/appdefinitions/app-id",
    ]);
    expect(remove).toHaveBeenCalledWith("/api/appdefinitions/app-id");
  });

  it("updates permissions through the legacy owner endpoint", async () => {
    const requester = axios.create();
    const existingUser = { aadId: "existing-user" } as AppUser;
    const newUser = { aadId: "new-user" } as AppUser;
    vi.spyOn(legacyTeamsDevPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(legacyTeamsDevPortalClient, "getApp")
      .mockResolvedValueOnce({ teamsAppId: "app-id", userList: [] })
      .mockResolvedValueOnce({ teamsAppId: "app-id", userList: [] });
    const post = vi.spyOn(requester, "post").mockResolvedValue({
      data: { teamsAppId: "app-id", userList: [newUser] },
    });

    await legacyTeamsDevPortalClient.removePermission("token", "app-id", existingUser);
    await legacyTeamsDevPortalClient.grantPermission("token", "app-id", newUser);

    expect(post).toHaveBeenCalledWith(
      "/api/appdefinitions/app-id/owner",
      expect.objectContaining({ userList: [newUser] })
    );
  });

  it.each([
    ["createApp", () => devPortalClient.createApp("token", "app")],
    ["updateApp", () => devPortalClient.updateApp("token", "app-id", Buffer.from("package"))],
    ["listApps", () => devPortalClient.listApps("token")],
  ])("wraps an empty response from %s", async (_operation, invoke) => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "post").mockResolvedValue(undefined);
    vi.spyOn(requester, "put").mockResolvedValue(undefined);
    vi.spyOn(requester, "get").mockResolvedValue(undefined);

    await expect(invoke()).rejects.toBeDefined();
  });

  it("reads paged validation history from the continuation response header", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    const get = vi
      .spyOn(requester, "get")
      .mockResolvedValueOnce({
        data: { items: [{ appValidationId: "validation-1" }] },
        headers: { "x-continuation-token": "next-page" },
      })
      .mockResolvedValueOnce({ data: { items: [{ appValidationId: "validation-2" }] } });

    const result = await devPortalClient.getAppValidationRequestList("token", "app-id");

    expect(result.appValidations.map((item) => item.id)).toEqual(["validation-1", "validation-2"]);
    expect(get.mock.calls[1][1]).toMatchObject({
      headers: { "x-ms-continuation": "next-page" },
    });
  });

  it("wraps an empty validation submission response", async () => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "post").mockResolvedValue(undefined);

    await expect(
      devPortalClient.submitAppValidationRequest("token", "app-id")
    ).rejects.toBeDefined();
  });

  it.each([
    [
      "createBotRegistration",
      () =>
        devPortalClient.createBotRegistration(
          "token",
          { botId: "bot-id", name: "bot", description: "", iconUrl: "" },
          false
        ),
    ],
    [
      "updateBotRegistration",
      () =>
        devPortalClient.updateBotRegistration("token", {
          botId: "bot-id",
          name: "bot",
          description: "",
          iconUrl: "",
        }),
    ],
    ["getBotRegistration", () => devPortalClient.getBotRegistration("token", "bot-id")],
  ])("handles an unsuccessful %s response", async (_operation, invoke) => {
    const requester = axios.create();
    vi.spyOn(devPortalClient, "createRequesterWithToken").mockReturnValue(requester);
    vi.spyOn(requester, "post").mockResolvedValue({ status: 400 });
    vi.spyOn(requester, "put").mockResolvedValue({ status: 400 });
    vi.spyOn(requester, "get").mockResolvedValue({ status: 400 });

    await expect(invoke()).rejects.toBeDefined();
  });
});
