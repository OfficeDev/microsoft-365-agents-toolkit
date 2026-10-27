// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { mkdtemp, readFile, remove, writeFile } from "fs-extra";
import os from "os";
import path from "path";
import { Platform, err, ok, UserError } from "@microsoft/teamsfx-api";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import Container from "typedi";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import { RetryHandler } from "../../../../src/common/retryHandler";
import { YamlParser } from "../../../../src/component/configManager/parser";
import "../../../../src/component/driver/index";
import { UpdateDcrDriver } from "../../../../src/component/driver/dcr/update";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";

const configurationId = "existing-dcr-id";
const resourcePath = `/v1.0/oAuthConfigurations/${configurationId}`;

function snapshot() {
  return {
    oAuthConfigId: configurationId,
    identityProvider: "Dynamic",
    clientName: "Existing client",
    m365AppId: "existing-app",
    applicableToApps: "SpecificApp",
    targetAudience: "HomeTenant",
    targetUrlsShouldStartWith: ["https://example.com/mcp"],
    mcpResourceUrl: "https://example.com/mcp",
    wellKnownAuthorizationServer:
      "https://login.example.com/.well-known/oauth-authorization-server",
    resource: "urn:example:resource",
    supportedAccountTypes: "Enterprise",
    scopes: ["existing-scope"],
    tokenExchangeMethodType: "PostRequestBody",
    isPKCEEnabled: true,
    authorizationEndpoint: "https://login.example.com/authorize",
    clientId: "service-managed-client",
    clientSecret: "not-a-real-secret",
  };
}

describe("dcr/update HTTP-boundary integration (verified routes, provisional broad field set)", () => {
  let http: MockAdapter;
  let context: DriverContext;
  let ui: MockedUserInteraction;
  let driver: UpdateDcrDriver;
  const retries = RetryHandler.RETRIES;

  beforeEach(() => {
    const requester = axios.create();
    http = new MockAdapter(requester);
    vi.spyOn(axios, "create").mockReturnValue(requester);
    RetryHandler.RETRIES = 1;
    vi.stubEnv(FeatureFlagName.SovereignCloudEnvironment, undefined);
    ui = new MockedUserInteraction();
    vi.spyOn(ui, "confirm").mockRejectedValue(
      new Error("DCR update must not request confirmation")
    );
    context = {
      m365TokenProvider: new MockedM365Provider(),
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui,
      logProvider: new MockedLogProvider(),
      progressBar: undefined,
      projectPath: "",
      platform: Platform.VSCode,
    };
    vi.spyOn(context.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("test-token"));
    http.onGet(resourcePath).reply(200, snapshot());
    http.onPatch(resourcePath).reply(204);
    driver = new UpdateDcrDriver();
  });

  afterEach(() => {
    http.restore();
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    RetryHandler.RETRIES = retries;
  });

  function execute(input: Record<string, unknown> = {}) {
    return driver.execute({ configurationId, ...input }, context);
  }

  function expectPatch(body: Record<string, unknown>) {
    expect(http.history.patch).toHaveLength(1);
    expect(http.history.patch[0].url).toBe(resourcePath);
    expect(JSON.parse(http.history.patch[0].data)).toEqual(body);
    expect(http.history.post).toHaveLength(0);
  }

  it("DCR-UPD-05: enables MSA with a sparse payload and no name or credential writes", async () => {
    const confirm = vi.spyOn(ui, "confirm");
    const result = await execute({ includePersonalMicrosoftAccounts: true });
    expect(result.result.isOk()).toBe(true);
    if (result.result.isOk()) expect(result.result.value.size).toBe(0);
    expectPatch({ supportedAccountTypes: "Enterprise, Consumer" });
    expect(result.summaries.join()).toContain(configurationId);
    expect(confirm).not.toHaveBeenCalled();
    expect(result.summaries.join()).toContain("includePersonalMicrosoftAccounts");
    expect(result.summaries.join()).not.toContain("not-a-real-secret");
  });

  it("DCR-UPD-06: disables MSA without changing other fields", async () => {
    http.onGet(resourcePath).reply(200, {
      ...snapshot(),
      supportedAccountTypes: "Enterprise, Consumer",
    });
    expect((await execute({ includePersonalMicrosoftAccounts: false })).result.isOk()).toBe(true);
    expectPatch({ supportedAccountTypes: "Enterprise" });
  });

  for (const { id, input, body } of [
    { id: "DCR-UPD-07", input: { name: "Renamed" }, body: { clientName: "Renamed" } },
    {
      id: "DCR-UPD-08",
      input: { name: "Entirely new name" },
      body: { clientName: "Entirely new name" },
    },
    {
      id: "DCR-UPD-10",
      input: { mcpResourceUrl: "https://new.example.com/full/mcp/path" },
      body: { mcpResourceUrl: "https://new.example.com/full/mcp/path" },
    },
    {
      id: "DCR-UPD-11",
      input: { wellKnownAuthorizationServer: "https://new.example.com/metadata" },
      body: { wellKnownAuthorizationServer: "https://new.example.com/metadata" },
    },
    {
      id: "DCR-UPD-12",
      input: { resource: "urn:example:new-resource" },
      body: { resource: "urn:example:new-resource" },
    },
    { id: "DCR-UPD-14", input: { appId: "new-app" }, body: { m365AppId: "new-app" } },
    {
      id: "DCR-UPD-17",
      input: { targetAudience: "AnyTenant" },
      body: { targetAudience: "AnyTenant" },
    },
  ]) {
    it(`${id}: submits only the requested ${Object.keys(input)[0]} change`, async () => {
      const result = await execute(input);
      expect(result.result.isOk()).toBe(true);
      expectPatch(body);
    });
  }

  for (const targets of [
    [],
    ["https://new.example.com/mcp"],
    ["https://one.example.com", "https://two.example.com"],
  ]) {
    it(`DCR-UPD-13: replaces targets with ${targets.length} entries`, async () => {
      expect((await execute({ targetUrlsShouldStartWith: targets })).result.isOk()).toBe(true);
      expectPatch({ targetUrlsShouldStartWith: targets });
    });
  }

  for (const input of [
    {},
    { name: "Existing client" },
    { includePersonalMicrosoftAccounts: false },
  ]) {
    it(`DCR-UPD-09: skips confirmation and update for matching ${JSON.stringify(input)}`, async () => {
      const confirm = vi.spyOn(ui, "confirm");
      const result = await execute(input);
      expect(result.result.isOk()).toBe(true);
      expect(http.history.get).toHaveLength(1);
      expect(http.history.patch).toHaveLength(0);
      expect(confirm).not.toHaveBeenCalled();
      expect(result.summaries.join()).toContain(configurationId);
    });
  }

  it("DCR-UPD-15: switches to SpecificApp retaining an existing app ID", async () => {
    http.onGet(resourcePath).reply(200, { ...snapshot(), applicableToApps: "AnyApp" });
    expect((await execute({ applicableToApps: "SpecificApp" })).result.isOk()).toBe(true);
    expectPatch({ applicableToApps: "SpecificApp" });
  });

  it("DCR-UPD-15: requires a binding when switching to SpecificApp", async () => {
    http.onGet(resourcePath).reply(200, {
      ...snapshot(),
      applicableToApps: "AnyApp",
      m365AppId: "",
    });
    const result = await execute({ applicableToApps: "SpecificApp" });
    expect(result.result.isErr()).toBe(true);
    expect(http.history.patch).toHaveLength(0);
    expect(
      (await execute({ applicableToApps: "SpecificApp", appId: "new-app" })).result.isOk()
    ).toBe(true);
    expectPatch({ applicableToApps: "SpecificApp", m365AppId: "new-app" });
  });

  it("DCR-UPD-16: removes the app binding only on an explicit AnyApp transition", async () => {
    expect((await execute({ applicableToApps: "AnyApp" })).result.isOk()).toBe(true);
    expectPatch({ applicableToApps: "AnyApp", m365AppId: "" });
  });

  for (const configurationId of [undefined, "", " ", 42, null]) {
    it(`DCR-UPD-03: rejects invalid ID ${String(configurationId)} before HTTP`, async () => {
      expect((await driver.execute({ configurationId }, context)).result.isErr()).toBe(true);
      expect(http.history.get).toHaveLength(0);
      expect(http.history.patch).toHaveLength(0);
    });
  }

  for (const input of [
    { name: "" },
    { name: "n".repeat(129) },
    { mcpResourceUrl: "http://example.com" },
    { wellKnownAuthorizationServer: "invalid" },
    { resource: "/relative" },
    { targetUrlsShouldStartWith: [42] },
    { targetAudience: "Other" },
    { applicableToApps: "Other" },
    { includePersonalMicrosoftAccounts: "true" },
    { clientId: "static-client" },
  ]) {
    it(`DCR-UPD-04: rejects invalid ${Object.keys(input)[0]} without HTTP`, async () => {
      expect((await execute(input)).result.isErr()).toBe(true);
      expect(http.history.get).toHaveLength(0);
      expect(http.history.patch).toHaveLength(0);
    });
  }

  for (const includePersonalMicrosoftAccounts of [true, false]) {
    it(`DCR-UPD-18: rejects explicit account ${includePersonalMicrosoftAccounts} in sovereign cloud`, async () => {
      vi.stubEnv(FeatureFlagName.SovereignCloudEnvironment, "GCC H");
      expect((await execute({ includePersonalMicrosoftAccounts })).result.isErr()).toBe(true);
      expect(http.history.get).toHaveLength(0);
    });
  }

  for (const status of [403, 404, 500]) {
    it(`DCR-UPD-19: returns GET ${status} without fallback registration`, async () => {
      http
        .onGet(resourcePath)
        .reply(status, { code: "ReadFailed" }, { "x-correlation-id": "read-correlation" });
      const result = await execute({ name: "Renamed" });
      expect(result.result.isErr()).toBe(true);
      if (result.result.isErr()) {
        expect(result.result.error.message).toContain("teams_graph_get_dcr");
        expect(result.result.error.message).toContain("read-correlation");
      }
      expect(http.history.patch).toHaveLength(0);
      expect(http.history.post).toHaveLength(0);
    });
  }

  for (const current of [
    null,
    [],
    {},
    { ...snapshot(), clientName: undefined },
    { ...snapshot(), clientName: "" },
    { ...snapshot(), m365AppId: undefined },
    { ...snapshot(), m365AppId: 42 },
    { ...snapshot(), targetAudience: undefined },
    { ...snapshot(), targetAudience: "Other" },
    { ...snapshot(), targetUrlsShouldStartWith: "https://example.com" },
    { ...snapshot(), targetUrlsShouldStartWith: [42] },
    { ...snapshot(), applicableToApps: "Other" },
    { ...snapshot(), identityProvider: undefined },
    { ...snapshot(), identityProvider: "Custom" },
    { ...snapshot(), identityProvider: "MicrosoftEntra" },
    { ...snapshot(), oAuthConfigId: "another-id" },
    { ...snapshot(), supportedAccountTypes: "Unknown" },
    { ...snapshot(), mcpResourceUrl: null },
    { ...snapshot(), wellKnownAuthorizationServer: 42 },
    { ...snapshot(), resource: [] },
  ]) {
    it(`DCR-UPD-19: rejects invalid or non-Dynamic snapshot ${JSON.stringify(current)?.slice(0, 45)}`, async () => {
      http.onGet(resourcePath).reply(200, current);
      expect((await execute({ name: "Renamed" })).result.isErr()).toBe(true);
      expect(http.history.patch).toHaveLength(0);
    });
  }

  it("DCR-UPD-19: rejects a shared-route snapshot with absent identity", async () => {
    const current: Partial<ReturnType<typeof snapshot>> = snapshot();
    delete current.identityProvider;
    http.onGet(resourcePath).reply(200, current);
    expect((await execute({ includePersonalMicrosoftAccounts: true })).result.isErr()).toBe(true);
    expect(http.history.patch).toHaveLength(0);
    expect(http.history.post).toHaveLength(0);
  });

  it("DCR-UPD-19: accepts an explicitly Dynamic snapshot without optional ID/account fields", async () => {
    const current: Partial<ReturnType<typeof snapshot>> = snapshot();
    delete current.oAuthConfigId;
    delete current.supportedAccountTypes;
    http.onGet(resourcePath).reply(200, current);
    expect((await execute({ includePersonalMicrosoftAccounts: false })).result.isOk()).toBe(true);
    expectPatch({ supportedAccountTypes: "Enterprise" });
  });

  for (const platform of [Platform.VSCode, Platform.CLI]) {
    it(`DCR-UPD-20: updates without a UI on ${platform}`, async () => {
      context.ui = undefined;
      context.platform = platform;
      context.nonInteractive = true;
      const log = vi.spyOn(context.logProvider, "info");
      const result = await execute({ name: "Renamed" });
      expect(result.result.isOk()).toBe(true);
      expectPatch({ clientName: "Renamed" });
      expect(result.summaries.join()).toContain(configurationId);
      expect(JSON.stringify(log.mock.calls)).toContain(configurationId);
    });
  }

  it("DCR-UPD-20: updates with a UI that has no confirm method without opening a step notification", async () => {
    if (!context.ui) throw new Error("Expected the test UI");
    context.ui.confirm = undefined;
    const notification = vi.spyOn(ui, "showMessage");
    const result = await execute({ name: "Renamed" });
    expect(result.result.isOk()).toBe(true);
    expectPatch({ clientName: "Renamed" });
    expect(notification).not.toHaveBeenCalled();
    expect(result.summaries.join()).toContain(configurationId);
  });

  it("DCR-UPD-27: a pending notification cannot block a committed update", async () => {
    const notification = vi.spyOn(ui, "showMessage").mockReturnValue(new Promise<never>(() => {}));
    const result = await Promise.race([
      execute({ includePersonalMicrosoftAccounts: true }),
      new Promise<undefined>((resolve) => setTimeout(() => resolve(undefined), 50)),
    ]);
    expect(result).not.toBeUndefined();
    if (result !== undefined) {
      expect(result.result.isOk()).toBe(true);
      expect(result.summaries.join()).toContain("includePersonalMicrosoftAccounts");
    }
    expect(notification).not.toHaveBeenCalled();
    expectPatch({ supportedAccountTypes: "Enterprise, Consumer" });
  });

  it("DCR-UPD-20: applies devtunnel changes without confirmation or exposing URL query values", async () => {
    const confirm = vi.spyOn(ui, "confirm");
    const log = vi.spyOn(context.logProvider, "info");
    const result = await execute({
      mcpResourceUrl: "https://example.devtunnels.ms/mcp?key=test-only-query-value",
    });
    expect(result.result.isOk()).toBe(true);
    expect(confirm).not.toHaveBeenCalled();
    expect(result.summaries.join()).toContain("mcpResourceUrl");
    expect(result.summaries.join()).not.toContain("test-only-query-value");
    expect(JSON.stringify(log.mock.calls)).not.toContain("test-only-query-value");
  });

  it("DCR-UPD-21: preserves authentication errors without HTTP", async () => {
    const error = new UserError("test", "TokenUnavailable", "Token unavailable");
    vi.spyOn(context.m365TokenProvider, "getAccessToken").mockResolvedValue(err(error));
    const result = await execute({ name: "Renamed" });
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) expect(result.result.error).toBe(error);
    expect(http.history.get).toHaveLength(0);
  });

  it("DCR-UPD-21: surfaces PATCH failure with diagnostics and no success summary", async () => {
    http
      .onPatch(resourcePath)
      .reply(400, { code: "UpdateFailed" }, { "x-correlation-id": "update-correlation" });
    const result = await execute({ targetUrlsShouldStartWith: [] });
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error.message).toContain("teams_graph_update_dcr");
      expect(result.result.error.message).toContain("update-correlation");
    }
    expect(result.summaries).toEqual([]);
    expect(http.history.post).toHaveLength(0);
  });

  it("DCR-UPD-21: reports unexpected authentication failures without mutation", async () => {
    vi.spyOn(context.m365TokenProvider, "getAccessToken").mockRejectedValue(
      new Error("Authentication failed")
    );
    expect((await execute({ name: "Renamed" })).result.isErr()).toBe(true);
    expect(http.history.patch).toHaveLength(0);
  });

  it("DCR-UPD-21: handles non-Error authentication failures explicitly", async () => {
    vi.spyOn(context.m365TokenProvider, "getAccessToken").mockRejectedValue(
      "Authentication unavailable"
    );
    expect((await execute({ name: "Renamed" })).result.isErr()).toBe(true);
    expect(http.history.patch).toHaveLength(0);
  });

  it("DCR-UPD-27: notification cancellation cannot turn a committed update into failure", async () => {
    const error = new UserError("test", "NotificationFailed", "Notification failed");
    vi.spyOn(ui, "showMessage").mockResolvedValue(err(error));
    const result = await execute({ name: "Renamed" });
    expect(result.result.isOk()).toBe(true);
    expectPatch({ clientName: "Renamed" });
    expect(result.summaries.join()).toContain(configurationId);
  });

  it("DCR-UPD-21: preserves the real-service immutable-URL rejection without a success or replacement", async () => {
    http.onPatch(resourcePath).reply(400, {
      error: {
        code: "BadRequest",
        innerError: { code: "ConfigUpdateError" },
        message:
          "The following fields cannot be updated for dynamic configurations: mcpResourceUrl",
      },
    });
    const result = await execute({ mcpResourceUrl: "https://new.example.com/mcp" });
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error.message).toContain("ConfigUpdateError");
    }
    expect(result.summaries).toEqual([]);
    expectPatch({ mcpResourceUrl: "https://new.example.com/mcp" });
  });

  it("DCR-UPD-22: ignores service-managed response state rather than echoing it into the request", async () => {
    http.onPatch(resourcePath).reply(200, { ...snapshot(), scopes: ["rediscovered-scope"] });
    expect((await execute({ mcpResourceUrl: "https://new.example.com/mcp" })).result.isOk()).toBe(
      true
    );
    expectPatch({ mcpResourceUrl: "https://new.example.com/mcp" });
    expect(http.history.get).toHaveLength(1);
  });

  it("DCR-UPD-23: resolves environment inputs and executes the registered action through a parsed lifecycle", async () => {
    vi.stubEnv("DCR_AUTH_ID_TEST", configurationId);
    const directory = await mkdtemp(path.join(os.tmpdir(), "atk-dcr-update-"));
    try {
      const yamlPath = path.join(directory, "m365agents.yml");
      const manifestPath = path.join(directory, "ai-plugin.json");
      const envPath = path.join(directory, ".env.test");
      const manifest = JSON.stringify({
        auth: { type: "OAuthPluginVault", reference_id: configurationId },
      });
      const environment = `DCR_AUTH_ID_TEST=${configurationId}\n`;
      await writeFile(manifestPath, manifest);
      await writeFile(envPath, environment);
      await writeFile(
        yamlPath,
        JSON.stringify({
          version: "v1.13",
          provision: [
            {
              uses: "dcr/update",
              with: {
                configurationId: "${{DCR_AUTH_ID_TEST}}",
                includePersonalMicrosoftAccounts: true,
              },
            },
          ],
        })
      );
      const parsed = await new YamlParser().parse(yamlPath, true);
      expect(parsed.isOk()).toBe(true);
      expect(Container.get("dcr/update")).toBeInstanceOf(UpdateDcrDriver);
      if (parsed.isErr() || !parsed.value.provision)
        throw new Error("Expected a provision lifecycle");
      expect(parsed.value.provision.resolvePlaceholders()).toEqual([]);
      context.projectPath = directory;
      const result = await parsed.value.provision.execute(context);
      expect(result.result.isOk()).toBe(true);
      expectPatch({ supportedAccountTypes: "Enterprise, Consumer" });
      expect(process.env.DCR_AUTH_ID_TEST).toBe(configurationId);
      expect(await readFile(manifestPath, "utf8")).toBe(manifest);
      expect(await readFile(envPath, "utf8")).toBe(environment);
      if (result.result.isOk()) expect(result.result.value.size).toBe(0);
    } finally {
      await remove(directory);
    }
  });

  it("DCR-UPD-23: escapes the configuration ID as a single resource-path segment", async () => {
    const id = "id/with?reserved#characters";
    const escaped = `/v1.0/oAuthConfigurations/${encodeURIComponent(id)}`;
    http.onGet(escaped).reply(200, { ...snapshot(), oAuthConfigId: id });
    http.onPatch(escaped).reply(204);
    expect((await execute({ configurationId: id, name: "Renamed" })).result.isOk()).toBe(true);
    expect(http.history.get[0].url).toBe(escaped);
    expect(http.history.patch[0].url).toBe(escaped);
  });
});
