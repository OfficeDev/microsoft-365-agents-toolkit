// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { Platform, UserError, err, ok } from "@microsoft/teamsfx-api";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { RetryHandler } from "../../../../src/common/retryHandler";
import { setTools, TOOLS } from "../../../../src/common/globalVars";
import { CreateDcrDriver } from "../../../../src/component/driver/dcr/create";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { CreateOauthDriver } from "../../../../src/component/driver/oauth/create";
import { CreateOauthArgs } from "../../../../src/component/driver/oauth/interface/createOauthArgs";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";

// These tests exercise only registration contracts, not live tokens or Copilot results.
const MCP_URL =
  // eslint-disable-next-line no-secrets/no-secrets -- Public requester-owned fixture URL.
  "https://ca-mcp-qetyc6ob.jollyfield-18c57c9a.eastus2.azurecontainerapps.io/mcp";
const TENANT_ID = "16801f2e-7163-4bd5-8542-f083fa59993c";
const ENTRA_AUTHORIZATION = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;
const ENTRA_TOKEN = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const SSO_ID = "pre-registered-sso";

type Mode = "static-oauth" | "dcr" | "sso";
const MODES: readonly Mode[] = ["static-oauth", "dcr", "sso"];

function outputKey(mode: Mode): string {
  return `TEST_${mode.replace("-", "_").toUpperCase()}_AUTH_ID`;
}

describe("Remote MCP mode isolation - registration boundary, not end-to-end", () => {
  let http: MockAdapter;
  let context: DriverContext;
  let previousTools = TOOLS;
  const retries = RetryHandler.RETRIES;

  beforeEach(() => {
    const requester = axios.create();
    http = new MockAdapter(requester, { onNoMatch: "throwException" });
    vi.spyOn(axios, "create").mockReturnValue(requester);
    RetryHandler.RETRIES = 1;
    vi.stubEnv("TEAMSFX_SOVEREIGN_CLOUD_ENVIRONMENT", undefined);
    vi.stubEnv("ENTRA_SSO_REGISTRATION_ID", SSO_ID);
    for (const mode of MODES) {
      vi.stubEnv(outputKey(mode), undefined);
    }
    context = {
      m365TokenProvider: new MockedM365Provider(),
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockedUserInteraction(),
      logProvider: new MockedLogProvider(),
      progressBar: undefined,
      projectPath: "",
      platform: Platform.VSCode,
    };
    previousTools = TOOLS;
    setTools({
      ui: new MockedUserInteraction(),
      logProvider: context.logProvider,
      tokenProvider: {
        azureAccountProvider: context.azureAccountProvider,
        m365TokenProvider: context.m365TokenProvider,
      },
    });
    vi.spyOn(context.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("test-only-token"));
    vi.spyOn(context.m365TokenProvider, "getJsonObject").mockResolvedValue(
      ok({ oid: "test-user" })
    );
  });

  afterEach(() => {
    http.restore();
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    setTools(previousTools);
    RetryHandler.RETRIES = retries;
  });

  function oauthArgs(mode: "static-oauth" | "sso"): CreateOauthArgs {
    return {
      name: `${mode} isolated registration`,
      appId: "test-m365-app",
      flow: "authorizationCode",
      baseUrl: MCP_URL,
      applicableToApps: "AnyApp",
      targetAudience: "HomeTenant",
      clientId: "fixture-aad-client",
      ...(mode === "static-oauth"
        ? {
            identityProvider: "Custom",
            clientSecret: "test-only-static-client-secret",
            authorizationUrl: ENTRA_AUTHORIZATION,
            tokenUrl: ENTRA_TOKEN,
            scope: "api://test-mcp/access_as_user",
          }
        : { identityProvider: "MicrosoftEntra" }),
    };
  }

  function execute(mode: Mode) {
    const outputs = new Map([["configurationId", outputKey(mode)]]);
    if (mode === "sso") {
      // Only this branch reads the optional SSO fixture setting.
      vi.stubEnv(outputKey(mode), process.env.ENTRA_SSO_REGISTRATION_ID);
      outputs.set("applicationIdUri", "TEST_SSO_APPLICATION_ID_URI");
      return new CreateOauthDriver().execute(oauthArgs(mode), context, outputs);
    }
    if (mode === "dcr") {
      return new CreateDcrDriver().execute(
        {
          name: "isolated dynamic registration",
          mcpResourceUrl: MCP_URL,
          targetUrlsShouldStartWith: [MCP_URL],
          applicableToApps: "AnyApp",
          targetAudience: "HomeTenant",
        },
        context,
        outputs
      );
    }
    return new CreateOauthDriver().execute(oauthArgs(mode), context, outputs);
  }

  function requestBody(index = 0): object {
    const parsed: unknown = JSON.parse(http.history.post[index].data);
    if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
      throw new Error("Expected a registration request object");
    }
    return parsed;
  }

  function existingSso() {
    return {
      oAuthConfigId: SSO_ID,
      identityProvider: "MicrosoftEntra",
      clientId: "fixture-aad-client",
      clientSecret: "",
      scopes: [],
      applicableToApps: "AnyApp",
      targetAudience: "HomeTenant",
      targetUrlsShouldStartWith: [MCP_URL],
      resourceIdentifierUri: "api://mini-owned-uri",
      redirectUris: ["https://mini.example.com/callback"],
      preAuthorizedApplications: ["mini-owned-preauthorization"],
    };
  }

  describe("DA + static OAuth", () => {
    it("MCP-AUTH-01: creates a Custom registration with Entra endpoints despite an SSO ID", async () => {
      http.onPost("/v1.0/oAuthConfigurations").reply(200, {
        configurationRegistrationId: { oAuthConfigId: "new-static-oauth" },
      });
      const result = await execute("static-oauth");
      expect(result.result.isOk()).toBe(true);
      expect(http.history.get).toHaveLength(0);
      expect(http.history.post).toHaveLength(1);
      expect(requestBody()).toMatchObject({
        identityProvider: "Custom",
        clientId: "fixture-aad-client",
        clientSecret: "test-only-static-client-secret",
        authorizationEndpoint: ENTRA_AUTHORIZATION,
        tokenExchangeEndpoint: ENTRA_TOKEN,
        scopes: ["api://test-mcp/access_as_user"],
      });
      expect(requestBody()).not.toHaveProperty("mcpResourceUrl");
      if (result.result.isOk()) {
        expect(result.result.value.get(outputKey("static-oauth"))).toBe("new-static-oauth");
        expect(result.result.value.get(outputKey("static-oauth"))).not.toBe(SSO_ID);
      }
      expect(process.env[outputKey("sso")]).toBeUndefined();
      expect(process.env[outputKey("dcr")]).toBeUndefined();
    });
  });

  describe("DA + DCR", () => {
    it("MCP-AUTH-02: creates through facade discovery without static Entra credentials or SSO reuse", async () => {
      http.onPost("/v1.0/dynamicConfigurations").reply(200, {
        configurationRegistrationId: { oAuthConfigId: "new-dynamic-registration" },
      });
      const result = await execute("dcr");
      expect(result.result.isOk()).toBe(true);
      expect(http.history.get).toHaveLength(0);
      expect(http.history.post).toHaveLength(1);
      expect(requestBody()).toMatchObject({
        mcpResourceUrl: MCP_URL,
        targetUrlsShouldStartWith: [MCP_URL],
        applicableToApps: "AnyApp",
        targetAudience: "HomeTenant",
      });
      for (const field of [
        "clientId",
        "clientSecret",
        "authorizationEndpoint",
        "tokenExchangeEndpoint",
      ]) {
        expect(requestBody()).not.toHaveProperty(field);
      }
      if (result.result.isOk()) {
        expect(result.result.value.get(outputKey("dcr"))).toBe("new-dynamic-registration");
        expect(result.result.value.get(outputKey("dcr"))).not.toBe(SSO_ID);
      }
      expect(process.env[outputKey("static-oauth")]).toBeUndefined();
      expect(process.env[outputKey("sso")]).toBeUndefined();
    });
  });

  describe("DA + Entra SSO", () => {
    it("MCP-AUTH-03: preserves an opaque padded SSO ID as one encoded URL segment", async () => {
      const id = "opaque/id+padding==";
      vi.stubEnv("ENTRA_SSO_REGISTRATION_ID", id);
      http.onGet(`/v1.0/oAuthConfigurations/${encodeURIComponent(id)}`).reply(200, {
        ...existingSso(),
        oAuthConfigId: id,
      });
      const result = await execute("sso");
      expect(result.result.isOk()).toBe(true);
      expect(http.history.get[0].url).toBe(`/v1.0/oAuthConfigurations/${encodeURIComponent(id)}`);
      expect(process.env[outputKey("sso")]).toBe(id);
      expect(http.history.post).toHaveLength(0);
    });

    it("MCP-AUTH-03: reuses a valid SSO record without creation or mini AAD mutation", async () => {
      const existing = existingSso();
      const before = structuredClone(existing);
      vi.stubEnv("TEST_SSO_APPLICATION_ID_URI", existing.resourceIdentifierUri);
      http.onGet(`/v1.0/oAuthConfigurations/${SSO_ID}`).reply(200, existing);
      const result = await execute("sso");
      expect(result.result.isOk()).toBe(true);
      expect(http.history.get).toHaveLength(1);
      expect(http.history.post).toHaveLength(0);
      expect(http.history.patch).toHaveLength(0);
      expect(context.m365TokenProvider.getJsonObject).not.toHaveBeenCalled();
      expect(process.env.TEST_SSO_APPLICATION_ID_URI).toBe("api://mini-owned-uri");
      expect(existing).toEqual(before);
      if (result.result.isOk()) expect(result.result.value.size).toBe(0);
      expect(process.env[outputKey("static-oauth")]).toBeUndefined();
      expect(process.env[outputKey("dcr")]).toBeUndefined();
    });

    it("MCP-AUTH-04: absent SSO ID retains normal creation and URI output mapping", async () => {
      vi.stubEnv("ENTRA_SSO_REGISTRATION_ID", undefined);
      http.onPost("/v1.0/oAuthConfigurations").reply(200, {
        configurationRegistrationId: { oAuthConfigId: "new-sso-registration" },
        resourceIdentifierUri: "api://new-sso-uri",
      });
      const result = await execute("sso");
      expect(result.result.isOk()).toBe(true);
      expect(http.history.get).toHaveLength(0);
      expect(http.history.post).toHaveLength(1);
      expect(requestBody()).toMatchObject({
        identityProvider: "MicrosoftEntra",
        useSingleSignOn: true,
        clientId: "fixture-aad-client",
        applicableToApps: "AnyApp",
        targetAudience: "HomeTenant",
      });
      expect(requestBody()).not.toHaveProperty("clientSecret");
      if (result.result.isOk()) {
        expect(result.result.value.get(outputKey("sso"))).toBe("new-sso-registration");
        expect(result.result.value.get("TEST_SSO_APPLICATION_ID_URI")).toBe("api://new-sso-uri");
      }
    });

    for (const status of [403, 404, 500]) {
      it(`MCP-AUTH-05: rejects supplied SSO ID on GET ${status} without replacement`, async () => {
        http.onGet(`/v1.0/oAuthConfigurations/${SSO_ID}`).reply(status, { code: "Unavailable" });
        const result = await execute("sso");
        expect(result.result.isErr()).toBe(true);
        expect(result.summaries).toEqual([]);
        expect(http.history.post).toHaveLength(0);
        expect(http.history.patch).toHaveLength(0);
      });
    }

    for (const changed of [
      { identityProvider: "Custom" },
      { identityProvider: "Dynamic" },
      { clientId: "a-different-aad-client" },
    ]) {
      it(`MCP-AUTH-05: rejects incompatible SSO record ${JSON.stringify(changed)}`, async () => {
        http.onGet(`/v1.0/oAuthConfigurations/${SSO_ID}`).reply(200, {
          ...existingSso(),
          ...changed,
        });
        expect((await execute("sso")).result.isErr()).toBe(true);
        expect(http.history.post).toHaveLength(0);
        expect(http.history.patch).toHaveLength(0);
      });
    }
  });

  it("MCP-AUTH-06: preserves legacy Custom reuse errors outside the SSO path", async () => {
    vi.stubEnv(outputKey("static-oauth"), "existing-custom");
    http.onGet("/v1.0/oAuthConfigurations/existing-custom").reply(404);
    expect((await execute("static-oauth")).result.isOk()).toBe(true);
    expect(http.history.post).toHaveLength(0);
    expect(http.history.patch).toHaveLength(0);
  });

  for (const mode of MODES.filter((item) => item !== "sso")) {
    it(`MCP-AUTH-07: ${mode} registration rejection cannot fall back to the SSO ID`, async () => {
      const route = mode === "dcr" ? "/v1.0/dynamicConfigurations" : "/v1.0/oAuthConfigurations";
      http.onPost(route).reply(401, { code: "RegistrationUnauthorized" });
      const result = await execute(mode);
      expect(result.result.isErr()).toBe(true);
      expect(http.history.post).toHaveLength(1);
      expect(http.history.get).toHaveLength(0);
      expect(result.summaries).toEqual([]);
    });
  }

  for (const mode of MODES) {
    it(`MCP-AUTH-07: ${mode} authentication failure cannot fall back or create a registration`, async () => {
      const error = new UserError("test", "AuthenticationFailed", "Authentication failed");
      vi.spyOn(context.m365TokenProvider, "getAccessToken").mockResolvedValue(err(error));
      const result = await execute(mode);
      expect(result.result.isErr()).toBe(true);
      if (result.result.isErr()) expect(result.result.error).toBe(error);
      expect(http.history.get).toHaveLength(0);
      expect(http.history.post).toHaveLength(0);
    });
  }
});
