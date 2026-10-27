const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

function diagnostics() {
  const file = path.join(__dirname, "mcp-auth-diagnostics.cjs");
  assert.ok(fs.existsSync(file), "Missing safe MCP authentication diagnostics");
  return require(file).diagnoseMcpAuthentication;
}

const environment = {
  M365_TENANT_ID: "16801f2e-7163-4bd5-8542-f083fa59993c",
  DA_MCP_OAUTH_CLIENTID: "7c5898c1-99fb-4d83-8474-02bf7985096c",
  DA_MCP_OAUTH_CLIENT_SECRET: "private-static-secret",
  DA_MCP_ENTRA_SSO_CLIENTID: "7c5898c1-99fb-4d83-8474-02bf7985096c",
  MCP_AUTH_SSO_ID: "private/registration=id",
  M365_ACCOUNT_NAME: "licensed@atktesting08.onmicrosoft.com",
  M365_ACCOUNT_PASSWORD: "private-password",
  MCP_AUTH_ENDPOINT: "https://mcp.example.test/mcp",
};

function response(status, body) {
  return { status, ok: status >= 200 && status < 300, json: async () => body };
}

test("VCB-188: rejected credentials and unreadable SSO IDs report safe protocol evidence", async () => {
  const run = diagnostics();
  const requests = [];
  const results = await run(environment, {
    teamsGraphEndpoint: "https://teams.microsoft.com/api/platform",
    getM365Token: async () => "private-access-token",
    fetch: async (url, options) => {
      requests.push({ url, options });
      if (options.method === "POST") {
        assert.ok(url.endsWith("/oauth2/v2.0/token"));
        return response(401, {
          error: "invalid_client",
          error_codes: [7000215],
          error_description: "private-static-secret private-password",
        });
      }
      assert.equal(
        url,
        "https://teams.microsoft.com/api/platform/v1.0/oAuthConfigurations/" +
          encodeURIComponent(environment.MCP_AUTH_SSO_ID),
      );
      return response(400, {
        error: {
          code: "BadRequest",
          innerError: {
            code: "ConfigReadError",
            message: "private/registration=id",
          },
        },
      });
    },
  });
  assert.equal(results.length, 3);
  assert.ok(results.every((item) => item.outcome === "rejected"));
  assert.equal(results[0].error, "invalid_client");
  assert.deepEqual(results[0].errorCodes, [7000215]);
  assert.equal(results[2].innerError, "ConfigReadError");
  const printed = JSON.stringify(results);
  for (const secret of [
    "private-static-secret",
    "private-password",
    "private/registration=id",
    "private-access-token",
  ])
    assert.ok(!printed.includes(secret));
  assert.ok(
    requests.every(
      ({ url, options }) => options.method !== "POST" || url.endsWith("/token"),
    ),
  );
});

test("VCB-188: readable SSO records are checked without replacing or mutating them", async () => {
  const run = diagnostics();
  const results = await run(environment, {
    teamsGraphEndpoint: "https://teams.microsoft.com/api/platform",
    getM365Token: async () => "private-access-token",
    fetch: async (url, options) =>
      options.method === "POST"
        ? response(200, { access_token: "private-issued-token" })
        : response(200, {
            identityProvider: "MicrosoftEntra",
            clientId: environment.DA_MCP_ENTRA_SSO_CLIENTID,
            useSingleSignOn: true,
            applicableToApps: "AnyApp",
            scopes: [
              `api://${environment.DA_MCP_ENTRA_SSO_CLIENTID}/access_as_user`,
            ],
            targetUrlsShouldStartWith: [environment.MCP_AUTH_ENDPOINT],
          }),
  });
  assert.ok(results.every((item) => item.outcome === "accepted"));
  assert.equal(results[2].providerMatches, true);
  assert.equal(results[2].clientMatches, true);
  assert.equal(results[2].usageMatches, true);
  assert.ok(!JSON.stringify(results).includes("private-issued-token"));
});

test("VCB-195: readable SSO identity does not hide unusable mode, scope or target restrictions", async () => {
  const results = await diagnostics()(environment, {
    teamsGraphEndpoint: "https://teams.microsoft.com/api/platform",
    getM365Token: async () => "private-access-token",
    fetch: async (url, options) =>
      options.method === "POST"
        ? response(200, { access_token: "private-issued-token" })
        : response(200, {
            identityProvider: "MicrosoftEntra",
            clientId: environment.DA_MCP_ENTRA_SSO_CLIENTID,
            useSingleSignOn: false,
            applicableToApps: "SpecificApp",
            scopes: ["api://different-client/access_as_user"],
            targetUrlsShouldStartWith: ["https://different.example.test/mcp"],
            clientSecret: "private-record-secret",
          }),
  });
  assert.equal(results[2].outcome, "rejected");
  assert.equal(results[2].usageMatches, false);
  assert.equal(results[2].usesSingleSignOn, false);
  assert.equal(results[2].appliesToAnyApp, false);
  assert.equal(results[2].scopeMatches, false);
  assert.equal(results[2].endpointMatches, false);
  for (const hidden of [
    "private-record-secret",
    "different-client",
    "different.example.test",
  ]) {
    assert.ok(!JSON.stringify(results).includes(hidden));
  }
});
