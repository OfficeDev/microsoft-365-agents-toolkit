function safeCode(value) {
  return typeof value === "string" &&
    /^[A-Za-z][A-Za-z0-9_.-]{0,80}$/.test(value)
    ? value
    : undefined;
}

async function diagnoseMcpAuthentication(environment, options) {
  const request = options.fetch ?? fetch;
  const results = [];
  const tenant = environment.M365_TENANT_ID;
  const client = environment.DA_MCP_OAUTH_CLIENTID;
  const secret = environment.DA_MCP_OAUTH_CLIENT_SECRET;
  if (!tenant || !client || !secret) {
    results.push({
      kind: "client-auth",
      outcome: "error",
      error: "MissingClientCredential",
    });
  } else {
    for (const method of ["post", "basic"]) {
      const body = new URLSearchParams({
        grant_type: "client_credentials",
        scope: "https://graph.microsoft.com/.default",
      });
      const headers = { "Content-Type": "application/x-www-form-urlencoded" };
      if (method === "post") {
        body.set("client_id", client);
        body.set("client_secret", secret);
      } else {
        headers.Authorization =
          "Basic " +
          Buffer.from(
            encodeURIComponent(client) + ":" + encodeURIComponent(secret),
          ).toString("base64");
      }
      try {
        const response = await request(
          "https://login.microsoftonline.com/" + tenant + "/oauth2/v2.0/token",
          { method: "POST", headers, body, signal: AbortSignal.timeout(30000) },
        );
        const data = await response.json();
        const issued = response.ok && typeof data.access_token === "string";
        const error = safeCode(data.error);
        results.push({
          kind: "client-auth",
          method,
          status: response.status,
          outcome: issued
            ? "accepted"
            : error === "invalid_client"
              ? "rejected"
              : "indeterminate",
          tokenIssued: issued,
          error,
          errorCodes: Array.isArray(data.error_codes)
            ? data.error_codes.filter(Number.isInteger)
            : [],
          secretHasOuterQuotes: /^["']|["']$/.test(secret),
          secretHasOuterWhitespace: secret.trim() !== secret,
          secretLooksLikeIdentifier:
            /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(secret),
        });
      } catch (error) {
        results.push({
          kind: "client-auth",
          method,
          outcome: "error",
          error:
            safeCode(error.code) ?? safeCode(error.name) ?? "TransportError",
        });
      }
    }
  }
  const id = environment.MCP_AUTH_SSO_ID;
  if (!id) {
    results.push({
      kind: "sso-read",
      outcome: "error",
      error: "MissingSsoRegistration",
    });
  } else {
    try {
      const token = await options.getM365Token();
      const response = await request(
        options.teamsGraphEndpoint.replace(/\/$/, "") +
          "/v1.0/oAuthConfigurations/" +
          encodeURIComponent(id),
        {
          method: "GET",
          headers: {
            Authorization: "Bearer " + token,
            "Client-Source": "agentstoolkit",
          },
          signal: AbortSignal.timeout(30000),
        },
      );
      const data = await response.json();
      const providerMatches =
        response.ok && data.identityProvider === "MicrosoftEntra";
      const clientMatches =
        response.ok && data.clientId === environment.DA_MCP_ENTRA_SSO_CLIENTID;
      results.push({
        kind: "sso-read",
        status: response.status,
        outcome: providerMatches && clientMatches ? "accepted" : "rejected",
        error: safeCode(data.error?.code),
        innerError: safeCode(data.error?.innerError?.code),
        providerMatches,
        clientMatches,
        idHasOuterQuotes: /^["']|["']$/.test(id),
        idHasOuterWhitespace: id.trim() !== id,
        idContainsAssignment:
          /^(?:MCP_AUTH_SSO_ID|ENTRA_SSO_REGISTRATION_ID)\s*=/.test(id),
      });
    } catch (error) {
      results.push({
        kind: "sso-read",
        outcome: "error",
        error:
          safeCode(error.errorCode) ??
          safeCode(error.code) ??
          safeCode(error.name) ??
          "TransportError",
      });
    }
  }
  return results;
}

module.exports = { diagnoseMcpAuthentication };

if (require.main === module) {
  (async () => {
    const path = require("node:path");
    const load = require("node:module").createRequire(
      require("node:fs").realpathSync(
        require("node:child_process")
          .execFileSync("which", ["atk"], { encoding: "utf8" })
          .trim(),
      ),
    );
    const { PublicClientApplication } = load("@azure/msal-node");
    const constants = load(
      path.join(
        path.dirname(load.resolve("@microsoft/teamsfx-core")),
        "common",
        "constants.js",
      ),
    );
    const results = await diagnoseMcpAuthentication(process.env, {
      teamsGraphEndpoint: constants.getResourceServiceEndpoint(
        constants.ResourceServiceType.TeamsGraph,
      ),
      getM365Token: async () => {
        const client = new PublicClientApplication({
          auth: {
            clientId: "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
            authority:
              "https://login.microsoftonline.com/" + process.env.M365_TENANT_ID,
          },
        });
        const result = await client.acquireTokenByUsernamePassword({
          username: process.env.M365_ACCOUNT_NAME,
          password: process.env.M365_ACCOUNT_PASSWORD,
          scopes: constants.TeamsGraphScopes(),
        });
        if (!result?.accessToken) throw new Error("MissingAccessToken");
        return result.accessToken;
      },
    });
    for (const result of results) console.log(JSON.stringify(result));
    if (
      results.some((result) => ["rejected", "error"].includes(result.outcome))
    ) {
      process.exitCode = 1;
    }
  })().catch((error) => {
    console.error(
      JSON.stringify({
        kind: "diagnostic-error",
        error:
          safeCode(error.code) ?? safeCode(error.name) ?? "DiagnosticError",
      }),
    );
    process.exitCode = 1;
  });
}
