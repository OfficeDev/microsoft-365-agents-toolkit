const fs = require("node:fs");
const path = require("node:path");
const dotenv = require("dotenv");
const { parseDocument } = require("yaml");
const { buildMcpAuthProfile } = require("./mcp-auth-fixture.cjs");
const { prepareVscuseConfig } = require("./prepare-vscuse-config.cjs");

function loadProfile(mode, ssoFile, staticFile, ssoVariant = "auto") {
  const consumer = dotenv.parse(fs.readFileSync(ssoFile));
  const options = { ssoVariant };
  if (mode === "static-oauth") {
    if (!staticFile)
      throw new Error("Static OAuth requires an explicit credential file");
    options.staticCredentials = dotenv.parse(fs.readFileSync(staticFile));
  } else if (
    mode === "entra-sso" &&
    (ssoVariant === "absent" ||
      (ssoVariant === "auto" && !consumer.ENTRA_SSO_REGISTRATION_ID))
  ) {
    const match = /^da-sso\.([A-Za-z0-9_-]+)\.env$/.exec(
      path.basename(ssoFile),
    );
    if (!match)
      throw new Error(
        "Cannot identify the mini environment from the consumer filename",
      );
    const miniRoot = path.resolve(path.dirname(ssoFile), "..");
    options.miniEnvironment = dotenv.parse(
      fs.readFileSync(path.join(miniRoot, "env", `.env.${match[1]}`)),
    );
    options.miniManifest = JSON.parse(
      fs.readFileSync(path.join(miniRoot, "aad.manifest.json"), "utf8"),
    );
  }
  return buildMcpAuthProfile(mode, consumer, options);
}

function fixtureMode(plan) {
  const modes = new Set(
    plan.steps.flatMap((step) =>
      (step.tags ?? [])
        .filter((tag) => tag.startsWith("mcp_auth_mode:"))
        .map((tag) => tag.slice("mcp_auth_mode:".length)),
    ),
  );
  if (modes.size !== 1)
    throw new Error("Expected one isolated MCP authentication mode");
  return [...modes][0];
}

function prepareMcpAuthRunConfig(
  configText,
  planText,
  mode,
  names,
  ssoIdPresent,
) {
  if (fixtureMode(JSON.parse(planText)) !== mode) {
    throw new Error("MCP plan/profile mode mismatch");
  }
  if (mode === "entra-sso" && typeof ssoIdPresent !== "boolean") {
    throw new Error(
      "SSO configuration must declare registration-ID presence explicitly",
    );
  }
  const document = parseDocument(prepareVscuseConfig(configText, planText));
  const environment = document.getIn(["docker", "environment"], true);
  if (!environment?.items)
    throw new Error("Missing Docker environment mapping");
  for (const pair of [...environment.items]) {
    const name = String(pair.key);
    if (/^(MCP_AUTH_|DA_MCP_|ENTRA_SSO_|TEST_AAD_APP_)/.test(name)) {
      document.deleteIn(["docker", "environment", name]);
    }
  }
  for (const name of names) {
    if (!/^(MCP_AUTH_|DA_MCP_)[A-Z0-9_]+$/.test(name)) {
      throw new Error("Unexpected MCP fixture environment variable");
    }
    if (
      (mode !== "entra-sso" &&
        /^(MCP_AUTH_SSO_|MCP_AUTH_AAD_MANIFEST_B64$|DA_MCP_ENTRA_SSO_)/.test(
          name,
        )) ||
      (mode !== "static-oauth" && name.startsWith("DA_MCP_OAUTH_"))
    ) {
      throw new Error("Cross-mode MCP authentication variable");
    }
    document.setIn(
      ["docker", "environment", name],
      mode === "entra-sso" && name === "MCP_AUTH_SSO_ID" && !ssoIdPresent
        ? ""
        : "${" + name + "}",
    );
  }
  document.setIn(["docker", "environment", "TEMPLATE_VERSION"], "local");
  return document.toString();
}

function redact(text, secrets) {
  let result = text;
  for (const secret of [...secrets]
    .filter((value) => value && value.length >= 4)
    .sort((a, b) => b.length - a.length)) {
    result = result.split(secret).join("[REDACTED]");
  }
  return result
    .replace(
      /\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\b/g,
      "[REDACTED_TOKEN]",
    )
    .replace(
      /([?&](?:code|access_token|refresh_token|id_token|client_secret)=)[^&\s"'<>]+/gi,
      "$1[REDACTED]",
    );
}

module.exports = { loadProfile, fixtureMode, prepareMcpAuthRunConfig, redact };
