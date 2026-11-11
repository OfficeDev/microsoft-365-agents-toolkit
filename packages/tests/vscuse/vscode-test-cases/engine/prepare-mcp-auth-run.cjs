const fs = require("node:fs");
const path = require("node:path");
const { randomUUID } = require("node:crypto");
const dotenv = require("dotenv");
const { parseDocument } = require("yaml");
const {
  buildMcpAuthProfile,
  expectedClient,
} = require("./mcp-auth-fixture.cjs");
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

function requiredEnvironment(environment, name) {
  const value = environment[name];
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`Missing MCP environment setting: ${name}`);
  }
  return value;
}

function loadEnvironmentProfile(mode, environment, ssoVariant = "auto") {
  if (!["static-oauth", "dcr", "entra-sso"].includes(mode)) {
    throw new Error("Unsupported MCP authentication mode");
  }
  const consumer = {
    ENTRA_TENANT_ID: requiredEnvironment(environment, "M365_TENANT_ID"),
    // Resource identity only, not a DCR client ID or credential.
    ENTRA_CLIENT_ID: expectedClient,
    MCP_ENDPOINT_URL: requiredEnvironment(environment, "MCP_AUTH_ENDPOINT"),
  };
  const options = { ssoVariant };
  if (mode === "static-oauth") {
    consumer.ENTRA_CLIENT_ID = requiredEnvironment(
      environment,
      "DA_MCP_OAUTH_CLIENTID",
    );
    options.staticCredentials = {
      ENTRA_TENANT_ID: consumer.ENTRA_TENANT_ID,
      ENTRA_CLIENT_ID: consumer.ENTRA_CLIENT_ID,
      ENTRA_CLIENT_SECRET: requiredEnvironment(
        environment,
        "DA_MCP_OAUTH_CLIENT_SECRET",
      ),
    };
  } else if (mode === "entra-sso") {
    consumer.ENTRA_CLIENT_ID = requiredEnvironment(
      environment,
      "DA_MCP_ENTRA_SSO_CLIENTID",
    );
    consumer.ENTRA_SSO_REGISTRATION_ID = environment.MCP_AUTH_SSO_ID;
    consumer.ENTRA_SSO_APPLICABLE_TO_APPS = "AnyApp";
    consumer.ENTRA_SSO_TARGET_AUDIENCE = "HomeTenant";
    if (
      ssoVariant === "absent" ||
      (ssoVariant === "auto" && !consumer.ENTRA_SSO_REGISTRATION_ID)
    ) {
      const encoded = environment.MCP_AUTH_AAD_MANIFEST_B64;
      if (typeof encoded !== "string" || !encoded.trim()) {
        throw new Error(
          "SSO registration ID is absent. Supply MCP_AUTH_SSO_ID for reuse, " +
            "or MCP_AUTH_AAD_MANIFEST_B64 and mini-owned metadata for full creation.",
        );
      }
      try {
        options.miniManifest = JSON.parse(
          Buffer.from(encoded, "base64").toString("utf8"),
        );
      } catch {
        throw new Error("Invalid MCP_AUTH_AAD_MANIFEST_B64");
      }
      options.miniEnvironment = {
        AAD_APP_CLIENT_ID: consumer.ENTRA_CLIENT_ID,
        AAD_APP_OBJECT_ID: requiredEnvironment(
          environment,
          "MCP_AUTH_AAD_APP_OBJECT_ID",
        ),
        AAD_APP_ACCESS_AS_USER_PERMISSION_ID: requiredEnvironment(
          environment,
          "MCP_AUTH_AAD_ACCESS_AS_USER_PERMISSION_ID",
        ),
      };
      consumer.ENTRA_SSO_APPLICATION_ID_URI = requiredEnvironment(
        environment,
        "ENTRA_SSO_APPLICATION_ID_URI",
      );
    }
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
  localTemplates = true,
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
  const isDcrUpdate = JSON.parse(planText).steps.some((step) =>
    step.tags?.includes("operation:dcr-update-stage"),
  );
  if (isDcrUpdate && mode !== "dcr") {
    throw new Error("DCR update requires the DCR authentication profile");
  }
  if (isDcrUpdate) {
    for (const name of ["M365_ACCOUNT_NAME", "M365_TENANT_ID"]) {
      document.setIn(["docker", "environment", name], "${" + name + "}");
    }
  }
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
  if (localTemplates) {
    document.setIn(["docker", "environment", "TEMPLATE_VERSION"], "local");
  }
  return document.toString();
}

function prepareMcpAuthCiRun(configText, planText, environment) {
  const mode = fixtureMode(JSON.parse(planText));
  const profile = loadEnvironmentProfile(mode, environment);
  const account = requiredEnvironment(environment, "M365_ACCOUNT_NAME").trim();
  if (!/@atktesting08\.onmicrosoft\.com$/i.test(account)) {
    throw new Error("MCP CI requires the licensed atktesting08 account");
  }
  const password = requiredEnvironment(environment, "M365_ACCOUNT_PASSWORD");
  const cleared = Object.fromEntries(
    Object.keys(environment)
      .filter((name) =>
        /^(MCP_AUTH_|DA_MCP_|ENTRA_SSO_|TEST_AAD_APP_)/.test(name),
      )
      .map((name) => [name, ""]),
  );
  const runtime = {
    ...cleared,
    ...profile.environment,
    M365_ACCOUNT_NAME: account,
    M365_ACCOUNT_PASSWORD: password,
    M365_TENANT_ID: profile.environment.MCP_AUTH_EXPECTED_TENANT,
  };
  const prepared = prepareMcpAuthRunConfig(
    configText,
    planText,
    mode,
    Object.keys(profile.environment),
    Boolean(profile.environment.MCP_AUTH_SSO_ID),
    false,
  );
  return { mode, environment: runtime, configText: prepared };
}

function writeGithubEnvironment(environment, destination) {
  const lines = [];
  const masks = new Set();
  for (const [name, value] of Object.entries(environment)) {
    if (!/^[A-Z][A-Z0-9_]*$/.test(name) || typeof value !== "string") {
      throw new Error("Invalid MCP GitHub environment entry");
    }
    if (
      value &&
      /PASSWORD|SECRET|SSO_ID|PROTECTED_ID_B64|AAD_MANIFEST_B64/.test(name)
    ) {
      for (const variant of [
        value,
        encodeURIComponent(value),
        JSON.stringify(value).slice(1, -1),
      ]) {
        if (masks.has(variant)) continue;
        masks.add(variant);
        const masked = variant
          .replaceAll("%", "%25")
          .replaceAll("\r", "%0D")
          .replaceAll("\n", "%0A");
        process.stdout.write(`::add-mask::${masked}\n`);
      }
    }
    let delimiter;
    do {
      delimiter = `MCP_${randomUUID()}`;
    } while (value.split(/\r?\n/).includes(delimiter));
    lines.push(`${name}<<${delimiter}\n${value}\n${delimiter}\n`);
  }
  fs.appendFileSync(destination, lines.join(""), "utf8");
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

module.exports = {
  loadProfile,
  loadEnvironmentProfile,
  fixtureMode,
  prepareMcpAuthRunConfig,
  prepareMcpAuthCiRun,
  redact,
};

if (require.main === module) {
  try {
    const args = process.argv.slice(2);
    if (args.length !== 3) {
      throw new Error(
        "Usage: node prepare-mcp-auth-run.cjs <config> <plan> <output>",
      );
    }
    const destination = requiredEnvironment(process.env, "GITHUB_ENV");
    const result = prepareMcpAuthCiRun(
      fs.readFileSync(args[0], "utf8"),
      fs.readFileSync(args[1], "utf8"),
      process.env,
    );
    writeGithubEnvironment(result.environment, destination);
    fs.writeFileSync(args[2], result.configText, "utf8");
    process.stdout.write(
      `Prepared isolated MCP authentication profile: ${result.mode}\n`,
    );
  } catch (error) {
    process.stderr.write(`${error.message}\n`);
    process.exitCode = 1;
  }
}
