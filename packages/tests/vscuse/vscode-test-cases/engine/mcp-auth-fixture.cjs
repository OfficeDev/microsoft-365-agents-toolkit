const { Buffer } = require("node:buffer");
const { randomUUID } = require("node:crypto");

const expectedTenant = "16801f2e-7163-4bd5-8542-f083fa59993c";
const expectedClient = "7c5898c1-99fb-4d83-8474-02bf7985096c";
// Well-known Microsoft enterprise token-store caller, not the resource application's client ID.
const enterpriseTokenStoreClient = "ab3be6b7-f5df-413d-ac2d-abf1e3fd9c0b";
const expectedCanary = "PURPLE-NARWHAL-7788";
const modes = new Set(["static-oauth", "dcr", "entra-sso"]);

function required(source, name) {
  const value = source[name];
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`Missing MCP fixture setting: ${name}`);
  }
  return value;
}

function buildMcpAuthProfile(mode, consumer, options = {}) {
  if (!modes.has(mode)) throw new Error("Unsupported MCP authentication mode");
  const tenant = required(consumer, "ENTRA_TENANT_ID");
  if (tenant !== expectedTenant) throw new Error("MCP fixture tenant mismatch");
  const client = required(consumer, "ENTRA_CLIENT_ID");
  if (client !== expectedClient)
    throw new Error("MCP fixture client identity mismatch");
  const endpoint = required(consumer, "MCP_ENDPOINT_URL");
  const url = new URL(endpoint);
  if (url.protocol !== "https:" || url.username || url.password) {
    throw new Error(
      "MCP fixture requires an HTTPS endpoint without credentials",
    );
  }
  const environment = {
    MCP_AUTH_MODE: mode,
    MCP_AUTH_ENDPOINT: endpoint,
    MCP_AUTH_EXPECTED_TENANT: tenant,
    MCP_AUTH_EXPECTED_CANARY: expectedCanary,
    MCP_AUTH_EXPECTED_JWT_CLIENT_ID:
      mode === "entra-sso" ? enterpriseTokenStoreClient : client,
    MCP_AUTH_EXPECTED_FLOW: mode === "entra-sso" ? "entra_sso" : "dcr_oauth",
    MCP_AUTH_RUN_ID: randomUUID(),
  };
  if (mode === "static-oauth") {
    const credentials = options.staticCredentials ?? {};
    if (
      required(credentials, "ENTRA_TENANT_ID") !== tenant ||
      required(credentials, "ENTRA_CLIENT_ID") !== client
    ) {
      throw new Error("Static OAuth identity does not match the MCP fixture");
    }
    Object.assign(environment, {
      DA_MCP_OAUTH_CLIENTID: client,
      DA_MCP_OAUTH_CLIENT_SECRET: required(credentials, "ENTRA_CLIENT_SECRET"),
      DA_MCP_OAUTH_SCOPE: required(credentials, "ENTRA_API_SCOPE"),
      MCP_AUTH_AUTHORIZATION_URL: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize`,
      MCP_AUTH_TOKEN_URL: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
    });
  } else if (mode === "entra-sso") {
    if (
      required(consumer, "ENTRA_SSO_APPLICABLE_TO_APPS") !== "AnyApp" ||
      required(consumer, "ENTRA_SSO_TARGET_AUDIENCE") !== "HomeTenant"
    ) {
      throw new Error("SSO fixture must use AnyApp/HomeTenant");
    }
    const variant = options.ssoVariant ?? "auto";
    if (!["auto", "absent", "invalid"].includes(variant)) {
      throw new Error("Unsupported SSO registration variant");
    }
    const supplied = consumer.ENTRA_SSO_REGISTRATION_ID;
    if (supplied !== undefined && typeof supplied !== "string") {
      throw new Error("Invalid SSO registration ID");
    }
    if (supplied && supplied.trim().length === 0) {
      throw new Error("Invalid SSO registration ID");
    }
    let registration = "";
    if (variant === "auto") {
      registration = supplied ?? "";
    } else if (variant === "invalid") {
      registration = "vscuse-invalid-sso-registration";
    }
    Object.assign(environment, {
      DA_MCP_ENTRA_SSO_CLIENTID: client,
      MCP_AUTH_CLIENT_ID: client,
      MCP_AUTH_SSO_ID: registration,
      MCP_AUTH_SSO_PROTECTED_ID_B64: Buffer.from(
        JSON.stringify(supplied || registration),
      ).toString("base64"),
      MCP_AUTH_SSO_VARIANT: variant,
    });
    if (!registration) {
      const mini = options.miniEnvironment ?? {};
      const manifest = options.miniManifest;
      if (
        !manifest ||
        typeof manifest !== "object" ||
        Array.isArray(manifest)
      ) {
        throw new Error("Full SSO setup requires the mini-owned AAD manifest");
      }
      if (required(mini, "AAD_APP_CLIENT_ID") !== client) {
        throw new Error("Mini AAD identity mismatch");
      }
      const context = {
        AAD_APP_CLIENT_ID: client,
        AAD_APP_OBJECT_ID: required(mini, "AAD_APP_OBJECT_ID"),
        AAD_APP_ACCESS_AS_USER_PERMISSION_ID: required(
          mini,
          "AAD_APP_ACCESS_AS_USER_PERMISSION_ID",
        ),
        TEAMSFX_ENV: "atktesting08",
        MCP_SERVER_URL: url.origin,
        ENTRA_SSO_APPLICATION_ID_URI: required(
          consumer,
          "ENTRA_SSO_APPLICATION_ID_URI",
        ),
      };
      const resolved = JSON.parse(
        JSON.stringify(manifest).replace(
          /\$\{\{([A-Z][A-Z0-9_]*)\}\}/g,
          (_match, name) => {
            if (!Object.hasOwn(context, name)) {
              throw new Error(`Unsupported mini manifest reference: ${name}`);
            }
            return context[name];
          },
        ),
      );
      if (
        resolved.appId !== client ||
        !Array.isArray(resolved.identifierUris)
      ) {
        throw new Error("Invalid mini AAD manifest identity or URI set");
      }
      resolved.identifierUris = [
        ...new Set([...resolved.identifierUris, "${{MCP_AUTH_NEW_SSO_URI}}"]),
      ];
      environment.MCP_AUTH_AAD_MANIFEST_B64 = Buffer.from(
        JSON.stringify(resolved),
      ).toString("base64");
    }
  }
  return { mode, environment };
}

const commonScript = String.raw`import base64
import json
import os
import re
from pathlib import Path
import yaml

project = Path(os.environ["PROJECT_DIR"]).resolve()
workflow_path = project / "m365agents.yml"
workflow = yaml.safe_load(workflow_path.read_text(encoding="utf-8"))
actions = workflow.get("provision", [])
registrations = [a for a in actions if a.get("uses") in ("oauth/register", "dcr/register")]
if len(registrations) != 1:
    raise AssertionError("Expected one independent MCP authentication registration")
action = registrations[0]
inputs = action["with"]
output_key = action["writeToEnvironmentFile"]["configurationId"]
if not isinstance(output_key, str) or not re.fullmatch(r"[A-Z][A-Z0-9_]*", output_key):
    raise AssertionError("Invalid MCP registration output mapping")
endpoint = os.environ["MCP_AUTH_ENDPOINT"]
plugin = json.loads((project / "appPackage" / "ai-plugin.json").read_text(encoding="utf-8"))
if plugin.get("functions") != [] or len(plugin.get("runtimes", [])) != 1:
    raise AssertionError("Expected one dynamic MCP runtime without static functions")
runtime = plugin["runtimes"][0]
if (runtime.get("type") != "RemoteMCPServer" or runtime.get("spec", {}).get("url") != endpoint
        or runtime.get("run_for_functions") != ["*"] or "mcp_tool_description" in runtime.get("spec", {})):
    raise AssertionError("Generated MCP runtime does not match the selected endpoint")
if runtime.get("auth") != {"type": "OAuthPluginVault", "reference_id": "$" + "{{" + output_key + "}}"}:
    raise AssertionError("MCP runtime and registration output reference do not agree")
if (project / "appPackage" / "mcp-tools-1.json").exists():
    raise AssertionError("Dynamic MCP project contains a static tools snapshot")
env_file = project / "env" / ".env.dev"
if not env_file.is_file():
    raise AssertionError("Generated project environment is missing")

def set_project_env(name, value):
    lines = env_file.read_text(encoding="utf-8").splitlines()
    kept = [line for line in lines if not line.startswith(name + "=")]
    kept.append(name + "=" + json.dumps(value))
    env_file.write_text("\n".join(kept) + "\n", encoding="utf-8")

def assert_mode(expected):
    if os.environ.get("MCP_AUTH_MODE") != expected:
        raise AssertionError("Authentication profile/plan mode mismatch")

shared_registration_id = ""
`;

function fixtureScript(mode) {
  if (!modes.has(mode)) throw new Error("Unsupported MCP authentication mode");
  const body = {
    "static-oauth": String.raw`assert_mode("static-oauth")
if action["uses"] != "oauth/register" or inputs.get("identityProvider") != "Custom":
    raise AssertionError("Expected static OAuth registration")
inputs["authorizationUrl"] = os.environ["MCP_AUTH_AUTHORIZATION_URL"]
inputs["tokenUrl"] = os.environ["MCP_AUTH_TOKEN_URL"]
inputs["refreshUrl"] = os.environ["MCP_AUTH_TOKEN_URL"]
inputs["scope"] = os.environ["DA_MCP_OAUTH_SCOPE"]
inputs["baseUrl"] = endpoint
set_project_env(output_key, "")
`,
    dcr: String.raw`assert_mode("dcr")
if action["uses"] != "dcr/register" or inputs.get("mcpResourceUrl") != endpoint:
    raise AssertionError("Expected MCP-facade DCR registration")
for forbidden in ("clientId", "clientSecret", "authorizationUrl", "tokenUrl"):
    if forbidden in inputs:
        raise AssertionError("DCR must not use static OAuth credentials or endpoints")
set_project_env(output_key, "")
`,
    "entra-sso": String.raw`assert_mode("entra-sso")
if action["uses"] != "oauth/register" or inputs.get("identityProvider") != "MicrosoftEntra":
    raise AssertionError("Expected Microsoft Entra SSO registration")
inputs["applicableToApps"] = "AnyApp"
inputs["targetAudience"] = "HomeTenant"
inputs["baseUrl"] = endpoint
registration_id = os.environ.get("MCP_AUTH_SSO_ID", "")
shared_registration_id = json.loads(base64.b64decode(os.environ["MCP_AUTH_SSO_PROTECTED_ID_B64"]))
if not isinstance(shared_registration_id, str):
    raise AssertionError("Invalid protected SSO ownership identity")
set_project_env(output_key, registration_id)
if registration_id:
    if any(a.get("uses", "").startswith("aadApp/") or a.get("uses") == "oauth/update" for a in actions):
        raise AssertionError("Shared SSO reuse must not reconfigure mini-owned identity")
else:
    manifest = json.loads(base64.b64decode(os.environ["MCP_AUTH_AAD_MANIFEST_B64"]))
    expected_client = os.environ["MCP_AUTH_CLIENT_ID"]
    if manifest.get("appId") != expected_client:
        raise AssertionError("SSO AAD client mismatch")
    action["writeToEnvironmentFile"]["applicationIdUri"] = "MCP_AUTH_NEW_SSO_URI"
    fixture_dir = project / ".auth-test"
    fixture_dir.mkdir(exist_ok=True)
    (fixture_dir / "aad.manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    check = """const uri=process.env.MCP_AUTH_NEW_SSO_URI;
const client=process.env.MCP_AUTH_CLIENT_ID;
if(!uri || !uri.startsWith('api://auth-') || !uri.endsWith('/'+client)) {
  throw new Error('Unexpected new SSO resource URI; AAD configuration was not applied');
}
"""
    (fixture_dir / "check-sso-uri.cjs").write_text(check, encoding="utf-8")
    index = actions.index(action) + 1
    actions[index:index] = [
        {"uses": "script", "with": {"run": "node .auth-test/check-sso-uri.cjs"}},
        {"uses": "aadApp/update", "with": {
            "manifestPath": "./.auth-test/aad.manifest.json",
            "outputFilePath": "./.auth-test/aad.manifest.resolved.json"
        }}
    ]
`,
  }[mode];
  const preparation =
    body +
    String.raw`
workflow_path.write_text(yaml.safe_dump(workflow, sort_keys=False), encoding="utf-8")
assignments = {}
for line in env_file.read_text(encoding="utf-8").splitlines():
    match = re.match(r"^\s*(?:export\s+)?([A-Z][A-Z0-9_]*)\s*=(.*)$", line)
    if match:
        assignments[match.group(1)] = match.group(2).strip()

def ownership(kind, key):
    if not isinstance(key, str) or not re.fullmatch(r"[A-Z][A-Z0-9_]*", key):
        raise AssertionError("Invalid resource ownership output mapping")
    value = assignments.get(key, "")
    return {"kind": kind, "envKey": key, "existedBefore": value not in ("", '""', "''")}

resources = [ownership("dcr-registration" if action["uses"] == "dcr/register" else "oauth-registration", output_key)]
for item in actions:
    outputs = item.get("writeToEnvironmentFile", {})
    if item.get("uses") == "teamsApp/create" and "teamsAppId" in outputs:
        resources.append(ownership("teams-app", outputs["teamsAppId"]))
    if item.get("uses") == "copilotAgent/publish" and "titleId" in outputs:
        resources.append(ownership("m365-title", outputs["titleId"]))
record = {"version": 1, "projectName": project.name, "runId": run_id,
          "mode": os.environ["MCP_AUTH_MODE"], "resources": resources}
if shared_registration_id:
    record["sharedRegistrationId"] = shared_registration_id
json.dump(record, ownership_stream)
`;
  return (
    commonScript +
    String.raw`
run_id = os.environ.get("MCP_AUTH_RUN_ID", "")
if not run_id:
    raise AssertionError("MCP authentication run ownership is missing")
ownership_dir = project / ".auth-test"
ownership_dir.mkdir(exist_ok=True)
with (ownership_dir / "ownership.json").open("x", encoding="utf-8") as ownership_stream:
` +
    preparation
      .split("\n")
      .map((line) => `    ${line}`)
      .join("\n") +
    '\nprint("VSCUSE_MCP_AUTH_FIXTURE_READY", run_id)\n'
  );
}

module.exports = {
  buildMcpAuthProfile,
  fixtureScript,
  expectedTenant,
  expectedClient,
  expectedCanary,
};
