const { fixtureScript } = require("./mcp-auth-fixture.cjs");
const fixedInitialEndpoint = "https://example.invalid/mcp";

async function runDcrUpdateStage(phase, projectPath, options = {}) {
  const fs = require("node:fs/promises");
  const path = require("node:path");
  const { createHash } = require("node:crypto");
  const assert = require("node:assert/strict");
  const load =
    options.load ??
    require("node:module").createRequire(
      require("node:fs").realpathSync(
        require("node:child_process")
          .execFileSync("which", ["atk"], { encoding: "utf8" })
          .trim(),
      ),
    );
  const yaml = load("js-yaml");
  const dotenv = load("dotenv");
  const project = path.resolve(projectPath);
  const initial = fixedInitialEndpoint;
  const endpoint = options.endpoint ?? process.env.MCP_AUTH_ENDPOINT;
  for (const value of [initial, endpoint]) {
    const url = new URL(value);
    assert.ok(
      url.protocol === "https:" && !url.username && !url.password && !url.hash,
      "DCR update requires credential-free HTTPS endpoints",
    );
  }
  assert.notEqual(initial, endpoint, "DCR update endpoints must differ");
  const directory = path.join(project, ".auth-test");
  const readJson = async (file) => JSON.parse(await fs.readFile(file, "utf8"));
  const ownership = await readJson(path.join(directory, "ownership.json"));
  const envText = await fs.readFile(
    path.join(project, "env", ".env.dev"),
    "utf8",
  );
  const env = dotenv.parse(envText);
  const workflowPath = path.join(project, "m365agents.yml");
  const workflow = yaml.load(await fs.readFile(workflowPath, "utf8"));
  const registrations = workflow.provision.filter(
    (item) => item.uses === "dcr/register",
  );
  assert.equal(
    registrations.length,
    1,
    "Expected one independent DCR registration",
  );
  const registration = registrations[0];
  const key = registration.writeToEnvironmentFile.configurationId;
  const id = env[key];
  assert.ok(
    id &&
      ownership.mode === "dcr" &&
      ownership.projectName === path.basename(project) &&
      ownership.runId === env.MCP_AUTH_RUN_ID,
    "Invalid project ownership or configuration ID",
  );
  assert.ok(
    ownership.resources.some(
      (item) =>
        item.kind === "dcr-registration" &&
        item.envKey === key &&
        item.existedBefore === false,
    ),
    "DCR registration is not owned",
  );
  const pluginPath = path.join(project, "appPackage", "ai-plugin.json");
  const plugin = await readJson(pluginPath);
  assert.equal(plugin.runtimes.length, 1, "Expected one MCP runtime");
  const runtime = plugin.runtimes[0];
  assert.equal(runtime.type, "RemoteMCPServer");
  assert.deepEqual(
    runtime.auth,
    { type: "OAuthPluginVault", reference_id: "${{" + key + "}}" },
    "Original configuration reference must be retained",
  );
  const environmentHash = createHash("sha256").update(envText).digest("hex");
  const baselinePath = path.join(directory, "dcr-update-baseline.json");
  const policyNames = [
    "clientName",
    "m365AppId",
    "applicableToApps",
    "targetAudience",
    "supportedAccountTypes",
    "wellKnownAuthorizationServer",
    "resource",
  ];
  const policy = (value) =>
    Object.fromEntries(
      policyNames
        .filter((name) => Object.hasOwn(value, name))
        .map((name) => [name, value[name]]),
    );
  const readSnapshot =
    options.readSnapshot ??
    (async (configurationId) => {
      const { PublicClientApplication } = load("@azure/msal-node");
      const coreEntry = load.resolve("@microsoft/teamsfx-core");
      const constants = load(
        path.join(path.dirname(coreEntry), "common", "constants.js"),
      );
      // Same public test/CLI login client as cli/commonlib/m365Login.ts; not a DCR client.
      const client = new PublicClientApplication({
        auth: {
          clientId: "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
          authority:
            "https://login.microsoftonline.com/" + process.env.M365_TENANT_ID,
        },
      });
      let token;
      try {
        token = await client.acquireTokenByUsernamePassword({
          username: process.env.M365_ACCOUNT_NAME,
          password: process.env.M365_ACCOUNT_PASSWORD,
          scopes: constants.TeamsGraphScopes(),
        });
      } catch (error) {
        throw new Error(
          "Teams Graph read authentication failed: " + error.errorCode,
        );
      }
      assert.ok(
        token?.accessToken,
        "Teams Graph read authentication returned no token",
      );
      const base = constants.getResourceServiceEndpoint(
        constants.ResourceServiceType.TeamsGraph,
      );
      const response = await fetch(
        base +
          "/v1.0/dynamicConfigurations/" +
          encodeURIComponent(configurationId),
        {
          headers: {
            Authorization: "Bearer " + token.accessToken,
            "Client-Source": "agentstoolkit",
          },
          signal: AbortSignal.timeout(30000),
        },
      );
      if (!response.ok)
        throw new Error("DCR readback failed: HTTP " + response.status);
      return response.json();
    });
  const snapshot = await readSnapshot(id);
  assert.ok(
    snapshot && typeof snapshot === "object" && !Array.isArray(snapshot),
    "Invalid DCR configuration readback",
  );
  assert.ok(
    typeof snapshot.clientName === "string" &&
      snapshot.clientName.length > 0 &&
      typeof snapshot.m365AppId === "string" &&
      ["AnyApp", "SpecificApp"].includes(snapshot.applicableToApps) &&
      ["HomeTenant", "AnyTenant"].includes(snapshot.targetAudience) &&
      Array.isArray(snapshot.targetUrlsShouldStartWith) &&
      snapshot.targetUrlsShouldStartWith.every(
        (value) => typeof value === "string",
      ) &&
      ["mcpResourceUrl", "wellKnownAuthorizationServer", "resource"].every(
        (name) => !(name in snapshot) || typeof snapshot[name] === "string",
      ) &&
      (snapshot.supportedAccountTypes === undefined ||
        ["Enterprise", "Enterprise, Consumer"].includes(
          snapshot.supportedAccountTypes,
        )),
    "Invalid DCR configuration policy readback",
  );
  assert.ok(
    snapshot.oAuthConfigId === undefined || snapshot.oAuthConfigId === id,
    "DCR configuration ID mismatch",
  );
  assert.ok(
    snapshot.identityProvider === undefined ||
      snapshot.identityProvider === "Dynamic",
    "DCR identity mismatch",
  );
  if (phase === "baseline") {
    assert.equal(
      snapshot.mcpResourceUrl,
      initial,
      "Initial endpoint was not persisted",
    );
    assert.deepEqual(
      snapshot.targetUrlsShouldStartWith,
      [initial],
      "Initial target URL mismatch",
    );
    assert.equal(runtime.spec.url, initial, "Initial plugin URL mismatch");
    const baseline = {
      configurationId: id,
      runId: ownership.runId,
      initial,
      endpoint,
      policy: policy(snapshot),
      environmentHash,
      auth: runtime.auth,
    };
    await fs.writeFile(baselinePath, JSON.stringify(baseline), { flag: "wx" });
    return ownership.runId;
  }
  const baseline = await readJson(baselinePath);
  assert.ok(
    baseline.configurationId === id &&
      baseline.runId === ownership.runId &&
      baseline.initial === initial &&
      baseline.endpoint === endpoint,
    "Baseline configuration identity mismatch",
  );
  assert.equal(
    environmentHash,
    baseline.environmentHash,
    "Project environment changed",
  );
  assert.deepEqual(
    runtime.auth,
    baseline.auth,
    "Plugin authentication reference changed",
  );
  assert.deepEqual(policy(snapshot), baseline.policy, "DCR policy changed");
  if (phase === "apply") {
    assert.equal(
      snapshot.mcpResourceUrl,
      initial,
      "Source endpoint no longer matches baseline",
    );
    assert.equal(runtime.spec.url, initial);
    assert.equal(
      workflow.provision.filter((item) => item.uses === "dcr/update").length,
      0,
      "Update action already exists",
    );
    const localPath = path.join(project, ".vscode", "mcp.json");
    const local = await readJson(localPath);
    const servers = Object.values(local.servers).filter(
      (server) => server.url === initial,
    );
    assert.equal(servers.length, 1, "Expected one matching local MCP URL");
    workflow.provision.splice(workflow.provision.indexOf(registration) + 1, 0, {
      uses: "dcr/update",
      with: {
        configurationId: "${{" + key + "}}",
        mcpResourceUrl: endpoint,
        targetUrlsShouldStartWith: [endpoint],
      },
    });
    runtime.spec.url = endpoint;
    servers[0].url = endpoint;
    await fs.writeFile(workflowPath, yaml.dump(workflow, { lineWidth: -1 }));
    await fs.writeFile(pluginPath, JSON.stringify(plugin, null, 2) + "\n");
    await fs.writeFile(localPath, JSON.stringify(local, null, 2) + "\n");
    return ownership.runId;
  }
  assert.ok(phase === "verify" || phase === "noop", "Unknown DCR update stage");
  assert.equal(
    snapshot.mcpResourceUrl,
    endpoint,
    "Updated endpoint was not persisted",
  );
  assert.deepEqual(
    snapshot.targetUrlsShouldStartWith,
    [endpoint],
    "Updated target URL mismatch",
  );
  assert.equal(runtime.spec.url, endpoint, "Updated plugin URL mismatch");
  assert.equal(
    registration.with.mcpResourceUrl,
    initial,
    "Registration baseline was rewritten",
  );
  const updates = workflow.provision.filter(
    (item) => item.uses === "dcr/update",
  );
  assert.equal(updates.length, 1);
  assert.deepEqual(updates[0].with, {
    configurationId: "${{" + key + "}}",
    mcpResourceUrl: endpoint,
    targetUrlsShouldStartWith: [endpoint],
  });
  const checkpoint = path.join(directory, "dcr-update-verified.json");
  if (phase === "verify") {
    await fs.writeFile(
      checkpoint,
      JSON.stringify({ configurationId: id, runId: ownership.runId }),
      { flag: "wx" },
    );
  } else {
    assert.deepEqual(
      await readJson(checkpoint),
      { configurationId: id, runId: ownership.runId },
      "Missing successful update checkpoint",
    );
  }
  return ownership.runId;
}

function dcrUpdateScript(phase) {
  if (!["prepare", "baseline", "apply", "verify", "noop"].includes(phase)) {
    throw new Error("Unknown DCR update phase");
  }
  const program = `const fixedInitialEndpoint = ${JSON.stringify(fixedInitialEndpoint)};
(${runDcrUpdateStage.toString()})(process.argv[2], process.cwd()).then(id => {
console.log("VSCUSE_DCR_UPDATE_READY", require("node:path").basename(process.cwd()), process.argv[2], id);
}).catch(error => { console.error("DCR update fixture failed:", error.message); process.exitCode = 1; });`;
  if (phase === "prepare") {
    return (
      String.raw`import json, os, urllib.parse, urllib.request

def read_public_metadata(url):
    with urllib.request.urlopen(url, timeout=30) as response:
        value = json.load(response)
    if not isinstance(value, dict):
        raise AssertionError("Invalid public discovery metadata")
    return value

real_endpoint = urllib.parse.urlsplit(os.environ["MCP_AUTH_ENDPOINT"])
if (real_endpoint.scheme != "https" or not real_endpoint.netloc
        or real_endpoint.username or real_endpoint.password or real_endpoint.fragment):
    raise AssertionError("Invalid real MCP endpoint")
prm_url = urllib.parse.urlunsplit((real_endpoint.scheme, real_endpoint.netloc,
    "/.well-known/oauth-protected-resource", "", ""))
dcr_prm = read_public_metadata(prm_url)
servers = dcr_prm.get("authorization_servers")
if not isinstance(servers, list) or len(servers) != 1 or not isinstance(servers[0], str):
    raise AssertionError("Expected one real authorization server")
issuer = urllib.parse.urlsplit(servers[0])
if (issuer.scheme != "https" or not issuer.netloc or issuer.username or issuer.password
        or issuer.query or issuer.fragment):
    raise AssertionError("Invalid real authorization server")
dcr_well_known = urllib.parse.urlunsplit((issuer.scheme, issuer.netloc,
    "/.well-known/oauth-authorization-server" + issuer.path.rstrip("/"), "", ""))
authorization = read_public_metadata(dcr_well_known)
registration_endpoint = authorization.get("registration_endpoint")
if not isinstance(registration_endpoint, str):
    raise AssertionError("Real authorization server does not support DCR")
registration = urllib.parse.urlsplit(registration_endpoint)
if registration.scheme != "https" or not registration.netloc or registration.username or registration.password:
    raise AssertionError("Invalid real DCR registration endpoint")
resource = dcr_prm.get("resource")
if not isinstance(resource, str) or not urllib.parse.urlsplit(resource).scheme or "#" in resource:
    raise AssertionError("Missing or invalid real OAuth resource")
` +
      fixtureScript("dcr") +
      `\nimport base64\ninitial = ${JSON.stringify(fixedInitialEndpoint)}\n` +
      String.raw`inputs["mcpResourceUrl"] = initial
inputs["targetUrlsShouldStartWith"] = [initial]
inputs["wellKnownAuthorizationServer"] = dcr_well_known
inputs["resource"] = dcr_prm["resource"]
runtime["spec"]["url"] = initial
local_path = project / ".vscode" / "mcp.json"
local_mcp = json.loads(local_path.read_text(encoding="utf-8"))
local_servers = [server for server in local_mcp["servers"].values()
                 if server.get("url") == endpoint]
if len(local_servers) != 1:
    raise AssertionError("Expected one real local MCP URL before preparation")
local_servers[0]["url"] = initial
workflow_path.write_text(yaml.safe_dump(workflow, sort_keys=False), encoding="utf-8")
(project / "appPackage" / "ai-plugin.json").write_text(json.dumps(plugin, indent=2) + "\n", encoding="utf-8")
local_path.write_text(json.dumps(local_mcp, indent=2) + "\n", encoding="utf-8")
` +
      `(project / ".auth-test" / "dcr-update.cjs").write_bytes(base64.b64decode("${Buffer.from(program).toString("base64")}"))\n` +
      `print("VSCUSE_DCR_UPDATE_READY", project.name, "prepare", run_id)\n`
    );
  }
  return (
    `import os, subprocess\nfrom pathlib import Path\n` +
    `project = Path(os.environ["PROJECT_DIR"]).resolve()\n` +
    `subprocess.run(["node", str(project / ".auth-test" / "dcr-update.cjs"), "${phase}"], cwd=project, check=True)\n`
  );
}

module.exports = { runDcrUpdateStage, dcrUpdateScript, fixedInitialEndpoint };
