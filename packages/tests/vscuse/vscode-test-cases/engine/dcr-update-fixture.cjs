const { fixtureScript } = require("./mcp-auth-fixture.cjs");

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
  const endpoint = options.endpoint ?? process.env.MCP_AUTH_ENDPOINT;
  const url = new URL(endpoint);
  assert.ok(
    url.protocol === "https:" && !url.username && !url.password && !url.hash,
    "DCR update requires a credential-free HTTPS endpoint",
  );
  assert.ok(
    ["baseline", "apply", "verify", "noop"].includes(phase),
    "Unknown DCR update stage",
  );
  const directory = path.join(project, ".auth-test");
  const readJson = async (file) => JSON.parse(await fs.readFile(file, "utf8"));
  const hash = (value) => createHash("sha256").update(value).digest("hex");
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
  assert.equal(
    registration.with.mcpResourceUrl,
    endpoint,
    "Registration endpoint changed",
  );
  assert.deepEqual(registration.with.targetUrlsShouldStartWith, [endpoint]);
  assert.equal(
    registration.with.includePersonalMicrosoftAccounts,
    false,
    "Registration policy baseline changed",
  );
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
  const pluginText = await fs.readFile(
    path.join(project, "appPackage", "ai-plugin.json"),
    "utf8",
  );
  const plugin = JSON.parse(pluginText);
  assert.equal(plugin.runtimes.length, 1, "Expected one MCP runtime");
  const runtime = plugin.runtimes[0];
  assert.equal(runtime.type, "RemoteMCPServer");
  assert.equal(runtime.spec.url, endpoint, "Plugin endpoint changed");
  assert.deepEqual(
    runtime.auth,
    { type: "OAuthPluginVault", reference_id: "${{" + key + "}}" },
    "Original configuration reference must be retained",
  );
  const localText = await fs.readFile(
    path.join(project, ".vscode", "mcp.json"),
    "utf8",
  );
  const local = JSON.parse(localText);
  assert.equal(
    Object.values(local.servers).filter((server) => server.url === endpoint)
      .length,
    1,
    "Expected one matching local MCP URL",
  );
  const fileHashes = {
    environment: hash(envText),
    plugin: hash(pluginText),
    localMcp: hash(localText),
    registration: hash(JSON.stringify(registration)),
  };
  const baselinePath = path.join(directory, "dcr-update-baseline.json");
  const environmentStateHash = hash(
    JSON.stringify(
      Object.keys(env)
        .sort()
        .map((key) => [key, env[key]]),
    ),
  );
  const publicEnvironment = Object.fromEntries(
    Object.keys(env)
      .filter((key) => !/SECRET|PASSWORD|TOKEN|KEY/i.test(key))
      .map((key) => [key, hash(env[key])]),
  );
  const policyNames = [
    "clientName",
    "m365AppId",
    "applicableToApps",
    "targetAudience",
    "mcpResourceUrl",
    "targetUrlsShouldStartWith",
    "wellKnownAuthorizationServer",
    "resource",
  ];
  const policy = (value) =>
    Object.fromEntries(
      policyNames
        .filter((name) => Object.hasOwn(value, name))
        .map((name) => [name, value[name]]),
    );
  let token;
  const readSnapshot =
    options.readSnapshot ??
    (async (configurationId) => {
      const coreEntry = load.resolve("@microsoft/teamsfx-core");
      const constants = load(
        path.join(path.dirname(coreEntry), "common", "constants.js"),
      );
      if (!token) {
        const { PublicClientApplication } = load("@azure/msal-node");
        // Same public test/CLI login client as cli/commonlib/m365Login.ts; not a DCR client.
        const client = new PublicClientApplication({
          auth: {
            clientId: "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
            authority:
              "https://login.microsoftonline.com/" + process.env.M365_TENANT_ID,
          },
        });
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
          "/v1.0/oAuthConfigurations/" +
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
  const readValidSnapshot = async () => {
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
        ["Enterprise", "Enterprise, Consumer"].includes(
          snapshot.supportedAccountTypes,
        ),
      "Invalid DCR configuration policy readback",
    );
    assert.ok(
      snapshot.oAuthConfigId === undefined || snapshot.oAuthConfigId === id,
      "DCR configuration ID mismatch",
    );
    assert.ok(snapshot.identityProvider === "Dynamic", "DCR identity mismatch");
    assert.equal(
      snapshot.mcpResourceUrl,
      endpoint,
      "DCR endpoint must remain unchanged",
    );
    assert.deepEqual(
      snapshot.targetUrlsShouldStartWith,
      [endpoint],
      "DCR target URLs must remain unchanged",
    );
    return snapshot;
  };
  let snapshot = await readValidSnapshot();
  if (phase === "baseline") {
    assert.equal(
      snapshot.supportedAccountTypes,
      "Enterprise",
      "Initial account policy must be Enterprise",
    );
    await fs.writeFile(
      baselinePath,
      JSON.stringify({
        configurationId: id,
        runId: ownership.runId,
        endpoint,
        supportedAccountTypes: snapshot.supportedAccountTypes,
        policy: policy(snapshot),
        fileHashes,
        environmentStateHash,
        publicEnvironment,
        auth: runtime.auth,
      }),
      { flag: "wx" },
    );
    return ownership.runId;
  }
  const baseline = await readJson(baselinePath);
  assert.ok(
    baseline.configurationId === id &&
      baseline.runId === ownership.runId &&
      baseline.endpoint === endpoint &&
      baseline.supportedAccountTypes === "Enterprise",
    "Baseline configuration identity mismatch",
  );
  if (fileHashes.environment !== baseline.fileHashes.environment) {
    const changed = [
      ...new Set([
        ...Object.keys(publicEnvironment),
        ...Object.keys(baseline.publicEnvironment),
      ]),
    ]
      .filter(
        (key) => publicEnvironment[key] !== baseline.publicEnvironment[key],
      )
      .sort();
    const reason =
      environmentStateHash === baseline.environmentStateHash
        ? "formatting only"
        : changed.join(", ") || "protected values";
    assert.fail("Project environment changed: " + reason);
  }
  for (const component of Object.keys(fileHashes)) {
    assert.ok(
      fileHashes[component] === baseline.fileHashes[component],
      `Project ${component} changed`,
    );
  }
  assert.deepEqual(
    runtime.auth,
    baseline.auth,
    "Plugin authentication reference changed",
  );
  const assertPolicyUnchanged = (value) => {
    const current = policy(value);
    const changed = policyNames.filter(
      (name) =>
        JSON.stringify(current[name]) !== JSON.stringify(baseline.policy[name]),
    );
    assert.ok(
      changed.length === 0,
      "Unrelated DCR policy changed: " + changed.join(", "),
    );
  };
  assertPolicyUnchanged(snapshot);
  if (phase === "apply") {
    assert.equal(
      snapshot.supportedAccountTypes,
      "Enterprise",
      "Source account policy no longer matches baseline",
    );
    assert.equal(
      workflow.provision.filter((item) => item.uses === "dcr/update").length,
      0,
      "Update action already exists",
    );
    workflow.provision.splice(workflow.provision.indexOf(registration) + 1, 0, {
      uses: "dcr/update",
      with: {
        configurationId: "${{" + key + "}}",
        includePersonalMicrosoftAccounts: true,
      },
    });
    await fs.writeFile(workflowPath, yaml.dump(workflow, { lineWidth: -1 }));
    return ownership.runId;
  }
  const updates = workflow.provision.filter(
    (item) => item.uses === "dcr/update",
  );
  assert.equal(updates.length, 1);
  assert.deepEqual(updates[0].with, {
    configurationId: "${{" + key + "}}",
    includePersonalMicrosoftAccounts: true,
  });
  const sleep =
    options.sleep ??
    ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
  for (
    let attempt = 0;
    phase === "verify" &&
    snapshot.supportedAccountTypes === "Enterprise" &&
    attempt < 14;
    attempt++
  ) {
    await sleep(1000);
    snapshot = await readValidSnapshot();
    assertPolicyUnchanged(snapshot);
  }
  assert.equal(
    snapshot.supportedAccountTypes,
    "Enterprise, Consumer",
    "Account-policy update was not persisted",
  );
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
  const program = `(${runDcrUpdateStage.toString()})(process.argv[2], process.cwd()).then(id => {
console.log("VSCUSE_DCR_UPDATE_READY", require("node:path").basename(process.cwd()), process.argv[2], id);
}).catch(error => {
console.error("DCR update fixture failed:", String(error?.message ?? error).split("\\n")[0]);
process.exitCode = 1;
});`;
  if (phase === "prepare") {
    return (
      fixtureScript("dcr") +
      `\ninputs["includePersonalMicrosoftAccounts"] = False\ninputs["targetUrlsShouldStartWith"] = [endpoint]\n` +
      `workflow_path.write_text(yaml.safe_dump(workflow, sort_keys=False), encoding="utf-8")\n` +
      `(project / ".auth-test" / "dcr-update.cjs").write_bytes(base64.b64decode("${Buffer.from(program).toString("base64")}"))\n` +
      `print("VSCUSE_DCR_UPDATE_READY", project.name, "prepare", run_id)\n`
    );
  }
  return (
    `import os, subprocess\nfrom pathlib import Path\n` +
    `project = Path(os.environ["PROJECT_DIR"]).resolve()\n` +
    `completed = subprocess.run(["node", str(project / ".auth-test" / "dcr-update.cjs"), "${phase}"], cwd=project, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)\n` +
    `print(completed.stdout, end="", flush=True)\nraise SystemExit(completed.returncode)\n`
  );
}

module.exports = { runDcrUpdateStage, dcrUpdateScript };
