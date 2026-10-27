const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const yaml = require("js-yaml");

const endpoint = "https://real.example.test/mcp";
const runId = "00000000-0000-4000-8000-000000000001";

function runtime() {
  const file = path.join(__dirname, "dcr-update-fixture.cjs");
  assert.ok(fsSync.existsSync(file), "Missing independent DCR update fixture");
  return require(file).runDcrUpdateStage;
}

async function fixture(context) {
  const project = await fs.mkdtemp(path.join(os.tmpdir(), "dcr-update-ui-"));
  context.after(() => fs.rm(project, { recursive: true, force: true }));
  for (const dir of ["env", "appPackage", ".vscode", ".auth-test"]) {
    await fs.mkdir(path.join(project, dir));
  }
  const registration = {
    uses: "dcr/register",
    with: {
      name: "Owned DCR",
      mcpResourceUrl: endpoint,
      targetUrlsShouldStartWith: [endpoint],
      includePersonalMicrosoftAccounts: false,
      wellKnownAuthorizationServer:
        "https://real.example.test/.well-known/oauth-authorization-server",
      resource: "urn:fixture:dcr-update",
    },
    writeToEnvironmentFile: { configurationId: "OWNED_DCR_ID" },
  };
  await fs.writeFile(
    path.join(project, "m365agents.yml"),
    yaml.dump({
      version: "v1.13",
      provision: [registration],
    }),
  );
  await fs.writeFile(
    path.join(project, "env", ".env.dev"),
    `OWNED_DCR_ID=owned-id\nMCP_AUTH_RUN_ID=${runId}\n`,
  );
  await fs.writeFile(
    path.join(project, ".auth-test", "ownership.json"),
    JSON.stringify({
      version: 1,
      projectName: path.basename(project),
      mode: "dcr",
      runId,
      resources: [
        {
          kind: "dcr-registration",
          envKey: "OWNED_DCR_ID",
          existedBefore: false,
        },
      ],
    }),
  );
  await fs.writeFile(
    path.join(project, "appPackage", "ai-plugin.json"),
    JSON.stringify({
      runtimes: [
        {
          type: "RemoteMCPServer",
          spec: { url: endpoint },
          auth: { type: "OAuthPluginVault", reference_id: "${{OWNED_DCR_ID}}" },
        },
      ],
    }),
  );
  await fs.writeFile(
    path.join(project, ".vscode", "mcp.json"),
    JSON.stringify({
      servers: { owned: { type: "http", url: endpoint } },
    }),
  );
  let state = {
    oAuthConfigId: "owned-id",
    identityProvider: "Dynamic",
    clientName: "Owned DCR",
    m365AppId: "",
    applicableToApps: "AnyApp",
    targetAudience: "HomeTenant",
    supportedAccountTypes: "Enterprise",
    mcpResourceUrl: endpoint,
    targetUrlsShouldStartWith: [endpoint],
    wellKnownAuthorizationServer:
      "https://real.example.test/.well-known/oauth-authorization-server",
    resource: "urn:fixture:dcr-update",
    clientSecret: "must-not-be-persisted",
  };
  const options = {
    load: require,
    endpoint,
    sleep: async () => {},
    readSnapshot: async (id) => {
      assert.equal(id, "owned-id");
      return state;
    },
  };
  return {
    project,
    options,
    registration,
    setState(value) {
      state = value;
    },
    state: () => state,
  };
}

test("VCB-185 / DCR-UI-UPD-02: baseline is exclusive, project-owned and credential-free", async (context) => {
  const run = runtime();
  const f = await fixture(context);
  await run("baseline", f.project, f.options);
  const text = await fs.readFile(
    path.join(f.project, ".auth-test", "dcr-update-baseline.json"),
    "utf8",
  );
  assert.ok(text.includes("owned-id"));
  assert.ok(!text.includes("must-not-be-persisted"));
  await assert.rejects(
    run("baseline", f.project, f.options),
    /exist|baseline/i,
  );
});

test("VCB-191 / DCR-UI-UPD-03: adds only the supported policy update without rewriting references", async (context) => {
  const run = runtime();
  const f = await fixture(context);
  const env = await fs.readFile(
    path.join(f.project, "env", ".env.dev"),
    "utf8",
  );
  const pluginBefore = await fs.readFile(
    path.join(f.project, "appPackage", "ai-plugin.json"),
    "utf8",
  );
  const localBefore = await fs.readFile(
    path.join(f.project, ".vscode", "mcp.json"),
    "utf8",
  );
  await run("baseline", f.project, f.options);
  await run("apply", f.project, f.options);
  const flow = yaml.load(
    await fs.readFile(path.join(f.project, "m365agents.yml"), "utf8"),
  );
  assert.deepEqual(flow.provision[0], f.registration);
  assert.deepEqual(flow.provision[1], {
    uses: "dcr/update",
    with: {
      configurationId: "${{OWNED_DCR_ID}}",
      includePersonalMicrosoftAccounts: true,
    },
  });
  const plugin = JSON.parse(
    await fs.readFile(
      path.join(f.project, "appPackage", "ai-plugin.json"),
      "utf8",
    ),
  );
  assert.equal(plugin.runtimes[0].spec.url, endpoint);
  assert.equal(plugin.runtimes[0].auth.reference_id, "${{OWNED_DCR_ID}}");
  assert.equal(
    await fs.readFile(path.join(f.project, "env", ".env.dev"), "utf8"),
    env,
  );
  const mcp = JSON.parse(
    await fs.readFile(path.join(f.project, ".vscode", "mcp.json"), "utf8"),
  );
  assert.equal(mcp.servers.owned.url, endpoint);
  assert.equal(
    await fs.readFile(
      path.join(f.project, "appPackage", "ai-plugin.json"),
      "utf8",
    ),
    pluginBefore,
  );
  assert.equal(
    await fs.readFile(path.join(f.project, ".vscode", "mcp.json"), "utf8"),
    localBefore,
  );
});

test("VCB-191 / DCR-UI-UPD-04: only persisted account-policy changes pass verification", async (context) => {
  const run = runtime();
  const f = await fixture(context);
  await run("baseline", f.project, f.options);
  await run("apply", f.project, f.options);
  await assert.rejects(
    run("verify", f.project, f.options),
    /policy|account|persist/i,
  );
  f.setState({
    ...f.state(),
    supportedAccountTypes: "Enterprise, Consumer",
    clientSecret: "rediscovered-secret",
  });
  await assert.rejects(run("noop", f.project, f.options), /ENOENT|checkpoint/i);
  await run("verify", f.project, f.options);
  await run("noop", f.project, f.options);
  f.setState({ ...f.state(), targetAudience: "AnyTenant" });
  await assert.rejects(
    run("noop", f.project, f.options),
    /policy|targetAudience/i,
  );
});

test("VCB-191: stale account-policy reads are bounded and unrelated changes fail immediately", async (context) => {
  const run = runtime();
  const f = await fixture(context);
  await run("baseline", f.project, f.options);
  await run("apply", f.project, f.options);
  let reads = 0;
  await run("verify", f.project, {
    ...f.options,
    readSnapshot: async () => {
      reads++;
      return reads < 3
        ? f.state()
        : { ...f.state(), supportedAccountTypes: "Enterprise, Consumer" };
    },
  });
  assert.equal(reads, 3);
  const other = await fixture(context);
  await run("baseline", other.project, other.options);
  await run("apply", other.project, other.options);
  let invalidReads = 0;
  await assert.rejects(
    run("verify", other.project, {
      ...other.options,
      readSnapshot: async () => {
        invalidReads++;
        return {
          ...other.state(),
          mcpResourceUrl: "https://foreign.example.test/mcp",
        };
      },
    }),
    /URL|policy|unchanged|invariant/i,
  );
  assert.equal(invalidReads, 1);
});

test("VCB-185 / DCR-UI-UPD-05: missing ownership, foreign IDs and read failures fail closed", async (context) => {
  const run = runtime();
  const f = await fixture(context);
  await assert.rejects(
    run("noop", f.project, f.options),
    /baseline|ENOENT|checkpoint/i,
  );
  await assert.rejects(
    run("baseline", f.project, {
      ...f.options,
      readSnapshot: async () => {
        throw new Error("HTTP 404");
      },
    }),
    /HTTP 404/,
  );
  await run("baseline", f.project, f.options);
  f.setState({ ...f.state(), oAuthConfigId: "foreign-id" });
  await assert.rejects(
    run("verify", f.project, f.options),
    /identity|ID|configuration/i,
  );
  await fs.rm(path.join(f.project, ".auth-test", "ownership.json"));
  await assert.rejects(run("apply", f.project, f.options), /ownership|ENOENT/i);
});

test("VCB-185: malformed policy readback cannot become a successful baseline", async (context) => {
  const run = runtime();
  for (const invalid of [
    { clientName: null },
    { resource: 42 },
    { wellKnownAuthorizationServer: null },
  ]) {
    const f = await fixture(context);
    f.setState({ ...f.state(), ...invalid });
    await assert.rejects(
      run("baseline", f.project, f.options),
      /configuration|readback/i,
    );
  }
});

test("VCB-185: the common OAuth read requires explicit Dynamic identity", async (context) => {
  const run = runtime();
  for (const identityProvider of [undefined, "Custom", "MicrosoftEntra"]) {
    const f = await fixture(context);
    f.setState({ ...f.state(), identityProvider });
    await assert.rejects(run("baseline", f.project, f.options), /identity/i);
  }
});

test("VCB-185: default readback uses the common OAuth resource without a mutation fallback", async (context) => {
  const f = await fixture(context);
  const coreEntry = path.join(f.project, "core", "index.js");
  const constantsPath = path.join(
    path.dirname(coreEntry),
    "common",
    "constants.js",
  );
  const load = (specifier) => {
    if (specifier === "@azure/msal-node") {
      return {
        PublicClientApplication: class {
          async acquireTokenByUsernamePassword() {
            return { accessToken: "fixture-token" };
          }
        },
      };
    }
    if (specifier === constantsPath) {
      return {
        TeamsGraphScopes: () => ["fixture-scope"],
        ResourceServiceType: { TeamsGraph: "teams-graph" },
        getResourceServiceEndpoint: () => "https://graph.example.test",
      };
    }
    return require(specifier);
  };
  load.resolve = (specifier) =>
    specifier === "@microsoft/teamsfx-core"
      ? coreEntry
      : require.resolve(specifier);
  const originalFetch = global.fetch;
  context.after(() => {
    global.fetch = originalFetch;
  });
  let requests = 0;
  global.fetch = async (url, options) => {
    requests++;
    assert.equal(
      url,
      "https://graph.example.test/v1.0/oAuthConfigurations/owned-id",
    );
    assert.equal(options.method, undefined);
    return new Response(JSON.stringify(f.state()), { status: 200 });
  };
  await runtime()("baseline", f.project, {
    ...f.options,
    load,
    readSnapshot: undefined,
  });
  assert.equal(requests, 1);
});

test("VCB-186: DCR profile does not read or export an initial endpoint setting", () => {
  const { loadEnvironmentProfile } = require("./prepare-mcp-auth-run.cjs");
  const environment = {
    M365_TENANT_ID: "16801f2e-7163-4bd5-8542-f083fa59993c",
    MCP_AUTH_ENDPOINT: endpoint,
  };
  Object.defineProperty(environment, "MCP_AUTH_INITIAL_ENDPOINT", {
    enumerable: true,
    get() {
      throw new Error("Initial endpoint must not be read");
    },
  });
  const profile = loadEnvironmentProfile("dcr", environment);
  assert.equal(
    Object.hasOwn(profile.environment, "MCP_AUTH_INITIAL_ENDPOINT"),
    false,
  );
});

test("VCB-191: preparation keeps the real endpoint and explicitly disables personal accounts", () => {
  const { dcrUpdateScript } = require("./dcr-update-fixture.cjs");
  const script = dcrUpdateScript("prepare");
  assert.ok(!script.includes("MCP_AUTH_INITIAL_ENDPOINT"));
  assert.ok(!script.includes("example.invalid"));
  assert.ok(
    script.includes('inputs["includePersonalMicrosoftAccounts"] = False'),
  );
  assert.ok(!script.includes("urllib.request"));
});

test("VCB-186: standalone update CI routing needs no new input or hosted service", async () => {
  const { parse } = require("yaml");
  const { prepareMcpAuthCiRun } = require("./prepare-mcp-auth-run.cjs");
  const config = await fs.readFile(
    path.join(__dirname, "..", "config.yaml"),
    "utf8",
  );
  const plan = JSON.stringify({
    steps: [{ tags: ["mcp_auth_mode:dcr", "operation:dcr-update-stage"] }],
  });
  const environment = {
    M365_TENANT_ID: "16801f2e-7163-4bd5-8542-f083fa59993c",
    MCP_AUTH_ENDPOINT: endpoint,
    M365_ACCOUNT_NAME: "licensed@atktesting08.onmicrosoft.com",
    M365_ACCOUNT_PASSWORD: "test-password",
  };
  const prepared = prepareMcpAuthCiRun(config, plan, environment);
  assert.equal(
    Object.hasOwn(
      parse(prepared.configText).docker.environment,
      "MCP_AUTH_INITIAL_ENDPOINT",
    ),
    false,
  );
  const workflow = await fs.readFile(
    path.join(
      __dirname,
      "..",
      "..",
      "..",
      "..",
      "..",
      ".github",
      "workflows",
      "ui-test-vscuse-common.yml",
    ),
    "utf8",
  );
  assert.ok(workflow.includes("da-mcp-server--da-mcp-dcr-update-copilot"));
  assert.ok(!workflow.includes("MCP_AUTH_INITIAL_ENDPOINT"));
  const caller = parse(
    await fs.readFile(
      path.join(
        __dirname,
        "..",
        "..",
        "..",
        "..",
        "..",
        ".github",
        "workflows",
        "ui-test-vscuse-template.yml",
      ),
      "utf8",
    ),
  );
  assert.equal(
    Object.hasOwn(caller.on.workflow_dispatch.inputs, "mcp_initial_endpoint"),
    false,
  );
  assert.ok(!workflow.includes("inputs.mcp_initial_endpoint"));
  for (const name of ["M365_ACCOUNT_NAME", "M365_TENANT_ID"]) {
    assert.equal(
      parse(prepared.configText).docker.environment[name],
      "${" + name + "}",
    );
  }
});
