const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const yaml = require("js-yaml");

const initial = "https://mock.example.test/mcp";
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
      mcpResourceUrl: initial,
      targetUrlsShouldStartWith: [initial],
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
          spec: { url: initial },
          auth: { type: "OAuthPluginVault", reference_id: "${{OWNED_DCR_ID}}" },
        },
      ],
    }),
  );
  await fs.writeFile(
    path.join(project, ".vscode", "mcp.json"),
    JSON.stringify({
      servers: { owned: { type: "http", url: initial } },
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
    mcpResourceUrl: initial,
    targetUrlsShouldStartWith: [initial],
    clientSecret: "must-not-be-persisted",
  };
  const options = {
    load: require,
    initialEndpoint: initial,
    endpoint,
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

test("VCB-185 / DCR-UI-UPD-03: edits only owned URL wiring and adds a sparse update", async (context) => {
  const run = runtime();
  const f = await fixture(context);
  const env = await fs.readFile(
    path.join(f.project, "env", ".env.dev"),
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
      mcpResourceUrl: endpoint,
      targetUrlsShouldStartWith: [endpoint],
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
});

test("VCB-185 / DCR-UI-UPD-04: only persisted URL changes pass verification", async (context) => {
  const run = runtime();
  const f = await fixture(context);
  await run("baseline", f.project, f.options);
  await run("apply", f.project, f.options);
  await assert.rejects(
    run("verify", f.project, f.options),
    /endpoint|URL|persist/i,
  );
  f.setState({
    ...f.state(),
    mcpResourceUrl: endpoint,
    targetUrlsShouldStartWith: [endpoint],
    clientSecret: "rediscovered-secret",
  });
  await run("verify", f.project, f.options);
  await run("noop", f.project, f.options);
  f.setState({ ...f.state(), targetAudience: "AnyTenant" });
  await assert.rejects(
    run("noop", f.project, f.options),
    /policy|targetAudience/i,
  );
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
  const f = await fixture(context);
  f.setState({ ...f.state(), clientName: null });
  await assert.rejects(
    run("baseline", f.project, f.options),
    /configuration|readback/i,
  );
});

test("VCB-186: DCR profile alone accepts a distinct initial HTTPS endpoint", () => {
  const { loadEnvironmentProfile } = require("./prepare-mcp-auth-run.cjs");
  const environment = {
    M365_TENANT_ID: "16801f2e-7163-4bd5-8542-f083fa59993c",
    MCP_AUTH_ENDPOINT: endpoint,
    MCP_AUTH_INITIAL_ENDPOINT: initial,
  };
  const profile = loadEnvironmentProfile("dcr", environment);
  assert.equal(profile.environment.MCP_AUTH_INITIAL_ENDPOINT, initial);
  for (const invalid of [
    endpoint,
    "http://mock.example.test/mcp",
    "https://user:pass@mock.example.test/mcp",
  ]) {
    assert.throws(() =>
      loadEnvironmentProfile("dcr", {
        ...environment,
        MCP_AUTH_INITIAL_ENDPOINT: invalid,
      }),
    );
  }
});

test("VCB-187: mock discovery is real-resource metadata without tokens or real tool results", async (context) => {
  const file = path.join(__dirname, "dcr-update-mock.cjs");
  assert.ok(fsSync.existsSync(file), "Missing controlled DCR update mock");
  const { createDcrUpdateMock } = require(file);
  const metadata = {
    resource: endpoint,
    authorization_servers: ["https://real.example.test"],
    scopes_supported: ["access_as_user"],
  };
  const server = createDcrUpdateMock(initial, metadata);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  context.after(() => new Promise((resolve) => server.close(resolve)));
  const local = "http://127.0.0.1:" + server.address().port;
  const response = await fetch(local + "/mcp", { method: "POST", body: "{}" });
  assert.equal(response.status, 401);
  assert.ok(
    response.headers
      .get("www-authenticate")
      .includes(
        "https://mock.example.test/.well-known/oauth-protected-resource",
      ),
  );
  const discovery = await fetch(
    local + "/.well-known/oauth-protected-resource/mcp",
  );
  assert.deepEqual(await discovery.json(), metadata);
  assert.ok(!(await response.text()).includes("PURPLE-NARWHAL-7788"));
  assert.equal((await fetch(local + "/token", { method: "POST" })).status, 404);
});

test("VCB-186: standalone update CI routing requires the public mock endpoint", async () => {
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
  assert.throws(
    () => prepareMcpAuthCiRun(config, plan, environment),
    /MCP_AUTH_INITIAL_ENDPOINT/,
  );
  const prepared = prepareMcpAuthCiRun(config, plan, {
    ...environment,
    MCP_AUTH_INITIAL_ENDPOINT: initial,
  });
  assert.equal(
    parse(prepared.configText).docker.environment.MCP_AUTH_INITIAL_ENDPOINT,
    "${MCP_AUTH_INITIAL_ENDPOINT}",
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
  assert.ok(workflow.includes("vars.MCP_AUTH_INITIAL_ENDPOINT"));
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
  assert.ok(caller.on.workflow_dispatch.inputs.mcp_initial_endpoint);
  assert.ok(workflow.includes("inputs.mcp_initial_endpoint"));
  for (const name of ["M365_ACCOUNT_NAME", "M365_TENANT_ID"]) {
    assert.equal(
      parse(prepared.configText).docker.environment[name],
      "${" + name + "}",
    );
  }
});
