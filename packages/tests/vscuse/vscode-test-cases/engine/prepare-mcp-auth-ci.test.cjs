const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const test = require("node:test");
const { parse } = require("yaml");
const { compileCaseBundle } = require("./compile-case-bundle.cjs");
const { createSemanticStepCompiler } = require("./semantic-step-compiler.cjs");
const {
  buildMcpAuthProfile,
  expectedTenant,
} = require("./mcp-auth-fixture.cjs");

const helper = require("./prepare-mcp-auth-run.cjs");
const client = "7c5898c1-99fb-4d83-8474-02bf7985096c";
const configText = `docker:
  image_name: pinned-image
  environment:
    TEMPLATE_VERSION: pinned-routing
    TEST_AAD_APP_ID: foreign-aad
    DA_MCP_OAUTH_CLIENT_SECRET: foreign-secret
execution:
  stop_on_error: true
`;

function inputs() {
  return {
    M365_TENANT_ID: expectedTenant,
    M365_ACCOUNT_NAME: "test019@atktesting08.onmicrosoft.com",
    M365_ACCOUNT_PASSWORD: "public-test-password",
    MCP_AUTH_ENDPOINT: "https://mcp.example.com/mcp",
    DA_MCP_OAUTH_CLIENTID: client,
    DA_MCP_OAUTH_CLIENT_SECRET: "public-test-client-secret",
    DA_MCP_ENTRA_SSO_CLIENTID: client,
    MCP_AUTH_SSO_ID: "opaque-registration/id+padding==",
  };
}

async function plans() {
  const result = await compileCaseBundle({
    sourcePath: "cases/da-mcp-auth-modes.yml",
    sourceText: await fs.readFile(
      path.join(__dirname, "..", "cases", "da-mcp-auth-modes.yml"),
      "utf8",
    ),
    compileStep: createSemanticStepCompiler(),
  });
  assert.equal(result.ok, true, JSON.stringify(result.diagnostics));
  return result.value.map(({ plan }) => plan);
}

test("VCB-182: static resource scope is derived without reading an independent scope setting", async () => {
  const consumer = {
    ENTRA_TENANT_ID: expectedTenant,
    ENTRA_CLIENT_ID: client,
    MCP_ENDPOINT_URL: "https://mcp.example.com/mcp",
  };
  const credentials = {
    ENTRA_TENANT_ID: expectedTenant,
    ENTRA_CLIENT_ID: client,
    ENTRA_CLIENT_SECRET: "public-test-secret",
  };
  Object.defineProperty(credentials, "ENTRA_API_SCOPE", {
    get() {
      throw new Error("Independent scope was read");
    },
  });
  const profile = buildMcpAuthProfile("static-oauth", consumer, {
    staticCredentials: credentials,
  });
  assert.equal(Object.hasOwn(profile.environment, "DA_MCP_OAUTH_SCOPE"), false);
  for (const plan of await plans()) {
    assert.equal(
      JSON.stringify(plan).includes("${{env:DA_MCP_OAUTH_SCOPE}}"),
      false,
    );
    const scope = plan.steps.find(
      (step) =>
        step.tool === "type_text" &&
        step.parameters.text ===
          "api://${{env:DA_MCP_OAUTH_CLIENTID}}/access_as_user",
    );
    assert.equal(Boolean(scope), helper.fixtureMode(plan) === "static-oauth");
  }
});

test("VCB-183: CI derives scoped runtime values and preserves workflow template routing", async () => {
  for (const plan of await plans()) {
    const result = helper.prepareMcpAuthCiRun(
      configText,
      JSON.stringify(plan),
      inputs(),
    );
    const mode = helper.fixtureMode(plan);
    assert.equal(result.mode, mode);
    assert.equal(result.environment.MCP_AUTH_MODE, mode);
    assert.equal(
      result.environment.MCP_AUTH_EXPECTED_FLOW,
      mode === "entra-sso" ? "entra_sso" : "dcr_oauth",
    );
    assert.equal(result.environment.M365_TENANT_ID, expectedTenant);
    assert.equal(
      result.environment.M365_ACCOUNT_NAME,
      inputs().M365_ACCOUNT_NAME,
    );
    assert.equal(Boolean(result.environment.MCP_AUTH_RUN_ID), false);
    assert.equal(Boolean(result.environment.DA_MCP_OAUTH_SCOPE), false);
    const config = parse(result.configText);
    assert.equal(config.docker.environment.TEMPLATE_VERSION, "pinned-routing");
    assert.equal(
      Object.hasOwn(config.docker.environment, "TEST_AAD_APP_ID"),
      false,
    );
    assert.equal(result.configText.includes("public-test-password"), false);
    assert.equal(
      result.configText.includes("public-test-client-secret"),
      false,
    );
    assert.equal(result.configText.includes("TEAMSFX_V4_ENABLED"), false);
    if (mode !== "static-oauth") {
      assert.equal(
        Object.hasOwn(config.docker.environment, "DA_MCP_OAUTH_CLIENT_SECRET"),
        false,
      );
    }
    if (mode !== "entra-sso") {
      assert.equal(
        Object.hasOwn(config.docker.environment, "MCP_AUTH_SSO_ID"),
        false,
      );
    }
  }
});

test("VCB-183: environment profiles never read credentials from another authentication mode", async () => {
  const compiled = await plans();
  for (const mode of ["static-oauth", "dcr", "entra-sso"]) {
    const environment = inputs();
    const forbidden = [
      "DA_MCP_OAUTH_SCOPE",
      "ENTRA_API_SCOPE",
      ...(mode !== "entra-sso"
        ? ["MCP_AUTH_SSO_ID", "MCP_AUTH_AAD_MANIFEST_B64"]
        : []),
      ...(mode !== "static-oauth"
        ? ["DA_MCP_OAUTH_CLIENTID", "DA_MCP_OAUTH_CLIENT_SECRET"]
        : []),
      ...(mode !== "entra-sso" ? ["DA_MCP_ENTRA_SSO_CLIENTID"] : []),
    ];
    for (const name of forbidden) {
      Object.defineProperty(environment, name, {
        enumerable: true,
        get() {
          throw new Error("Foreign input read: " + name);
        },
      });
    }
    const result = helper.loadEnvironmentProfile(mode, environment);
    assert.equal(result.mode, mode);
    assert.equal(Boolean(result.environment.DA_MCP_OAUTH_SCOPE), false);
    assert.equal(
      Boolean(result.environment.MCP_AUTH_SSO_ID),
      mode === "entra-sso",
    );
    const plan = compiled.find((item) => helper.fixtureMode(item) === mode);
    const prepared = helper.prepareMcpAuthCiRun(
      configText,
      JSON.stringify(plan),
      environment,
    );
    assert.equal(
      Boolean(prepared.environment.DA_MCP_OAUTH_CLIENT_SECRET),
      mode === "static-oauth",
    );
    assert.equal(
      Boolean(prepared.environment.MCP_AUTH_SSO_ID),
      mode === "entra-sso",
    );
  }
});

test("VCB-183: SSO reuse avoids manifest access and absent IDs preserve full creation inputs", () => {
  const reuse = inputs();
  Object.defineProperty(reuse, "MCP_AUTH_AAD_MANIFEST_B64", {
    get() {
      throw new Error("Shared AAD configuration was read on reuse");
    },
  });
  const reused = helper.loadEnvironmentProfile("entra-sso", reuse);
  assert.equal(reused.environment.MCP_AUTH_SSO_ID, reuse.MCP_AUTH_SSO_ID);
  assert.equal(
    Object.hasOwn(reused.environment, "MCP_AUTH_AAD_MANIFEST_B64"),
    false,
  );
  const inaccessible = helper.loadEnvironmentProfile("entra-sso", {
    ...inputs(),
    MCP_AUTH_SSO_ID: "inaccessible-registration",
  });
  assert.equal(
    inaccessible.environment.MCP_AUTH_SSO_ID,
    "inaccessible-registration",
  );
  assert.equal(
    Object.hasOwn(inaccessible.environment, "MCP_AUTH_AAD_MANIFEST_B64"),
    false,
  );
  const create = inputs();
  create.MCP_AUTH_SSO_ID = "";
  assert.throws(
    () => helper.loadEnvironmentProfile("entra-sso", create),
    /MCP_AUTH_AAD_MANIFEST_B64/,
  );
  const manifest = {
    appId: "${{AAD_APP_CLIENT_ID}}",
    identifierUris: [
      "api://${{AAD_APP_CLIENT_ID}}",
      "${{ENTRA_SSO_APPLICATION_ID_URI}}",
    ],
    web: { redirectUris: ["https://mini.example.com/callback"] },
    api: { preAuthorizedApplications: [{ appId: "shared-token-store" }] },
  };
  Object.assign(create, {
    MCP_AUTH_AAD_MANIFEST_B64: Buffer.from(JSON.stringify(manifest)).toString(
      "base64",
    ),
    MCP_AUTH_AAD_APP_OBJECT_ID: "fixture-object",
    MCP_AUTH_AAD_ACCESS_AS_USER_PERMISSION_ID: "fixture-permission",
    ENTRA_SSO_APPLICATION_ID_URI: `api://auth-existing/${client}`,
  });
  const created = helper.loadEnvironmentProfile("entra-sso", create);
  const actual = JSON.parse(
    Buffer.from(created.environment.MCP_AUTH_AAD_MANIFEST_B64, "base64"),
  );
  assert.equal(created.environment.MCP_AUTH_SSO_ID, "");
  assert.deepEqual(actual.web, manifest.web);
  assert.deepEqual(actual.api, manifest.api);
  assert.ok(actual.identifierUris.includes("${{MCP_AUTH_NEW_SSO_URI}}"));
});

test("VCB-183: missing credentials and the wrong account or tenant fail before runner execution", async () => {
  assert.equal(typeof helper.prepareMcpAuthCiRun, "function");
  const [plan] = await plans();
  for (const change of [
    { M365_ACCOUNT_PASSWORD: "" },
    { M365_ACCOUNT_NAME: "other@example.invalid" },
    { M365_TENANT_ID: "other-tenant" },
    { MCP_AUTH_ENDPOINT: "" },
    { DA_MCP_OAUTH_CLIENT_SECRET: "" },
    { DA_MCP_OAUTH_CLIENTID: "" },
  ]) {
    assert.throws(() =>
      helper.prepareMcpAuthCiRun(configText, JSON.stringify(plan), {
        ...inputs(),
        ...change,
      }),
    );
  }
});

test("VCB-183: the actual CI CLI masks secrets and writes delimited host values plus reference-only Docker config", async () => {
  const directory = await fs.mkdtemp(
    path.join(os.tmpdir(), "mcp-ci-preparation-"),
  );
  try {
    const [plan] = await plans();
    const inputConfig = path.join(directory, "config.yaml");
    const planFile = path.join(directory, "plan.json");
    const output = path.join(directory, "config.mcp.yaml");
    const envFile = path.join(directory, "github-env");
    await fs.writeFile(inputConfig, configText);
    await fs.writeFile(planFile, JSON.stringify(plan));
    const result = spawnSync(
      process.execPath,
      [
        path.join(__dirname, "prepare-mcp-auth-run.cjs"),
        inputConfig,
        planFile,
        output,
      ],
      {
        encoding: "utf8",
        env: {
          ...process.env,
          ...inputs(),
          GITHUB_ENV: envFile,
          DA_MCP_OAUTH_SCOPE: "stale-scope",
          M365_ACCOUNT_PASSWORD: "public-password\nsecond-line%value",
        },
      },
    );
    assert.equal(result.status, 0, result.stderr);
    assert.ok(result.stdout.includes("::add-mask::"));
    const encodedPassword = encodeURIComponent(
      "public-password\nsecond-line%value",
    ).replaceAll("%", "%25");
    assert.ok(result.stdout.includes(`::add-mask::${encodedPassword}\n`));
    assert.equal(
      result.stdout
        .split("\n")
        .filter((line) => !line.startsWith("::add-mask::"))
        .join("\n")
        .includes("public-password"),
      false,
    );
    const text = await fs.readFile(envFile, "utf8");
    assert.ok(text.includes("MCP_AUTH_EXPECTED_FLOW<<"));
    assert.ok(text.includes("public-password\nsecond-line%value"));
    assert.ok(text.includes("DA_MCP_OAUTH_SCOPE<<"));
    const prepared = await fs.readFile(output, "utf8");
    assert.equal(prepared.includes("public-password"), false);
    assert.equal(prepared.includes("public-test-client-secret"), false);
    assert.equal(prepared.includes("DA_MCP_OAUTH_SCOPE"), false);
    assert.equal(prepared.includes("MCP_AUTH_SSO_ID"), false);
    const failed = spawnSync(
      process.execPath,
      [
        path.join(__dirname, "prepare-mcp-auth-run.cjs"),
        inputConfig,
        planFile,
        path.join(directory, "failed.yaml"),
      ],
      {
        encoding: "utf8",
        env: {
          ...process.env,
          ...inputs(),
          GITHUB_ENV: path.join(directory, "failed-env"),
          DA_MCP_OAUTH_CLIENT_SECRET: "",
        },
      },
    );
    assert.equal(failed.status, 1);
    assert.ok(failed.stderr.includes("DA_MCP_OAUTH_CLIENT_SECRET"));
    await assert.rejects(fs.access(path.join(directory, "failed.yaml")));
    await assert.rejects(fs.access(path.join(directory, "failed-env")));
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("VCB-183: workflow reads public vars and gates the prepared MCP config without changing generic routing", async () => {
  const workflow = parse(
    await fs.readFile(
      path.resolve(
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
    ),
  );
  const job = workflow.jobs.main;
  for (const name of [
    "M365_TENANT_ID",
    "DA_MCP_OAUTH_CLIENTID",
    "DA_MCP_ENTRA_SSO_CLIENTID",
    "MCP_AUTH_ENDPOINT",
  ]) {
    assert.equal(job.env[name], "${{ vars." + name + " }}", name);
  }
  assert.equal(
    job.env.DA_MCP_OAUTH_CLIENT_SECRET,
    "${{ secrets.DA_MCP_OAUTH_CLIENT_SECRET }}",
  );
  const preparation = job.steps.find((step) => step.id === "prepare-mcp-auth");
  assert.ok(preparation);
  for (const id of [
    "da-mcp-static-entra-copilot",
    "da-mcp-dcr-copilot",
    "da-mcp-entra-sso-copilot",
  ]) {
    assert.ok(preparation.if.includes(id), id);
  }
  assert.ok(
    preparation.env.MCP_AUTH_SSO_ID.includes(
      "secrets.ENTRA_SSO_REGISTRATION_ID",
    ),
  );
  assert.ok(preparation.env.MCP_AUTH_SSO_ID.includes("matrix.test_plan"));
  assert.ok(
    preparation.env.DA_MCP_OAUTH_CLIENT_SECRET.includes("matrix.test_plan"),
  );
  assert.equal(
    preparation.env.M365_ACCOUNT_NAME,
    "${{ vars.M365_ACCOUNT_NAME_ENABLECOPILOTACCESS }}",
  );
  assert.equal(
    preparation.env.M365_ACCOUNT_PASSWORD,
    "${{ secrets.M365_ACCOUNT_PASSWORD_ENABLECOPILOTACCESS }}",
  );
  assert.ok(preparation.run.includes("prepare-mcp-auth-run.cjs"));
  const random = job.steps.find(
    (step) => step.name === "Set m365 account randomly",
  );
  assert.ok(random.if.includes("da-mcp-dcr-copilot"));
  assert.ok(random.if.includes("!contains"));
  const selected = JSON.parse(preparation.if.match(/fromJSON\('([^']+)'\)/)[1]);
  const skipped = JSON.parse(random.if.match(/fromJSON\('([^']+)'\)/)[1]);
  assert.equal(selected.length, 3);
  assert.deepEqual(skipped, selected);
  const execution = job.steps.find((step) => step.name === "Run test");
  assert.ok(execution.run.includes("steps.prepare-mcp-auth.outcome"));
  assert.ok(execution.run.includes("config.mcp.yaml"));
  assert.ok(execution.run.includes("prepare-vscuse-config.cjs"));
  assert.equal(execution.run.includes("run-mcp-auth-integration.ps1"), false);
});
