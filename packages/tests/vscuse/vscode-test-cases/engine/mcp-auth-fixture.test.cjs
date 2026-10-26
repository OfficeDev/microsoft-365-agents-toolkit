const assert = require("node:assert/strict");
const test = require("node:test");
const fs = require("node:fs/promises");
const path = require("node:path");
const os = require("node:os");
const {
  buildMcpAuthProfile,
  fixtureScript,
  expectedTenant,
  expectedCanary,
} = require("./mcp-auth-fixture.cjs");
const { compileCaseBundle } = require("./compile-case-bundle.cjs");
const { createSemanticStepCompiler } = require("./semantic-step-compiler.cjs");
const { parse } = require("yaml");
const {
  prepareMcpAuthRunConfig,
  loadProfile,
  redact,
} = require("./prepare-mcp-auth-run.cjs");

const client = "7c5898c1-99fb-4d83-8474-02bf7985096c";
function consumer() {
  return {
    ENTRA_TENANT_ID: expectedTenant,
    ENTRA_CLIENT_ID: client,
    MCP_ENDPOINT_URL: "https://mcp.example.com/mcp",
    ENTRA_SSO_APPLICABLE_TO_APPS: "AnyApp",
    ENTRA_SSO_TARGET_AUDIENCE: "HomeTenant",
    ENTRA_SSO_REGISTRATION_ID: "opaque/id+padding==",
    ENTRA_SSO_APPLICATION_ID_URI: `api://auth-existing/${client}`,
  };
}

test("MCP-AUTH-10: every mode rejects a stale resource application identity", () => {
  for (const mode of ["static-oauth", "dcr", "entra-sso"]) {
    assert.throws(
      () =>
        buildMcpAuthProfile(
          mode,
          {
            ...consumer(),
            ENTRA_CLIENT_ID: "11111111-2222-3333-4444-555555555555",
          },
          {
            staticCredentials: {
              ENTRA_TENANT_ID: expectedTenant,
              ENTRA_CLIENT_ID: "11111111-2222-3333-4444-555555555555",
              ENTRA_CLIENT_SECRET: "test-only-secret",
              ENTRA_API_SCOPE: "api://old-test-client/access_as_user",
            },
          },
        ),
      /client.*mismatch/i,
      mode,
    );
  }
});

test("MCP-AUTH-10: latest profiles isolate run ownership and credential exports", () => {
  const credentials = {
    ENTRA_TENANT_ID: expectedTenant,
    ENTRA_CLIENT_ID: client,
    ENTRA_CLIENT_SECRET: "test-only-secret",
    ENTRA_API_SCOPE: `api://${client}/access_as_user`,
  };
  const profiles = ["static-oauth", "dcr", "entra-sso"].map((mode) =>
    buildMcpAuthProfile(mode, consumer(), { staticCredentials: credentials }),
  );
  assert.equal(
    new Set(profiles.map((profile) => profile.environment.MCP_AUTH_RUN_ID))
      .size,
    3,
  );
  for (const profile of profiles) {
    assert.match(profile.environment.MCP_AUTH_RUN_ID, /^[0-9a-f-]{36}$/);
  }
  assert.equal(profiles[0].environment.DA_MCP_OAUTH_CLIENTID, client);
  assert.equal(profiles[2].environment.DA_MCP_ENTRA_SSO_CLIENTID, client);
  assert.ok(!Object.hasOwn(profiles[1].environment, "DA_MCP_OAUTH_CLIENTID"));
  assert.ok(!Object.hasOwn(profiles[1].environment, "MCP_AUTH_SSO_ID"));
  assert.equal(profiles[0].environment.MCP_AUTH_EXPECTED_JWT_CLIENT_ID, client);
  assert.equal(profiles[1].environment.MCP_AUTH_EXPECTED_JWT_CLIENT_ID, client);
  assert.equal(
    profiles[2].environment.MCP_AUTH_EXPECTED_JWT_CLIENT_ID,
    "ab3be6b7-f5df-413d-ac2d-abf1e3fd9c0b",
  );
  assert.deepEqual(
    profiles.map((profile) => profile.environment.MCP_AUTH_EXPECTED_FLOW),
    ["dcr_oauth", "dcr_oauth", "entra_sso"],
  );
});

test("MCP-AUTH-04: naturally missing or empty SSO ID retains full creation setup", () => {
  for (const value of [undefined, ""]) {
    const env = consumer();
    if (value === undefined) delete env.ENTRA_SSO_REGISTRATION_ID;
    else env.ENTRA_SSO_REGISTRATION_ID = value;
    const profile = buildMcpAuthProfile("entra-sso", env, {
      miniEnvironment: {
        AAD_APP_CLIENT_ID: client,
        AAD_APP_OBJECT_ID: "fixture-object-id",
        AAD_APP_ACCESS_AS_USER_PERMISSION_ID: "fixture-scope-id",
      },
      miniManifest: {
        appId: client,
        identifierUris: ["${{ENTRA_SSO_APPLICATION_ID_URI}}"],
        web: { redirectUris: ["https://mini.example.com/callback"] },
        api: { preAuthorizedApplications: [{ appId: "token-store" }] },
      },
    });
    assert.equal(profile.environment.MCP_AUTH_SSO_ID, "");
    assert.ok(profile.environment.MCP_AUTH_AAD_MANIFEST_B64);
  }
});

test("MCP-AUTH-05: an inaccessible supplied SSO ID is not converted into creation setup", () => {
  const profile = buildMcpAuthProfile("entra-sso", {
    ...consumer(),
    ENTRA_SSO_REGISTRATION_ID: "inaccessible-registration",
  });
  assert.equal(
    profile.environment.MCP_AUTH_SSO_ID,
    "inaccessible-registration",
  );
  assert.ok(!Object.hasOwn(profile.environment, "MCP_AUTH_AAD_MANIFEST_B64"));
});

test("MCP-AUTH-04: the file loader defaults to full setup when the latest export omits its SSO ID", async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "mcp-latest-profile-"));
  try {
    await fs.mkdir(path.join(root, ".generated"));
    await fs.mkdir(path.join(root, "env"));
    const publicProfile = consumer();
    delete publicProfile.ENTRA_SSO_REGISTRATION_ID;
    const source = path.join(root, ".generated", "da-sso.atktesting08.env");
    await fs.writeFile(
      source,
      Object.entries(publicProfile)
        .map(([name, value]) => `${name}=${value}`)
        .join("\n"),
    );
    await fs.writeFile(
      path.join(root, "env", ".env.atktesting08"),
      [
        `AAD_APP_CLIENT_ID=${client}`,
        "AAD_APP_OBJECT_ID=fixture-object",
        "AAD_APP_ACCESS_AS_USER_PERMISSION_ID=fixture-scope",
      ].join("\n"),
    );
    await fs.writeFile(
      path.join(root, "aad.manifest.json"),
      JSON.stringify({
        appId: client,
        identifierUris: ["${{ENTRA_SSO_APPLICATION_ID_URI}}"],
      }),
    );
    const profile = loadProfile("entra-sso", source);
    assert.equal(profile.environment.MCP_AUTH_SSO_ID, "");
    assert.ok(profile.environment.MCP_AUTH_AAD_MANIFEST_B64);
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});

function source(mode) {
  const auth = {
    "static-oauth": "oauth",
    dcr: "oauth-dynamic",
    "entra-sso": "entra-sso",
  }[mode];
  return `version: 1
cases:
  - id: auth-test
    scenarioId: SCN-DA-CREATE-WITH-MCP-SERVER
    workItemIds: [1001]
    steps: [create, check, fixture]
steps:
  create:
    type: scaffold
    with:
      template: da/mcp-server
      answers:
        - question: projectType
          value: copilot-agent-type
        - question: daTemplate
          value: add-action
        - question: actionSource
          value: mcp
        - question: mcpServerUrl
          type: text
          value: https://mcp.example.com/mcp
        - question: authType
          value: ${auth}
        - question: workspaceFolder
          value: default
        - question: appName
          type: text
          value: "\${{var:app_name:mcp_auth_#####}}"
  check:
    type: checks
    with:
      - type: file
        path: appPackage/ai-plugin.json
        expect:
          contains: ['"RemoteMCPServer"']
  fixture:
    type: configureMcpAuth
    with:
      mode: ${mode}
`;
}

for (const mode of ["static-oauth", "dcr", "entra-sso"]) {
  test(`VCB-168: compiles checked ${mode} preparation without embedding credential values`, async () => {
    const result = await compileCaseBundle({
      sourcePath: "cases/auth-test.yml",
      sourceText: source(mode),
      compileStep: createSemanticStepCompiler(),
    });
    assert.equal(result.ok, true, JSON.stringify(result.diagnostics));
    const step = result.value[0].plan.steps.find((item) =>
      item.tags.includes("operation:mcp-auth-fixture"),
    );
    assert.ok(step);
    assert.ok(step.tags.includes(`mcp_auth_mode:${mode}`));
    const body = step.parameters.text;
    const encoded = body.match(/SCRIPT_B64="([^"]+)"/)[1];
    const script = Buffer.from(encoded, "base64").toString("utf8");
    assert.equal(script.includes("MCP_AUTH_SSO_ID"), mode === "entra-sso");
    assert.ok(!body.includes("opaque/id+padding=="));
  });
}

test("VCB-168: rejects mismatched fixture/scaffold authentication modes", async () => {
  const result = await compileCaseBundle({
    sourcePath: "cases/auth-test.yml",
    sourceText: source("dcr").replace("mode: dcr", "mode: entra-sso"),
    compileStep: createSemanticStepCompiler(),
  });
  assert.equal(result.ok, false);
});

test("VCB-168: rejects repeated preparation instead of clearing a case's created state", async () => {
  const result = await compileCaseBundle({
    sourcePath: "cases/auth-test.yml",
    sourceText: source("dcr").replace(
      "steps: [create, check, fixture]",
      "steps: [create, check, fixture, fixture]",
    ),
    compileStep: createSemanticStepCompiler(),
  });
  assert.equal(result.ok, false);
  assert.equal(result.diagnostics[0].code, "VCB_MCP_AUTH_FIXTURE_INVALID");
});

for (const mode of ["static-oauth", "dcr"]) {
  test(`VCB-169: ${mode} never reads or exports the SSO registration ID`, () => {
    const env = consumer();
    Object.defineProperty(env, "ENTRA_SSO_REGISTRATION_ID", {
      get() {
        throw new Error("Non-SSO mode read the SSO ID");
      },
    });
    const credentials = {
      ENTRA_TENANT_ID: expectedTenant,
      ENTRA_CLIENT_ID: client,
      ENTRA_CLIENT_SECRET: "test-only-static-secret",
      ENTRA_API_SCOPE: `api://${client}/access_as_user`,
    };
    const result = buildMcpAuthProfile(mode, env, {
      staticCredentials: credentials,
    });
    assert.ok(!Object.hasOwn(result.environment, "MCP_AUTH_SSO_ID"));
    if (mode === "dcr") {
      assert.ok(
        !Object.hasOwn(result.environment, "DA_MCP_OAUTH_CLIENT_SECRET"),
      );
      assert.ok(!Object.hasOwn(result.environment, "DA_MCP_OAUTH_CLIENTID"));
    } else {
      assert.ok(
        result.environment.MCP_AUTH_AUTHORIZATION_URL.startsWith(
          `https://login.microsoftonline.com/${expectedTenant}/`,
        ),
      );
    }
  });
}

test("VCB-169: SSO preserves the opaque ID and requires no private credentials on reuse", () => {
  const result = buildMcpAuthProfile("entra-sso", consumer());
  assert.equal(result.environment.MCP_AUTH_SSO_ID, "opaque/id+padding==");
  assert.ok(!Object.hasOwn(result.environment, "DA_MCP_OAUTH_CLIENT_SECRET"));
  assert.ok(!Object.hasOwn(result.environment, "MCP_AUTH_AAD_MANIFEST_B64"));
});

test("VCB-169: absent-ID setup preserves the mini manifest and adds only its new URI", () => {
  const env = consumer();
  const manifest = {
    id: "${{AAD_APP_OBJECT_ID}}",
    appId: "${{AAD_APP_CLIENT_ID}}",
    identifierUris: [
      "api://${{AAD_APP_CLIENT_ID}}",
      "${{ENTRA_SSO_APPLICATION_ID_URI}}",
    ],
    web: { redirectUris: ["https://mini.example.com/callback"] },
    api: {
      preAuthorizedApplications: [
        { appId: "token-store", delegatedPermissionIds: ["scope"] },
      ],
    },
  };
  const before = structuredClone(manifest);
  const result = buildMcpAuthProfile("entra-sso", env, {
    ssoVariant: "absent",
    miniEnvironment: {
      AAD_APP_CLIENT_ID: client,
      AAD_APP_OBJECT_ID: "object-id",
      AAD_APP_ACCESS_AS_USER_PERMISSION_ID: "scope-id",
    },
    miniManifest: manifest,
  });
  const prepared = JSON.parse(
    Buffer.from(
      result.environment.MCP_AUTH_AAD_MANIFEST_B64,
      "base64",
    ).toString("utf8"),
  );
  assert.equal(result.environment.MCP_AUTH_SSO_ID, "");
  assert.ok(result.environment.MCP_AUTH_SSO_PROTECTED_ID_B64);
  assert.equal(
    JSON.parse(
      Buffer.from(
        result.environment.MCP_AUTH_SSO_PROTECTED_ID_B64,
        "base64",
      ).toString("utf8"),
    ),
    env.ENTRA_SSO_REGISTRATION_ID,
  );
  assert.deepEqual(manifest, before);
  assert.deepEqual(prepared.web, before.web);
  assert.deepEqual(prepared.api, before.api);
  assert.deepEqual(prepared.identifierUris, [
    `api://${client}`,
    `api://auth-existing/${client}`,
    "${{MCP_AUTH_NEW_SSO_URI}}",
  ]);
});

test("VCB-169: wrong tenants, mismatched static clients and unsafe SSO scopes fail explicitly", () => {
  assert.throws(() =>
    buildMcpAuthProfile("dcr", { ...consumer(), ENTRA_TENANT_ID: "other" }),
  );
  assert.throws(() =>
    buildMcpAuthProfile("static-oauth", consumer(), { staticCredentials: {} }),
  );
  assert.throws(() =>
    buildMcpAuthProfile("entra-sso", {
      ...consumer(),
      ENTRA_SSO_TARGET_AUDIENCE: "AnyTenant",
    }),
  );
  assert.throws(() => fixtureScript("unknown"));
});

test("MCP-AUTH-10: integration cases use product feature defaults without injecting overrides", async () => {
  const sourceText = await fs.readFile(
    path.join(__dirname, "..", "cases", "da-mcp-auth-modes.yml"),
    "utf8",
  );
  const result = await compileCaseBundle({
    sourcePath: "cases/da-mcp-auth-modes.yml",
    sourceText,
    compileStep: createSemanticStepCompiler(),
  });
  assert.equal(result.ok, true, JSON.stringify(result.diagnostics));
  const configText = await fs.readFile(
    path.join(__dirname, "..", "config.yaml"),
    "utf8",
  );
  const flags = [
    "TEAMSFX_MCP_FOR_DA_DT",
    "TEAMSFX_MCP_FOR_DA_DCR",
    "TEAMSFX_V4_ENABLED",
  ];
  for (const { caseId, plan } of result.value) {
    for (const flag of flags) {
      assert.ok(
        !plan.plan_metadata.tags.some((tag) =>
          tag.startsWith(`feature_flag:${flag}=`),
        ),
        caseId,
      );
    }
    const fixture = plan.steps.find((step) =>
      step.tags.includes("operation:mcp-auth-fixture"),
    );
    const mode = fixture.tags
      .find((tag) => tag.startsWith("mcp_auth_mode:"))
      .slice("mcp_auth_mode:".length);
    const config = parse(
      prepareMcpAuthRunConfig(
        configText,
        JSON.stringify(plan),
        mode,
        [],
        false,
      ),
    );
    for (const flag of flags) {
      assert.ok(
        !Object.hasOwn(config.docker.environment, flag),
        `${caseId}: ${flag}`,
      );
    }
  }
});

test("VCB-170: auth cases require visible invocation evidence for all three tools", async () => {
  const sourceText = await fs.readFile(
    path.join(__dirname, "..", "cases", "da-mcp-auth-modes.yml"),
    "utf8",
  );
  const result = await compileCaseBundle({
    sourcePath: "cases/da-mcp-auth-modes.yml",
    sourceText,
    compileStep: createSemanticStepCompiler(),
  });
  assert.equal(result.ok, true, JSON.stringify(result.diagnostics));
  assert.equal(result.value.length, 3);
  for (const item of result.value) {
    for (const [name, result] of [
      ["whoami", `"tenant_id": "${expectedTenant}"`],
      ["whoami", '"client_id": "${{env:MCP_AUTH_EXPECTED_JWT_CLIENT_ID}}"'],
      ["auth_flow_probe", `"canary": "${expectedCanary}"`],
      ["auth_flow_probe", '"auth_flow": "${{env:MCP_AUTH_EXPECTED_FLOW}}"'],
      ["add", '"sum": 27'],
    ]) {
      assert.ok(
        item.plan.steps.some(
          (step) =>
            step.tags.includes("assertion:tool-called") &&
            step.description.includes(name),
        ),
        `${item.caseId}: ${name}`,
      );
      assert.ok(
        item.plan.steps.some(
          (step) =>
            step.tags.includes("expectation:contains") &&
            step.description.includes(result),
        ),
        `${item.caseId}: ${name} must assert its named result field`,
      );
    }
  }
});

for (const invalid of ["../whoami", "42"]) {
  test(`VCB-170: rejects an invalid tool-invocation expectation ${invalid}`, async () => {
    const sourceText = await fs.readFile(
      path.join(__dirname, "..", "cases", "da-mcp-auth-modes.yml"),
      "utf8",
    );
    const result = await compileCaseBundle({
      sourcePath: "cases/da-mcp-auth-modes.yml",
      sourceText: sourceText.replace(
        "toolCalled: whoami",
        `toolCalled: ${invalid}`,
      ),
      compileStep: createSemanticStepCompiler(),
    });
    assert.equal(result.ok, false);
    assert.equal(result.diagnostics[0].code, "VCB_CHECK_ASSERTION_INVALID");
  });
}

test("VCB-169: mode configs remove ambient auth state and contain references rather than credentials", () => {
  const plan = JSON.stringify({
    plan_metadata: { tags: ["feature_flag:TEAMSFX_MCP_FOR_DA_DCR=true"] },
    steps: [{ tags: ["mcp_auth_mode:dcr"] }],
  });
  const config = prepareMcpAuthRunConfig(
    "docker:\n  environment:\n    ENTRA_SSO_REGISTRATION_ID: old-id\n    DA_MCP_OAUTH_CLIENT_SECRET: old-secret\n",
    plan,
    "dcr",
    ["MCP_AUTH_MODE", "MCP_AUTH_ENDPOINT"],
  );
  const environment = parse(config).docker.environment;
  assert.equal(environment.MCP_AUTH_MODE, "${MCP_AUTH_MODE}");
  assert.equal(environment.MCP_AUTH_ENDPOINT, "${MCP_AUTH_ENDPOINT}");
  assert.equal(environment.TEAMSFX_MCP_FOR_DA_DCR, "true");
  assert.equal(environment.TEMPLATE_VERSION, "local");
  assert.ok(!Object.hasOwn(environment, "ENTRA_SSO_REGISTRATION_ID"));
  assert.ok(!Object.hasOwn(environment, "DA_MCP_OAUTH_CLIENT_SECRET"));
  assert.throws(() => prepareMcpAuthRunConfig(config, plan, "entra-sso", []));
  assert.throws(() =>
    prepareMcpAuthRunConfig(config, plan, "dcr", ["MCP_AUTH_SSO_ID"]),
  );
});

test("MCP-AUTH-04: optional SSO-ID config resolves an absent Windows environment variable to empty", () => {
  const plan = JSON.stringify({
    plan_metadata: { tags: [] },
    steps: [{ tags: ["mcp_auth_mode:entra-sso"] }],
  });
  const config = prepareMcpAuthRunConfig(
    "docker:\n  environment: {}\n",
    plan,
    "entra-sso",
    ["MCP_AUTH_SSO_ID", "MCP_AUTH_SSO_PROTECTED_ID_B64"],
    false,
  );
  const environment = parse(config).docker.environment;
  assert.equal(environment.MCP_AUTH_SSO_ID, "");
  assert.equal(
    environment.MCP_AUTH_SSO_PROTECTED_ID_B64,
    "${MCP_AUTH_SSO_PROTECTED_ID_B64}",
  );
});

test("VCB-169: report redaction removes secrets, JWTs and callback authorization codes", () => {
  const text =
    "password=test-password callback?code=real-code&state=ok bearer eyJabcdefghijk.abcdefgh.abcdefgh";
  const result = redact(text, ["test-password"]);
  assert.ok(!result.includes("test-password"));
  assert.ok(!result.includes("real-code"));
  assert.ok(!result.includes("eyJabcdefghijk"));
  assert.ok(result.includes("state=ok"));
});
