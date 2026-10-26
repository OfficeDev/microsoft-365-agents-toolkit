// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const assert = require("node:assert/strict");
const fs = require("fs-extra");
const os = require("node:os");
const path = require("node:path");
const sinon = require("sinon");
const { Cleaner } = require("../../src/commonlib/cleaner.ts");
const { AadManager } = require("../../src/commonlib/aadManager.ts");
const { AzureHelper } = require("../../src/commonlib/azureHelper.ts");
const { TeamsAppHelper } = require("../../src/commonlib/teamsAppHelper.ts");
const { M365TitleHelper } = require("../../src/commonlib/m365TitleHelper.ts");
const { ProjectEnvReader } = require("../../src/commonlib/projectEnvReader.ts");
const {
  tryMcpAuthTeardown,
} = require("../../src/commonlib/mcpAuthTeardown.ts");
const { GraphApiCleanHelper } = require("../../src/utils/cleanHelper.ts");

describe("MCP authentication resource ownership", function () {
  let sandbox;
  let project;
  let previousRunId;
  let deleted;
  let environment;
  let ownership;

  beforeEach(async function () {
    sandbox = sinon.createSandbox();
    project = await fs.mkdtemp(path.join(os.tmpdir(), "mcp-owned-teardown-"));
    previousRunId = process.env.MCP_AUTH_RUN_ID;
    process.env.MCP_AUTH_RUN_ID = "current-test-run";
    deleted = { aad: [], groups: [], teams: [], titles: [], registrations: [] };
    environment = {
      AAD_APP_CLIENT_ID: "7c5898c1-99fb-4d83-8474-02bf7985096c",
      BOT_ID: "shared-monitoring-app",
      SHARED_SP_ID: "shared-service-principal",
      MONITORING_APP_ID: "shared-monitoring-app",
      MCP_AUTH_ID: "shared-registration",
      TEAMS_APP_ID: "new-teams-app",
      M365_TITLE_ID: "new-title",
    };
    ownership = {
      version: 1,
      projectName: path.basename(project),
      runId: "current-test-run",
      mode: "entra-sso",
      sharedRegistrationId: "shared-registration",
      resources: [
        { kind: "aad-app", envKey: "AAD_APP_CLIENT_ID", existedBefore: false },
        {
          kind: "service-principal",
          envKey: "SHARED_SP_ID",
          existedBefore: false,
        },
        {
          kind: "monitoring-app",
          envKey: "MONITORING_APP_ID",
          existedBefore: false,
        },
        {
          kind: "oauth-registration",
          envKey: "MCP_AUTH_ID",
          existedBefore: true,
        },
        { kind: "teams-app", envKey: "TEAMS_APP_ID", existedBefore: false },
        { kind: "m365-title", envKey: "M365_TITLE_ID", existedBefore: false },
      ],
    };
    await fs.outputFile(
      path.join(project, "m365agents.yml"),
      "version: v1.12\n",
    );
    await fs.outputFile(
      path.join(project, "env", ".env.dev"),
      Object.entries(environment)
        .map(([name, value]) => `${name}=${value}`)
        .join("\n"),
    );
    await fs.outputJSON(
      path.join(project, ".auth-test", "ownership.json"),
      ownership,
    );
    sandbox.stub(ProjectEnvReader, "readAllEnvFiles").resolves([environment]);
    const directDeleteAad = async (id) => {
      deleted.aad.push(id);
      return true;
    };
    sandbox.stub(AadManager, "init").resolves({
      deleteAadAppById: directDeleteAad,
      deleteAadAppsByClientId: async (id) => {
        if (id) await directDeleteAad(id);
        return [true];
      },
    });
    sandbox.stub(AzureHelper, "init").returns({
      deleteResourceGroup: async (id) => {
        deleted.groups.push(id);
        return true;
      },
    });
    sandbox.stub(TeamsAppHelper, "init").resolves({
      deleteTeamsAppById: async (id) => {
        deleted.teams.push(id);
        return true;
      },
      deleteBotById: async () => true,
      cancelStagedTeamsAppById: async () => true,
    });
    sandbox.stub(M365TitleHelper, "init").resolves({
      unacquire: async (id) => {
        deleted.titles.push(id);
        return true;
      },
    });
  });

  afterEach(async function () {
    sandbox.restore();
    if (previousRunId === undefined) delete process.env.MCP_AUTH_RUN_ID;
    else process.env.MCP_AUTH_RUN_ID = previousRunId;
    await fs.remove(project);
  });

  async function clean() {
    const result = await Cleaner.clean(project);
    if (Array.isArray(result)) await Promise.all(result);
    return result;
  }

  async function saveFixture() {
    await fs.outputJSON(
      path.join(project, ".auth-test", "ownership.json"),
      ownership,
    );
    await fs.outputFile(
      path.join(project, "env", ".env.dev"),
      Object.entries(environment)
        .map(([name, value]) => `${name}=${value}`)
        .join("\n"),
    );
  }

  function permissiveOperations() {
    return {
      deleteTeamsApp: async (id) => {
        deleted.teams.push(id);
      },
      unacquireM365Title: async (id) => {
        deleted.titles.push(id);
      },
      deleteRegistration: async (kind, id) => {
        deleted.registrations.push({ kind, id });
      },
    };
  }

  it("MCP-AUTH-11: protects shared AAD, SP, monitoring and supplied SSO even with direct-ID cleanup", async function () {
    await clean();
    assert.deepEqual(deleted.aad, []);
    assert.deepEqual(deleted.groups, []);
    assert.deepEqual(deleted.registrations, []);
    assert.deepEqual(deleted.teams, ["new-teams-app"]);
    assert.deepEqual(deleted.titles, ["new-title"]);
    assert.equal(AadManager.init.called, false);
    assert.equal(AzureHelper.init.called, false);
  });

  it("MCP-AUTH-11: skips resources belonging to another run", async function () {
    ownership.runId = "another-test-run";
    await fs.outputJSON(
      path.join(project, ".auth-test", "ownership.json"),
      ownership,
    );
    await clean();
    assert.deepEqual(deleted, {
      aad: [],
      groups: [],
      teams: [],
      titles: [],
      registrations: [],
    });
  });

  it("MCP-AUTH-11: missing ownership in an MCP run fails without generic cleanup", async function () {
    await fs.remove(path.join(project, ".auth-test", "ownership.json"));
    await assert.rejects(clean(), /ownership/i);
    assert.equal(AadManager.init.called, false);
    assert.equal(AzureHelper.init.called, false);
  });

  it("MCP-AUTH-11: malformed ownership fails without falling back to shared-resource deletion", async function () {
    await fs.outputFile(
      path.join(project, ".auth-test", "ownership.json"),
      "{broken",
    );
    await assert.rejects(clean());
    assert.equal(AadManager.init.called, false);
    assert.equal(AzureHelper.init.called, false);
  });

  it("MCP-AUTH-11: direct Graph deletion helpers skip shared AAD, SP and monitoring directory objects", async function () {
    const helper = new GraphApiCleanHelper("unit-test-token");
    const request = sandbox.stub(helper, "execute").resolves({});
    await helper.deleteAad("shared-aad-object-id");
    await helper.deleteEnterpriseApplication("shared-sp-object-id");
    await helper.deleteDeletedItem("shared-monitor-object-id");
    assert.equal(request.called, false);
  });

  it("MCP-AUTH-11: direct AAD-manager deletion cannot bypass the MCP-run identity guard", async function () {
    const manager = new AadManager("unit-test-token");
    const remove = sandbox.stub(manager.axios, "delete").resolves({});
    const lookup = sandbox
      .stub(manager.axios, "get")
      .resolves({ data: { value: [] } });
    await manager.deleteAadAppById("shared-aad-object-id", 0);
    await manager.deleteAadAppsByClientId(environment.AAD_APP_CLIENT_ID);
    assert.equal(remove.called, false);
    assert.equal(lookup.called, false);
  });

  it("MCP-AUTH-11: non-MCP direct cleanup keeps its existing behavior", async function () {
    delete process.env.MCP_AUTH_RUN_ID;
    const helper = new GraphApiCleanHelper("unit-test-token");
    const request = sandbox.stub(helper, "execute").resolves({});
    await helper.deleteAad("ordinary-test-aad");
    await helper.deleteEnterpriseApplication("ordinary-test-sp");
    await helper.deleteDeletedItem("ordinary-deleted-object");
    assert.equal(request.callCount, 5);
  });

  for (const [mode, variant, expectedRegistrationDeletes] of [
    ["static-oauth", "fresh", 1],
    ["dcr", "fresh", 1],
    ["entra-sso", "absent", 1],
    ["entra-sso", "present", 0],
    ["entra-sso", "invalid", 0],
  ]) {
    it(`MCP-AUTH-11: ${mode}/${variant} delegates only this run's registration to a direct-ID helper`, async function () {
      ownership.mode = mode;
      const registration = ownership.resources.find(
        (item) => item.kind === "oauth-registration",
      );
      registration.kind =
        mode === "dcr" ? "dcr-registration" : "oauth-registration";
      registration.existedBefore =
        variant === "present" || variant === "invalid";
      environment.MCP_AUTH_ID = registration.existedBefore
        ? `${variant}-supplied-registration`
        : `new-${mode}-registration`;
      delete ownership.sharedRegistrationId;
      if (registration.existedBefore) {
        ownership.sharedRegistrationId = environment.MCP_AUTH_ID;
      }
      await saveFixture();
      await tryMcpAuthTeardown(project, permissiveOperations());
      assert.equal(deleted.registrations.length, expectedRegistrationDeletes);
      if (expectedRegistrationDeletes) {
        assert.equal(deleted.registrations[0].id, environment.MCP_AUTH_ID);
      }
      assert.deepEqual(deleted.aad, []);
    });
  }

  it("MCP-AUTH-11: a protected SSO ID cannot become owned just by changing its baseline flag", async function () {
    ownership.resources.find(
      (item) => item.kind === "oauth-registration",
    ).existedBefore = false;
    await saveFixture();
    await tryMcpAuthTeardown(project, permissiveOperations());
    assert.deepEqual(deleted.registrations, []);
  });

  it("MCP-AUTH-11: missing ownership stays fail-closed without an ambient run ID", async function () {
    delete process.env.MCP_AUTH_RUN_ID;
    await fs.remove(path.join(project, ".auth-test", "ownership.json"));
    await fs.outputJSON(path.join(project, ".vscode", "mcp.json"), {
      servers: {
        fixture: { type: "http", url: "https://mcp.example.com/mcp" },
      },
    });
    await fs.outputFile(
      path.join(project, "m365agents.yml"),
      "provision:\n  - uses: oauth/register\n    with:\n      baseUrl: https://mcp.example.com/mcp\n",
    );
    await assert.rejects(clean(), /ownership/i);
    assert.equal(AadManager.init.called, false);
  });

  it("MCP-AUTH-11: a developer-only stdio MCP config does not change unrelated cleanup", async function () {
    delete process.env.MCP_AUTH_RUN_ID;
    await fs.remove(path.join(project, ".auth-test", "ownership.json"));
    await fs.outputJSON(path.join(project, ".vscode", "mcp.json"), {
      servers: { toolkit: { type: "stdio", command: "atk" } },
    });
    await fs.outputFile(
      path.join(project, "m365agents.yml"),
      "provision:\n  - uses: oauth/register\n",
    );
    await clean();
    assert.equal(AadManager.init.called, true);
    assert.ok(deleted.aad.length > 0);
  });

  it("MCP-AUTH-11: a title is not owned when its parent Teams app was pre-existing", async function () {
    ownership.resources.find(
      (item) => item.kind === "teams-app",
    ).existedBefore = true;
    await saveFixture();
    await clean();
    assert.deepEqual(deleted.teams, []);
    assert.deepEqual(deleted.titles, []);
  });

  it("MCP-AUTH-11: unsupported registration cleanup is reported as retained, not deleted", async function () {
    delete ownership.sharedRegistrationId;
    ownership.resources.find(
      (item) => item.kind === "oauth-registration",
    ).existedBefore = false;
    environment.MCP_AUTH_ID = "new-sso-registration";
    await saveFixture();
    const result = await clean();
    assert.equal(result.retained.length, 1);
    assert.equal(result.retained[0].id, "new-sso-registration");
    assert.deepEqual(deleted.registrations, []);
  });

  it("MCP-AUTH-11: failed owned-resource deletion is surfaced without generic cleanup fallback", async function () {
    TeamsAppHelper.init.resolves({ deleteTeamsAppById: async () => false });
    await assert.rejects(clean(), /Failed to delete/);
    assert.equal(AadManager.init.called, false);
  });
});
