// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform, ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { ensureDir, mkdtemp, readFile, remove, writeFile } from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mcpAuthScaffoldDeps } from "../../../src/v4/mcp/mcpAuthScaffold";
import { setTools, TOOLS } from "../../../src/common/globalVars";
import { RetryHandler } from "../../../src/common/retryHandler";
import { DriverDefinition } from "../../../src/component/configManager/interface";
import { Lifecycle } from "../../../src/component/configManager/lifecycle";
import { YamlParser } from "../../../src/component/configManager/parser";
import "../../../src/component/driver/index";
import { DriverContext } from "../../../src/component/driver/interface/commonArgs";
import { envUtil } from "../../../src/component/utils/envUtil";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../plugins/solution/util";
import {
  isRecord,
  loadV4Package,
  readJsonObject,
  recordArrayProperty,
  recordProperty,
  runV4Package,
  text,
} from "./helpers/scenarioHarness";

// Configuration data only: all service traffic is intercepted by the HTTP fake.
const MCP_URL =
  // eslint-disable-next-line no-secrets/no-secrets -- Public fixture URL, not a credential.
  "https://ca-mcp-qetyc6ob.jollyfield-18c57c9a.eastus2.azurecontainerapps.io/mcp";
const MCP_ORIGIN = new URL(MCP_URL).origin;
const CONFIGURATION_ID = "test-dcr-configuration";
const RESOURCE_PATH = `/v1.0/oAuthConfigurations/${CONFIGURATION_ID}`;
const createPackage = loadV4Package("create", "da/mcp-server");
const addPackage = loadV4Package("modify", "add-mcp-server");

type Project = Awaited<ReturnType<typeof runV4Package>>;

async function createProject(): Promise<Project> {
  return runV4Package(createPackage, {
    answers: {
      mcpServerType: "remote",
      "derived.mcp.serverTypes.catalog": "{}",
      mcpServerUrl: MCP_URL,
      authType: "oauth-dynamic",
    },
    callerFloor: { appName: "DCR lifecycle test", language: "common" },
  });
}

function requestObject(data: unknown): Record<string, unknown> {
  if (typeof data !== "string") throw new Error("Expected a JSON request body");
  const parsed: unknown = JSON.parse(data);
  if (!isRecord(parsed)) throw new Error("Expected a JSON object");
  return parsed;
}

describe("DCR MCP auth lifecycle - L1 composition, not UI or live service coverage", () => {
  let http: MockAdapter;
  let context: DriverContext;
  let ui: MockedUserInteraction;
  let stored: Record<string, unknown> | undefined;
  let rejectUpdate: boolean;
  let directory: string | undefined;
  let previousTools = TOOLS;
  const retries = RetryHandler.RETRIES;

  beforeEach(() => {
    vi.stubEnv("TEAMSFX_MCP_FOR_DA_DT", "true");
    vi.stubEnv("TEAMSFX_MCP_FOR_DA_DCR", "true");
    vi.stubEnv("TEAMSFX_V4_ENABLED", "true");
    vi.stubEnv("TEAMSFX_SOVEREIGN_CLOUD_ENVIRONMENT", undefined);
    vi.stubEnv("TEAMSFX_ENV", "dev");
    vi.stubEnv("TEAMS_APP_ID", "test-m365-app");
    vi.spyOn(mcpAuthScaffoldDeps, "probeMCPServerAuth").mockRejectedValue(
      new Error("Dynamic scaffolding must not discover OAuth metadata")
    );
    vi.spyOn(mcpAuthScaffoldDeps, "resolveMCPOAuthMetadata").mockRejectedValue(
      new Error("Dynamic scaffolding must leave discovery to TGS")
    );
    const requester = axios.create();
    http = new MockAdapter(requester, { onNoMatch: "throwException" });
    vi.spyOn(axios, "create").mockReturnValue(requester);
    RetryHandler.RETRIES = 1;
    stored = undefined;
    rejectUpdate = false;
    directory = undefined;
    ui = new MockedUserInteraction();
    vi.spyOn(ui, "confirm").mockRejectedValue(
      new Error("DCR updates must not request confirmation")
    );
    context = {
      m365TokenProvider: new MockedM365Provider(),
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui,
      progressBar: undefined,
      logProvider: new MockedLogProvider(),
      projectPath: "",
      platform: Platform.VSCode,
    };
    previousTools = TOOLS;
    setTools({
      ui,
      logProvider: context.logProvider,
      tokenProvider: {
        azureAccountProvider: context.azureAccountProvider,
        m365TokenProvider: context.m365TokenProvider,
      },
    });
    vi.spyOn(context.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("test-token"));

    http.onPost("/v1.0/dynamicConfigurations").reply((request) => {
      if (stored) throw new Error("A second registration is not allowed in this fixture");
      const input = requestObject(request.data);
      stored = {
        ...input,
        oAuthConfigId: CONFIGURATION_ID,
        identityProvider: "Dynamic",
        supportedAccountTypes: input.supportedAccountTypes ?? "Enterprise",
        scopes: ["access_as_user"],
        tokenExchangeMethodType: "PostRequestBody",
        isPKCEEnabled: true,
        authorizationEndpoint: `${MCP_ORIGIN}/authorize`,
        tokenExchangeEndpoint: `${MCP_ORIGIN}/token`,
        clientId: "service-managed-client",
        clientSecret: "test-only-service-managed-secret",
      };
      return [200, { configurationRegistrationId: { oAuthConfigId: CONFIGURATION_ID } }];
    });
    http.onGet(RESOURCE_PATH).reply(() => [200, structuredClone(currentState())]);
    http.onPatch(RESOURCE_PATH).reply((request) => {
      const input = requestObject(request.data);
      if (rejectUpdate) return [400, { code: "ControlledUpdateRejection" }];
      stored = { ...currentState(), ...input };
      return [204];
    });
  });

  afterEach(async () => {
    http.restore();
    RetryHandler.RETRIES = retries;
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    setTools(previousTools);
    if (directory) await remove(directory);
  });

  function currentState(): Record<string, unknown> {
    if (!stored) throw new Error("Expected the fixture to have a registered configuration");
    return stored;
  }

  async function prepare(project: Project) {
    directory = await mkdtemp(path.join(os.tmpdir(), "atk-dcr-lifecycle-"));
    context.projectPath = directory;
    for (const [relativePath, body] of project.files) {
      const filePath = path.join(directory, ...relativePath.split("/"));
      await ensureDir(path.dirname(filePath));
      await writeFile(filePath, body);
    }
    const yamlPath = path.join(directory, "m365agents.yml");
    const parsed = await new YamlParser().parse(yamlPath, true);
    if (parsed.isErr() || !parsed.value.provision) {
      throw new Error("Expected a valid generated provision lifecycle");
    }
    const registration = parsed.value.provision.driverDefs.find(
      (action) => action.uses === "dcr/register"
    );
    const idKey = registration?.writeToEnvironmentFile?.configurationId;
    if (!registration || !idKey) throw new Error("Expected the DCR registration output mapping");
    vi.stubEnv(idKey, undefined);
    const envPath = path.join(directory, "env", ".env.dev");
    const manifestPath = path.join(directory, "appPackage", "ai-plugin.json");

    async function executeAuthSlice() {
      const current = await new YamlParser().parse(yamlPath, true);
      if (current.isErr() || !current.value.provision) {
        throw new Error("Expected valid lifecycle YAML");
      }
      const actions = current.value.provision.driverDefs.filter((action) =>
        ["dcr/register", "dcr/update"].includes(action.uses)
      );
      expect(actions.length).toBeGreaterThan(0);
      return new Lifecycle("provision", actions, current.value.version).execute(context);
    }

    async function writeUpdate(input: Record<string, unknown>) {
      const update: DriverDefinition = {
        uses: "dcr/update",
        with: { configurationId: `\${{${idKey}}}`, ...input },
      };
      await writeFile(
        yamlPath,
        JSON.stringify({ version: "v1.13", provision: [registration, update] })
      );
    }

    async function references() {
      return {
        environment: await readFile(envPath, "utf8"),
        manifest: await readFile(manifestPath, "utf8"),
      };
    }

    return { yamlPath, idKey, executeAuthSlice, writeUpdate, references };
  }

  it("SCN-DCR-MCP-01: real create output wires the fixture URL and configuration reference without credentials", async () => {
    const project = await createProject();
    const fixture = await prepare(project);
    const parsed = await new YamlParser().parse(fixture.yamlPath, true);
    expect(parsed.isOk()).toBe(true);
    if (parsed.isErr() || !parsed.value.provision) throw new Error("Expected provision YAML");
    const action = parsed.value.provision.driverDefs.find((item) => item.uses === "dcr/register");
    if (!action || !isRecord(action.with)) throw new Error("Expected DCR registration inputs");
    expect(action.with.mcpResourceUrl).toBe(MCP_URL);
    expect(action.with.targetUrlsShouldStartWith).toEqual([MCP_URL]);
    expect(action.writeToEnvironmentFile).toEqual({ configurationId: fixture.idKey });
    const plugin = readJsonObject(project.files, "appPackage/ai-plugin.json");
    const runtime = recordArrayProperty(plugin, "runtimes")[0];
    expect(runtime.type).toBe("RemoteMCPServer");
    expect(recordProperty(runtime, "spec")).toEqual({ url: MCP_URL });
    expect(recordProperty(runtime, "auth")).toEqual({
      type: "OAuthPluginVault",
      reference_id: `\${{${fixture.idKey}}}`,
    });
    expect(runtime.run_for_functions).toEqual(["*"]);
    expect(plugin.functions).toEqual([]);
    expect(project.files.has("appPackage/mcp-tools-1.json")).toBe(false);
    expect(project.secrets.size).toBe(0);
    expect(text(project.files, "m365agents.yml")).toContain(`mcpResourceUrl: ${MCP_URL}`);
    expect(text(project.files, "m365agents.yml")).not.toMatch(
      /clientId:|clientSecret:|wellKnownAuthorizationServer:|oauth\/register|PLEASE_FILL_IN/
    );
    expect(JSON.stringify(readJsonObject(project.files, ".vscode/mcp.json"))).toContain(MCP_URL);
    expect(mcpAuthScaffoldDeps.probeMCPServerAuth).not.toHaveBeenCalled();
    expect(mcpAuthScaffoldDeps.resolveMCPOAuthMetadata).not.toHaveBeenCalled();
    expect(http.history.post).toHaveLength(0);
  });

  it("SCN-DCR-MCP-02: real add/re-add converges on one plugin and registration while preserving existing content", async () => {
    const seedFiles = {
      "appPackage/manifest.json": JSON.stringify({
        copilotAgents: {
          declarativeAgents: [{ id: "agent", file: "declarativeAgent.json" }],
        },
      }),
      "appPackage/declarativeAgent.json": JSON.stringify({
        name: "Existing DA",
        actions: [{ id: "existing", file: "existing-plugin.json" }],
      }),
      "appPackage/existing-plugin.json": JSON.stringify({ namespace: "existing" }),
      "m365agents.yml": [
        "version: v1.13",
        "provision:",
        "  - uses: teamsApp/create",
        "    with:",
        "      name: Existing DA",
        "    writeToEnvironmentFile:",
        "      teamsAppId: TEAMS_APP_ID",
      ].join("\n"),
      "env/.env.dev": "TEAMSFX_ENV=dev\nKEEP_EXISTING=yes\n",
    };
    const options = {
      answers: {
        mcpServerUrl: MCP_URL,
        teamsManifestPath: "appPackage/manifest.json",
        authType: "oauth-dynamic",
      },
      callerFloor: { appName: "Existing DA", language: "common" },
    };
    const first = await runV4Package(addPackage, {
      ...options,
      existing: Object.keys(seedFiles),
      seedFiles,
    });
    const second = await runV4Package(addPackage, {
      ...options,
      existing: [...first.files.keys()],
      seedFiles: Object.fromEntries(first.files),
    });
    const actions = recordArrayProperty(
      readJsonObject(second.files, "appPackage/declarativeAgent.json"),
      "actions"
    );
    expect(actions).toHaveLength(2);
    expect(actions.filter((action) => action.file === "existing-plugin.json")).toHaveLength(1);
    expect(new Set(actions.map((action) => action.file)).size).toBe(2);
    const added = actions.find((action) => action.file !== "existing-plugin.json");
    if (!added || typeof added.file !== "string") throw new Error("Expected the new MCP action");
    const plugin = readJsonObject(second.files, `appPackage/${added.file}`);
    const runtime = recordArrayProperty(plugin, "runtimes")[0];
    expect(runtime.type).toBe("RemoteMCPServer");
    expect(recordProperty(runtime, "spec")).toEqual({ url: MCP_URL });
    expect(recordProperty(runtime, "auth").type).toBe("OAuthPluginVault");
    expect(text(second.files, "appPackage/existing-plugin.json")).toBe(
      seedFiles["appPackage/existing-plugin.json"]
    );
    const yaml = text(second.files, "m365agents.yml");
    expect(yaml.match(/uses: dcr\/register/g)).toHaveLength(1);
    expect(yaml.match(/configurationId:/g)).toHaveLength(1);
    expect(yaml).toContain(`mcpResourceUrl: ${MCP_URL}`);
    expect(yaml).not.toMatch(/clientId:|clientSecret:|oauth\/register/);
    expect(text(second.files, "env/.env.dev")).toContain("KEEP_EXISTING=yes");
    expect(second.secrets.size).toBe(0);
    expect(http.history.post).toHaveLength(0);
  });

  it("SCN-DCR-MCP-03: generated auth lifecycle registers once and preserves ID on re-provision", async () => {
    const fixture = await prepare(await createProject());
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    expect(http.history.post).toHaveLength(1);
    expect(requestObject(http.history.post[0].data)).toMatchObject({
      mcpResourceUrl: MCP_URL,
      targetUrlsShouldStartWith: [MCP_URL],
    });
    expect(process.env[fixture.idKey]).toBe(CONFIGURATION_ID);
    const before = await fixture.references();
    expect(before.environment).toContain(`${fixture.idKey}=${CONFIGURATION_ID}`);
    expect(before.manifest).toContain(`\${{${fixture.idKey}}}`);
    vi.stubEnv(fixture.idKey, undefined);
    const reloaded = await envUtil.readEnv(context.projectPath, "dev", false, false);
    if (reloaded.isErr()) throw reloaded.error;
    expect(reloaded.value[fixture.idKey]).toBe(CONFIGURATION_ID);
    vi.stubEnv(fixture.idKey, reloaded.value[fixture.idKey]);
    const second = await fixture.executeAuthSlice();
    expect(second.result.isOk()).toBe(true);
    if (second.result.isOk()) expect(second.result.value.size).toBe(0);
    expect(http.history.post).toHaveLength(1);
    expect(await fixture.references()).toEqual(before);
    expect(http.history.get).toHaveLength(0);
  });

  it("SCN-DCR-MCP-04: account-policy round trip changes only MSA and retains generated references", async () => {
    const fixture = await prepare(await createProject());
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    const before = structuredClone(currentState());
    const references = await fixture.references();
    for (const { enabled, accountTypes } of [
      { enabled: true, accountTypes: "Enterprise, Consumer" },
      { enabled: false, accountTypes: "Enterprise" },
    ]) {
      await fixture.writeUpdate({ includePersonalMicrosoftAccounts: enabled });
      expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
      expect(currentState()).toEqual({ ...before, supportedAccountTypes: accountTypes });
      expect(await fixture.references()).toEqual(references);
    }
    expect(http.history.patch.map((request) => requestObject(request.data))).toEqual([
      { supportedAccountTypes: "Enterprise, Consumer" },
      { supportedAccountTypes: "Enterprise" },
    ]);
    expect(http.history.post).toHaveLength(1);
    expect(ui.confirm).not.toHaveBeenCalled();
  });

  it("SCN-DCR-MCP-05: update rejection and retry preserve the configuration without replacement", async () => {
    const fixture = await prepare(await createProject());
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    const before = structuredClone(currentState());
    const references = await fixture.references();
    const notification = vi.spyOn(ui, "showMessage");
    await fixture.writeUpdate({ includePersonalMicrosoftAccounts: true });
    rejectUpdate = true;
    expect((await fixture.executeAuthSlice()).result.isErr()).toBe(true);
    expect(currentState()).toEqual(before);
    expect(notification).not.toHaveBeenCalled();
    expect(await fixture.references()).toEqual(references);
    rejectUpdate = false;
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    expect(currentState()).toEqual({ ...before, supportedAccountTypes: "Enterprise, Consumer" });
    expect(http.history.post).toHaveLength(1);
    expect(http.history.patch).toHaveLength(2);
    expect(await fixture.references()).toEqual(references);
  });

  it("SCN-DCR-MCP-06: repeated update is a no-op and omission preserves enabled account support", async () => {
    const fixture = await prepare(await createProject());
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    const references = await fixture.references();
    await fixture.writeUpdate({ includePersonalMicrosoftAccounts: true });
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    const unchanged = await fixture.executeAuthSlice();
    expect(unchanged.result.isOk()).toBe(true);
    expect(unchanged.summaries.flat().join()).toContain("already matches");
    expect(http.history.patch).toHaveLength(1);
    const beforePolicyChange = structuredClone(currentState());
    await fixture.writeUpdate({ targetAudience: "AnyTenant" });
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    expect(currentState()).toEqual({ ...beforePolicyChange, targetAudience: "AnyTenant" });
    expect(requestObject(http.history.patch[1].data)).toEqual({
      targetAudience: "AnyTenant",
    });
    expect(await fixture.references()).toEqual(references);
    expect(http.history.post).toHaveLength(1);
    expect(ui.confirm).not.toHaveBeenCalled();
  });

  it("SCN-DCR-MCP-07: invalid update YAML fails before auth and can be repaired without re-registration", async () => {
    const fixture = await prepare(await createProject());
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    const references = await fixture.references();
    vi.mocked(context.m365TokenProvider.getAccessToken).mockClear();
    http.resetHistory();
    for (const input of [
      { configurationId: "", includePersonalMicrosoftAccounts: true },
      { name: "Renamed configuration" },
      { mcpResourceUrl: MCP_URL },
      { wellKnownAuthorizationServer: "https://example.com/metadata" },
      { targetUrlsShouldStartWith: [MCP_URL] },
    ]) {
      await fixture.writeUpdate(input);
      const invalid = await new YamlParser().parse(fixture.yamlPath, true);
      expect(invalid.isErr()).toBe(true);
      expect(context.m365TokenProvider.getAccessToken).not.toHaveBeenCalled();
      expect(http.history.get).toHaveLength(0);
      expect(http.history.patch).toHaveLength(0);
      expect(await fixture.references()).toEqual(references);
    }
    await fixture.writeUpdate({ includePersonalMicrosoftAccounts: true });
    expect((await fixture.executeAuthSlice()).result.isOk()).toBe(true);
    expect(currentState().supportedAccountTypes).toBe("Enterprise, Consumer");
    expect(currentState().oAuthConfigId).toBe(CONFIGURATION_ID);
    expect(http.history.post).toHaveLength(0);
    expect(await fixture.references()).toEqual(references);
  });
});
