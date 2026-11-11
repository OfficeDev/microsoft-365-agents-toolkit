// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ensureDir, mkdtemp, readJson, readdir, remove, writeJson } from "fs-extra";
import os from "os";
import * as path from "path";
import { SpecParser } from "@microsoft/m365-spec-parser";
import { SystemError, UserError } from "@microsoft/teamsfx-api";
import { ok } from "neverthrow";
import { featureFlagManager } from "../../../../src/common/featureFlags";
import { StepContext } from "../../../../src/v4/pipeline/runScaffoldPipeline";
import { STEP_REGISTRY } from "../../../../src/v4/runtime/runtimeRegistry";
import {
  STEP_GENERATE_OPENAPI_PLUGIN_FILES,
  STEP_GENERATE_TEAMS_AI_CUSTOM_API_FILES,
  openApiGeneratePluginFiles,
  openApiGenerateTeamsAiCustomApiFiles,
} from "../../../../src/v4/runtime/steps/openApi";
import { generateTeamsAiCustomApiFiles } from "../../../../src/v4/runtime/steps/openApiCustomApi";
import { ProgrammingLanguage } from "../../../../src/question/constants";
import { assert, beforeEach, expect, vi } from "vitest";
import { isMap, isSeq, parseDocument } from "yaml";

interface MockParserOperation {
  api: string;
  isValid: boolean;
  operationId?: string;
  summary?: string;
  description?: string;
  server?: string;
  auth?: {
    name?: string;
    authScheme?: Record<string, unknown>;
  };
}

interface MockSpecParserState {
  filteredSpec: unknown;
  listOperations: MockParserOperation[];
  pluginManifest: unknown;
  pluginAdaptiveCards: Record<string, string>;
  pluginConversationStarters: Record<string, unknown>[];
  validationStatus: string;
}

const mockSpecParserState = vi.hoisted<MockSpecParserState>(() => ({
  filteredSpec: undefined,
  listOperations: [],
  pluginManifest: undefined,
  pluginAdaptiveCards: {},
  pluginConversationStarters: [{ text: "Find pets" }],
  validationStatus: "Ok",
}));

const mockGenerateAdaptiveCard = vi.hoisted(() =>
  vi.fn(() => [
    {
      type: "AdaptiveCard",
      $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.5",
      body: [
        {
          type: "Container",
          $data: "${pets}",
          items: [
            {
              type: "TextBlock",
              text: "name: ${if(name, name, 'N/A')}",
              wrap: true,
            },
          ],
        },
      ],
    },
    "pets",
    { pets: [{ name: "Milo" }] },
    [],
  ])
);
const mockAxiosGet = vi.hoisted(() => vi.fn());

const SPEC_PATH = path.resolve(__dirname, "../../scenarios/fixtures/repairs-openapi.yaml");

vi.mock("axios", () => ({
  default: { get: mockAxiosGet },
}));

vi.mock("@microsoft/m365-spec-parser", () => {
  class SpecParserError extends Error {
    constructor(
      message: string,
      public readonly errorType: string
    ) {
      super(message);
    }
  }

  class SpecParser {
    async list(): Promise<{ APIs: MockParserOperation[] }> {
      return { APIs: mockSpecParserState.listOperations };
    }

    async validate(): Promise<{ status: string }> {
      return { status: mockSpecParserState.validationStatus };
    }

    async generate(): Promise<void> {
      return undefined;
    }

    async generateAdaptiveCardInPlugin(): Promise<void> {
      return undefined;
    }

    async getFilteredSpecs(): Promise<unknown[]> {
      return [undefined, mockSpecParserState.filteredSpec];
    }

    async generateForCopilot(
      _manifestPath: string,
      _apiOperations: string[],
      apiSpecPath: string,
      pluginPath: string
    ): Promise<{ allSuccess: boolean; warnings: unknown[] }> {
      const fs = await import("fs-extra");
      const nodePath = await import("path");
      await fs.writeFile(apiSpecPath, "openapi: 3.0.0\n");
      const pluginManifest = mockSpecParserState.pluginManifest ?? {
        functions: [{ name: "getPets", description: "Get pets" }],
      };
      if (typeof pluginManifest === "string") {
        await fs.writeFile(pluginPath, pluginManifest);
      } else {
        await fs.writeJson(pluginPath, pluginManifest);
      }
      for (const [relativePath, content] of Object.entries(
        mockSpecParserState.pluginAdaptiveCards
      )) {
        const filePath = nodePath.join(nodePath.dirname(pluginPath), relativePath);
        await fs.ensureDir(nodePath.dirname(filePath));
        await fs.writeFile(filePath, content);
      }
      return { allSuccess: true, warnings: [] };
    }
  }

  return {
    AdaptiveCardGenerator: { generateAdaptiveCard: mockGenerateAdaptiveCard },
    AdaptiveCardUpdateStrategy: { CreateNew: "CreateNew", KeepExisting: "KeepExisting" },
    ConstantString: {
      AllOperationMethods: ["get", "post", "put", "patch", "delete", "head", "options"],
      RegistrationIdPostfix: "REGISTRATION_ID",
    },
    ErrorType: { NoSupportedApi: "NoSupportedApi", SpecNotValid: "SpecNotValid" },
    ProjectType: { Copilot: "copilot", TeamsAi: "teams-ai" },
    SpecParser,
    SpecParserError,
    Utils: {
      format(template: string, ...args: string[]): string {
        return `${template} ${args.join(",")}`;
      },
      getAuthArray(security: unknown): unknown[] {
        return Array.isArray(security) ? security : [];
      },
      getSafeRegistrationIdEnvName(value: string): string {
        return value.replace(/[^a-zA-Z0-9]/g, "_").toUpperCase();
      },
      checkServerUrl(): unknown[] {
        return [];
      },
      isAPIKeyAuthButNotInCookie(authScheme: Record<string, unknown> | undefined): boolean {
        return authScheme?.type === "apiKey" && authScheme.in !== "cookie";
      },
      isBearerTokenAuth(authScheme: Record<string, unknown> | undefined): boolean {
        return authScheme?.type === "http" && authScheme.scheme === "bearer";
      },
      isOAuthWithAuthCodeFlow(authScheme: Record<string, unknown> | undefined): boolean {
        return authScheme?.type === "oauth2";
      },
    },
    ValidationStatus: { Error: "Error", Valid: "Valid" },
    WarningType: {
      GenerateCardFailed: "GenerateCardFailed",
      GenerateJsonDataFailed: "GenerateJsonDataFailed",
      OperationIdContainsSpecialCharacters: "OperationIdContainsSpecialCharacters",
      OperationIdMissing: "OperationIdMissing",
      UnsupportedAuthType: "UnsupportedAuthType",
    },
  };
});

function makeCtx(initial: Record<string, string> = {}): {
  ctx: StepContext;
  files: Map<string, Buffer>;
} {
  const files = new Map<string, Buffer>();
  for (const [filePath, body] of Object.entries(initial)) {
    files.set(filePath, Buffer.from(body, "utf8"));
  }
  return {
    files,
    ctx: {
      read: (filePath) => files.get(filePath),
      write: (filePath, data) => {
        files.set(filePath, data);
      },
      writeEnvironment: () => Promise.resolve(ok(undefined)),
    },
  };
}

function text(files: Map<string, Buffer>, filePath: string): string {
  return files.get(filePath)?.toString("utf8") ?? "";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === "object" && !Array.isArray(value);
}

function isRecordArray(value: unknown): value is Record<string, unknown>[] {
  return Array.isArray(value) && value.every(isRecord);
}

function readJsonObject(files: Map<string, Buffer>, filePath: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(text(files, filePath));
  if (!isRecord(parsed)) {
    assert.fail(`${filePath} should contain a JSON object`);
  }
  return parsed;
}

function teamsAiSpec(): Record<string, unknown> {
  return {
    info: { description: "Pet store APIs" },
    paths: {
      "/pets/{petId}": {
        get: {
          operationId: "getPets",
          summary: "Get pets by id",
          parameters: [
            {
              name: "petId",
              in: "path",
              required: true,
              description: "Pet id",
              schema: { type: "string" },
            },
            { name: "includeDetails", in: "query", schema: { type: "boolean" } },
          ],
          requestBody: {
            required: true,
            description: "Filter body",
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  required: ["kind"],
                  properties: { kind: { type: "string", description: "Pet kind" } },
                },
              },
            },
          },
          security: [{ apiKey: [] }],
          responses: {
            "200": {
              description: "Pets response",
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    properties: {
                      pets: {
                        type: "array",
                        items: {
                          type: "object",
                          properties: {
                            name: { type: "string" },
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        },
      },
    },
  };
}

beforeEach(() => {
  mockSpecParserState.filteredSpec = teamsAiSpec();
  mockSpecParserState.listOperations = [
    {
      api: "GET /pets",
      isValid: true,
      operationId: "getPets",
      summary: "Find pets",
      server: "https://api.example.com",
      auth: { name: "petKey", authScheme: { type: "apiKey", in: "header" } },
    },
  ];
  mockSpecParserState.pluginManifest = undefined;
  mockSpecParserState.pluginAdaptiveCards = {};
  mockSpecParserState.pluginConversationStarters = [{ text: "Find pets" }];
  mockSpecParserState.validationStatus = "Ok";
  mockGenerateAdaptiveCard.mockClear();
  mockAxiosGet.mockReset();
  // The Kiota branch spawns an external binary; pin it off unless a test opts in.
  vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(false);
});

describe("OpenAPI runtime steps (v4)", () => {
  it.each([
    { marker: "  # Renamed packaging comment\n", anchor: true },
    { marker: "", anchor: true },
    { marker: "", anchor: false },
  ])("AC-31: inserts auth in provision independent of comments: %j", async ({ marker, anchor }) => {
    const yaml =
      "# Project settings\nversion: v1.12\nprovision:\n  - uses: teamsApp/create\n" +
      marker +
      (anchor ? "  - uses: teamsApp/zipAppPackage\n" : "") +
      "deploy:\n  # Keep deploy here\n  - uses: cli/runNpmCommand\n    with:\n      args: build\n      env:\n        ID: 9007199254740993\n";
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
      "m365agents.yml": yaml,
      "m365agents.local.yml": yaml,
    });
    const result = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"] },
      ctx
    );
    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    for (const file of ["m365agents.yml", "m365agents.local.yml"]) {
      const output = text(files, file);
      const document = parseDocument(output, { intAsBigInt: true });
      assert.isEmpty(document.errors);
      const provision = document.get("provision", true);
      assert.isTrue(isSeq(provision));
      if (!isSeq(provision)) assert.fail("Expected provision sequence");
      const actions = provision.items.map((item) => (isMap(item) ? item.get("uses") : undefined));
      assert.deepEqual(actions, [
        "teamsApp/create",
        "apiKey/register",
        ...(anchor ? ["teamsApp/zipAppPackage"] : []),
      ]);
      assert.deepEqual(document.toJS().deploy, [
        {
          uses: "cli/runNpmCommand",
          with: { args: "build", env: { ID: BigInt("9007199254740993") } },
        },
      ]);
      assert.include(output, "ID: 9007199254740993");
      assert.include(output, "# Project settings");
      assert.include(output, "# Keep deploy here");
      if (marker) assert.include(output, marker.trim());
    }
  });

  it("AC-31: treats auth names containing YAML syntax as scalar values", async () => {
    const authName = 'pet: key # "quoted"\nother: value';
    const operation = mockSpecParserState.listOperations[0];
    operation.auth = { name: authName, authScheme: { type: "oauth2" } };
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
      "m365agents.yml": "provision: []\n",
    });
    const result = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"] },
      ctx
    );
    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    const document = parseDocument(text(files, "m365agents.yml"));
    assert.isEmpty(document.errors);
    assert.strictEqual(document.getIn(["provision", 0, "with", "name"]), authName);
    assert.strictEqual(document.getIn(["provision", 0, "with", "appId"]), "${{TEAMS_APP_ID}}");
    assert.strictEqual(document.getIn(["provision", 0, "with", "flow"]), "authorizationCode");
  });

  it.each(["provision: [", "deploy: []\n", "provision: {}\n", "provision: null\n"])(
    "AC-31: rejects malformed auth YAML without rewriting it: %s",
    async (yaml) => {
      const { ctx, files } = makeCtx({
        "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
        "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
        "m365agents.yml": yaml,
      });
      const result = await openApiGeneratePluginFiles.apply(
        { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"] },
        ctx
      );
      assert.isTrue(result.isErr());
      assert.strictEqual(result._unsafeUnwrapErr().name, "OpenApiAuthYamlInvalid");
      assert.strictEqual(text(files, "m365agents.yml"), yaml);
    }
  );

  it("AC-31: leaves YAML bytes unchanged when no auth registration is needed", async () => {
    mockSpecParserState.listOperations[0].auth = undefined;
    const yaml = "# no auth\nprovision: []\n\ndeploy: []\n";
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
      "m365agents.yml": yaml,
    });
    const result = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"] },
      ctx
    );
    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    assert.strictEqual(text(files, "m365agents.yml"), yaml);
  });

  it.each([ProgrammingLanguage.TS, ProgrammingLanguage.JS, ProgrammingLanguage.PY])(
    "AC-30: preserves OpenAPI generated bytes for %s",
    async (language) => {
      const extension = language === ProgrammingLanguage.JS ? "js" : "ts";
      const appPath =
        language === ProgrammingLanguage.PY ? "src/app.py" : `src/app/app.${extension}`;
      const handlersPath =
        language === ProgrammingLanguage.PY ? "src/handlers.py" : `src/app/handlers.${extension}`;
      const { ctx, files } = makeCtx({
        "appPackage/manifest.json": JSON.stringify({ bots: [{ commandLists: [] }] }),
        [appPath]:
          "// Replace with function definition code\n//Replace with functions to be imported\n",
        [handlersPath]: "// Replace with function handler code\n{{OPENAPI_SPEC_PATH}}",
      });
      const result = await openApiGenerateTeamsAiCustomApiFiles.apply(
        { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"], language },
        ctx
      );
      assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
      expect(
        Object.fromEntries(
          [...files]
            .sort(([left], [right]) => left.localeCompare(right))
            .map(([name, bytes]) => [name, bytes.toString("utf8")])
        )
      ).toMatchSnapshot();
    }
  );

  it("AC-30: ships OpenAPI source fragments as imported data", async () => {
    const asset = await import("../../../../src/v4/runtime/steps/assets/openApiCustomApi.json");
    assert.hasAllKeys(asset.default.functionDefinitionCode, ["typescript", "javascript", "python"]);
    assert.hasAllKeys(asset.default.functionHandlerCode, ["typescript", "javascript", "python"]);
    assert.isArray(asset.default.instructions);
  });

  it("registers both OpenAPI steps", () => {
    assert.strictEqual(
      STEP_REGISTRY.get(STEP_GENERATE_OPENAPI_PLUGIN_FILES),
      openApiGeneratePluginFiles
    );
    assert.strictEqual(
      STEP_REGISTRY.get(STEP_GENERATE_TEAMS_AI_CUSTOM_API_FILES),
      openApiGenerateTeamsAiCustomApiFiles
    );
  });

  it("validateParams reports missing or unsupported OpenAPI step parameters", () => {
    assert.strictEqual(
      openApiGeneratePluginFiles.validateParams({ apiOperations: ["GET /pets"] }),
      "missing string parameter 'apiSpecLocation'"
    );
    assert.strictEqual(
      openApiGeneratePluginFiles.validateParams({ apiSpecLocation: "openapi.yml" }),
      "missing string[] parameter 'apiOperations'"
    );
    assert.strictEqual(
      openApiGenerateTeamsAiCustomApiFiles.validateParams({
        apiSpecLocation: "openapi.yml",
        apiOperations: ["GET /pets"],
        language: "csharp",
      }),
      "missing supported language parameter 'language'"
    );
  });

  it("SCN-CREATE-RAG-CUSTOM-API-06: emits recoverable adaptive-card warnings", async () => {
    mockGenerateAdaptiveCard.mockImplementationOnce(() => {
      throw new Error("unsupported response shape");
    });
    const warnings: string[] = [];
    const state = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ bots: [{ commandLists: [] }] }),
      "src/app/app.ts": "// Replace with function definition code\n",
      "src/app/handlers.ts": "// Replace with function handler code\n{{OPENAPI_SPEC_PATH}}",
    });
    state.ctx.warn = (warning) => warnings.push(warning.content);

    const result = await openApiGenerateTeamsAiCustomApiFiles.apply(
      {
        apiSpecLocation: SPEC_PATH,
        apiOperations: ["GET /pets"],
        language: ProgrammingLanguage.TS,
      },
      state.ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    assert.lengthOf(warnings, 1);
    assert.include(warnings[0], "getPets");
    assert.include(warnings[0], "unsupported response shape");
  });

  it("SCN-CREATE-RAG-CUSTOM-API-06: emits recoverable mock-data warnings", async () => {
    mockGenerateAdaptiveCard.mockImplementationOnce(() => [
      { type: "AdaptiveCard", body: [] },
      "$",
      {},
      [{ content: "example data is incomplete" }],
    ]);
    const warnings: string[] = [];
    const state = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ bots: [{ commandLists: [] }] }),
      "src/app/app.ts": "// Replace with function definition code\n",
      "src/app/handlers.ts": "// Replace with function handler code\n{{OPENAPI_SPEC_PATH}}",
    });
    state.ctx.warn = (warning) => warnings.push(warning.content);

    const result = await openApiGenerateTeamsAiCustomApiFiles.apply(
      {
        apiSpecLocation: SPEC_PATH,
        apiOperations: ["GET /pets"],
        language: ProgrammingLanguage.TS,
      },
      state.ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    assert.lengthOf(warnings, 1);
    assert.include(warnings[0], "example data is incomplete");
  });

  it("SCN-CREATE-RAG-CUSTOM-API-01: filters malformed parameters and preserves array schemas", async () => {
    mockSpecParserState.filteredSpec = {
      info: {},
      paths: {
        "/ignored": undefined,
        "/pets": {
          get: {
            operationId: "list-items",
            parameters: [
              null,
              {
                name: "tags",
                in: "query",
                schema: { type: "array", items: { type: "string" } },
              },
              { name: "untyped", in: "query" },
              { name: "ignored", in: "toString", schema: { type: "string" } },
            ],
            responses: {},
          },
        },
      },
    };
    const state = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ bots: [{ commandLists: [] }] }),
      "src/app/app.ts": "// Replace with function definition code\n",
      "src/app/handlers.ts": "// Replace with function handler code\n{{OPENAPI_SPEC_PATH}}",
    });

    const result = await openApiGenerateTeamsAiCustomApiFiles.apply(
      {
        apiSpecLocation: SPEC_PATH,
        apiOperations: ["GET /pets"],
        language: ProgrammingLanguage.TS,
      },
      state.ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    const functions = text(state.files, "src/app/functions.json");
    assert.include(functions, '"items": {');
    assert.include(functions, '"untyped": {');
    assert.notInclude(functions, '"ignored": {');
  });

  it("SCN-CREATE-RAG-CUSTOM-API-01: rejects a selected operation without an operation id", async () => {
    mockSpecParserState.filteredSpec = {
      info: {},
      paths: { "/pets": { get: { responses: {} } } },
    };
    const state = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ bots: [{ commandLists: [] }] }),
      "src/app/app.ts": "// Replace with function definition code\n",
      "src/app/handlers.ts": "// Replace with function handler code\n{{OPENAPI_SPEC_PATH}}",
    });

    const result = await openApiGenerateTeamsAiCustomApiFiles.apply(
      {
        apiSpecLocation: SPEC_PATH,
        apiOperations: ["GET /pets"],
        language: ProgrammingLanguage.TS,
      },
      state.ctx
    );

    assert.isTrue(result.isErr());
    assert.include(result._unsafeUnwrapErr().message, "no operationId");
  });

  it("does not generate language-specific files for an unsupported language", async () => {
    const tempRoot = await mkdtemp(path.join(os.tmpdir(), "m365atk-openapi-custom-api-"));
    try {
      await ensureDir(path.join(tempRoot, "appPackage"));
      await writeJson(path.join(tempRoot, "appPackage", "manifest.json"), {
        bots: [{ commandLists: [] }],
      });

      const warnings = await generateTeamsAiCustomApiFiles(
        { info: { title: "API", version: "1.0.0" }, paths: {} },
        "csharp",
        tempRoot,
        "openapi.yaml"
      );

      assert.isEmpty(warnings);
      assert.deepStrictEqual(
        await readJson(path.join(tempRoot, "src", "app", "functions.json")),
        {}
      );
      assert.deepStrictEqual(await readdir(path.join(tempRoot, "src", "app")), ["functions.json"]);
      assert.deepStrictEqual(await readdir(path.join(tempRoot, "src")), ["app"]);
    } finally {
      await remove(tempRoot);
    }
  });

  it("returns SystemError results for invalid resolved params and missing render output", async () => {
    const invalidParams = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: "openapi.yml", apiOperations: "GET /pets" },
      makeCtx().ctx
    );
    assert.isTrue(invalidParams.isErr());
    assert.instanceOf(invalidParams._unsafeUnwrapErr(), SystemError);
    assert.strictEqual(invalidParams._unsafeUnwrapErr().name, "OpenApiGenerateParams");

    const missingBaseFile = await openApiGenerateTeamsAiCustomApiFiles.apply(
      {
        apiSpecLocation: "openapi.yml",
        apiOperations: ["GET /pets"],
        language: ProgrammingLanguage.TS,
      },
      makeCtx().ctx
    );
    assert.isTrue(missingBaseFile.isErr());
    assert.strictEqual(missingBaseFile._unsafeUnwrapErr().name, "OpenApiGeneratedBaseFileMissing");
  });

  it("generates plugin files, registers the action, and injects API key registration yaml", async () => {
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
      "m365agents.yml":
        "provision:\n  # Build app package with latest env value\n  - uses: teamsApp/zipAppPackage\n",
      "m365agents.local.yml": "provision: []\n",
    });

    const result = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"] },
      ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    const agent = readJsonObject(files, "appPackage/declarativeAgent.json");
    if (!isRecordArray(agent.actions) || !isRecordArray(agent.conversation_starters)) {
      assert.fail("declarative agent should contain action and conversation starter arrays");
    }
    assert.deepInclude(agent.actions, { id: "action_1", file: "ai-plugin.json" });
    assert.deepInclude(agent.conversation_starters, { text: "Find pets" });
    assert.include(
      text(files, "appPackage/apiSpecificationFile/openapi.yaml.original"),
      "title: Repairs API"
    );
    assert.include(text(files, "m365agents.yml"), "uses: apiKey/register");
    assert.include(text(files, "m365agents.yml"), "registrationId: PETKEY_REGISTRATION_ID");
    assert.include(text(files, "m365agents.local.yml"), "uses: apiKey/register");
  });

  it("does not mislabel a non-spec-parser generator crash as a spec-parser failure", async () => {
    const generate = vi
      .spyOn(SpecParser.prototype, "generateForCopilot")
      .mockRejectedValue(new Error("disk full"));
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
    });

    try {
      const result = await openApiGeneratePluginFiles.apply(
        { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"] },
        ctx
      );

      assert.isTrue(result.isErr());
      const error = result._unsafeUnwrapErr();
      assert.instanceOf(error, SystemError);
      assert.strictEqual(error.name, "OpenApiGenerateFailed");
      assert.include(error.message, "disk full");
      assert.isFalse(files.has("appPackage/ai-plugin.json"));
    } finally {
      generate.mockRestore();
    }
  });

  it("preserves existing conversation starters and uses descriptions up to the six-item limit", async () => {
    mockSpecParserState.listOperations = [
      {
        api: "GET /pets",
        isValid: true,
        summary: "   ",
        description: "Find pets by description",
      },
      {
        api: "POST /pets",
        isValid: true,
        summary: "Create a pet",
      },
    ];
    const existingStarters = Array.from({ length: 5 }, (_, index) => ({
      text: `Existing starter ${index + 1}`,
    }));
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({
        name: "Agent",
        conversation_starters: existingStarters,
      }),
    });

    const result = await openApiGeneratePluginFiles.apply(
      {
        apiSpecLocation: SPEC_PATH,
        apiOperations: ["GET /pets", "POST /pets"],
      },
      ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    const agent = readJsonObject(files, "appPackage/declarativeAgent.json");
    if (!isRecordArray(agent.conversation_starters)) {
      assert.fail("declarative agent should contain a conversation starter array");
    }
    assert.deepStrictEqual(agent.conversation_starters, [
      ...existingStarters,
      { text: "Find pets by description" },
    ]);
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-09: removes starter duplicates before applying the limit", async () => {
    mockSpecParserState.listOperations = [
      { api: "GET /pets", isValid: true, summary: "Existing starter 1" },
      { api: "POST /pets", isValid: true, summary: "Create a pet" },
    ];
    const existingStarters = Array.from({ length: 5 }, (_, index) => ({
      text: `Existing starter ${index + 1}`,
    }));
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({
        name: "Agent",
        conversation_starters: existingStarters,
      }),
    });

    const result = await openApiGeneratePluginFiles.apply(
      {
        apiSpecLocation: SPEC_PATH,
        apiOperations: ["GET /pets", "POST /pets"],
      },
      ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    const agent = readJsonObject(files, "appPackage/declarativeAgent.json");
    if (!isRecordArray(agent.conversation_starters)) {
      assert.fail("declarative agent should contain a conversation starter array");
    }
    assert.deepStrictEqual(agent.conversation_starters, [
      ...existingStarters,
      { text: "Create a pet" },
    ]);
  });

  it("preserves adaptive card files generated for OpenAPI plugin response semantics", async () => {
    mockSpecParserState.pluginManifest = {
      functions: [
        {
          name: "getPets",
          description: "Get pets",
          capabilities: {
            response_semantics: {
              static_template: { file: "adaptiveCards/getPets.json" },
            },
          },
        },
      ],
      capabilities: { conversation_starters: [{ text: "Find pets" }] },
    };
    mockSpecParserState.pluginAdaptiveCards = {
      "adaptiveCards/getPets.json": JSON.stringify({ type: "AdaptiveCard", version: "1.0" }),
    };
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
    });

    const result = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"] },
      ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    assert.strictEqual(
      text(files, "appPackage/adaptiveCards/getPets.json"),
      '{"type":"AdaptiveCard","version":"1.0"}'
    );
  });

  it("preserves the original OpenAPI description from a URL for plugin regeneration", async () => {
    mockAxiosGet.mockResolvedValue({
      data: Buffer.from("openapi: 3.0.1\ninfo:\n  title: Remote API\n"),
    });
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
    });

    const result = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: "https://example.com/openapi.yaml", apiOperations: ["GET /pets"] },
      ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    expect(mockAxiosGet).toHaveBeenCalledWith("https://example.com/openapi.yaml", {
      responseType: "arraybuffer",
    });
    assert.include(
      text(files, "appPackage/apiSpecificationFile/openapi.yaml.original"),
      "title: Remote API"
    );
  });

  it("surfaces the manifest read error when the generated declarative agent is invalid JSON", async () => {
    const malformedAgent = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: "openapi.yml", apiOperations: ["GET /pets"] },
      makeCtx({
        "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
        "appPackage/declarativeAgent.json": "{",
      }).ctx
    );
    assert.isTrue(malformedAgent.isErr());
    // Wrapper-owned action registration preserves the shared manifest read error name.
    assert.strictEqual(malformedAgent._unsafeUnwrapErr().name, "JSONSyntaxError");
  });

  it("injects OAuth registration yaml for OAuth-protected OpenAPI operations", async () => {
    mockSpecParserState.listOperations = [
      {
        api: "GET /pets",
        isValid: true,
        server: "https://api.example.com",
        auth: { name: "petOAuth", authScheme: { type: "oauth2" } },
      },
    ];
    const { ctx, files } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
      "m365agents.yml": "provision: []\n",
    });

    const result = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets"] },
      ctx
    );

    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    assert.include(text(files, "m365agents.yml"), "uses: oauth/register");
    assert.include(text(files, "m365agents.yml"), "configurationId: PETOAUTH_REGISTRATION_ID");
  });

  it("returns UserError when selected authenticated OpenAPI operations span servers", async () => {
    mockSpecParserState.listOperations = [
      {
        api: "GET /pets",
        isValid: true,
        server: "https://api.one.example.com",
        auth: { name: "petKey", authScheme: { type: "apiKey", in: "header" } },
      },
      {
        api: "POST /pets",
        isValid: true,
        server: "https://api.two.example.com",
        auth: { name: "otherKey", authScheme: { type: "http", scheme: "bearer" } },
      },
    ];
    const { ctx } = makeCtx({
      "appPackage/manifest.json": JSON.stringify({ name: "manifest" }),
      "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
    });

    const result = await openApiGeneratePluginFiles.apply(
      { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /pets", "POST /pets"] },
      ctx
    );

    assert.isTrue(result.isErr());
    assert.instanceOf(result._unsafeUnwrapErr(), UserError);
    assert.strictEqual(result._unsafeUnwrapErr().name, "OpenApiMultipleAuthServers");
  });

  it("returns Teams AI validation and filtered-spec errors", async () => {
    mockSpecParserState.validationStatus = "Error";
    const validation = await openApiGenerateTeamsAiCustomApiFiles.apply(
      {
        apiSpecLocation: "openapi.yaml",
        apiOperations: ["GET /pets"],
        language: ProgrammingLanguage.JS,
      },
      makeCtx({
        "appPackage/manifest.json": JSON.stringify({ bots: [{}] }),
        "src/app/app.js": "// Replace with function definition code\n",
        "src/app/handlers.js": "{{OPENAPI_SPEC_PATH}}\n// Replace with function handler code\n",
      }).ctx
    );
    assert.isTrue(validation.isErr());
    assert.strictEqual(validation._unsafeUnwrapErr().name, "OpenApiSpecInvalid");

    mockSpecParserState.validationStatus = "Ok";
    mockSpecParserState.filteredSpec = undefined;
    const missingSpec = await openApiGenerateTeamsAiCustomApiFiles.apply(
      {
        apiSpecLocation: "openapi.yaml",
        apiOperations: ["GET /pets"],
        language: ProgrammingLanguage.JS,
      },
      makeCtx({
        "appPackage/manifest.json": JSON.stringify({ bots: [{}] }),
        "src/app/app.js": "// Replace with function definition code\n",
        "src/app/handlers.js": "{{OPENAPI_SPEC_PATH}}\n// Replace with function handler code\n",
      }).ctx
    );
    assert.isTrue(missingSpec.isErr());
    assert.strictEqual(missingSpec._unsafeUnwrapErr().name, "OpenApiTeamsAiFilteredSpecMissing");
  });
});
