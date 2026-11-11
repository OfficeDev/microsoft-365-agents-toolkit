// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import os from "os";
import { mkdtemp, readFile, remove, writeJson } from "fs-extra";
import AdmZip from "adm-zip";
import { parse } from "yaml";
import { UserError } from "@microsoft/teamsfx-api";
import { assert } from "vitest";
import { collectInputs } from "../../../src/v4/collectInputs/collectInputs";
import { openCreateQuestions } from "../../../src/v4/distribution/createQuestions";
import { evaluateExpression } from "../../../src/v4/expression/evaluateExpression";
import { createDefaultCreateOptionsProviders } from "../../../src/v4/providers/createOptionsProviders";
import { createDefaultCreateInputValidators } from "../../../src/v4/validators/createInputValidators";
import { createExpressionPort } from "../../../src/v4/runtime/whitelist";
import { REQUIRE_EMPTY_TARGET } from "../../../src/v4/pipeline/runScaffoldPipeline";
import { createInMemoryRuntime } from "../../../src/v4/runtime/inMemoryRuntime";
import { scaffold } from "../../../src/v4/runtime/scaffold";
import {
  assertContainsInOrder,
  isRecord,
  loadV4Package,
  recordArrayProperty,
  readJsonObject,
  recordProperty,
  runV4Package,
  text,
} from "./helpers/scenarioHarness";

/**
 * T3 scenario tier: the `custom-copilot-rag-custom-api` create package scaffolded under
 * `InMemoryRuntime`.
 *
 * Spec: docs/03-specs/scenarios/teams/create-custom-copilot-rag-custom-api.md
 * (SCN-CREATE-RAG-CUSTOM-API-01..06)
 */

const SPEC_PATH = path.resolve(__dirname, "fixtures/repairs-openapi.yaml");
const templatePackage = loadV4Package("create", "custom-copilot-rag-custom-api");
const appName = "My API Agent";

async function run(language: "typescript" | "javascript" | "python" = "typescript") {
  return runV4Package(templatePackage, {
    answers: { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /repairs"] },
    callerFloor: { appName, language },
  });
}

describe("SCN-TEAMS-CREATE-CUSTOM-COPILOT-RAG-CUSTOM-API (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-RAG-CUSTOM-API-07: shipped questions reject HTTP before rendering and generate collected HTTPS operations", async () => {
    const zip = new AdmZip();
    zip.addFile(
      "v4/create/custom-copilot-rag-custom-api/questions.json",
      Buffer.from(JSON.stringify(templatePackage.questions))
    );
    zip.addLocalFolder(
      path.resolve(templatePackage.packageDir, "../../_shared/questions"),
      "v4/_shared/questions"
    );
    const questions = openCreateQuestions(zip.toBuffer(), {
      kind: "create",
      templateId: "custom-copilot-rag-custom-api",
    })._unsafeUnwrap();
    if (!isRecord(templatePackage.descriptor)) {
      assert.fail("expected a descriptor object");
    }
    const properties = recordProperty(
      recordProperty(templatePackage.descriptor, "optionsSchema"),
      "properties"
    );
    const directory = await mkdtemp(path.join(os.tmpdir(), "custom-api-policy-"));
    try {
      for (const protocol of ["https", "http"]) {
        const document: unknown = parse(await readFile(SPEC_PATH, "utf8"));
        if (!isRecord(document)) {
          assert.fail("expected an OpenAPI document object");
        }
        document.servers = [{ url: `${protocol}://api.example.com` }];
        const source = path.join(directory, `${protocol}.json`);
        await writeJson(source, document);
        const providers = createDefaultCreateOptionsProviders(
          async () => ({ requiresAuth: false, tools: [] }),
          async () => []
        );
        const validators = createDefaultCreateInputValidators();
        const expressionPort = createExpressionPort(() => false);
        const runtime = createInMemoryRuntime();
        const collected = await collectInputs(
          questions,
          { properties },
          {
            apiSpecLocation: source,
            apiOperations: ["GET /repairs"],
            llmService: "llm-service-openai",
            openAIKey: "",
          },
          {
            ui: {
              ask: async () => assert.fail("all scalar answers are prefilled"),
              askMulti: async () => assert.fail("selected operations are prefilled"),
            },
            optionsProvider: (id) => providers[id],
            validator: (id) => validators[id],
            evaluate: (node, scope) => evaluateExpression(node, scope, expressionPort),
          }
        );
        if (protocol === "http") {
          assert.isTrue(collected.isErr(), "HTTP must fail at collection, not after rendering");
          assert.instanceOf(collected._unsafeUnwrapErr(), UserError);
          assert.equal(collected._unsafeUnwrapErr().name, "OpenApiSpecInvalid");
          assert.equal(runtime.files.size, 0);
          continue;
        }
        const answers = collected._unsafeUnwrap();
        assert.equal(answers["derived.openapi.teamsAiOperations.apiSpecLocation"], source);
        assert.notProperty(answers, "derived.openapi.operations.apiSpecLocation");
        const generated = await scaffold(
          {
            descriptor: templatePackage.descriptor,
            pipeline: templatePackage.pipeline,
            content: templatePackage.content,
            answers,
            callerFloor: { appName, language: "typescript" },
            targetDir: { path: "/out", existing: [] },
          },
          runtime
        );
        assert.isTrue(generated.isOk(), generated.isErr() ? generated.error.message : "");
        assert.include(
          text(runtime.files, "appPackage/apiSpecificationFile/openapi.json"),
          "https://api.example.com"
        );
        assert.property(readJsonObject(runtime.files, "src/app/functions.json"), "listRepairs");
        assert.include(text(runtime.files, "src/app/handlers.ts"), "listRepairsHandler");
        assert.isTrue(runtime.files.has("src/adaptiveCards/listRepairs.json"));
      }
    } finally {
      await remove(directory);
    }
  });

  it("SCN-CREATE-RAG-CUSTOM-API-01: TypeScript scaffold writes and generates the custom API file set", async () => {
    const { files, outcome } = await run("typescript");
    assert.include(outcome.written, "package.json");
    assert.include(outcome.written, "src/index.ts");
    assert.include(outcome.written, "src/config.ts");
    assert.include(outcome.written, "src/app/app.ts");
    assert.include(outcome.written, "src/app/handlers.ts");
    assert.include(outcome.written, "appPackage/manifest.json");
    assert.include(outcome.written, "infra/azure.bicep");
    assert.include(outcome.written, "m365agents.yml");

    assert.isTrue(files.has("appPackage/apiSpecificationFile/openapi.yaml"));
    assert.isTrue(files.has("src/app/functions.json"));
    assert.isTrue(files.has("src/app/instructions.txt"));
    assert.isTrue(files.has("src/adaptiveCards/listRepairs.json"));
    assert.include(text(files, "src/adaptiveCards/listRepairs.json"), '"$data": "${results}"');
    assert.include(
      text(files, "src/adaptiveCards/listRepairs.json"),
      "title: ${if(title, title, 'N/A')}"
    );
    assert.include(text(files, "src/adaptiveCards/listRepairs.data.json"), '"title"');
    assert.include(text(files, "src/app/app.ts"), "functionDefs.listRepairs.name");
    assert.include(text(files, "src/app/handlers.ts"), "listRepairsHandler");
    assert.include(text(files, "src/app/handlers.ts"), "openapi.yaml");
    const functions = readJsonObject(files, "src/app/functions.json");
    assert.isTrue("listRepairs" in functions);
  });

  it("SCN-CREATE-RAG-CUSTOM-API-02: package and manifest render appName-derived values", async () => {
    const { files } = await run("typescript");
    const pkg = readJsonObject(files, "package.json");
    assert.strictEqual(pkg.name, "myapiagent");

    const manifest = readJsonObject(files, "appPackage/manifest.json");
    const name = recordProperty(manifest, "name");
    assert.strictEqual(name.short, "My API Agent${{APP_NAME_SUFFIX}}");
  });

  it("SCN-CREATE-RAG-CUSTOM-API-03: post-render OpenAPI step runs after require-empty-target", async () => {
    const { outcome } = await run("typescript");
    assert.deepStrictEqual(outcome.stepsRun, [
      "require-empty-target",
      "openapi/generate-teams-ai-custom-api-files",
    ]);
  });

  it("SCN-CREATE-RAG-CUSTOM-API-04: JavaScript scaffold selects and updates the JavaScript subtree", async () => {
    const { files, outcome } = await run("javascript");
    assert.include(outcome.written, "src/index.js");
    assert.include(outcome.written, "src/config.js");
    assert.include(outcome.written, "src/app/app.js");
    assert.include(outcome.written, "src/app/handlers.js");
    assert.notInclude(outcome.written, "src/index.ts");
    assert.include(text(files, "src/app/app.js"), "functionDefs.listRepairs.name");
    assert.include(text(files, "src/app/handlers.js"), "listRepairsHandler");
  });

  it("SCN-CREATE-RAG-CUSTOM-API-05: Python scaffold selects and updates the Python subtree", async () => {
    const { files, outcome } = await run("python");
    const app = text(files, "src/app.py");
    assert.include(outcome.written, "src/app.py");
    assert.include(outcome.written, "src/handlers.py");
    assert.isTrue(files.has("src/functions.json"));
    assert.isTrue(files.has("src/adaptiveCards/listRepairs.json"));
    assert.notInclude(outcome.written, "package.json");
    assert.include(app, 'function_defs["listRepairs"]["name"]');
    assert.include(app, "from openai import OpenAIError");
    assert.include(app, 'message = error_body.get("message")');
    assertContainsInOrder(app, [
      "def get_openai_error_message(error: OpenAIError) -> str:",
      "try:",
      "chat_result = await prompt.send(",
      "except OpenAIError as e:",
      'print(f"Error sending chat prompt: {get_openai_error_message(e)}")',
      'await ctx.send(MessageActivityInput(text="An error occurred while processing your request."))',
      "return",
    ]);
    assert.include(text(files, "src/handlers.py"), "client.listRepairs");
    assert.include(text(files, "src/handlers.py"), "openapi.yaml");

    const manifest = readJsonObject(files, "appPackage/manifest.json");
    const bots = recordArrayProperty(manifest, "bots");
    assert.deepStrictEqual(bots[0].commandLists, [
      {
        scopes: ["personal"],
        commands: [{ title: "List repairs", description: "List repairs" }],
      },
    ]);
  });

  it("SCN-CREATE-RAG-CUSTOM-API-06: a non-empty target fails require-empty-target first", async () => {
    const runtime = createInMemoryRuntime();
    const result = await scaffold(
      {
        descriptor: templatePackage.descriptor,
        pipeline: templatePackage.pipeline,
        content: templatePackage.content,
        answers: { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /repairs"] },
        callerFloor: { appName, language: "typescript" },
        targetDir: { path: "/out", existing: ["README.md"] },
      },
      runtime
    );
    assert.isTrue(result.isErr());
    const error = result._unsafeUnwrapErr();
    assert.instanceOf(error, UserError);
    assert.strictEqual(error.name, REQUIRE_EMPTY_TARGET);
    assert.strictEqual(runtime.files.size, 0);
  });
});
