/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import * as fs from "fs";
import {
  StubSpecParserAdapter,
  createSpecParserAdapter,
} from "../../../src/templates/openApi/specParserAdapter";
import { makeOpenApiScaffoldFn } from "../../../src/templates/openApi/scaffoldFn";
import { openApiTemplateDescriptors } from "../../../src/templates/descriptors/openApi";
import { QuestionNames } from "../../../src/questions/questionNames";
import { createMockContext } from "../testHelper";

describe("SpecParserAdapter", () => {
  describe("StubSpecParserAdapter", () => {
    const adapter = new StubSpecParserAdapter();

    it("should validate any spec as valid", async () => {
      const result = await adapter.validate("/path/to/spec.yaml");
      expect(result.valid).to.be.true;
      expect(result.errors).to.be.empty;
    });

    it("should return stub operations", async () => {
      const ops = await adapter.listOperations("/path/to/spec.yaml");
      expect(ops).to.have.length(2);
      expect(ops[0].id).to.equal("GET /api/items");
      expect(ops[1].id).to.equal("POST /api/items");
    });

    it("should return empty files for generate", async () => {
      const result = await adapter.generate(
        "/path/to/spec.yaml",
        ["GET /api/items"],
        "/tmp/out",
        "Copilot"
      );
      expect(result.files.size).to.equal(0);
      expect(result.warnings).to.have.length(1);
      expect(result.warnings[0]).to.include("Stub adapter");
    });
  });

  describe("createSpecParserAdapter", () => {
    it("should return a RealSpecParserAdapter", () => {
      const adapter = createSpecParserAdapter();
      expect(adapter).to.have.property("validate").that.is.a("function");
      expect(adapter).to.have.property("listOperations").that.is.a("function");
      expect(adapter).to.have.property("generate").that.is.a("function");
    });
  });
});

describe("makeOpenApiScaffoldFn", () => {
  let ctx: ReturnType<typeof createMockContext>;

  beforeEach(() => {
    ctx = createMockContext();
  });

  afterEach(() => {
    sinon.restore();
  });

  it("should return error when apiSpecPath is missing", async () => {
    const scaffoldFn = makeOpenApiScaffoldFn("Copilot", "test-template");
    const result = await scaffoldFn(ctx, {
      language: "typescript",
      projectName: "test",
      destinationPath: "/tmp/test",
    });
    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().code).to.equal("MissingApiSpecPath");
  });

  it("should return error when apiOperations is empty", async () => {
    const scaffoldFn = makeOpenApiScaffoldFn("Copilot", "test-template");
    const result = await scaffoldFn(ctx, {
      language: "typescript",
      projectName: "test",
      destinationPath: "/tmp/test",
      apiSpecPath: "/tmp/spec.yaml",
      apiOperations: [],
    });
    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().code).to.equal("MissingApiOperations");
  });

  it("should return error when spec validation fails", async () => {
    const failingAdapter = new StubSpecParserAdapter();
    sinon.stub(failingAdapter, "validate").resolves({
      valid: false,
      errors: ["Invalid schema at line 10"],
      warnings: [],
    });
    // Stub scaffoldTemplates to succeed
    const scaffoldTemplatesMod = await import("../../../src/templates/scaffold/scaffolder");
    sinon.stub(scaffoldTemplatesMod, "scaffoldTemplates").resolves({
      isOk: () => true,
      isErr: () => false,
      value: ["file1.ts"],
      _unsafeUnwrap: () => ["file1.ts"],
      map: (fn: any) => ({
        isOk: () => true,
        isErr: () => false,
        value: fn(["file1.ts"]),
        _unsafeUnwrap: () => fn(["file1.ts"]),
      }),
    } as any);

    const scaffoldFn = makeOpenApiScaffoldFn("Copilot", "test-template", failingAdapter);
    const result = await scaffoldFn(ctx, {
      language: "typescript",
      projectName: "test",
      destinationPath: "/tmp/test",
      apiSpecPath: "/tmp/spec.yaml",
      apiOperations: ["GET /items"],
    });
    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().code).to.equal("InvalidApiSpec");
  });

  it("should accept a custom adapter", async () => {
    const customAdapter = new StubSpecParserAdapter();
    const generateStub = sinon.stub(customAdapter, "generate").resolves({
      files: new Map([["generated.ts", "// generated"]]),
      warnings: [],
    });

    // Stub scaffoldTemplates and fs
    const scaffoldTemplatesMod = await import("../../../src/templates/scaffold/scaffolder");
    const { ok } = await import("neverthrow");
    sinon.stub(scaffoldTemplatesMod, "scaffoldTemplates").resolves(ok(["file1.ts"]));
    sinon.stub(fs.promises, "mkdir").resolves(undefined);
    sinon.stub(fs.promises, "writeFile").resolves();

    const scaffoldFn = makeOpenApiScaffoldFn("TeamsAi", "test-template", customAdapter);
    const result = await scaffoldFn(ctx, {
      language: "typescript",
      projectName: "test",
      destinationPath: "/tmp/test",
      apiSpecPath: "/tmp/spec.yaml",
      apiOperations: ["GET /items"],
    });
    expect(result.isOk()).to.be.true;
    expect(generateStub.calledOnce).to.be.true;
    expect(generateStub.firstCall.args[3]).to.equal("TeamsAi");
  });

  it("should include DeclarativeCopilot in replaceMap for Copilot project type", async () => {
    const customAdapter = new StubSpecParserAdapter();
    const scaffoldTemplatesMod = await import("../../../src/templates/scaffold/scaffolder");
    const { ok } = await import("neverthrow");
    const scaffoldStub = sinon
      .stub(scaffoldTemplatesMod, "scaffoldTemplates")
      .resolves(ok(["file1.ts"]));
    sinon.stub(fs.promises, "mkdir").resolves(undefined);
    sinon.stub(fs.promises, "writeFile").resolves();

    const scaffoldFn = makeOpenApiScaffoldFn("Copilot", "test-template", customAdapter);
    await scaffoldFn(ctx, {
      language: "typescript",
      projectName: "test",
      destinationPath: "/tmp/test",
      apiSpecPath: "/tmp/spec.yaml",
      apiOperations: ["GET /items"],
    });

    const tplInfos = scaffoldStub.firstCall.args[1];
    expect(tplInfos[0].replaceMap).to.have.property("DeclarativeCopilot", "true");
  });
});

describe("OpenAPI Template Descriptors", () => {
  it("should have 3 descriptors", () => {
    expect(openApiTemplateDescriptors).to.have.length(3);
  });

  it("should have correct IDs", () => {
    const ids = openApiTemplateDescriptors.map((d) => d.id);
    expect(ids).to.include("da/api-plugin-from-spec");
    expect(ids).to.include("ai-agent/rag-from-spec");
    expect(ids).to.include("me/from-spec");
  });

  it("should have IDs matching category/variant convention", () => {
    for (const d of openApiTemplateDescriptors) {
      expect(d.id).to.match(/^[a-z-]+\/[a-z0-9-]+$/);
    }
  });

  it("should have 'openapi' tag on all descriptors", () => {
    for (const d of openApiTemplateDescriptors) {
      expect(d.tags).to.include("openapi");
    }
  });

  it("da/api-plugin-from-spec should have spec questions", () => {
    const da = openApiTemplateDescriptors.find((d) => d.id === "da/api-plugin-from-spec")!;
    expect(da.category).to.equal("declarative-agent");
    const qNames = da.questions!.map((q) => q.question.name);
    expect(qNames).to.include(QuestionNames.apiSpecPath);
    expect(qNames).to.include(QuestionNames.apiOperations);
    expect(qNames).to.not.include(QuestionNames.llmProvider);
  });

  it("ai-agent/rag-from-spec should have spec + LLM questions", () => {
    const ai = openApiTemplateDescriptors.find((d) => d.id === "ai-agent/rag-from-spec")!;
    expect(ai.category).to.equal("ai-agent");
    const qNames = ai.questions!.map((q) => q.question.name);
    expect(qNames).to.include(QuestionNames.apiSpecPath);
    expect(qNames).to.include(QuestionNames.apiOperations);
    expect(qNames).to.include(QuestionNames.llmProvider);
    expect(qNames).to.include(QuestionNames.azureOpenAiKey);
    expect(qNames).to.include(QuestionNames.openAiKey);
  });

  it("me/from-spec should have spec questions only", () => {
    const me = openApiTemplateDescriptors.find((d) => d.id === "me/from-spec")!;
    expect(me.category).to.equal("message-extension");
    const qNames = me.questions!.map((q) => q.question.name);
    expect(qNames).to.include(QuestionNames.apiSpecPath);
    expect(qNames).to.include(QuestionNames.apiOperations);
    expect(qNames).to.not.include(QuestionNames.llmProvider);
  });
});
