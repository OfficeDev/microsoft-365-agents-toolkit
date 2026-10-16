// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import * as fs from "node:fs";
import * as path from "node:path";
import * as vm from "node:vm";
import * as ts from "typescript";
import { quicktypeJSONSchema } from "../convert";

interface ProbeConverter {
  toProbe(json: string): unknown;
  probeToJson(value: unknown): string;
}

describe("Generated converter runtime", () => {
  let converter: ProbeConverter;

  before(async () => {
    const schema = {
      type: "object",
      additionalProperties: false,
      properties: {
        value: {
          anyOf: [
            { type: "string" },
            {
              type: "array",
              items: {
                type: "object",
                properties: { enabled: { type: "boolean" } },
                required: ["enabled"],
                additionalProperties: false,
              },
            },
          ],
        },
        nullable: { type: ["string", "null"] },
        mode: { type: "string", enum: ["first", "second"] },
        attributes: { type: "object", additionalProperties: { type: "number" } },
        timestamp: { type: "string", format: "date-time" },
        nested: {
          type: "object",
          properties: { label: { type: "string" } },
          required: ["label"],
          additionalProperties: false,
        },
      },
      required: ["value"],
    };
    const { lines } = await quicktypeJSONSchema("typescript", "Probe", JSON.stringify(schema));
    const { outputText, diagnostics } = ts.transpileModule(lines.join("\n"), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
      reportDiagnostics: true,
    });
    assert.isEmpty(diagnostics);
    converter = vm.runInNewContext(`${outputText}\nConvert;`, { exports: {}, Error, Date });
  });

  it("keeps every checked-in converter runtime in sync with the generation template", () => {
    const root = path.join(__dirname, "..");
    const template = fs.readFileSync(path.join(root, "converter-runtime.txt"), "utf8").trim();
    let converters = 0;

    function checkFolder(folder: string): void {
      for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
        const file = path.join(folder, entry.name);
        if (entry.isDirectory()) {
          checkFolder(file);
        } else if (entry.name.endsWith(".ts") && entry.name !== "index.ts") {
          const source = fs.readFileSync(file, "utf8");
          const start = source.indexOf("class ManifestConversionError");
          const end = source.indexOf("const typeMap:", start);
          assert.isAtLeast(start, 0, file);
          assert.isAbove(end, start, file);
          assert.strictEqual(source.slice(start, end).trim(), template, file);
          converters++;
        }
      }
    }

    checkFolder(path.join(root, "src", "generated-types"));
    assert.isAbove(converters, 0);
  });

  it("keeps the complete preview model and type map in sync with its schema snapshot", async () => {
    const root = path.join(__dirname, "..", "src");
    const schema = fs.readFileSync(
      path.join(root, "json-schemas", "teams", "vDevPreview", "MicrosoftTeams.schema.json"),
      "utf8"
    );
    const { lines } = await quicktypeJSONSchema("typescript", "TeamsManifestVDevPreview", schema);
    const source = fs.readFileSync(
      path.join(root, "generated-types", "teams", "TeamsManifestVDevPreview.ts"),
      "utf8"
    );
    assert.strictEqual(source, lines.join("\n"));
  });

  for (const value of ["text", [], [{ enabled: true }]]) {
    it(`keeps successful union alternatives: ${JSON.stringify(value)}`, () => {
      const json = JSON.stringify({ value });
      const result = converter.toProbe(json);

      assert.deepEqual(JSON.parse(converter.probeToJson(result)), JSON.parse(json));
    });
  }

  it("keeps nested failures through a multi-member union", () => {
    assert.throws(
      () => converter.toProbe(JSON.stringify({ value: [{ enabled: "private-value" }] })),
      /value\[0\]\.enabled.*boolean.*got string/
    );
  });

  it("describes all alternatives when no union member matches the input shape", () => {
    const json = JSON.stringify({ value: false });

    assert.throws(() => converter.toProbe(json), /one of \[.*array.*string.*\].*got boolean/);
  });

  it("describes optional enum values rather than an object descriptor", () => {
    assert.throws(
      () => converter.toProbe(JSON.stringify({ value: "text", mode: "invalid" })),
      /mode.*"first".*"second".*got string/
    );
  });

  it("handles null type descriptors in diagnostics", () => {
    assert.throws(
      () => converter.toProbe(JSON.stringify({ value: "text", nullable: 1 })),
      /nullable.*null.*string.*got number/
    );
  });

  it("preserves explicit null and missing optional properties", () => {
    const json = JSON.stringify({ value: "text", nullable: null });
    const result = converter.toProbe(json);

    assert.deepEqual(JSON.parse(converter.probeToJson(result)), JSON.parse(json));
  });

  it("formats map keys without confusing literal dots with nested properties", () => {
    assert.throws(
      () => converter.toProbe(JSON.stringify({ value: "text", attributes: { "a.b": "invalid" } })),
      /attributes\["a\.b"\].*number.*got string/
    );
  });

  it("preserves date conversion and serialization", () => {
    const json = JSON.stringify({ value: "text", timestamp: "2026-10-15T00:00:00.000Z" });
    const result = converter.toProbe(json);

    assert.deepEqual(JSON.parse(converter.probeToJson(result)), JSON.parse(json));
  });

  it("reports invalid dates without leaking the input", () => {
    assert.throws(
      () => converter.toProbe(JSON.stringify({ value: "text", timestamp: "private-value" })),
      /timestamp.*Date.*got string/
    );
  });

  it("does not swallow unrelated exceptions while trying union members", () => {
    const failure = new Error("Getter failed");
    const value = {
      value: "text",
      nested: {
        get label() {
          throw failure;
        },
      },
    };
    try {
      converter.probeToJson(value);
    } catch (error) {
      assert.strictEqual(error, failure);
      return;
    }
    assert.fail("Expected the original exception");
  });
});
