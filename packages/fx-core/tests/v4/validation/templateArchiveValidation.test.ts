// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError, UserError } from "@microsoft/teamsfx-api";
import { spawnSync } from "child_process";
import * as path from "path";
import AdmZip from "adm-zip";
import { assert } from "vitest";
import {
  validateDeclarativePackageArchive,
  validateDeclarativeTemplateArchive,
} from "../../../src/v4/validation/templateArchiveValidation";

const V4_ROOT = path.resolve(__dirname, "../../../../../templates/v4");
let fullArchiveCache: Buffer | undefined;
const runtimeErrors = {
  user: (name: string, message: string) => new UserError({ source: "Scaffold", name, message }),
  system: (name: string, message: string) => new SystemError({ source: "Scaffold", name, message }),
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function fullArchive(): Buffer {
  if (fullArchiveCache === undefined) {
    const zip = new AdmZip();
    zip.addLocalFolder(V4_ROOT, "v4");
    fullArchiveCache = zip.toBuffer();
  }
  return fullArchiveCache;
}

function archiveWithUnsafePosixContentPath(): Buffer {
  const zip = new AdmZip(fullArchive());
  const safeEntryName = "v4/create/da/mcp-server/content/_payload.txt";
  const unsafeEntryName = "v4/create/da/mcp-server/content//payload.txt";
  zip.addFile(safeEntryName, Buffer.from("unsafe", "utf8"));

  // addFile normalizes repeated separators, so rewrite the same-length name in both ZIP headers.
  const archive = zip.toBuffer();
  const safeEntryNameBytes = Buffer.from(safeEntryName, "utf8");
  const unsafeEntryNameBytes = Buffer.from(unsafeEntryName, "utf8");
  assert.equal(safeEntryNameBytes.length, unsafeEntryNameBytes.length);

  let replacementCount = 0;
  let offset = archive.indexOf(safeEntryNameBytes);
  while (offset >= 0) {
    unsafeEntryNameBytes.copy(archive, offset);
    replacementCount++;
    offset = archive.indexOf(safeEntryNameBytes, offset + safeEntryNameBytes.length);
  }
  assert.equal(replacementCount, 2, "expected entry names in local and central ZIP headers");
  return archive;
}

function selectorOnlyArchive(): AdmZip {
  const zip = new AdmZip();
  zip.addLocalFolder(path.join(V4_ROOT, "schema"), "v4/schema");
  zip.addLocalFile(path.join(V4_ROOT, "create", "selector.json"), "v4/create");
  zip.addLocalFile(path.join(V4_ROOT, "modify", "selector.json"), "v4/modify");
  return zip;
}

describe("v4/validation/templateArchiveValidation", () => {
  it("AC-29: validates source archives before product API build output exists", () => {
    const templatesRoot = path.dirname(V4_ROOT);
    const child = spawnSync(
      process.execPath,
      [
        "--no-preserve-symlinks",
        "--eval",
        `
          const assert = require("node:assert/strict");
          const fs = require("node:fs");
          const bytes = fs.readFileSync(0);
          const Module = require("node:module");
          const sourceRequire = Module.createRequire(require("node:path").resolve("../packages/fx-core/src/v4/validation/templateArchiveValidation.ts"));
          const resolution = () => {
            const semver = sourceRequire.resolve("semver");
            const semverRequire = Module.createRequire(semver);
            const lru = semverRequire.resolve("lru-cache");
            return { semver, realSemver: fs.realpathSync.native(semver), lru,
              semverVersion: semverRequire("./package.json").version,
              lruVersion: semverRequire("lru-cache/package.json").version };
          };
          const nativeResolution = resolution();
          require("tsx/cjs");
          const resolveFilename = Module._resolveFilename;
          const loadedModules = [];
          Module._resolveFilename = function(request, ...args) {
            assert.ok(!request.startsWith("@microsoft/teamsfx-api"),
              "Build validation must not load product API: " + request);
            const resolved = resolveFilename.call(this, request, ...args);
            if (request === "semver" || request === "lru-cache") {
              loadedModules.push({ request, parent: args[0]?.filename, resolved });
            }
            return resolved;
          };
          try {
          assert.throws(() => require.resolve("@microsoft/teamsfx-api"), /must not load product API/);
          const { validateDeclarativeTemplateArchive, validateDeclarativePackageArchive } =
            require("../packages/fx-core/src/v4/validation/templateArchiveValidation.ts");
          const { CURRENT_V4_ENGINE_VERSION } = require("../packages/fx-core/src/v4/engineVersion.ts");
          const failure = Object.assign(new Error("caller-owned"), { source: "TemplatesBuild" });
          const errors = { user: () => failure, system: () => failure };
          const archive = validateDeclarativeTemplateArchive(bytes, "build", CURRENT_V4_ENGINE_VERSION, errors);
          assert.ok(archive.isOk());
          const opened = validateDeclarativePackageArchive(bytes,
            { kind: "create", templateId: "da/mcp-server" }, "load", CURRENT_V4_ENGINE_VERSION, errors);
          assert.ok(opened.isOk());
          assert.equal(opened.value.template.descriptor.id, "da/mcp-server");
          assert.equal(opened.value.template.content, opened.value.content);
          const invalid = validateDeclarativeTemplateArchive(Buffer.from("invalid"), "build", CURRENT_V4_ENGINE_VERSION, errors);
          assert.equal(invalid.error, failure);
          const { prepareTemplate, PACKAGE_PARSE_ERROR } = require("../packages/fx-core/src/v4/validation/packageParse.ts");
          const parseError = (name, message) => {
            assert.equal(name, PACKAGE_PARSE_ERROR);
            assert.ok(message.length > 0);
            return failure;
          };
          for (const raw of [
            { descriptor: [], pipeline: {} },
            { descriptor: {}, pipeline: {} },
            { descriptor: { languageOptions: false }, pipeline: { pipeline: "default", steps: [] } },
          ]) {
            assert.equal(prepareTemplate({ ...raw, content: [] }, parseError).error, failure);
          }
          console.log("validated without product API build output");
          } catch (error) {
            console.error(JSON.stringify({ nativeResolution, loadedModules }));
            console.error(String(error.stack ?? error.message).replaceAll("    at ", "    frame: "));
            process.exitCode = 1;
          }
        `,
      ],
      {
        cwd: templatesRoot,
        input: fullArchive(),
        encoding: "utf8",
        timeout: 60000,
        env: {
          ...process.env,
          NODE_OPTIONS: "--preserve-symlinks",
          NODE_PATH: path.resolve(templatesRoot, "../node_modules/.pnpm/node_modules"),
        },
      }
    );
    assert.equal(child.status, 0, child.stderr || child.error?.message);
    assert.include(child.stdout, "validated without product API build output");
  });

  it("CLEAN-07: consuming the new OpenAPI output requires its introduction version", () => {
    const zip = new AdmZip(fullArchive());
    const descriptorPath = "v4/create/da/api-plugin-from-existing-api/descriptor.json";
    const descriptor: unknown = JSON.parse(zip.readAsText(descriptorPath));
    if (!isRecord(descriptor)) {
      assert.fail("expected descriptor");
    }
    descriptor.minEngineVersion = "6.11.0";
    zip.updateFile(descriptorPath, Buffer.from(JSON.stringify(descriptor)));
    const result = validateDeclarativePackageArchive(
      zip.toBuffer(),
      { kind: "create", templateId: "da/api-plugin-from-existing-api" },
      "load",
      "6.11.0",
      runtimeErrors
    );
    assert.equal(result._unsafeUnwrapErr().name, "TemplatePackageCapabilityFloor");
  });

  it("validates the complete archive and opens one package from the same final bytes", () => {
    const bytes = fullArchive();

    const archiveResult = validateDeclarativeTemplateArchive(
      bytes,
      "build",
      "6.11.0",
      runtimeErrors
    );
    const packageResult = validateDeclarativePackageArchive(
      bytes,
      { kind: "create", templateId: "da/mcp-server" },
      "load",
      "6.11.0",
      runtimeErrors
    );

    if (archiveResult.isErr()) {
      throw archiveResult.error;
    }
    assert.include(archiveResult.value, "create/da/mcp-server");
    if (packageResult.isErr()) {
      throw packageResult.error;
    }
    const loadedPackage = packageResult.value;
    assert.isTrue(isRecord(loadedPackage.descriptor));
    if (!isRecord(loadedPackage.descriptor)) {
      throw new Error("expected descriptor object");
    }
    assert.equal(loadedPackage.descriptor.id, "da/mcp-server");
    assert.isAbove(loadedPackage.content.length, 0);
  });

  it("rejects bytes that are not a zip archive", () => {
    const result = validateDeclarativeTemplateArchive(
      Buffer.from("not-a-zip"),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplatePackageCorrupt");
  });

  it("rejects non-zip bytes when opening one package", () => {
    const result = validateDeclarativePackageArchive(
      Buffer.from("not-a-zip"),
      { kind: "create", templateId: "da/mcp-server" },
      "load",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplatePackageCorrupt");
  });

  it("rejects an archive with a missing schema file", () => {
    const zip = selectorOnlyArchive();
    zip.deleteFile("v4/schema/pipeline.schema.json");

    const result = validateDeclarativePackageArchive(
      zip.toBuffer(),
      { kind: "create", templateId: "da/mcp-server" },
      "load",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "PackageFileMissing");
    assert.include(result._unsafeUnwrapErr().message, "pipeline.schema.json");
  });

  it("rejects a schema file containing malformed JSON", () => {
    const zip = selectorOnlyArchive();
    zip.updateFile("v4/schema/pipeline.schema.json", Buffer.from("{"));

    const result = validateDeclarativeTemplateArchive(
      zip.toBuffer(),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "PackageFileInvalid");
    assert.include(result._unsafeUnwrapErr().message, "pipeline.schema.json");
  });

  it("rejects a schema file whose JSON value is not a schema", () => {
    const zip = selectorOnlyArchive();
    zip.updateFile("v4/schema/pipeline.schema.json", Buffer.from("[]"));

    const result = validateDeclarativeTemplateArchive(
      zip.toBuffer(),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplateSchemaInvalid");
    assert.include(result._unsafeUnwrapErr().message, "pipeline.schema.json");
  });

  it("rejects a schema set that AJV cannot compile", () => {
    const zip = selectorOnlyArchive();
    zip.updateFile(
      "v4/schema/pipeline.schema.json",
      Buffer.from(JSON.stringify({ type: "not-a-json-schema-type" }))
    );

    const result = validateDeclarativeTemplateArchive(
      zip.toBuffer(),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplateSchemaInvalid");
    assert.include(result._unsafeUnwrapErr().message, "could not be compiled");
  });

  it("rejects a selector containing malformed JSON", () => {
    const zip = selectorOnlyArchive();
    zip.updateFile("v4/modify/selector.json", Buffer.from("{"));

    const result = validateDeclarativeTemplateArchive(
      zip.toBuffer(),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "PackageFileInvalid");
    assert.include(result._unsafeUnwrapErr().message, "v4/modify/selector.json");
  });

  it("rejects an invalid selector when opening one package", () => {
    const zip = selectorOnlyArchive();
    const selector = zip.getEntry("v4/create/selector.json");
    if (selector === null) {
      throw new Error("expected create selector entry");
    }
    selector.setData(Buffer.from(JSON.stringify({ routes: "not-an-array" }), "utf8"));

    const result = validateDeclarativePackageArchive(
      zip.toBuffer(),
      { kind: "create", templateId: "da/mcp-server" },
      "load",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplatePackageSchema");
    assert.include(result._unsafeUnwrapErr().message, "v4/create/selector.json");
  });

  it.each(["descriptor", "questions", "pipeline"])(
    "rejects a package whose %s JSON is malformed",
    (fileName) => {
      const zip = new AdmZip(fullArchive());
      zip.updateFile(`v4/create/da/mcp-server/${fileName}.json`, Buffer.from("{"));

      const result = validateDeclarativePackageArchive(
        zip.toBuffer(),
        { kind: "create", templateId: "da/mcp-server" },
        "load",
        "6.11.0",
        runtimeErrors
      );

      assert.isTrue(result.isErr());
      assert.equal(result._unsafeUnwrapErr().name, "PackageFileInvalid");
      assert.include(result._unsafeUnwrapErr().message, `${fileName}.json`);
    }
  );

  it("AC-26: POSIX-absolute content paths are rejected before rendering", () => {
    const result = validateDeclarativeTemplateArchive(
      archiveWithUnsafePosixContentPath(),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplatePackageUnsafePath");
    assert.include(result._unsafeUnwrapErr().message, "content//payload.txt");
  });

  it("AC-29: archive validation uses the caller-owned error factory", () => {
    const zip = selectorOnlyArchive();
    const selector = zip.getEntry("v4/create/selector.json");
    if (selector === null) {
      throw new Error("expected create selector entry");
    }
    selector.setData(Buffer.from(JSON.stringify({ routes: "not-an-array" }), "utf8"));

    const errors = {
      source: "TemplatesBuild",
      user(name: string, message: string) {
        return Object.assign(new Error(message), {
          name,
          source: this.source,
          timestamp: new Date(),
        });
      },
      system(name: string, message: string) {
        return Object.assign(new Error(message), {
          name,
          source: this.source,
          timestamp: new Date(),
        });
      },
    };

    const result = validateDeclarativeTemplateArchive(zip.toBuffer(), "build", "6.11.0", errors);

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().source, "TemplatesBuild");
  });

  it("AC-25: descriptor-less package metadata is not omitted from archive validation", () => {
    const zip = new AdmZip(fullArchive());
    zip.addFile(
      "v4/create/aa-orphan/questions.json",
      Buffer.from(JSON.stringify({ questions: [] }), "utf8")
    );
    zip.addFile(
      "v4/create/aa-orphan/pipeline.json",
      Buffer.from(JSON.stringify({ pipeline: "default", steps: [] }), "utf8")
    );

    const result = validateDeclarativeTemplateArchive(
      zip.toBuffer(),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplatePackageRequiredFile");
    assert.include(result._unsafeUnwrapErr().message, "create/aa-orphan");
    assert.include(result._unsafeUnwrapErr().message, "descriptor.json");
  });

  it("AC-26: drive-qualified content paths are rejected before rendering", () => {
    const zip = new AdmZip(fullArchive());
    zip.addFile("v4/create/da/mcp-server/content/z:/payload.txt", Buffer.from("unsafe", "utf8"));

    const result = validateDeclarativeTemplateArchive(
      zip.toBuffer(),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplatePackageUnsafePath");
    assert.include(result._unsafeUnwrapErr().message, "z:/payload.txt");
  });

  it("AC-27: selectors are schema-validated even when the archive has no packages", () => {
    const zip = selectorOnlyArchive();
    const selector = zip.getEntry("v4/create/selector.json");
    if (selector === null) {
      throw new Error("expected create selector entry");
    }
    selector.setData(Buffer.from(JSON.stringify({ routes: "not-an-array" }), "utf8"));

    const result = validateDeclarativeTemplateArchive(
      zip.toBuffer(),
      "build",
      "6.11.0",
      runtimeErrors
    );

    assert.isTrue(result.isErr());
    assert.equal(result._unsafeUnwrapErr().name, "TemplatePackageSchema");
    assert.include(result._unsafeUnwrapErr().message, "v4/create/selector.json");
  });
});
