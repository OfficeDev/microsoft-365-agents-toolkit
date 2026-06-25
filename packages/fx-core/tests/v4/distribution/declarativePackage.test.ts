// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import AdmZip from "adm-zip";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { SystemError } from "@microsoft/teamsfx-api";
import { loadPackageDir } from "../../../src/v4/distribution/packageDir";
import { openDeclarativePackage } from "../../../src/v4/distribution/declarativePackage";
import { createRealRuntime } from "../../../src/v4/runtime/realRuntime";
import { scaffold } from "../../../src/v4/runtime/scaffold";

/**
 * The zip-bytes declarative reader: pull one authored package's
 * `v4/<kind>/<templateId>/` subtree (descriptor + pipeline + content) out of the
 * channel `templates.zip` — the on-the-wire / bundled-floor sibling of
 * `loadPackageDir` (which reads the same shape from a loose authoring dir).
 *
 * The keystone property is parity (DECL-06): the package read from the zip is
 * byte-identical to the one read from its source dir, so `scaffold` runs the
 * SAME from the shipped floor as from the dev tree (DECL-07).
 *
 * Reads only the `v4/` declarative subtree; the zip's coexisting v3 mirror
 * (`<lang>/<scenario>/`) is never touched (INV-7 — v4 depends on no v3 content).
 *
 * Spec: docs/03-specs/operations/scaffolding/open-template-package.md (the
 * declarative-subtree variant of the same consume boundary).
 */

const PKG_DIR = path.resolve(__dirname, "../../../../../templates/v4/create/da/mcp-server");

const LOCATOR = { kind: "create", templateId: "da/mcp-server" };
const NAMESPACE = "apigithubc";

/**
 * Build a channel-shaped zip: the v4 authored package under
 * `v4/create/da/mcp-server/` plus a coexisting v3-mirror entry the declarative
 * reader must ignore.
 */
function buildChannelZip(): Buffer {
  const zip = new AdmZip();
  zip.addLocalFolder(PKG_DIR, "v4/create/da/mcp-server");
  zip.addFile("common/da-mcp/should-be-ignored.txt", Buffer.from("v3 mirror content"));
  return zip.toBuffer();
}

describe("openDeclarativePackage (v4, zip declarative-subtree reader)", () => {
  const bytes = buildChannelZip();

  it("DECL-01: opens the v4 subtree, ignoring root JSON and the v3 mirror", () => {
    const result = openDeclarativePackage(bytes, LOCATOR);
    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
    const loaded = result._unsafeUnwrap();

    const descriptor = loaded.descriptor as { id?: string };
    assert.strictEqual(descriptor.id, "da/mcp-server");
    const pipeline = loaded.pipeline as { pipeline?: string };
    assert.strictEqual(pipeline.pipeline, "default");

    const paths = loaded.content.map((entry) => entry.path);
    assert.include(paths, "appPackage/ai-plugin.json.tpl");
    assert.include(paths, "m365agents.yml.tpl");
    // content is ONLY the content/ subtree — package JSON is parsed away, not shipped as a file.
    assert.notInclude(paths, "descriptor.json");
    assert.notInclude(paths, "pipeline.json");
    assert.notInclude(paths, "questions.json");
    // the coexisting v3 mirror never leaks into a v4 read.
    assert.notInclude(paths, "should-be-ignored.txt");
    assert.isFalse(
      paths.some((p) => p.includes("da-mcp")),
      "no v3-mirror path may appear"
    );
  });

  it("DECL-02: content paths are content-relative, forward-slash, sorted", () => {
    const loaded = openDeclarativePackage(bytes, LOCATOR)._unsafeUnwrap();
    const paths = loaded.content.map((entry) => entry.path);

    assert.isFalse(
      paths.some((p) => p.startsWith("content/") || p.includes("\\")),
      "paths are relative to the content root, forward-slash normalized"
    );
    const sorted = [...paths].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
    assert.deepEqual(paths, sorted, "entries are deterministically sorted");
  });

  it("DECL-03: an unknown templateId is a missing-descriptor SystemError", () => {
    const result = openDeclarativePackage(bytes, { kind: "create", templateId: "da/nope" });
    assert.isTrue(result.isErr(), "expected an error for an absent package");
    const error = result._unsafeUnwrapErr();
    assert.instanceOf(error, SystemError);
    assert.strictEqual(error.name, "PackageFileMissing");
  });

  it("DECL-04: invalid archive bytes are a SystemError", () => {
    const result = openDeclarativePackage(Buffer.from("not a zip at all"), LOCATOR);
    assert.isTrue(result.isErr(), "expected an error for non-zip bytes");
    assert.strictEqual(result._unsafeUnwrapErr().name, "TemplatePackageCorrupt");
  });

  it("DECL-05: a Zip-Slip content entry is rejected", () => {
    const zip = new AdmZip();
    zip.addLocalFolder(PKG_DIR, "v4/create/da/mcp-server");
    zip.addFile("v4/create/da/mcp-server/content/evil.txt", Buffer.from("pwned"));
    // adm-zip canonicalizes `..` at add time, so set the raw traversal name directly.
    const evil = zip
      .getEntries()
      .find((entry) => entry.entryName.replace(/\\/g, "/").endsWith("content/evil.txt"));
    assert.isDefined(evil);
    evil!.entryName = "v4/create/da/mcp-server/content/../evil.txt";

    const result = openDeclarativePackage(zip.toBuffer(), LOCATOR);
    assert.isTrue(result.isErr(), "expected the Zip-Slip guard to fire");
    assert.strictEqual(result._unsafeUnwrapErr().name, "TemplatePackageUnsafePath");
  });

  it("DECL-06: the zip read is byte-identical to the loose-dir read (parity keystone)", () => {
    const fromZip = openDeclarativePackage(bytes, LOCATOR)._unsafeUnwrap();
    const fromDirResult = loadPackageDir(PKG_DIR);
    assert.isTrue(fromDirResult.isOk(), "loose-dir read must succeed");
    const fromDir = fromDirResult._unsafeUnwrap();

    assert.deepEqual(fromZip.descriptor, fromDir.descriptor, "descriptor must match");
    assert.deepEqual(fromZip.pipeline, fromDir.pipeline, "pipeline must match");

    const zipByPath = new Map(fromZip.content.map((entry) => [entry.path, entry.data]));
    const dirByPath = new Map(fromDir.content.map((entry) => [entry.path, entry.data]));
    assert.deepEqual(
      [...zipByPath.keys()].sort(),
      [...dirByPath.keys()].sort(),
      "the two reads must expose the same content paths"
    );
    for (const [contentPath, dirData] of dirByPath) {
      const zipData = zipByPath.get(contentPath);
      assert.isDefined(zipData, `zip read missing '${contentPath}'`);
      assert.isTrue(zipData!.equals(dirData), `bytes of '${contentPath}' must match`);
    }
  });

  it("DECL-07: scaffolds onto disk straight from the channel bytes (end-to-end)", async () => {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "atk-v4-decl-"));
    try {
      const loaded = openDeclarativePackage(bytes, LOCATOR)._unsafeUnwrap();
      const result = await scaffold(
        {
          descriptor: loaded.descriptor,
          pipeline: loaded.pipeline,
          content: loaded.content,
          answers: {
            mcpServerType: "remote",
            mcpServerUrl: "https://api.github.com/mcp",
            authType: "none",
          },
          callerFloor: { appName: "MyMcpAgent", language: "common" },
          targetDir: { path: tempDir, existing: [] },
        },
        createRealRuntime(tempDir)
      );
      assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");

      const aiPluginPath = path.join(tempDir, "appPackage/ai-plugin.json");
      assert.isTrue(fs.existsSync(aiPluginPath), "ai-plugin.json must be on disk");
      assert.isFalse(
        fs.existsSync(path.join(tempDir, "m365agents.yml.tpl")),
        ".tpl must be stripped"
      );
      const aiPlugin: { namespace?: string } = JSON.parse(fs.readFileSync(aiPluginPath, "utf8"));
      assert.strictEqual(aiPlugin.namespace, NAMESPACE);
    } finally {
      fs.removeSync(tempDir);
    }
  });
});
