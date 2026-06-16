// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform, SystemError } from "@microsoft/teamsfx-api";
import chai, { assert } from "chai";
import chaiAsPromised from "chai-as-promised";
import fs from "fs-extra";
import { err, ok } from "neverthrow";
import os from "os";
import path from "path";
import { createSandbox } from "sinon";
import { TelemetryProperty } from "../../../src/common/telemetry";
import { GeneratorContext } from "../../../src/component/generator/generatorAction";
import {
  renderTemplateFileData,
  renderTemplateFileName,
} from "../../../src/component/generator/utils";
import {
  renderTemplateEntries,
  scaffoldFromV4Channel,
  v4TemplateBridgeDeps,
} from "../../../src/component/generator/v4TemplateBridge";
import { TemplateFileEntry, TemplateSource } from "../../../src/v4";

chai.use(chaiAsPromised);

// Build a GeneratorContext whose rename/data/filter functions mirror exactly
// what DefaultTemplateGenerator.scaffolding constructs, so the render contract
// is validated against the real v3 rendering functions.
function makeContext(
  folderName: string,
  destination: string,
  replaceMap: { [key: string]: string },
  extraFilter?: (fileName: string) => boolean
): GeneratorContext {
  return {
    name: folderName,
    language: "common",
    destination,
    logProvider: {
      debug: () => {},
      info: () => {},
      warning: () => {},
      error: () => {},
    } as any,
    platform: Platform.VSCode,
    fileNameReplaceFn: (fileName, fileData) =>
      renderTemplateFileName(fileName, fileData, replaceMap)
        .replace(/\\/g, "/")
        .replace(`${folderName}/`, ""),
    fileDataReplaceFn: (fileName, fileData) =>
      renderTemplateFileData(fileName, fileData, replaceMap),
    filterFn: (fileName) =>
      fileName.replace(/\\/g, "/").startsWith(`${folderName}/`) &&
      (extraFilter ? extraFilter(fileName) : true),
    onActionError: () => Promise.resolve(),
  };
}

describe("v4TemplateBridge.renderTemplateEntries", () => {
  const sandbox = createSandbox();
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "v4bridge-"));
  });

  afterEach(async () => {
    sandbox.restore();
    await fs.remove(tmpDir);
  });

  it("writes entries with the locator prefix re-added then stripped, returns written paths", async () => {
    const folderName = "declarative-agent-basic";
    const ctx = makeContext(folderName, tmpDir, {});
    const entries: TemplateFileEntry[] = [
      { path: "manifest.json", data: Buffer.from('{"a":1}') },
      { path: "src/index.ts", data: Buffer.from("console.log(1);") },
    ];

    const outputs = await renderTemplateEntries(ctx, entries);

    assert.deepEqual(outputs.sort(), ["manifest.json", "src/index.ts"]);
    assert.strictEqual(
      (await fs.readFile(path.join(tmpDir, "manifest.json"))).toString(),
      '{"a":1}'
    );
    assert.strictEqual(
      (await fs.readFile(path.join(tmpDir, "src/index.ts"))).toString(),
      "console.log(1);"
    );
  });

  it("renders .tpl mustache data and strips the .tpl suffix from the name", async () => {
    const folderName = "bot";
    const ctx = makeContext(folderName, tmpDir, { appName: "MyApp" });
    const entries: TemplateFileEntry[] = [
      { path: "config.json.tpl", data: Buffer.from('{"name":"{{appName}}"}') },
    ];

    const outputs = await renderTemplateEntries(ctx, entries);

    assert.deepEqual(outputs, ["config.json"]);
    assert.strictEqual(
      (await fs.readFile(path.join(tmpDir, "config.json"))).toString(),
      '{"name":"MyApp"}'
    );
  });

  it("does not render data for non-.tpl files (binary preserved verbatim)", async () => {
    const folderName = "bot";
    const ctx = makeContext(folderName, tmpDir, { appName: "MyApp" });
    const binary = Buffer.from([0x00, 0x01, 0x02, 0xff, 0xfe]);
    const entries: TemplateFileEntry[] = [{ path: "assets/icon.png", data: binary }];

    await renderTemplateEntries(ctx, entries);

    const written = await fs.readFile(path.join(tmpDir, "assets/icon.png"));
    assert.isTrue(written.equals(binary));
  });

  it("excludes entries rejected by the context filterFn", async () => {
    const folderName = "bot";
    const ctx = makeContext(
      folderName,
      tmpDir,
      {},
      (fileName) => !fileName.endsWith(".env.sandbox")
    );
    const entries: TemplateFileEntry[] = [
      { path: "keep.txt", data: Buffer.from("keep") },
      { path: ".env.sandbox", data: Buffer.from("secret") },
    ];

    const outputs = await renderTemplateEntries(ctx, entries);

    assert.deepEqual(outputs, ["keep.txt"]);
    assert.isFalse(await fs.pathExists(path.join(tmpDir, ".env.sandbox")));
  });

  it("respects the trailing-slash locator boundary via the re-added prefix", async () => {
    // folderName "da" must not pick up a sibling whose name starts with "da-".
    // The bridge re-adds "${name}/" so filterFn's startsWith("da/") is exact.
    const folderName = "da";
    const ctx = makeContext(folderName, tmpDir, {});
    const entries: TemplateFileEntry[] = [{ path: "file.txt", data: Buffer.from("x") }];

    const outputs = await renderTemplateEntries(ctx, entries);

    assert.deepEqual(outputs, ["file.txt"]);
    // entryName seen by filterFn is "da/file.txt", which does NOT start with "da-".
    assert.isFalse("da/file.txt".startsWith("da-/"));
  });

  it("writes entries verbatim under the name prefix when no optional fns are set", async () => {
    const ctx: GeneratorContext = {
      name: "bot",
      language: "common",
      destination: tmpDir,
      logProvider: {
        debug: () => {},
        info: () => {},
        warning: () => {},
        error: () => {},
      } as any,
      platform: Platform.VSCode,
      onActionError: () => Promise.resolve(),
    };
    const entries: TemplateFileEntry[] = [{ path: "a.txt", data: Buffer.from("a") }];

    const outputs = await renderTemplateEntries(ctx, entries);

    assert.deepEqual(outputs, ["bot/a.txt"]);
    assert.strictEqual((await fs.readFile(path.join(tmpDir, "bot/a.txt"))).toString(), "a");
  });

  it("rejects an entry whose path escapes the destination (zip-slip)", async () => {
    const folderName = "bot";
    const ctx = makeContext(folderName, tmpDir, {});
    // entryName "bot/../evil.txt" passes the startsWith("bot/") filter but the
    // name-replace strips "bot/" leaving "../evil.txt", which escapes tmpDir.
    const entries: TemplateFileEntry[] = [{ path: "../evil.txt", data: Buffer.from("pwned") }];

    await assert.isRejected(
      renderTemplateEntries(ctx, entries),
      /resolves outside the destination directory/
    );
    assert.isFalse(await fs.pathExists(path.join(path.dirname(tmpDir), "evil.txt")));
  });

  it("allows an in-root filename that starts with '..' (not a traversal segment)", async () => {
    const folderName = "bot";
    const ctx = makeContext(folderName, tmpDir, {});
    // "bot/..foo" name-replaces to "..foo": its relative path starts with ".."
    // but stays inside tmpDir, so it must NOT be rejected.
    const entries: TemplateFileEntry[] = [{ path: "..foo", data: Buffer.from("ok") }];

    const outputs = await renderTemplateEntries(ctx, entries);

    assert.deepEqual(outputs, ["..foo"]);
    assert.strictEqual((await fs.readFile(path.join(tmpDir, "..foo"))).toString(), "ok");
  });
});

describe("v4TemplateBridge.scaffoldFromV4Channel", () => {
  const sandbox = createSandbox();
  let tmpDir: string;
  const locator = { language: "common", scenario: "declarative-agent-basic" };
  const source: TemplateSource = {
    origin: "bundled",
    version: "6.10.1",
    digest: "sha256:abc",
    location: "/floor/templates.zip",
  };

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "v4bridge-"));
    sandbox.stub(v4TemplateBridgeDeps, "createTemplateSourcePort").returns({} as any);
    sandbox.stub(v4TemplateBridgeDeps, "loadBundledFloor").returns({} as any);
  });

  afterEach(async () => {
    sandbox.restore();
    await fs.remove(tmpDir);
  });

  it("resolves, reads, renders and records source telemetry on the happy path", async () => {
    const ctx = makeContext("declarative-agent-basic", tmpDir, {});
    const entries: TemplateFileEntry[] = [{ path: "manifest.json", data: Buffer.from('{"a":1}') }];
    sandbox.stub(v4TemplateBridgeDeps, "resolveTemplateSource").resolves(ok(source));
    sandbox.stub(v4TemplateBridgeDeps, "loadResolvedPackage").returns(ok(Buffer.from("zip-bytes")));
    sandbox.stub(v4TemplateBridgeDeps, "openTemplatePackage").returns(ok(entries));
    const telemetryProps: Record<string, string> = {};

    const result = await scaffoldFromV4Channel(ctx, locator, telemetryProps);

    assert.deepEqual(result, source);
    assert.deepEqual(ctx.outputs, ["manifest.json"]);
    assert.strictEqual(
      (await fs.readFile(path.join(tmpDir, "manifest.json"))).toString(),
      '{"a":1}'
    );
    assert.strictEqual(telemetryProps[TelemetryProperty.TemplatePackageSource], "bundled");
    assert.strictEqual(telemetryProps[TelemetryProperty.TemplatePackageVersion], "6.10.1");
    assert.strictEqual(telemetryProps[TelemetryProperty.TemplatePackageDigest], "sha256:abc");
  });

  it("throws and records no source telemetry when resolution fails", async () => {
    const ctx = makeContext("declarative-agent-basic", tmpDir, {});
    const resolveError = new SystemError("v4", "ResolveFailed", "no tag");
    sandbox.stub(v4TemplateBridgeDeps, "resolveTemplateSource").resolves(err(resolveError));
    const telemetryProps: Record<string, string> = {};

    await assert.isRejected(scaffoldFromV4Channel(ctx, locator, telemetryProps), "no tag");
    assert.isUndefined(telemetryProps[TelemetryProperty.TemplatePackageSource]);
  });

  it("throws but still records source telemetry when reading the package fails", async () => {
    const ctx = makeContext("declarative-agent-basic", tmpDir, {});
    sandbox.stub(v4TemplateBridgeDeps, "resolveTemplateSource").resolves(ok(source));
    sandbox
      .stub(v4TemplateBridgeDeps, "loadResolvedPackage")
      .returns(err(new SystemError("v4", "DigestMismatch", "bad digest")));
    const telemetryProps: Record<string, string> = {};

    await assert.isRejected(scaffoldFromV4Channel(ctx, locator, telemetryProps), "bad digest");
    assert.strictEqual(telemetryProps[TelemetryProperty.TemplatePackageVersion], "6.10.1");
    assert.isUndefined(ctx.outputs);
  });

  it("throws when the package cannot be opened", async () => {
    const ctx = makeContext("declarative-agent-basic", tmpDir, {});
    sandbox.stub(v4TemplateBridgeDeps, "resolveTemplateSource").resolves(ok(source));
    sandbox.stub(v4TemplateBridgeDeps, "loadResolvedPackage").returns(ok(Buffer.from("zip")));
    sandbox
      .stub(v4TemplateBridgeDeps, "openTemplatePackage")
      .returns(err(new SystemError("v4", "OpenFailed", "corrupt zip")));

    await assert.isRejected(scaffoldFromV4Channel(ctx, locator, {}), "corrupt zip");
  });
});
