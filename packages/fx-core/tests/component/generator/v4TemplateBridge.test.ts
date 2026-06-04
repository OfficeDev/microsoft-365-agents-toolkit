// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import os from "os";
import path from "path";
import { createSandbox } from "sinon";
import { TemplateFileEntry } from "../../../src/v4";
import { GeneratorContext } from "../../../src/component/generator/generatorAction";
import {
  renderTemplateFileData,
  renderTemplateFileName,
} from "../../../src/component/generator/utils";
import { renderTemplateEntries } from "../../../src/component/generator/v4TemplateBridge";

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
    tmpDir = path.join(
      os.tmpdir(),
      `v4bridge-${Date.now()}-${Math.random().toString(36).slice(2)}`
    );
    await fs.ensureDir(tmpDir);
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
});
