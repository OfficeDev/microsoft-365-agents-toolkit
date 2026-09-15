// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { UserError } from "@microsoft/teamsfx-api";
import { TemplateFileEntry } from "../../../src/v4/model/dataModel";
import {
  EXISTING_FILE_SKIPPED_WARNING,
  REQUIRE_EMPTY_TARGET,
  runScaffoldPipeline,
} from "../../../src/v4/pipeline/runScaffoldPipeline";
import { createRealRuntime } from "../../../src/v4/runtime/realRuntime";
import { ScaffoldRequest, scaffold } from "../../../src/v4/runtime/scaffold";
import { mcpAuthScaffoldDeps } from "../../../src/v4/mcp/mcpAuthScaffold";
import { afterEach, assert, beforeEach, expect, vi } from "vitest";

/**
 * The on-disk `ScaffoldRuntime` face (ADR-0018): the same `da/mcp-server` create
 * package the T3 in-memory scenario scaffolds, but materialized onto a real
 * temp directory through `createRealRuntime`. This proves the production seam —
 * the render phase writes real bytes, a read-modify-write step reads them back
 * from disk, `.tpl` suffixes are stripped into real filenames, and the
 * create-empty contract guards a real non-empty directory.
 *
 * Spec: docs/03-specs/scenarios/da/create-mcp-server.md (the SCN-CREATE-MCP-*
 * contract, here re-validated against a real filesystem sink).
 *
 * v4-owned (INV-7), including the MCP auth YAML action mutator shared by create and add.
 */

const PKG_DIR = path.resolve(__dirname, "../../../../../templates/v4/create/da/mcp-server");

const MCP_SERVER_URL = "https://api.github.com/mcp"; // namespace derives to apigithubc
const NAMESPACE = "apigithubc";
const AUTH_ENV_VAR = "MCP_DA_AUTH_ID_APIGITHUBC";
const CLIENT_ID_ENV_VAR = "MCP_DA_OAUTH_CLIENT_ID_APIGITHUBC";
const CLIENT_SECRET_ENV_VAR = "SECRET_MCP_DA_OAUTH_CLIENT_SECRET_APIGITHUBC";
const CLIENT_ID = "on-disk-client-id";
const CLIENT_SECRET = "on-disk-client-secret";

const descriptor: unknown = JSON.parse(
  fs.readFileSync(path.join(PKG_DIR, "descriptor.json"), "utf8")
);
const pipeline: unknown = JSON.parse(fs.readFileSync(path.join(PKG_DIR, "pipeline.json"), "utf8"));

/** Load the package's `content/**` as the opened-entry list (forward-slash paths, raw bytes). */
function loadContent(): TemplateFileEntry[] {
  const root = path.join(PKG_DIR, "content");
  const entries: TemplateFileEntry[] = [];
  const walk = (dir: string): void => {
    for (const name of fs.readdirSync(dir)) {
      const full = path.join(dir, name);
      if (fs.statSync(full).isDirectory()) {
        walk(full);
      } else {
        entries.push({
          path: path.relative(root, full).replace(/\\/g, "/"),
          data: fs.readFileSync(full),
        });
      }
    }
  };
  walk(root);
  return entries;
}

const content = loadContent();

interface RunOptions {
  authType?: string;
  existing?: string[];
}

/** Build one scaffold request rooted at `dir` (the on-disk runtime's output root). */
function makeRequest(dir: string, options: RunOptions = {}): ScaffoldRequest {
  const authType = options.authType ?? "none";
  return {
    descriptor,
    pipeline,
    content,
    answers: {
      mcpServerType: "remote",
      mcpServerUrl: MCP_SERVER_URL,
      authType,
      ...(authType === "oauth"
        ? { oauthClientId: CLIENT_ID, oauthClientSecret: CLIENT_SECRET }
        : {}),
      ...(authType === "entra-sso" ? { entraClientId: CLIENT_ID } : {}),
    },
    callerFloor: { appName: "MyMcpAgent", language: "common" },
    targetDir: { path: dir, existing: options.existing ?? [] },
  };
}

describe("createRealRuntime (v4, on-disk ScaffoldRuntime)", () => {
  beforeEach(() => {
    // The oauth/oauth-dynamic auth step probes the server for metadata; stub the network so the
    // on-disk run stays offline and deterministic. entra-sso/none never probe.
    vi.spyOn(mcpAuthScaffoldDeps, "probeMCPServerAuth").mockResolvedValue({
      requiresAuth: true,
      authMetadataUrl: "https://auth.example.com/.well-known/oauth-protected-resource",
    });
    vi.spyOn(mcpAuthScaffoldDeps, "resolveMCPOAuthMetadata").mockResolvedValue({
      authorizationUrl: "https://auth.example.com/authorize",
      tokenUrl: "https://auth.example.com/token",
      wellKnownUrl: "https://auth.example.com/.well-known/oauth-authorization-server",
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "atk-v4-mcp-"));
  });

  afterEach(() => {
    fs.removeSync(tempDir);
  });

  /** Scaffold the package into `tempDir` against a fresh on-disk runtime. */
  async function run(options: RunOptions = {}) {
    const runtime = createRealRuntime(tempDir);
    return scaffold(makeRequest(tempDir, options), runtime);
  }

  function diskPath(rel: string): string {
    return path.join(tempDir, rel);
  }
  function diskExists(rel: string): boolean {
    return fs.existsSync(diskPath(rel));
  }
  function diskText(rel: string): string {
    return fs.readFileSync(diskPath(rel), "utf8");
  }

  it.each(["late.txt", "late.txt.tpl"])(
    "IO-01: preserves a file created after the snapshot for %s",
    async (entryPath) => {
      const warnings: string[] = [];
      const runtime = createRealRuntime(tempDir, undefined, undefined, (warning) => {
        assert.equal(warning.type, EXISTING_FILE_SKIPPED_WARNING);
        warnings.push(warning.content);
      });
      const targetDir = { path: tempDir, existing: [] };
      fs.writeFileSync(diskPath("late.txt"), "original");
      const outcome = (
        await runScaffoldPipeline(
          { pipeline: "default", steps: [] },
          [{ path: entryPath, data: Buffer.from("replacement") }],
          {},
          targetDir,
          runtime.port
        )
      )._unsafeUnwrap();
      assert.equal(diskText("late.txt"), "original");
      assert.deepEqual(outcome.written, []);
      assert.deepEqual(
        outcome.skipped.map((file) => file.path),
        ["late.txt"]
      );
      assert.deepEqual(
        warnings,
        outcome.skipped.map((file) => file.warning)
      );
    }
  );

  it("IO-01: delegates case-alias collisions to the actual filesystem", async () => {
    fs.writeFileSync(diskPath("CONFIG.json"), "original");
    const aliases = fs.existsSync(diskPath("config.json"));
    if (aliases) {
      assert.equal(
        fs.statSync(diskPath("CONFIG.json")).ino,
        fs.statSync(diskPath("config.json")).ino
      );
    }
    const outcome = (
      await runScaffoldPipeline(
        { pipeline: "default", steps: [] },
        [{ path: "config.json.tpl", data: Buffer.from("new bytes") }],
        {},
        { path: tempDir, existing: ["CONFIG.json"] },
        createRealRuntime(tempDir).port
      )
    )._unsafeUnwrap();
    assert.equal(diskText("CONFIG.json"), "original");
    assert.equal(diskText("config.json"), aliases ? "original" : "new bytes");
    assert.deepEqual(outcome.written, aliases ? [] : ["config.json"]);
    assert.deepEqual(
      outcome.skipped.map((file) => file.path),
      aliases ? ["config.json"] : []
    );
  });

  it("IO-01: exclusive creation preserves bytes while ordinary writes still replace them", () => {
    const { port } = createRealRuntime(tempDir);
    assert.isTrue(port.writeNew("nested/file.txt", Buffer.from("first")));
    assert.isFalse(port.writeNew("nested/file.txt", Buffer.from("second")));
    assert.equal(port.read("nested/file.txt")?.toString(), "first");
    port.write("nested/file.txt", Buffer.from("step"));
    assert.equal(port.read("nested/file.txt")?.toString(), "step");
    assert.isUndefined(port.read("missing/child.txt"));
  });

  it.each(["EACCES", "EIO"])(
    "IO-01: writeNew propagates %s instead of reporting a collision",
    (code) => {
      const failure = Object.assign(new Error("write failed"), { code });
      const { port } = createRealRuntime(tempDir);
      vi.spyOn(fs, "writeFileSync").mockImplementationOnce(() => {
        throw failure;
      });
      expect(() => port.writeNew("file.txt", Buffer.from("bytes"))).toThrow(failure);
    }
  );

  it.each(["EACCES", "ENOTDIR"])("IO-03: propagates lstat %s before any I/O", (code) => {
    const failure = Object.assign(new Error("access denied"), { code });
    const { port } = createRealRuntime(tempDir);
    vi.spyOn(fs, "lstatSync").mockImplementation(() => {
      throw failure;
    });
    for (const operation of [
      () => port.read("file.txt"),
      () => port.write("file.txt", Buffer.from("bytes")),
      () => port.writeNew("file.txt", Buffer.from("bytes")),
    ]) {
      expect(operation).toThrow(failure);
    }
  });

  it("IO-03: rejects lexical escapes", () => {
    const { port } = createRealRuntime(tempDir);
    for (const entryPath of [".", "../outside.txt", "nested/../../outside.txt"]) {
      expect(() => port.read(entryPath)).toThrowError(
        expect.objectContaining({ name: "ScaffoldPathEscape" })
      );
    }
  });

  it.each(["read", "write", "writeNew"])(
    "IO-03: rejects %s through a directory link below the root",
    (operation) => {
      const outside = fs.mkdtempSync(path.join(os.tmpdir(), "atk-v4-outside-"));
      const link = diskPath("linked");
      try {
        fs.writeFileSync(path.join(outside, "keep.txt"), "original");
        fs.symlinkSync(outside, link, "junction");
        const runtime = createRealRuntime(tempDir);
        expect(() => {
          if (operation === "read") {
            runtime.port.read("linked/keep.txt");
          } else if (operation === "write") {
            runtime.port.write("linked/keep.txt", Buffer.from("changed"));
          } else {
            runtime.port.writeNew("linked/new.txt", Buffer.from("changed"));
          }
        }).toThrowError(expect.objectContaining({ name: "ScaffoldPathEscape" }));
        assert.strictEqual(fs.readFileSync(path.join(outside, "keep.txt"), "utf8"), "original");
        assert.isFalse(fs.existsSync(path.join(outside, "new.txt")));
      } finally {
        fs.removeSync(link);
        fs.removeSync(outside);
      }
    }
  );

  it.each(["inward", "dangling"])("IO-03: rejects %s directory links", (kind) => {
    const target = diskPath("nested");
    if (kind === "inward") {
      fs.ensureDirSync(target);
      fs.writeFileSync(path.join(target, "file.txt"), "original");
    }
    const link = diskPath("linked");
    fs.symlinkSync(target, link, "junction");
    try {
      const { port } = createRealRuntime(tempDir);
      for (const operation of [
        () => port.read("linked/file.txt"),
        () => port.write("linked/file.txt", Buffer.from("changed")),
        () => port.writeNew("linked/new.txt", Buffer.from("changed")),
      ]) {
        expect(operation).toThrowError(expect.objectContaining({ name: "ScaffoldPathEscape" }));
      }
      if (kind === "inward") assert.equal(diskText("nested/file.txt"), "original");
    } finally {
      fs.removeSync(link);
    }
  });

  it.for([false, true])("IO-03: rejects final file links, dangling=%s", (dangling, context) => {
    const target = diskPath("target.txt");
    if (!dangling) fs.writeFileSync(target, "original");
    const link = diskPath("linked.txt");
    try {
      fs.symlinkSync(target, link, "file");
    } catch (error) {
      if (
        process.platform === "win32" &&
        error instanceof Error &&
        "code" in error &&
        error.code === "EPERM"
      ) {
        context.skip("Windows file symlinks require Developer Mode or symbolic-link privilege");
        return;
      }
      throw error;
    }
    try {
      const { port } = createRealRuntime(tempDir);
      for (const operation of [
        () => port.read("linked.txt"),
        () => port.write("linked.txt", Buffer.from("changed")),
        () => port.writeNew("linked.txt", Buffer.from("changed")),
      ]) {
        expect(operation).toThrowError(expect.objectContaining({ name: "ScaffoldPathEscape" }));
      }
      if (!dangling) assert.equal(diskText("target.txt"), "original");
    } finally {
      fs.unlinkSync(link);
    }
  });

  it("IO-03: trusts a selected root directory alias", () => {
    const physical = diskPath("physical");
    const alias = diskPath("alias");
    fs.ensureDirSync(physical);
    fs.symlinkSync(physical, alias, "junction");
    try {
      const { port } = createRealRuntime(alias);
      assert.isTrue(port.writeNew("nested/file.txt", Buffer.from("first")));
      port.write("nested/file.txt", Buffer.from("step"));
      assert.equal(port.read("nested/file.txt")?.toString(), "step");
      assert.equal(diskText("physical/nested/file.txt"), "step");
    } finally {
      fs.removeSync(alias);
    }
  });

  it.each(["env", "custom-env"])(
    "IO-03: guards writeEnvironment through %s directory links",
    async (folder) => {
      const outside = fs.mkdtempSync(path.join(os.tmpdir(), "atk-v4-env-"));
      const link = diskPath(folder);
      const yaml = `version: 1.9.0\nenvironmentFolderPath: ./${folder}\n`;
      fs.writeFileSync(diskPath("m365agents.yml"), yaml);
      fs.symlinkSync(outside, link, "junction");
      try {
        await expect(
          createRealRuntime(tempDir).port.writeEnvironment("dev", { CLIENT_ID: "new" })
        ).rejects.toMatchObject({ name: "ScaffoldPathEscape" });
        assert.deepEqual(fs.readdirSync(outside), []);
        assert.equal(diskText("m365agents.yml"), yaml);
      } finally {
        fs.removeSync(link);
        fs.removeSync(outside);
      }
    }
  );

  it.for(["env/.env.dev", "env/.env.dev.user", "m365agents.yml"])(
    "IO-03: guards writeEnvironment against targeted file link %s",
    async (entryPath, context) => {
      const target = diskPath("target.txt");
      const original = entryPath === "m365agents.yml" ? "version: 1.9.0\n" : "VALUE=original\n";
      fs.writeFileSync(target, original);
      fs.ensureDirSync(diskPath("env"));
      if (entryPath !== "m365agents.yml")
        fs.writeFileSync(diskPath("m365agents.yml"), "version: 1.9.0\n");
      const link = diskPath(entryPath);
      try {
        fs.symlinkSync(target, link, "file");
      } catch (error) {
        if (
          process.platform === "win32" &&
          error instanceof Error &&
          "code" in error &&
          error.code === "EPERM"
        ) {
          context.skip("Windows file symlinks require Developer Mode or symbolic-link privilege");
          return;
        }
        throw error;
      }
      try {
        await expect(
          createRealRuntime(tempDir).port.writeEnvironment("dev", {
            CLIENT_ID: "new",
            SECRET_TOKEN: "test-secret",
          })
        ).rejects.toMatchObject({ name: "ScaffoldPathEscape" });
        assert.equal(diskText("target.txt"), original);
      } finally {
        fs.unlinkSync(link);
      }
    }
  );

  it.each(["env/.env.dev", "env/.env.dev.user", "custom-env"])(
    "IO-03: rejects final env junctions and a dangling custom env directory: %s",
    async (entryPath) => {
      const link = diskPath(entryPath);
      const destination = diskPath("destination");
      fs.ensureDirSync(path.dirname(link));
      if (entryPath !== "custom-env") fs.ensureDirSync(destination);
      const yaml = `version: 1.9.0\nenvironmentFolderPath: ./${entryPath === "custom-env" ? "custom-env" : "env"}\n`;
      fs.writeFileSync(diskPath("m365agents.yml"), yaml);
      fs.symlinkSync(destination, link, "junction");
      try {
        await expect(
          createRealRuntime(tempDir).port.writeEnvironment("dev", { CLIENT_ID: "new" })
        ).rejects.toMatchObject({ name: "ScaffoldPathEscape" });
        assert.equal(diskText("m365agents.yml"), yaml);
      } finally {
        fs.removeSync(link);
      }
    }
  );

  it("ON-DISK-01: materializes the package onto a real directory, `.tpl` stripped", async () => {
    const result = await run();
    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");

    for (const rel of [
      "appPackage/ai-plugin.json",
      "appPackage/declarativeAgent.json",
      "appPackage/manifest.json",
      "appPackage/color.png",
      "appPackage/outline.png",
      "appPackage/instruction.txt",
      "m365agents.yml",
      ".vscode/mcp.json",
      ".vscode/extensions.json",
      "env/.env.dev",
      ".gitignore",
      "README.md",
    ]) {
      assert.isTrue(diskExists(rel), `expected '${rel}' on disk`);
    }

    // The render phase strips the `.tpl` suffix — no template artifacts remain.
    assert.isFalse(diskExists("m365agents.yml.tpl"), "m365agents.yml.tpl must not remain");
    assert.isFalse(
      diskExists("appPackage/ai-plugin.json.tpl"),
      "ai-plugin.json.tpl must not remain"
    );
    assert.isFalse(diskExists(".gitignore.tpl"), ".gitignore.tpl must not remain");

    // The rendered bytes on disk are valid JSON with the derived namespace.
    const aiPlugin: { namespace?: string } = JSON.parse(diskText("appPackage/ai-plugin.json"));
    assert.strictEqual(aiPlugin.namespace, NAMESPACE);
  });

  it("ON-DISK-02: an oauth run persists regular credentials and an encrypted user-env secret", async () => {
    const result = await run({ authType: "oauth" });
    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");

    // The inject step read the rendered yml back from disk, mutated it, rewrote it.
    const yml = diskText("m365agents.yml");
    assert.include(yml, "uses: oauth/register");
    assert.include(yml, `name: ${NAMESPACE}`);

    const env = diskText("env/.env.dev");
    assert.include(env, `${AUTH_ENV_VAR}=`);
    assert.include(env, `${CLIENT_ID_ENV_VAR}=${CLIENT_ID}`);
    assert.notInclude(env, CLIENT_SECRET);

    const userEnv = diskText("env/.env.dev.user");
    assert.include(userEnv, `${CLIENT_SECRET_ENV_VAR}=crypto_`);
    assert.notInclude(userEnv, CLIENT_SECRET);
    assert.include(yml, "projectId:");
  });

  it("ON-DISK-03: an entra-sso run injects oauth/register (Entra) and persists the env var", async () => {
    const result = await run({ authType: "entra-sso" });
    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");

    assert.include(diskText("m365agents.yml"), "uses: oauth/register");
    assert.include(diskText("m365agents.yml"), "identityProvider: MicrosoftEntra");
    assert.include(diskText("env/.env.dev"), `${AUTH_ENV_VAR}=`);
  });

  it("ON-DISK-04: a none run writes no auth action and no credential var", async () => {
    const result = await run();
    assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");

    const yml = diskText("m365agents.yml");
    assert.notInclude(yml, "oauth/register");
    assert.notInclude(yml, "microsoftEntra/register");
    assert.notInclude(diskText("env/.env.dev"), AUTH_ENV_VAR);
  });

  it("ON-DISK-05: a non-empty target fails the create contract and writes nothing", async () => {
    fs.writeFileSync(path.join(tempDir, "existing.txt"), "keep me");

    const result = await run({ existing: ["existing.txt"] });

    assert.isTrue(result.isErr(), "expected the create-empty contract to fail");
    const error = result.isErr() ? result.error : undefined;
    assert.instanceOf(error, UserError);
    assert.strictEqual(error?.name, REQUIRE_EMPTY_TARGET);

    // The render phase never ran: none of our files landed, the pre-existing file is untouched.
    assert.isFalse(diskExists("appPackage/ai-plugin.json"));
    assert.isFalse(diskExists("m365agents.yml"));
    assert.strictEqual(diskText("existing.txt"), "keep me");
  });

  it("ON-DISK-06: re-running into a fresh directory is byte-identical (deterministic)", async () => {
    const first = await run();
    assert.isTrue(first.isOk(), first.isErr() ? first.error.message : "expected ok");
    const firstBytes = fs.readFileSync(diskPath("appPackage/ai-plugin.json"));

    const secondDir = fs.mkdtempSync(path.join(os.tmpdir(), "atk-v4-mcp2-"));
    try {
      const second = await scaffold(makeRequest(secondDir), createRealRuntime(secondDir));
      assert.isTrue(second.isOk(), second.isErr() ? second.error.message : "expected ok");
      const secondBytes = fs.readFileSync(path.join(secondDir, "appPackage/ai-plugin.json"));
      assert.isTrue(
        firstBytes.equals(secondBytes),
        "ai-plugin.json must be byte-identical across runs"
      );
    } finally {
      fs.removeSync(secondDir);
    }
  });
});
