// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { createHash, randomUUID } from "crypto";
import fs from "fs-extra";
import os from "os";
import path from "path";
import properLock from "proper-lockfile";
import type { Result } from "neverthrow";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  getPreparedSharePointProductId,
  createPreparedPublicationErrorData,
  isPreparedCatalogUrl,
  isAmbiguousPublicationState,
  loadPreparedPublication,
  preparePublication,
  PreparedPublicationJournal,
  updatePreparedPublication,
  verifyPreparedApplicationPackage,
  verifyPreparedArtifact,
  verifyPreparedSources,
  verifyPreparedSharePointPackage,
  withPreparedPublicationIdentityLock,
} from "../../../../../src/component/driver/deploy/spfx/utility/preparedPublication";
import { SHAREPOINT_PACKAGE_LIMITS } from "../../../../../src/component/driver/deploy/spfx/utility/packageFile";
import { assembleSharePointApplicationPackage } from "../../../../../src/sharepointPackage";
import {
  createAppPackage,
  createSppkg,
  publicationTenantContext,
} from "./preparedPublicationFixtures";

function expectError(result: Result<unknown, { code: string }>, code: string): void {
  expect(result.isErr()).toBe(true);
  if (result.isErr()) expect(result.error.code).toBe(code);
}

describe("Prepared publication validation and durable recovery", () => {
  let root: string;
  const initialLimits = { ...SHAREPOINT_PACKAGE_LIMITS };

  beforeEach(async () => {
    root = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), "prepared-recovery-")));
    await fs.outputFile(path.join(root, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.outputFile(path.join(root, "appPackage", "application.zip"), createAppPackage());
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    Object.assign(SHAREPOINT_PACKAGE_LIMITS, initialLimits);
    await fs.remove(root);
  });

  function input() {
    return {
      projectPath: root,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    };
  }

  async function prepared() {
    const result = await preparePublication(input());
    if (result.isErr()) throw result.error;
    const loaded = await loadPreparedPublication(
      root,
      result.value.publicationId,
      result.value.preparedPublicationPath
    );
    if (loaded.isErr()) throw loaded.error;
    return { output: result.value, ...loaded.value };
  }

  async function lineageIds(journalPath: string): Promise<string[]> {
    return (await fs.readdir(path.dirname(path.dirname(journalPath)), { withFileTypes: true }))
      .filter((entry) => entry.isDirectory() && /^[0-9a-f-]{36}$/i.test(entry.name))
      .map((entry) => entry.name);
  }

  async function changeJson(
    part: string,
    change: (object: Record<string, unknown>) => void
  ): Promise<void> {
    const zip = new AdmZip(await fs.readFile(path.join(root, "appPackage", "application.zip")));
    const object = JSON.parse(zip.readAsText(part));
    change(object);
    zip.updateFile(part, Buffer.from(JSON.stringify(object)));
    await fs.writeFile(path.join(root, "appPackage", "application.zip"), zip.toBuffer());
  }

  async function changePlugin(change: (object: Record<string, unknown>) => void): Promise<void> {
    const zip = new AdmZip(await fs.readFile(path.join(root, "sharepoint", "solution.sppkg")));
    const bytes = zip.readFile("ClientSideAssets/agent.zip");
    if (!bytes) throw new Error("Missing fixture embedded package.");
    const nested = new AdmZip(bytes);
    const plugin = JSON.parse(nested.readAsText("ai-plugin.json"));
    change(plugin);
    nested.updateFile("ai-plugin.json", Buffer.from(JSON.stringify(plugin)));
    zip.updateFile("ClientSideAssets/agent.zip", nested.toBuffer());
    await fs.writeFile(path.join(root, "sharepoint", "solution.sppkg"), zip.toBuffer());
  }

  it.each(["../solution.sppkg", "/absolute/solution.sppkg"])(
    "SPFX-PREP-04: rejects an escaping or absolute input (%s) before staging",
    async (packagePath) => {
      expectError(await preparePublication({ ...input(), packagePath }), "InvalidInputPath");
      expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
    }
  );

  it("SPFX-PREP-04: refuses input, source and artifact junctions that resolve outside the project", async () => {
    const outside = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-outside-"));
    try {
      await fs.writeFile(path.join(outside, "solution.sppkg"), createSppkg());
      const current = await prepared();
      await fs.ensureSymlink(outside, path.join(root, "escape"), "junction");
      expectError(
        await preparePublication({ ...input(), packagePath: "escape/solution.sppkg" }),
        "InvalidInputPath"
      );
      expectError(
        await verifyPreparedArtifact(root, {
          ...current.journal.sharePointArtifact,
          path: "escape/solution.sppkg",
        }),
        "InvalidArtifactPath"
      );
      current.journal.sourceSharePoint.path = "escape/solution.sppkg";
      expectError(await verifyPreparedSources(root, current.journal), "InvalidSourcePath");
    } finally {
      await fs.remove(outside);
    }
  });

  it("SPFX-PREP-04: refuses a journal directory redirected through a junction", async () => {
    const current = await prepared();
    const outside = await fs.mkdtemp(path.join(os.tmpdir(), "journal-outside-"));
    try {
      await fs.writeJson(path.join(outside, "journal.json"), current.journal);
      const directory = path.dirname(current.path);
      await fs.remove(directory);
      await fs.ensureSymlink(outside, directory, "junction");
      expectError(
        await loadPreparedPublication(
          root,
          current.output.publicationId,
          current.output.preparedPublicationPath
        ),
        "InvalidJournalPath"
      );
    } finally {
      await fs.remove(outside);
    }
  });

  it("SPFX-PREP-04: refuses missing inputs and an oversized input before parsing", async () => {
    expectError(
      await preparePublication({ ...input(), appPackagePath: "missing.zip" }),
      "InputReadFailed"
    );
    SHAREPOINT_PACKAGE_LIMITS.inputBytes = 1;
    const close = vi.spyOn(fs, "close");
    expectError(await getPreparedSharePointProductId(root, input().packagePath), "InputTooLarge");
    expect(close).toHaveBeenCalledOnce();
    expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
  });

  it("SPFX-PREP-04: rejects an unsupported source filename without inspecting its content", async () => {
    const name = "sharepoint/solution with spaces.sppkg";
    await fs.copy(path.join(root, input().packagePath), path.join(root, name));
    expectError(
      await preparePublication({ ...input(), packagePath: name }),
      "UnsupportedSharePointPackageFileName"
    );
  });

  it("SPFX-PREP-01: product discovery reports project, input and inspection failures distinctly", async () => {
    expectError(
      await getPreparedSharePointProductId(path.join(root, "missing"), input().packagePath),
      "ProjectPathUnavailable"
    );
    expectError(await getPreparedSharePointProductId(root, "missing.sppkg"), "InputReadFailed");
    await fs.writeFile(path.join(root, input().packagePath), Buffer.from("invalid"));
    expectError(await getPreparedSharePointProductId(root, input().packagePath), "InvalidArchive");
    expectError(await preparePublication(input()), "InvalidArchive");
    expectError(
      await preparePublication({ ...input(), projectPath: path.join(root, "missing") }),
      "ProjectPathUnavailable"
    );
  });

  it.each([
    {
      name: "unparsable catalog URL",
      context: { catalogUrl: "invalid" },
      code: "InvalidPreparedTenantContext",
    },
    {
      name: "unparsable tenant URL",
      context: { tenantUrl: "invalid" },
      code: "InvalidPreparedTenantContext",
    },
    {
      name: "HTTP tenant",
      context: { tenantUrl: "http://tenant.example.test" },
      code: "InvalidPreparedTenantContext",
    },
    {
      name: "catalog credentials",
      context: { catalogUrl: "https://user:password@tenant.example.test/sites/appcatalog" },
      code: "InvalidPreparedTenantContext",
    },
    {
      name: "tenant query",
      context: { tenantUrl: "https://tenant.example.test?unexpected=true" },
      code: "InvalidPreparedTenantContext",
    },
    {
      name: "foreign catalog origin",
      context: { catalogUrl: "https://foreign.example.test/sites/appcatalog" },
      code: "InvalidPreparedTenantContext",
    },
    {
      name: "malformed tenant ID",
      context: { tenantId: "invalid" },
      code: "InvalidPreparedTenantContext",
    },
    {
      name: "different selected product",
      context: { productId: randomUUID() },
      code: "SelectedProductMismatch",
    },
  ])("SPFX-PREP-05: refuses $name before freezing a lineage", async ({ context, code }) => {
    expectError(
      await preparePublication({
        ...input(),
        tenantContext: { ...publicationTenantContext, ...context },
      }),
      code
    );
    expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
  });

  it.each([
    {
      name: "missing runtime spec",
      runtime: { type: "RemoteMCPServer" },
      code: "AmbiguousPluginEndpoint",
    },
    {
      name: "non-string endpoint",
      runtime: { type: "RemoteMCPServer", spec: { url: 12 } },
      code: "AmbiguousPluginEndpoint",
    },
    {
      name: "foreign endpoint",
      runtime: { type: "RemoteMCPServer", spec: { url: "https://foreign.test/mcp" } },
      code: "AmbiguousPluginEndpoint",
    },
    {
      name: "unregistered placeholder",
      runtime: { type: "RemoteMCPServer", spec: { url: "{{OTHER_URL}}" } },
      code: "UnresolvedPluginPlaceholder",
    },
  ])(
    "SPFX-PREP-05: refuses $name instead of performing broad text substitution",
    async ({ runtime, code }) => {
      await changePlugin((plugin) => {
        plugin.runtimes = [runtime];
      });
      expectError(await preparePublication(input()), code);
      expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
    }
  );

  describe("SPFX-PREP-24: tenant values in generated plugin fields", () => {
    const tenantOrigin = "https://tenant.example.test";
    const sha256 = (bytes: Buffer) => createHash("sha256").update(bytes).digest("hex");
    const uiTool = (name: string, frameDomains: string[], resourceDomains: string[]) => ({
      name,
      _meta: { "ui/csp": { frameDomains, resourceDomains } },
    });
    const runtimeWith = (tools: unknown[], auth: unknown = { type: "None" }) => ({
      type: "RemoteMCPServer",
      auth,
      spec: { url: "{{TENANT_MCP_URL}}", mcp_tool_description: { tools } },
    });

    async function withTools(tools: unknown[], auth?: unknown): Promise<void> {
      await changePlugin((plugin) => {
        plugin.runtimes = [runtimeWith(tools, auth)];
      });
    }

    async function changePluginText(change: (text: string) => string): Promise<void> {
      const sppkgPath = path.join(root, "sharepoint/solution.sppkg");
      const zip = new AdmZip(await fs.readFile(sppkgPath));
      const embedded = zip.readFile("ClientSideAssets/agent.zip");
      if (!embedded) {
        throw new Error("Missing fixture embedded package.");
      }
      const nested = new AdmZip(embedded);
      nested.updateFile(
        "ai-plugin.json",
        Buffer.from(change(nested.readAsText("ai-plugin.json")), "utf8")
      );
      zip.updateFile("ClientSideAssets/agent.zip", nested.toBuffer());
      await fs.writeFile(sppkgPath, zip.toBuffer());
    }

    async function stagedPlugin(appPackagePath: string) {
      const zip = new AdmZip(await fs.readFile(path.join(root, appPackagePath)));
      const entry = zip
        .getEntries()
        .find((candidate) => /^plugins\/spfx-api-[0-9a-f]{16}\.json$/.test(candidate.entryName));
      if (!entry) {
        throw new Error("Missing staged plugin.");
      }
      const bytes = entry.getData();
      return { bytes, plugin: JSON.parse(bytes.toString("utf8")) };
    }

    it("resolves approved UI origins to the verified tenant origin and nothing else", async () => {
      await withTools([
        uiTool("first", ["{{TENANT_ORIGIN}}"], ["{{ TENANT_ORIGIN }}", "https://cdn.example.test"]),
        uiTool(
          "second",
          ["https://static.example.test", "{{TENANT_ORIGIN}}"],
          ["{{TENANT_ORIGIN}}"]
        ),
      ]);
      const sourceBefore = await fs.readFile(path.join(root, "sharepoint/solution.sppkg"));

      const current = await prepared();
      const staged = await stagedPlugin(current.output.appPackagePath);

      const [first, second] = staged.plugin.runtimes[0].spec.mcp_tool_description.tools;
      expect(first._meta["ui/csp"]).toEqual({
        frameDomains: [tenantOrigin],
        resourceDomains: [tenantOrigin, "https://cdn.example.test"],
      });
      expect(second._meta["ui/csp"]).toEqual({
        frameDomains: ["https://static.example.test", tenantOrigin],
        resourceDomains: [tenantOrigin],
      });
      expect(staged.plugin.runtimes[0].spec.url).toBe(
        `${tenantOrigin}/_api/mcp/spfx?solutionId=${publicationTenantContext.productId}`
      );
      expect(staged.bytes.toString("utf8")).not.toContain("{{");
      expect(sha256(staged.bytes)).toBe(current.journal.plugin.sha256);
      const sourceAfter = await fs.readFile(path.join(root, "sharepoint/solution.sppkg"));
      expect(sourceAfter.equals(sourceBefore)).toBe(true);
    });

    it("preserves partial tokens inside free text", async () => {
      const description = "Documentation mentions {{TENANT_ORIGIN}} and {{OTHER_VALUE}} literally.";
      await withTools([{ name: "t", description }]);

      const current = await prepared();
      const staged = await stagedPlugin(current.output.appPackagePath);

      expect(staged.plugin.runtimes[0].spec.mcp_tool_description.tools[0].description).toBe(
        description
      );
    });

    it.each([
      {
        name: "connectDomains",
        tool: { name: "t", _meta: { "ui/csp": { connectDomains: ["{{TENANT_ORIGIN}}"] } } },
      },
      {
        name: "a nested ui.csp object",
        tool: { name: "t", _meta: { ui: { csp: { frameDomains: ["{{TENANT_ORIGIN}}"] } } } },
      },
      { name: "a tool description", tool: { name: "t", description: "{{TENANT_ORIGIN}}" } },
      { name: "an unregistered array", tool: { name: "t", tags: ["{{TENANT_ORIGIN}}"] } },
    ])("refuses the tenant origin token in $name before staging", async ({ tool }) => {
      await withTools([tool]);

      expectError(await preparePublication(input()), "UnqualifiedTenantOriginField");
      expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
    });

    it("refuses the unresolved managed auth reference before staging", async () => {
      await withTools([{ name: "t" }], {
        type: "OAuthPluginVault",
        reference_id: "{{MCP_DA_AUTH_REFERENCE_ID}}",
      });

      expectError(await preparePublication(input()), "UnresolvedPluginPlaceholder");
      expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
    });

    it.each([
      {
        name: "an unregistered whole-value string",
        tool: { name: "t", description: "{{OTHER_VALUE}}" },
      },
      {
        name: "an unregistered whole-value array element",
        tool: { name: "t", tags: ["{{OTHER_VALUE}}"] },
      },
      {
        name: "an unknown token in an approved UI field",
        tool: uiTool("t", ["{{OTHER_VALUE}}"], []),
      },
      {
        name: "a token embedded in an approved UI value",
        tool: uiTool("t", ["https://{{TENANT_ORIGIN}}"], []),
      },
    ])("refuses $name before staging", async ({ tool }) => {
      await withTools([tool]);

      expectError(await preparePublication(input()), "UnresolvedPluginPlaceholder");
      expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
    });

    it("refuses an excessively nested plugin with a typed error", async () => {
      const depth = 10_000;
      await changePluginText((text) => {
        const body = text.trimEnd();
        // Distinct keys keep the entry under the archive compression-ratio guard.
        let opening = "";
        for (let level = 0; level < depth; level++) {
          opening += `{"k${level}":`;
        }
        return `${body.slice(0, -1)},"extra":${opening}1${"}".repeat(depth)}}`;
      });

      expectError(await preparePublication(input()), "InvalidPluginManifest");
      expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
    });

    it("journals new preparations with the tenant-value transformation identity", async () => {
      const current = await prepared();

      expect(current.journal.transformation.id).toBe("spfx-copilot-plugin-v5");
    });

    it("refuses to resume an unsubmitted lineage prepared before tenant values were resolved", async () => {
      const current = await prepared();
      await fs.writeJson(current.path, {
        ...current.journal,
        transformation: { id: "spfx-copilot-plugin-v4" },
      });

      expectError(await preparePublication(input()), "UnresolvedPreparedPublication");
      expect(await lineageIds(current.path)).toHaveLength(1);
    });

    it("re-prepares an already prepared application package for another tenant", async () => {
      await withTools([uiTool("t", ["{{TENANT_ORIGIN}}"], ["{{TENANT_ORIGIN}}"])]);
      const first = await prepared();
      const stagedBytes = await fs.readFile(path.join(root, first.output.appPackagePath));
      await fs.outputFile(path.join(root, "appPackage/application.zip"), stagedBytes);
      const otherOrigin = "https://other.example.test";

      const second = await preparePublication({
        ...input(),
        tenantContext: {
          tenantId: "dddddddd-eeee-4fff-8000-111111111111",
          tenantUrl: otherOrigin,
          catalogUrl: `${otherOrigin}/sites/appcatalog`,
          productId: publicationTenantContext.productId,
        },
      });

      expect(second.isOk()).toBe(true);
      if (second.isErr()) {
        return;
      }
      const staged = await stagedPlugin(second.value.appPackagePath);
      const [tool] = staged.plugin.runtimes[0].spec.mcp_tool_description.tools;
      expect(tool._meta["ui/csp"]).toEqual({
        frameDomains: [otherOrigin],
        resourceDomains: [otherOrigin],
      });
      expect(staged.bytes.toString("utf8")).not.toContain(tenantOrigin);
    });
  });

  it.each([undefined, 30, "1.99"])(
    "SPFX-PREP-05: refuses an unsupported Teams schema version (%s) before staging",
    async (manifestVersion) => {
      await changeJson("manifest.json", (manifest) => {
        manifest.manifestVersion = manifestVersion;
      });
      expectError(await preparePublication(input()), "UnsupportedTeamsManifestVersion");
      expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
    }
  );

  it.each([undefined, 9, "v9.0"])(
    "SPFX-PREP-05: refuses an unsupported agent schema version (%s)",
    async (version) => {
      await changeJson("declarativeAgent.json", (agent) => {
        agent.version = version;
      });
      expectError(await preparePublication(input()), "UnsupportedAgentManifestVersion");
    }
  );

  it.each(["app", "agent", "plugin"])(
    "SPFX-PREP-05: validates final %s schema constraints before staging",
    async (surface) => {
      if (surface === "app") {
        await changeJson("manifest.json", (manifest) => {
          manifest.name = { short: "x".repeat(200) };
        });
      } else if (surface === "agent") {
        await changeJson("declarativeAgent.json", (agent) => {
          agent.name = "x".repeat(2000);
        });
      } else {
        await changePlugin((plugin) => {
          plugin.schema_version = "v9.0";
        });
      }
      expectError(await preparePublication(input()), "ApplicationSchemaValidationFailed");
      expect(await fs.pathExists(path.join(root, ".fx"))).toBe(false);
    }
  );

  it.each([
    { name: "escaping reference", file: "../plugin.json", content: undefined },
    { name: "missing file", file: "plugins/absent.json", content: undefined },
    { name: "malformed JSON", file: "plugins/extra.json", content: "{" },
    { name: "non-object JSON", file: "plugins/extra.json", content: "null" },
    { name: "missing schema version", file: "plugins/extra.json", content: "{}" },
  ])(
    "SPFX-PREP-05: does not freeze a package with a $name in an unrelated action",
    async ({ file, content }) => {
      await changeJson("declarativeAgent.json", (agent) => {
        agent.actions = [{ id: "extra", file }];
      });
      if (content !== undefined) {
        const zip = new AdmZip(await fs.readFile(path.join(root, input().appPackagePath)));
        zip.addFile(file, Buffer.from(content));
        await fs.writeFile(path.join(root, input().appPackagePath), zip.toBuffer());
      }
      expectError(await preparePublication(input()), "ApplicationSchemaValidationFailed");
    }
  );

  it.each(["unknown", "submitted"])(
    "SPFX-PREP-08: retains a version-1 %s journal and applies its explicit compatibility rule",
    async (state) => {
      const id = randomUUID();
      const legacyPath = path.join(root, ".fx", "sharepoint-publications", id, "journal.json");
      const legacy = { journalVersion: 1, state };
      await fs.outputJson(legacyPath, legacy);
      const result = await preparePublication(input());
      if (state === "unknown")
        expectError(result, "LegacyPreparedPublicationRequiresReconciliation");
      else expect(result.isErr() ? result.error : undefined).toBeUndefined();
      expect(await fs.readJson(legacyPath)).toEqual(legacy);
    }
  );

  it.each(["missing", "malformed", "invalid"])(
    "SPFX-PREP-08: refuses new work when an existing journal is %s",
    async (condition) => {
      const directory = path.join(root, ".fx", "sharepoint-publications", randomUUID());
      await fs.ensureDir(directory);
      if (condition !== "missing") {
        await fs.writeFile(
          path.join(directory, "journal.json"),
          condition === "malformed" ? "{" : "{}"
        );
      }
      expectError(await preparePublication(input()), "PublicationRecoveryCheckFailed");
      expect(await fs.pathExists(directory)).toBe(true);
    }
  );

  it("SPFX-PREP-08: preserves a recovery enumeration error instead of assuming no active publication", async () => {
    await prepared();
    const original = fs.readdir.bind(fs);
    vi.spyOn(fs, "readdir").mockImplementation(async (directory, options) => {
      if (directory === path.join(root, ".fx", "sharepoint-publications")) {
        throw Object.assign(new Error("Access denied"), { code: "EACCES" });
      }
      return original(directory, options);
    });
    expectError(await preparePublication(input()), "PublicationRecoveryCheckFailed");
  });

  it.each(["sharepoint.sppkg", "application.zip", "journal.json"])(
    "SPFX-PREP-08: cleans incomplete staging after an atomic %s commit fails, without changing sources",
    async (fileName) => {
      const source = await fs.readFile(path.join(root, input().packagePath));
      const application = await fs.readFile(path.join(root, input().appPackagePath));
      const rename = fs.rename.bind(fs);
      vi.spyOn(fs, "rename").mockImplementation(async (from, to) => {
        if (path.basename(String(to)) === fileName) throw new Error("Simulated storage failure.");
        await rename(from, to);
      });
      expectError(await preparePublication(input()), "PublicationWriteFailed");
      const entries = await fs.readdir(path.join(root, ".fx", "sharepoint-publications"));
      expect(entries.every((entry) => entry.startsWith("_identity-"))).toBe(true);
      expect(await fs.readFile(path.join(root, input().packagePath))).toEqual(source);
      expect(await fs.readFile(path.join(root, input().appPackagePath))).toEqual(application);
    }
  );

  it("SPFX-PREP-08: preserves a staging-directory failure as a preparation error", async () => {
    const ensure = fs.ensureDir.bind(fs);
    vi.spyOn(fs, "ensureDir").mockImplementation(async (directory) => {
      if (/^[0-9a-f-]{36}$/i.test(path.basename(String(directory)))) throw new Error("No space.");
      return ensure(directory);
    });
    expectError(await preparePublication(input()), "PublicationPrepareFailed");
  });

  it.each(["lock acquired", "SharePoint artifact committed"])(
    "SPFX-PREP-08: cancellation after %s leaves no committed lineage",
    async (stage) => {
      const controller = new AbortController();
      if (stage === "lock acquired") {
        const acquire = properLock.lock.bind(properLock);
        vi.spyOn(properLock, "lock").mockImplementation(async (file, options) => {
          const release = await acquire(file, options);
          controller.abort();
          return release;
        });
      } else {
        const rename = fs.rename.bind(fs);
        vi.spyOn(fs, "rename").mockImplementation(async (from, to) => {
          await rename(from, to);
          if (path.basename(String(to)) === "sharepoint.sppkg") controller.abort();
        });
      }
      expectError(
        await preparePublication({ ...input(), signal: controller.signal }),
        "PublicationCancelled"
      );
      const entries = await fs.readdir(path.join(root, ".fx", "sharepoint-publications"));
      expect(entries.every((entry) => entry.startsWith("_identity-"))).toBe(true);
    }
  );

  it("SPFX-PREP-08: reports a lock-release failure after preserving the committed journal", async () => {
    const acquire = properLock.lock.bind(properLock);
    vi.spyOn(properLock, "lock").mockImplementation(async (file, options) => {
      const release = await acquire(file, options);
      return async () => {
        await release();
        throw new Error("Release failed.");
      };
    });
    expectError(await preparePublication(input()), "PublicationIdentityLockReleaseFailed");
    const records = (await fs.readdir(path.join(root, ".fx", "sharepoint-publications"))).filter(
      (entry) => !entry.startsWith("_")
    );
    expect(records).toHaveLength(1);
    expect(
      (
        await fs.readJson(
          path.join(root, ".fx", "sharepoint-publications", records[0], "journal.json")
        )
      ).state
    ).toBe("prepared");
  });

  it.each(["sharePoint", "application"])(
    "SPFX-PREP-08: re-preparation detects a deleted %s artifact instead of returning stale outputs",
    async (kind) => {
      const current = await prepared();
      const artifact =
        kind === "sharePoint"
          ? current.journal.sharePointArtifact
          : current.journal.applicationArtifact;
      await fs.remove(path.join(root, artifact.path));
      expectError(await preparePublication(input()), "ArtifactReadFailed");
      expect(await fs.readJson(current.path)).toEqual(current.journal);
    }
  );

  it.each([
    { name: "foreign app identity", section: "identity", key: "appId", value: randomUUID() },
    {
      name: "foreign product identity",
      section: "identity",
      key: "productId",
      value: randomUUID(),
    },
    {
      name: "foreign catalog",
      section: "identity",
      key: "catalogUrl",
      value: "https://foreign.test/catalog",
    },
    { name: "future transform", section: "transformation", key: "id", value: "future" },
    { name: "malformed plugin binding", section: "plugin", key: "bindingSha256", value: "invalid" },
    { name: "non-string plugin binding", section: "plugin", key: "bindingSha256", value: 123 },
    {
      name: "invalid upload state",
      section: "sharePointArtifact",
      key: "uploadState",
      value: "invalid",
    },
    {
      name: "invalid submission state",
      section: "applicationArtifact",
      key: "submissionState",
      value: "invalid",
    },
    {
      name: "non-record submission",
      section: "applicationArtifact",
      key: "submission",
      value: "x",
    },
    {
      name: "submission without a scope",
      section: "applicationArtifact",
      key: "submission",
      value: { titleId: "title", appId: "app" },
    },
    {
      name: "submission with an empty scope",
      section: "applicationArtifact",
      key: "submission",
      value: { scope: "" },
    },
    {
      name: "submission with an empty title identifier",
      section: "applicationArtifact",
      key: "submission",
      value: { scope: "Tenant", titleId: "" },
    },
    {
      name: "submission with a non-string app identifier",
      section: "applicationArtifact",
      key: "submission",
      value: { scope: "Tenant", appId: 7 },
    },
  ])(
    "SPFX-PREP-08: refuses a journal with $name without rewriting it",
    async ({ section, key, value }) => {
      const current = await prepared();
      const raw = await fs.readJson(current.path);
      raw[section][key] = value;
      await fs.writeJson(current.path, raw);
      expectError(
        await loadPreparedPublication(
          root,
          current.output.publicationId,
          current.output.preparedPublicationPath
        ),
        "InvalidJournal"
      );
      expect(await fs.readJson(current.path)).toEqual(raw);
    }
  );

  it("SPFX-PREP-08: loads the previous transformation without a binding hash and keeps replay guarded", async () => {
    const current = await prepared();
    current.journal.transformation.id = "spfx-copilot-plugin-v2";
    delete current.journal.plugin.bindingSha256;
    await fs.writeJson(current.path, current.journal);
    const loaded = await loadPreparedPublication(
      root,
      current.output.publicationId,
      current.output.preparedPublicationPath
    );
    expect(loaded.isErr() ? loaded.error : loaded.value.journal).toEqual(current.journal);
    expectError(await preparePublication(input()), "UnresolvedPreparedPublication");
    expect(await fs.readJson(current.path)).toEqual(current.journal);
  });

  it.each(["spfx-copilot-plugin-v3", "spfx-copilot-plugin-v4"])(
    "SPFX-PREP-08: reuses a completed %s lineage without creating another publication",
    async (transformationId) => {
      const current = await prepared();
      current.journal.transformation.id = transformationId;
      current.journal.state = "submitted";
      current.journal.sharePointArtifact.uploadState = "confirmed";
      current.journal.sharePointArtifact.deployState = "confirmed";
      current.journal.applicationArtifact.submissionState = "confirmed";
      current.journal.applicationArtifact.submission = {
        scope: "Tenant",
        titleId: "published-title",
        appId: "published-app",
      };
      await fs.writeJson(current.path, current.journal);

      const replay = await preparePublication(input());

      expect(replay.isErr() ? replay.error : undefined).toBeUndefined();
      if (replay.isOk()) {
        expect(replay.value.publicationId).toBe(current.output.publicationId);
        expect(replay.value.appPackageSha256).toBe(current.output.appPackageSha256);
      }
      expect(await fs.readJson(current.path)).toEqual(current.journal);
      expect(await lineageIds(current.path)).toEqual([current.output.publicationId]);
    }
  );

  describe("source application identity", () => {
    const rebuild = { entryTime: new Date(2021, 5, 15, 12, 30, 40), reverseOrder: true };
    const applicationPath = () => path.join(root, input().appPackagePath);
    const rawSha256 = (bytes: Buffer) => createHash("sha256").update(bytes).digest("hex");

    async function writeRebuiltApplication(version = "2.3.4"): Promise<void> {
      const rebuilt = createAppPackage(version, rebuild);
      if (version === "2.3.4") {
        expect(rawSha256(rebuilt)).not.toBe(rawSha256(await fs.readFile(applicationPath())));
      }
      await fs.writeFile(applicationPath(), rebuilt);
    }

    async function submitted() {
      const current = await prepared();
      current.journal.state = "submitted";
      current.journal.sharePointArtifact.uploadState = "confirmed";
      current.journal.sharePointArtifact.deployState = "confirmed";
      current.journal.applicationArtifact.submissionState = "confirmed";
      current.journal.applicationArtifact.submission = {
        scope: "Tenant",
        titleId: "published-title",
        appId: "published-app",
      };
      await fs.writeJson(current.path, current.journal);
      return current;
    }

    it("SPFX-PREP-21: resumes the same lineage after the application package is rebuilt with identical content", async () => {
      const current = await prepared();
      await writeRebuiltApplication();

      const resumed = await preparePublication(input());

      expect(resumed.isErr() ? resumed.error : undefined).toBeUndefined();
      if (resumed.isOk()) {
        expect(resumed.value.publicationId).toBe(current.output.publicationId);
        expect(resumed.value.sourceAppPackageSha256).toBe(current.output.sourceAppPackageSha256);
        expect(resumed.value.appPackageSha256).toBe(current.output.appPackageSha256);
      }
      expect(await fs.readJson(current.path)).toEqual(current.journal);
      expect(await lineageIds(current.path)).toEqual([current.output.publicationId]);
    });

    it("SPFX-PREP-21: keeps blocking an unresolved lineage when the application content changed", async () => {
      const current = await prepared();
      await writeRebuiltApplication("2.3.5");

      expectError(await preparePublication(input()), "UnresolvedPreparedPublication");
      expect(await fs.readJson(current.path)).toEqual(current.journal);
      expect(await lineageIds(current.path)).toEqual([current.output.publicationId]);
    });

    it("SPFX-PREP-23: reports the same source application digest as the offline assembler", async () => {
      const current = await prepared();
      const offline = await assembleSharePointApplicationPackage({
        applicationPackage: await fs.readFile(applicationPath()),
        sharePointPackage: await fs.readFile(path.join(root, input().packagePath)),
      });

      expect(offline.isErr() ? offline.error : undefined).toBeUndefined();
      if (offline.isOk()) {
        expect(offline.value.sourceApplicationSha256).toBe(current.output.sourceAppPackageSha256);
      }
      expect(current.output.sourceAppPackageSha256).not.toBe(
        rawSha256(await fs.readFile(applicationPath()))
      );
    });

    it("SPFX-PREP-21: verifies the application source by content and refuses changed or invalid sources", async () => {
      const current = await prepared();

      await writeRebuiltApplication();
      const rebuilt = await verifyPreparedSources(root, current.journal);
      expect(rebuilt.isErr() ? rebuilt.error : undefined).toBeUndefined();

      await writeRebuiltApplication("2.3.5");
      expectError(await verifyPreparedSources(root, current.journal), "SourceHashMismatch");

      await fs.writeFile(applicationPath(), Buffer.from("not an archive"));
      expectError(await verifyPreparedSources(root, current.journal), "SourceHashMismatch");
    });

    it("SPFX-PREP-22: reuses a submitted lineage after the application package is rebuilt with identical content", async () => {
      const current = await submitted();
      await writeRebuiltApplication();

      const replay = await preparePublication(input());

      expect(replay.isErr() ? replay.error : undefined).toBeUndefined();
      if (replay.isOk()) {
        expect(replay.value.publicationId).toBe(current.output.publicationId);
        expect(replay.value.sourceAppPackageSha256).toBe(current.output.sourceAppPackageSha256);
      }
      expect(await fs.readJson(current.path)).toEqual(current.journal);
      expect(await lineageIds(current.path)).toEqual([current.output.publicationId]);
    });

    it("SPFX-PREP-22: starts a new lineage when the application content changed after a submission", async () => {
      const current = await submitted();
      await writeRebuiltApplication("2.3.5");

      const next = await preparePublication(input());

      expect(next.isErr() ? next.error : undefined).toBeUndefined();
      if (next.isOk()) {
        expect(next.value.publicationId).not.toBe(current.output.publicationId);
        expect(next.value.sourceAppPackageSha256).not.toBe(current.output.sourceAppPackageSha256);
      }
      expect(await fs.readJson(current.path)).toEqual(current.journal);
      expect(await lineageIds(current.path)).toHaveLength(2);
    });
  });

  it.each([
    { name: "no submission record", submission: undefined },
    { name: "a scope-only submission record", submission: { scope: "Tenant" } },
    {
      name: "a submission record without the app identifier",
      submission: { scope: "Tenant", titleId: "t" },
    },
    {
      name: "a submission record without the title identifier",
      submission: { scope: "Tenant", appId: "a" },
    },
  ])(
    "SPFX-PREP-08: refuses a confirmed submission with $name and keeps the journal unchanged",
    async ({ submission }) => {
      const current = await prepared();
      const raw = await fs.readJson(current.path);
      raw.state = "submitted";
      raw.applicationArtifact.submissionState = "confirmed";
      raw.applicationArtifact.submittedAppId = "published-by-the-graph-draft";
      if (submission === undefined) {
        delete raw.applicationArtifact.submission;
      } else {
        raw.applicationArtifact.submission = submission;
      }
      await fs.writeJson(current.path, raw);

      expectError(
        await loadPreparedPublication(
          root,
          current.output.publicationId,
          current.output.preparedPublicationPath
        ),
        "InvalidJournal"
      );
      expectError(await preparePublication(input()), "PublicationRecoveryCheckFailed");
      expect(await fs.readJson(current.path)).toEqual(raw);
    }
  );

  it("SPFX-PREP-08: refuses a journal that becomes unreadable between lock acquisition and update", async () => {
    const current = await prepared();
    const acquire = properLock.lock.bind(properLock);
    vi.spyOn(properLock, "lock").mockImplementation(async (file, options) => {
      const release = await acquire(file, options);
      await fs.writeFile(current.path, "{");
      return release;
    });
    const update = vi.fn((journal: PreparedPublicationJournal) => journal);
    expectError(
      await updatePreparedPublication(
        root,
        current.output.publicationId,
        current.output.preparedPublicationPath,
        update
      ),
      "JournalReadFailed"
    );
    expect(update).not.toHaveBeenCalled();
    expect(await fs.readFile(current.path, "utf8")).toBe("{");
  });

  it("SPFX-PREP-08: refuses unsupported journal versions at the identity-lock boundary", async () => {
    const current = await prepared();
    const legacy = { ...current.journal };
    Object.assign(legacy, { journalVersion: 1 });
    const operation = vi.fn();
    await expect(
      withPreparedPublicationIdentityLock(root, legacy, operation)
    ).rejects.toMatchObject({ code: "LegacyPreparedPublicationRequiresReconciliation" });
    expect(operation).not.toHaveBeenCalled();
    expectError(
      await updatePreparedPublication(root, randomUUID(), "invalid", (journal) => journal),
      "InvalidJournalPath"
    );
  });

  it.each([
    {
      name: "non-object root",
      part: "manifest.json",
      value: "null",
      code: "InvalidPackageMetadata",
    },
    {
      name: "malformed root JSON",
      part: "manifest.json",
      value: "{",
      code: "InvalidPackageMetadata",
    },
    {
      name: "non-object agent",
      part: "declarativeAgent.json",
      value: "[]",
      code: "InvalidPackageMetadata",
    },
    {
      name: "missing actions",
      part: "declarativeAgent.json",
      value: "{}",
      code: "InvalidAgentActions",
    },
  ])(
    "SPFX-PREP-06: rehashed bytes still cannot bypass $name validation",
    async ({ part, value, code }) => {
      const current = await prepared();
      const zip = new AdmZip(
        await fs.readFile(path.join(root, current.journal.applicationArtifact.path))
      );
      zip.updateFile(part, Buffer.from(value));
      const bytes = zip.toBuffer();
      current.journal.applicationArtifact.sha256 = createHash("sha256").update(bytes).digest("hex");
      expectError(verifyPreparedApplicationPackage(current.journal, bytes), code);
    }
  );

  it.each(["non-string agent file", "multiple agents", "missing connector"])(
    "SPFX-PREP-06: rehashed bytes cannot bypass reference validation (%s)",
    async (condition) => {
      const current = await prepared();
      const zip = new AdmZip(
        await fs.readFile(path.join(root, current.journal.applicationArtifact.path))
      );
      const manifest = JSON.parse(zip.readAsText("manifest.json"));
      if (condition === "non-string agent file")
        manifest.copilotAgents.declarativeAgents[0].file = 123;
      else if (condition === "multiple agents")
        manifest.copilotAgents.declarativeAgents.push({ file: "other.json" });
      else delete manifest.copilotAgents;
      zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
      const bytes = zip.toBuffer();
      current.journal.applicationArtifact.sha256 = createHash("sha256").update(bytes).digest("hex");
      expectError(
        verifyPreparedApplicationPackage(current.journal, bytes),
        condition === "non-string agent file"
          ? "InvalidAgentReference"
          : condition === "multiple agents"
            ? "UnsupportedApplicationPackage"
            : "MissingAgentConnector"
      );
    }
  );

  it("SPFX-PREP-08: write-level uncertainty cannot be hidden behind a prepared top-level state", async () => {
    const current = await prepared();
    expect(isAmbiguousPublicationState(current.journal)).toBe(false);
    for (const state of ["inFlight", "unknown"] as const) {
      for (const stage of ["uploadState", "deployState"] as const) {
        const journal = structuredClone(current.journal);
        journal.sharePointArtifact[stage] = state;
        expect(isAmbiguousPublicationState(journal)).toBe(true);
      }
      const journal = structuredClone(current.journal);
      journal.applicationArtifact.submissionState = state;
      expect(isAmbiguousPublicationState(journal)).toBe(true);
    }
  });

  it("SPFX-PREP-06: verification retains its byte caps even for hash-bound sources, artifacts and journals", async () => {
    const current = await prepared();
    SHAREPOINT_PACKAGE_LIMITS.inputBytes = 1;
    expectError(await verifyPreparedSources(root, current.journal), "InputTooLarge");
    expectError(
      await verifyPreparedArtifact(root, current.journal.applicationArtifact),
      "InputTooLarge"
    );
    SHAREPOINT_PACKAGE_LIMITS.inputBytes = initialLimits.inputBytes;
    SHAREPOINT_PACKAGE_LIMITS.metadataBytes = 10;
    expectError(
      await loadPreparedPublication(
        root,
        current.output.publicationId,
        current.output.preparedPublicationPath
      ),
      "JournalReadFailed"
    );
    const bytes = await fs.readFile(path.join(root, current.journal.applicationArtifact.path));
    expectError(verifyPreparedApplicationPackage(current.journal, bytes), "MetadataTooLarge");
  });

  it("SPFX-PREP-06: rejects invalid archive bytes even when their journal hashes are updated", async () => {
    const current = await prepared();
    const bytes = Buffer.from("not an archive");
    const hash = createHash("sha256").update(bytes).digest("hex");
    expectError(
      verifyPreparedApplicationPackage(current.journal, bytes),
      "ApplicationArtifactHashMismatch"
    );
    current.journal.applicationArtifact.sha256 = hash;
    expectError(
      verifyPreparedApplicationPackage(current.journal, bytes),
      "InvalidApplicationPackage"
    );
    current.journal.sharePointArtifact.sha256 = hash;
    expectError(await verifyPreparedSharePointPackage(current.journal, bytes), "InvalidArchive");
  });

  it("SPFX-PREP-08: accepts the exact plugin lineage from the previous transformation", async () => {
    const current = await prepared();
    const oldPath = `plugins/spfx-api-${current.journal.plugin.sha256.slice(0, 16)}.json`;
    const zip = new AdmZip(
      await fs.readFile(path.join(root, current.journal.applicationArtifact.path))
    );
    const agent = JSON.parse(zip.readAsText("declarativeAgent.json"));
    const plugin = zip.readFile(agent.actions[0].file);
    if (!plugin) throw new Error("Missing prepared plugin.");
    zip.deleteFile(agent.actions[0].file);
    agent.actions = [{ id: path.posix.basename(oldPath, ".json"), file: oldPath }];
    zip.updateFile("declarativeAgent.json", Buffer.from(JSON.stringify(agent)));
    zip.addFile(oldPath, plugin);
    const bytes = zip.toBuffer();
    current.journal.transformation.id = "spfx-copilot-plugin-v2";
    delete current.journal.plugin.bindingSha256;
    current.journal.applicationArtifact.sha256 = createHash("sha256").update(bytes).digest("hex");
    const verified = verifyPreparedApplicationPackage(current.journal, bytes);
    expect(verified.isErr() ? verified.error : undefined).toBeUndefined();
  });

  it("SPFX-PREP-02/06: verifies a prepared connector-only contribution without an agent manifest", async () => {
    await changeJson("manifest.json", (manifest) => {
      delete manifest.copilotAgents;
      manifest.manifestVersion = "devPreview";
      manifest.$schema =
        "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json";
      manifest.agentConnectors = [];
    });
    const current = await prepared();
    const bytes = await fs.readFile(path.join(root, current.journal.applicationArtifact.path));
    expect(verifyPreparedApplicationPackage(current.journal, bytes).isOk()).toBe(true);
  });

  it.each([
    { name: "current", transformationId: undefined },
    { name: "previous v4", transformationId: "spfx-copilot-plugin-v4" },
  ])(
    "SPFX-PREP-12/13: resolves prepared plugin actions from a nested agent manifest directory ($name lineage)",
    async ({ transformationId }) => {
      const application = new AdmZip(await fs.readFile(path.join(root, input().appPackagePath)));
      const manifest = JSON.parse(application.readAsText("manifest.json"));
      manifest.copilotAgents.declarativeAgents[0].file = "agents/agent.json";
      application.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
      application.addFile(
        "agents/agent.json",
        application.readFile("declarativeAgent.json") ?? Buffer.alloc(0)
      );
      application.deleteFile("declarativeAgent.json");
      await fs.writeFile(path.join(root, input().appPackagePath), application.toBuffer());

      const current = await prepared();
      const bytes = await fs.readFile(path.join(root, current.journal.applicationArtifact.path));
      const output = new AdmZip(bytes);
      const agent = JSON.parse(output.readAsText("agents/agent.json"));
      const reference = agent.actions.find((action: { id: string }) =>
        action.id.startsWith("spfx-api-")
      );
      expect(reference).toBeDefined();
      if (reference) {
        const resolvedEntry = path.posix.normalize(
          path.posix.join(path.posix.dirname("agents/agent.json"), reference.file)
        );
        expect(reference.file).toBe(`../plugins/${reference.id}.json`);
        expect(output.getEntry(resolvedEntry)).not.toBeNull();
      }
      if (transformationId) {
        current.journal.transformation.id = transformationId;
      }
      expect(verifyPreparedApplicationPackage(current.journal, bytes).isOk()).toBe(true);
    }
  );

  it("SPFX-PREP-11: reports SharePoint verification only when upload and deployment are confirmed", async () => {
    const current = await prepared();
    for (const [state, uploadState, deployState] of [
      ["unknown", "unknown", "notStarted"],
      ["rejected", "rejected", "notStarted"],
      ["unknown", "confirmed", "unknown"],
      ["rejected", "confirmed", "rejected"],
    ] as const) {
      const journal = structuredClone(current.journal);
      journal.state = state;
      journal.sharePointArtifact.uploadState = uploadState;
      journal.sharePointArtifact.deployState = deployState;
      expect(
        createPreparedPublicationErrorData(
          new Error("failed"),
          journal,
          current.output.preparedPublicationPath,
          state === "unknown" ? "sharePointUpload" : "sharePointDeploy"
        ).lastSharePointVerifiedState
      ).toBeUndefined();
    }

    const verifiedJournal = structuredClone(current.journal);
    verifiedJournal.state = "unknown";
    verifiedJournal.sharePointArtifact.uploadState = "confirmed";
    verifiedJournal.sharePointArtifact.deployState = "confirmed";
    expect(
      createPreparedPublicationErrorData(
        new Error("submission uncertain"),
        verifiedJournal,
        current.output.preparedPublicationPath,
        "applicationSubmission"
      ).lastSharePointVerifiedState
    ).toBe("sharePointVerified");
  });

  it("SPFX-PREP-11: excludes untrusted error names and non-error values from recovery diagnostics", () => {
    const projected = createPreparedPublicationErrorData(
      {
        name: "untrusted error text with credentials",
        response: {
          status: 403,
          data: { code: "<invalid>" },
          headers: { "request-id": "invalid value" },
        },
      },
      undefined,
      undefined,
      "preparedPublication"
    );
    expect(projected).toMatchObject({
      errorType: "Error",
      httpStatus: 403,
      serviceCode: undefined,
      requestId: undefined,
    });
    expect(
      createPreparedPublicationErrorData(undefined, undefined, undefined, "preparedPublication")
        .errorType
    ).toBe("Error");
    expect(isPreparedCatalogUrl("not a URL", publicationTenantContext.tenantUrl)).toBe(false);
  });
});
