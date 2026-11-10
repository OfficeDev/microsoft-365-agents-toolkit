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
import {
  createAppPackage,
  createSppkg,
  productId,
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
