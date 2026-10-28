// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, Platform, TeamsAppManifest, UserError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import axios, { AxiosRequestConfig } from "axios";
import MockAdapter from "axios-mock-adapter";
import { createHash } from "crypto";
import FormData from "form-data";
import fs from "fs-extra";
import nodeFs from "fs";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import { CopilotAgentPublishDriver } from "../../../../src/component/driver/copilotAgent/publish";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { M365TitleAcquireDriver } from "../../../../src/component/driver/m365/acquire";
import { AppScope, PackageService } from "../../../../src/component/m365/packageService";
import { prepareIncrementalSkillUpload } from "../../../../src/component/m365/skills/incrementalSkillUpload";
import {
  MockedAzureAccountProvider,
  MockedM365Provider,
  MockLogProvider,
} from "../../../core/utils";

const endpoint = "https://titles.prod.mos.microsoft.com";
const agentId = "11111111-1111-4111-8111-111111111111";
const folderId = "11111111-1111-1111-1111-111111111111";
const extraFolderId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
// eslint-disable-next-line no-secrets/no-secrets -- Published MOS encoding vector, not a credential.
const driveId = "b!LCwsLBoaS0uMjJ2dnZ2dnT8_Pz8rK0xMjY0ODg4ODg5KSkpKPDxNTY6OHx8fHx8f";
const skillMd = "---\nname: example\ndescription: Example skill\n---\nUse the attached data.";
const sharePointIds = {
  siteId: "2c2c2c2c-1a1a-4b4b-8c8c-9d9d9d9d9d9d",
  webId: "3f3f3f3f-2b2b-4c4c-8d8d-0e0e0e0e0e0e",
  listId: "4a4a4a4a-3c3c-4d4d-8e8e-1f1f1f1f1f1f",
  uniqueId: `{${folderId}}`,
  siteUrl: "https://contoso.sharepoint.com/sites/skills",
  tenantId: "77777777-7777-4777-8777-777777777777",
};

function packageBytes(
  withSkills = true,
  changed = false,
  extraFiles = 0,
  version = "1.0.0"
): Buffer {
  const manifest = new TeamsAppManifest();
  manifest.id = agentId;
  manifest.version = version;
  manifest.copilotAgents = {
    declarativeAgents: [{ id: "agent", file: "agent.json" }],
  };
  const zip = new AdmZip();
  zip.addFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
  zip.addFile(
    "agent.json",
    Buffer.from(
      JSON.stringify({
        version: "v1.6",
        name: "Agent",
        description: "Scenario agent",
        instructions: "Use skills.",
        ...(withSkills ? { agent_skills: [{ folder: "skills/example" }] } : {}),
      })
    )
  );
  zip.addFile("color.png", Buffer.from([0, 255, 128]));
  zip.addFile("unrelated.txt", Buffer.from("retain me"));
  if (withSkills) {
    zip.addFile("skills/example/SKILL.md", Buffer.from(skillMd));
    zip.addFile("skills/example/assets/data.bin", Buffer.from(changed ? [1, 254] : [0, 255]));
    if (!changed) zip.addFile("skills/example/empty", Buffer.alloc(0));
    for (let index = 0; index < extraFiles; index++) {
      zip.addFile(`skills/example/assets/file-${index}.txt`, Buffer.from(`content-${index}`));
    }
  }
  return zip.toBuffer();
}

function rootSkillPackage(mixed = false, distinct = false): AdmZip {
  const zip = new AdmZip(packageBytes());
  const manifest = JSON.parse(zip.readAsText("manifest.json"));
  manifest.manifestVersion = "1.30";
  delete manifest.packageName;
  manifest.agentSkills = [{ folder: "./skills/example" }];
  if (distinct) {
    manifest.agentSkills.push({ folder: "skills/extra" });
    zip.addFile(
      "skills/extra/SKILL.md",
      Buffer.from("---\nname: extra\ndescription: Extra skill\n---\nExtra instructions.")
    );
  }
  if (mixed) {
    const agent = JSON.parse(zip.readAsText("agent.json"));
    agent.agent_skills[0].expose_skill_to_copilot = false;
    agent["x-agent_skills"] = [{ folder: "./skills/example", expose_skill_to_copilot: true }];
    zip.updateFile("agent.json", Buffer.from(JSON.stringify(agent)));
  } else {
    delete manifest.copilotAgents;
    zip.deleteFile("agent.json");
  }
  zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
  return zip;
}

function multipartZip(data: unknown): Buffer {
  if (!(data instanceof FormData)) throw new Error("Expected multipart package");
  const body = data.getBuffer();
  const boundary = Buffer.from(`\r\n--${data.getBoundary()}`);
  let cursor = 0;
  while (cursor < body.length) {
    const headersEnd = body.indexOf("\r\n\r\n", cursor);
    if (headersEnd < 0) break;
    const start = headersEnd + 4;
    const end = body.indexOf(boundary, start);
    if (end < 0) break;
    const part = body.subarray(start, end);
    if (part.subarray(0, 4).equals(Buffer.from([80, 75, 3, 4]))) return part;
    cursor = end + boundary.length;
  }
  throw new Error("No ZIP in multipart package");
}

interface WireItem {
  id: string;
  name: string;
  eTag: string;
  size?: number;
  folder?: Record<string, never>;
  file?: { hashes: { sha1Hash?: string; quickXorHash?: string } };
}

class UploadHttpBoundary {
  readonly children = new Map<string, WireItem[]>([["skill-root", []]]);
  readonly requests: string[] = [];
  readonly submitted: Buffer[] = [];
  readonly uploads: { name: string; bytes: Buffer }[] = [];
  readonly deleted: string[] = [];
  readonly skillNames = ["example"];
  readonly monitors = new Map<string, string[]>();
  draftItemId = sharePointIds.uniqueId;
  failUploadName?: string;
  directoryStatus = 200;
  finalStatus = 202;
  activeRequests = 0;
  maxActiveRequests = 0;
  private id = 0;
  private published = false;
  private draftNumber = 0;

  reply(config: AxiosRequestConfig): [number, unknown] {
    const url = new URL(config.url ?? "", config.baseURL ?? endpoint);
    let path = url.pathname.replace(/^\/v1\.0/, "");
    if (path.startsWith("/drives/")) {
      const prefix = `/drives/${driveId}/items/`;
      expect(path.startsWith(prefix)).toBe(true);
      path = `/drives/drive/items/${decodeURIComponent(path.substring(prefix.length))}`;
      path = path.replace(`/items/${this.draftItemId}`, "/items/skill-root");
      path = path.replace(`/items/{${extraFolderId}}`, "/items/extra-skill-root");
    }
    const method = config.method?.toUpperCase();
    this.requests.push(`${method} ${path}`);
    if (method === "GET" && path === "/config/v1/environment") {
      return [200, { titlesServiceUrl: endpoint }];
    }
    if (method === "POST" && path === "/builder/v1/directories") {
      const body = JSON.parse(config.data);
      expect(body).toEqual({
        agentId,
        directoryType: "skill",
        name: expect.any(String),
        skipPublishedContentCopy: true,
      });
      expect(this.skillNames).toContain(body.name);
      if (body.name === "extra") {
        if (!this.children.has("extra-skill-root")) this.children.set("extra-skill-root", []);
        return [
          this.directoryStatus,
          {
            folder: extraFolderId,
            sharePointIds: { ...sharePointIds, uniqueId: `{${extraFolderId}}` },
            copyStatusUrls: [],
          },
        ];
      }
      if (this.published) {
        this.draftNumber++;
        this.draftItemId = `{22222222-2222-4222-8222-${String(this.draftNumber).padStart(12, "0")}}`;
        this.children.clear();
        this.children.set("skill-root", []);
        this.published = false;
      }
      return [
        this.directoryStatus,
        {
          folder: folderId,
          sharePointIds: { ...sharePointIds, uniqueId: this.draftItemId },
          copyStatusUrls: this.monitors.get(body.name),
        },
      ];
    }
    const list = /^\/drives\/drive\/items\/([^/]+)\/children$/.exec(path);
    if (list && method === "GET") return [200, { value: this.children.get(list[1]) ?? [] }];
    if (list && method === "POST") {
      const body: unknown = JSON.parse(config.data);
      if (!body || typeof body !== "object" || !("name" in body) || typeof body.name !== "string") {
        throw new Error("Invalid Graph folder body");
      }
      expect(body).toEqual({
        name: body.name,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      });
      const item: WireItem = {
        id: `folder-${this.id++}`,
        name: body.name,
        eTag: '"v1"',
        folder: {},
      };
      this.children.set(list[1], [...(this.children.get(list[1]) ?? []), item]);
      this.children.set(item.id, []);
      return [201, item];
    }
    const create = /^\/drives\/drive\/items\/([^/:]+):\/(.+):\/content$/.exec(path);
    const update = /^\/drives\/drive\/items\/([^/]+)\/content$/.exec(path);
    if (method === "PUT" && (create || update)) {
      const existing =
        update && [...this.children.values()].flat().find((item) => item.id === update[1]);
      const name = create ? decodeURIComponent(create[2]) : existing ? existing.name : "";
      if (name === this.failUploadName) return [403, { error: { code: "accessDenied" } }];
      if (!Buffer.isBuffer(config.data)) throw new Error("Expected raw Graph upload bytes");
      this.uploads.push({ name, bytes: Buffer.from(config.data) });
      const item: WireItem = {
        id: existing ? existing.id : `file-${this.id++}`,
        name,
        eTag: `"version-${this.id}"`,
        size: config.data.length,
        file: { hashes: { sha1Hash: createHash("sha1").update(config.data).digest("hex") } },
      };
      if (create) this.children.set(create[1], [...(this.children.get(create[1]) ?? []), item]);
      else
        for (const [id, items] of this.children) {
          this.children.set(
            id,
            items.map((previous) => (previous.id === item.id ? item : previous))
          );
        }
      return [existing ? 200 : 201, item];
    }
    const deletion = /^\/drives\/drive\/items\/([^/]+)$/.exec(path);
    if (method === "DELETE" && deletion) {
      this.deleted.push(deletion[1]);
      for (const [id, items] of this.children) {
        this.children.set(
          id,
          items.filter((item) => item.id !== deletion[1])
        );
      }
      this.children.delete(deletion[1]);
      return [204, undefined];
    }
    if (method === "POST" && path === "/builder/v1/users/packages") {
      expect(config.params).toMatchObject({ isRAIValidationAsync: false });
      this.submitted.push(multipartZip(config.data));
      if (this.finalStatus === 409) {
        return [409, { error: { code: "Conflict", message: "private conflict response" } }];
      }
      this.published = true;
      return [202, { statusId: "status", titlePreview: { titleId: "title" } }];
    }
    if (method === "GET" && path === "/builder/v1/users/packages/status/status") {
      expect(config.params).toEqual({ isRAIValidationAsync: false });
      return [
        200,
        {
          titleId: "title",
          appId: agentId,
          validationInfo: { status: "successful", completedDate: "2026-10-10T00:00:00Z" },
        },
      ];
    }
    if (method === "GET" && path === "/marketplace/v1/users/titles/title/sharingInfo") {
      return [200, { unifiedStoreLink: "https://example.com/share" }];
    }
    throw new Error(`Unexpected HTTP request: ${method} ${path}`);
  }
}

describe("incremental skill upload PackageService scenarios", () => {
  let bytes: Buffer;
  let boundary: UploadHttpBoundary;
  let http: MockAdapter[];
  let provider: MockedM365Provider;
  let service: PackageService;

  beforeEach(() => {
    vi.stubEnv(FeatureFlagName.IncrementalSkillUpload, "true");
    vi.stubEnv("TEAMSFX_BUILDER_API", "true");
    bytes = packageBytes();
    provider = new MockedM365Provider();
    boundary = new UploadHttpBoundary();
    http = [];
    const create = axios.create;
    vi.spyOn(axios, "create").mockImplementation((options) => {
      const client = create(options);
      const adapter = new MockAdapter(client, { onNoMatch: "throwException" });
      adapter.onAny().reply(async (config) => {
        boundary.activeRequests++;
        boundary.maxActiveRequests = Math.max(boundary.maxActiveRequests, boundary.activeRequests);
        try {
          await Promise.resolve();
          const [status, data] = boundary.reply(config);
          return [status, data, { "content-type": "application/json" }];
        } finally {
          boundary.activeRequests--;
        }
      });
      client.interceptors.response.use((response) => ({
        ...response,
        request: { method: response.config.method },
      }));
      http.push(adapter);
      return client;
    });
    vi.spyOn(fs, "readFile").mockImplementation(async () => bytes);
    const readFileSync = nodeFs.readFileSync;
    vi.spyOn(nodeFs, "readFileSync").mockImplementation((file, options) =>
      file.toString() === __filename || file.toString().endsWith("app.zip")
        ? bytes
        : readFileSync(file, options)
    );
    const stats = fs.statSync(__filename);
    stats.size = bytes.length;
    vi.spyOn(fs, "statSync").mockReturnValue(stats);
    service = new PackageService(endpoint, new MockLogProvider(), provider);
  });

  afterEach(() => {
    http.forEach((adapter) => adapter.restore());
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
  });

  it.each([
    { mixed: false, distinct: false },
    { mixed: true, distinct: false },
    { mixed: true, distinct: true },
  ])(
    "SCN-ISU-09 AC-30 AC-31: root skills publish in memory ($mixed mixed, $distinct distinct)",
    async ({ mixed, distinct }) => {
      const source = rootSkillPackage(mixed, distinct);
      bytes = source.toBuffer();
      const original = Buffer.from(bytes);
      if (distinct) boundary.skillNames.push("extra");
      const write = vi.spyOn(fs, "writeFile");
      const writeSync = vi.spyOn(fs, "writeFileSync");
      await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
      expect(boundary.submitted).toHaveLength(1);
      expect(
        boundary.requests.filter((request) => request === "POST /builder/v1/directories")
      ).toHaveLength(distinct ? 2 : 1);
      expect(boundary.uploads).toHaveLength(distinct ? 4 : 3);
      expect(
        boundary.uploads.filter((upload) => upload.bytes.equals(Buffer.from(skillMd)))
      ).toHaveLength(1);
      const transport = new AdmZip(boundary.submitted[0]);
      const root = JSON.parse(transport.readAsText("manifest.json"));
      const localRoot = JSON.parse(source.readAsText("manifest.json"));
      expect(root).toEqual({
        ...localRoot,
        agentSkills: [{ folder: folderId }, ...(distinct ? [{ folder: extraFolderId }] : [])],
      });
      if (mixed) {
        expect(JSON.parse(transport.readAsText("agent.json"))).toEqual({
          ...JSON.parse(source.readAsText("agent.json")),
          agent_skills: [{ folder: folderId, expose_skill_to_copilot: false }],
          "x-agent_skills": [{ folder: folderId, expose_skill_to_copilot: true }],
        });
      } else expect(transport.getEntry("agent.json")).toBeNull();
      expect(JSON.parse(transport.readAsText("packageMetadata.json"))).toEqual({
        builderInfo: { builderName: "TeamsToolKit" },
        skills: [
          { folder: folderId, name: "example", description: "Example skill" },
          ...(distinct
            ? [{ folder: extraFolderId, name: "extra", description: "Extra skill" }]
            : []),
        ],
      });
      expect(transport.getEntries().some((entry) => entry.entryName.startsWith("skills/"))).toBe(
        false
      );
      for (const path of ["color.png", "unrelated.txt"]) {
        expect(transport.readFile(path)).toEqual(source.readFile(path));
      }
      const submission = boundary.requests.indexOf("POST /builder/v1/users/packages");
      expect(submission).toBeGreaterThan(0);
      expect(
        boundary.requests.slice(submission + 1).some((request) => request.startsWith("PUT "))
      ).toBe(false);
      expect(boundary.maxActiveRequests).toBe(1);
      expect(bytes).toEqual(original);
      expect(write).not.toHaveBeenCalled();
      expect(writeSync).not.toHaveBeenCalled();
    }
  );

  it.each(
    [
      { mixed: false, malformed: false },
      { mixed: true, malformed: false },
      { mixed: false, malformed: true },
    ].flatMap((scenario) => ["false", undefined].map((flag) => ({ ...scenario, flag })))
  )(
    "AC-32 SCN-ISU-09: root skills preserve the legacy path ($mixed mixed, $malformed malformed, $flag flag)",
    async ({ mixed, malformed, flag }) => {
      vi.stubEnv(FeatureFlagName.IncrementalSkillUpload, flag);
      const zip = rootSkillPackage(mixed);
      if (malformed) {
        const manifest = JSON.parse(zip.readAsText("manifest.json"));
        manifest.agentSkills = [{ folder: "../unsafe" }];
        zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
      }
      bytes = zip.toBuffer();
      const original = Buffer.from(bytes);
      const token = vi.spyOn(provider, "getAccessToken");
      await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
      expect(boundary.submitted).toEqual([original]);
      expect(token).not.toHaveBeenCalled();
      expect(
        boundary.requests.some(
          (request) => request.includes("/directories") || request.includes("/drives/")
        )
      ).toBe(false);
      expect(boundary.uploads).toEqual([]);
    }
  );

  it.each([undefined, []].map((skills) => ({ skills })))(
    "AC-32: empty or absent root skills leave no-skill packages unchanged ($skills)",
    async ({ skills }) => {
      const zip = rootSkillPackage();
      const manifest = JSON.parse(zip.readAsText("manifest.json"));
      manifest.agentSkills = skills;
      zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
      bytes = zip.toBuffer();
      const token = vi.spyOn(provider, "getAccessToken");
      await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
      expect(boundary.submitted).toEqual([bytes]);
      expect(token).not.toHaveBeenCalled();
      expect(
        boundary.requests.some(
          (request) => request.includes("/directories") || request.includes("/drives/")
        )
      ).toBe(false);
    }
  );

  it.each(
    [
      null,
      {},
      "invalid",
      [null],
      [{}],
      [{ folder: 1 }],
      [{ folder: "" }],
      [{ folder: "../unsafe" }],
      [{ folder: "skills/missing" }],
    ].flatMap((skills) => ["1.30", "99.0"].map((version) => ({ skills, version })))
  )(
    "AC-33: invalid root declarations fail before remote mutation ($skills, schema $version)",
    async ({ skills, version }) => {
      const zip = rootSkillPackage();
      const manifest = JSON.parse(zip.readAsText("manifest.json"));
      manifest.manifestVersion = version;
      manifest.agentSkills = skills;
      zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
      bytes = zip.toBuffer();
      const token = vi.spyOn(provider, "getAccessToken");
      await expect(
        service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal)
      ).rejects.toMatchObject({ name: "InvalidSkillPackage" });
      expect(token).not.toHaveBeenCalled();
      expect(
        boundary.requests.some(
          (request) => request.includes("/directories") || request.includes("/drives/")
        )
      ).toBe(false);
      expect(boundary.submitted).toEqual([]);
    }
  );

  it.each(["missing-skill-md", "conflicting-name", "remaining-reference", "bad-metadata"])(
    "AC-33: root skill safety checks reject %s before remote mutations",
    async (failure) => {
      const zip = rootSkillPackage(failure === "conflicting-name");
      if (failure === "missing-skill-md") zip.deleteFile("skills/example/SKILL.md");
      if (failure === "conflicting-name") {
        const manifest = JSON.parse(zip.readAsText("manifest.json"));
        manifest.agentSkills = [{ folder: "other/example" }];
        zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
        zip.addFile("other/example/SKILL.md", Buffer.from(skillMd));
      }
      if (failure === "remaining-reference") {
        const manifest = JSON.parse(zip.readAsText("manifest.json"));
        manifest.icons.color = "skills/example/assets/data.bin";
        zip.updateFile("manifest.json", Buffer.from(JSON.stringify(manifest)));
      }
      if (failure === "bad-metadata") {
        zip.addFile(
          "packageMetadata.json",
          Buffer.from(
            JSON.stringify({
              skills: [{ folder: "skills/example", name: "wrong" }],
            })
          )
        );
      }
      bytes = zip.toBuffer();
      const token = vi.spyOn(provider, "getAccessToken");
      await expect(
        service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal)
      ).rejects.toMatchObject({ name: "InvalidSkillPackage" });
      expect(token).not.toHaveBeenCalled();
      expect(
        boundary.requests.some(
          (request) => request.includes("/directories") || request.includes("/drives/")
        )
      ).toBe(false);
      expect(boundary.submitted).toEqual([]);
    }
  );

  it("SCN-ISU-01 AC-03 AC-11 AC-20 AC-21: first upload skips copying without UI and submits only in-memory skill references", async () => {
    const original = Buffer.from(bytes);
    const write = vi.spyOn(fs, "writeFile");
    const writeSync = vi.spyOn(fs, "writeFileSync");
    const result = await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    expect(result).toBeDefined();
    expect(boundary.uploads.map((upload) => upload.name).sort()).toEqual([
      "SKILL.md",
      "data.bin",
      "empty",
    ]);
    expect(boundary.uploads.find((upload) => upload.name === "data.bin")?.bytes).toEqual(
      Buffer.from([0, 255])
    );
    expect(boundary.uploads.find((upload) => upload.name === "empty")?.bytes).toEqual(
      Buffer.alloc(0)
    );
    expect(boundary.submitted).toHaveLength(1);
    const transport = new AdmZip(boundary.submitted[0]);
    const source = new AdmZip(original);
    expect(JSON.parse(transport.readAsText("agent.json")).agent_skills).toEqual([
      { folder: folderId },
    ]);
    expect(JSON.parse(transport.readAsText("packageMetadata.json"))).toEqual({
      builderInfo: { builderName: "TeamsToolKit" },
      skills: [{ folder: folderId, name: "example", description: "Example skill" }],
    });
    expect(transport.getEntries().some((entry) => entry.entryName.startsWith("skills/"))).toBe(
      false
    );
    for (const name of ["manifest.json", "color.png", "unrelated.txt"]) {
      expect(transport.readFile(name)).toEqual(source.readFile(name));
    }
    const submission = boundary.requests.indexOf("POST /builder/v1/users/packages");
    expect(
      boundary.requests.slice(submission + 1).every((request) => !request.startsWith("PUT "))
    ).toBe(true);
    expect(bytes).toEqual(original);
    expect(write).not.toHaveBeenCalled();
    expect(writeSync).not.toHaveBeenCalled();
    expect(boundary.requests.some((request) => request.startsWith("DELETE "))).toBe(false);
  });

  it("SCN-ISU-02 AC-04: retained-draft synchronization trusts remote SHA-1 and QuickXor metadata without PUTs", async () => {
    const prepared = await prepareIncrementalSkillUpload(bytes, endpoint, "mos-token", provider);
    expect(prepared.isOk()).toBe(true);
    const empty = [...boundary.children.values()].flat().find((item) => item.name === "empty");
    if (!empty) throw new Error("Missing uploaded empty file");
    empty.file = { hashes: { quickXorHash: "AAAAAAAAAAAAAAAAAAAAAAAAAAA=" } };
    boundary.requests.length = 0;
    boundary.uploads.length = 0;
    bytes = packageBytes(true, false, 0, "2.0.0");
    await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    expect(boundary.uploads).toEqual([]);
    expect(boundary.requests.some((request) => request.startsWith("PUT "))).toBe(false);
    expect(
      boundary.requests.some(
        (request) => request.startsWith("GET ") && request.endsWith("/content")
      )
    ).toBe(false);
    expect(boundary.submitted).toHaveLength(1);
    expect(JSON.parse(new AdmZip(boundary.submitted[0]).readAsText("manifest.json")).version).toBe(
      "2.0.0"
    );
  });

  it("SCN-ISU-02 AC-05: changed same-size bytes upload and deleted files are removed before submission", async () => {
    const prepared = await prepareIncrementalSkillUpload(bytes, endpoint, "mos-token", provider);
    expect(prepared.isOk()).toBe(true);
    const empty = [...boundary.children.values()].flat().find((item) => item.name === "empty");
    bytes = packageBytes(true, true, 0, "2.0.0");
    boundary.uploads.length = 0;
    await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    expect(boundary.uploads).toEqual([{ name: "data.bin", bytes: Buffer.from([1, 254]) }]);
    expect(boundary.deleted).toEqual([empty?.id]);
    expect(boundary.submitted).toHaveLength(1);
  });

  it("SCN-ISU-06 AC-22: an unchanged skill uploads every file to a fresh draft after publication", async () => {
    await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    const oldDestination = boundary.draftItemId;
    boundary.requests.length = 0;
    boundary.uploads.length = 0;
    bytes = packageBytes(true, false, 0, "2.0.0");
    await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    expect(boundary.draftItemId).not.toBe(oldDestination);
    expect(boundary.uploads.map((upload) => upload.name).sort()).toEqual([
      "SKILL.md",
      "data.bin",
      "empty",
    ]);
    const graphRequests = http
      .flatMap((adapter) => adapter.history.get)
      .filter((request) => request.url?.includes("/drives/"));
    expect(
      graphRequests.some((request) =>
        request.url?.includes(encodeURIComponent(boundary.draftItemId))
      )
    ).toBe(true);
    expect(boundary.submitted).toHaveLength(2);
    const transport = new AdmZip(boundary.submitted[1]);
    expect(JSON.parse(transport.readAsText("agent.json")).agent_skills).toEqual([
      { folder: folderId },
    ]);
    expect(boundary.requests.some((request) => request.startsWith("DELETE "))).toBe(false);
  });

  it("SCN-ISU-06 AC-23: a later run resumes an interrupted copy-free draft without resetting it", async () => {
    boundary.failUploadName = "data.bin";
    await expect(
      service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal)
    ).rejects.toMatchObject({ name: "SkillUploadAuthorizationFailed" });
    expect(boundary.uploads.map((upload) => upload.name)).toEqual(["SKILL.md"]);
    expect(boundary.submitted).toEqual([]);
    const destination = boundary.draftItemId;
    boundary.failUploadName = undefined;
    boundary.uploads.length = 0;
    boundary.requests.length = 0;
    service = new PackageService(endpoint, new MockLogProvider(), provider);
    await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    expect(boundary.draftItemId).toBe(destination);
    expect(boundary.uploads.map((upload) => upload.name).sort()).toEqual(["data.bin", "empty"]);
    expect(boundary.requests.filter((request) => request.includes("/directories"))).toEqual([
      "POST /builder/v1/directories",
    ]);
    expect(boundary.requests.some((request) => request.startsWith("DELETE "))).toBe(false);
    expect(boundary.submitted).toHaveLength(1);
  });

  it("SCN-ISU-03 AC-09 AC-21: an ambiguous create is not replayed and a later run can reuse the copy-free draft", async () => {
    boundary.directoryStatus = 503;
    await expect(
      service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal)
    ).rejects.toMatchObject({ name: "SkillUploadRequestFailed" });
    expect(boundary.requests.filter((request) => request.includes("/directories"))).toHaveLength(1);
    expect(boundary.uploads).toEqual([]);
    expect(boundary.submitted).toEqual([]);
    const destination = boundary.draftItemId;
    boundary.directoryStatus = 200;
    service = new PackageService(endpoint, new MockLogProvider(), provider);
    await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    expect(boundary.draftItemId).toBe(destination);
    expect(boundary.requests.filter((request) => request.includes("/directories"))).toHaveLength(2);
    expect(boundary.uploads).toHaveLength(3);
    expect(boundary.submitted).toHaveLength(1);
    expect(boundary.requests.some((request) => request.startsWith("DELETE "))).toBe(false);
  });

  it("SCN-ISU-08 AC-07 AC-16: copying reported for a later skill stops before any file writes or monitor requests", async () => {
    const zip = new AdmZip(bytes);
    const agent = JSON.parse(zip.readAsText("agent.json"));
    agent.agent_skills.push({ folder: "skills/second" });
    zip.updateFile("agent.json", Buffer.from(JSON.stringify(agent)));
    zip.addFile(
      "skills/second/SKILL.md",
      Buffer.from("---\nname: second\ndescription: Second skill\n---\nUse the data.")
    );
    bytes = zip.toBuffer();
    boundary.skillNames.push("second");
    const monitor = "https://contoso.sharepoint.com/copy?tempauth=private-monitor-value";
    boundary.monitors.set("second", [monitor]);
    const logger = new MockLogProvider();
    const warning = vi.spyOn(logger, "warning");
    const info = vi.spyOn(logger, "info");
    service = new PackageService(endpoint, logger, provider);
    const upload = service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    await expect(upload).rejects.toMatchObject({
      name: "SkillUploadCopyNotSupported",
      source: "M365",
    });
    const failure: unknown = await upload.catch((error: unknown) => error);
    expect(JSON.stringify([failure, warning.mock.calls, info.mock.calls])).not.toContain(
      "private-monitor-value"
    );
    expect(warning).toHaveBeenCalledWith(expect.stringContaining("skip-copy"));
    expect(boundary.uploads).toEqual([]);
    expect(boundary.submitted).toEqual([]);
    expect(boundary.requests.filter((request) => request.includes("/directories"))).toHaveLength(2);
    expect(boundary.requests.some((request) => /^(PUT|DELETE) /.test(request))).toBe(false);
    expect(boundary.requests.some((request) => request.includes("/copy"))).toBe(false);
    expect(
      http.flatMap((adapter) => adapter.history.get).some((request) => request.url === monitor)
    ).toBe(false);
  });

  it("SCN-ISU-03 AC-09: Graph failure prevents final package submission and legacy fallback", async () => {
    boundary.failUploadName = "data.bin";
    await expect(
      service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal)
    ).rejects.toBeDefined();
    expect(boundary.submitted).toEqual([]);
    expect(
      boundary.requests.some(
        (request) => request.startsWith("POST ") && request.includes("/packages")
      )
    ).toBe(false);
    expect(boundary.requests.filter((request) => request.startsWith("PUT "))).toHaveLength(2);
  });

  it.each([false, true])(
    "AC-01 AC-02: bypass preserves original ZIP bytes (no skills: %s)",
    async (noSkills) => {
      bytes = packageBytes(!noSkills);
      if (!noSkills) vi.stubEnv(FeatureFlagName.IncrementalSkillUpload, "false");
      const token = vi.spyOn(provider, "getAccessToken");
      await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
      expect(boundary.submitted).toEqual([bytes]);
      expect(token).not.toHaveBeenCalled();
      expect(
        boundary.requests.some(
          (request) =>
            request.includes("/directories") ||
            request.includes("/drives/") ||
            request.includes("/sites/")
        )
      ).toBe(false);
    }
  );

  it.each(["missing", "failed"])(
    "AC-10: %s Graph token provider prevents all mutations",
    async (mode) => {
      const failure = new UserError("test", "ConsentRequired", "Consent required");
      if (mode === "missing") service = new PackageService(endpoint, new MockLogProvider());
      else vi.spyOn(provider, "getAccessToken").mockResolvedValue(err(failure));
      const upload = service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
      if (mode === "missing") {
        await expect(upload).rejects.toMatchObject({
          name: "SkillUploadAuthorizationFailed",
          source: "M365",
        });
      } else {
        await expect(upload).rejects.toBe(failure);
      }
      expect(boundary.requests.every((request) => request.startsWith("GET "))).toBe(true);
      expect(boundary.submitted).toEqual([]);
    }
  );

  it("AC-18: 256 additional files use sequential requests and never enter the reference ZIP", async () => {
    bytes = packageBytes(true, false, 256);
    await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    expect(boundary.uploads).toHaveLength(259);
    expect(boundary.maxActiveRequests).toBe(1);
    expect(boundary.activeRequests).toBe(0);
    expect(boundary.submitted).toHaveLength(1);
    const transport = new AdmZip(boundary.submitted[0]);
    expect(transport.getEntries().some((entry) => entry.entryName.startsWith("skills/"))).toBe(
      false
    );
    expect(JSON.parse(transport.readAsText("packageMetadata.json")).skills).toEqual([
      { folder: folderId, name: "example", description: "Example skill" },
    ]);
    const submission = boundary.requests.indexOf("POST /builder/v1/users/packages");
    expect(
      boundary.requests.slice(0, submission).filter((request) => request.startsWith("PUT "))
    ).toHaveLength(259);
  });

  it("AC-15: publishAgent rejects original bytes above 10 MB before preparing a smaller transport", async () => {
    // Keep the mocked stat size small to require validation of the actual read buffer.
    bytes = Buffer.alloc(10 * 1024 * 1024 + 1);
    const token = vi.spyOn(provider, "getAccessToken");
    await expect(
      service.publishAgent("mos-token", "app.zip", AppScope.Personal)
    ).rejects.toMatchObject({ name: "AppPackageSizeExceeded" });
    expect(token).not.toHaveBeenCalled();
    expect(boundary.requests.every((request) => request.startsWith("GET "))).toBe(true);
    expect(boundary.submitted).toEqual([]);
    expect(boundary.uploads).toEqual([]);
  });

  it("AC-09: a final HTTP 409 for an actual version-2 upload never replays or bumps the version", async () => {
    await service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    bytes = packageBytes(true, false, 0, "2.0.0");
    const original = Buffer.from(bytes);
    boundary.requests.length = 0;
    boundary.uploads.length = 0;
    boundary.finalStatus = 409;
    const upload = service.sideLoadingV2("mos-token", "app.zip", AppScope.Personal);
    await expect(upload).rejects.toMatchObject({
      name: "SkillUploadConflict",
      source: "M365",
    });
    const failure: unknown = await upload.catch((error: unknown) => error);
    expect(JSON.stringify(failure)).not.toContain("private conflict response");
    expect(
      boundary.requests.filter((request) => request === "POST /builder/v1/users/packages")
    ).toHaveLength(1);
    expect(boundary.requests.some((request) => request.includes("/packages/status/"))).toBe(false);
    expect(boundary.submitted).toHaveLength(2);
    const submissions = http
      .flatMap((adapter) => adapter.history.post)
      .filter((request) => request.url === "/builder/v1/users/packages");
    expect(submissions).toHaveLength(2);
    expect(
      submissions.every((request) => request.timeout === 0 && request.maxRedirects === 0)
    ).toBe(true);
    expect(JSON.parse(new AdmZip(boundary.submitted[1]).readAsText("manifest.json")).version).toBe(
      "2.0.0"
    );
    expect(boundary.uploads).toHaveLength(3);
    expect(bytes).toEqual(original);
  });

  it("AC-14: legacy sideLoading entry delegates through the same real incremental pipeline", async () => {
    await service.sideLoading("mos-token", __filename, AppScope.Personal);
    expect(boundary.uploads).toHaveLength(3);
    expect(boundary.submitted).toHaveLength(1);
    expect(new AdmZip(boundary.submitted[0]).readFile("skills/example/SKILL.md")).toBeNull();
  });

  it.each(
    ["extendToM365", "publish"].flatMap((operation) =>
      [false, true].map((rootSkills) => ({ operation, rootSkills }))
    )
  )(
    "SCN-ISU-07 SCN-ISU-09 AC-14 AC-24 AC-30: real $operation driver executes noninteractively without UI ($rootSkills root skills)",
    async ({ operation, rootSkills }) => {
      if (rootSkills) bytes = rootSkillPackage().toBuffer();
      const token = vi.spyOn(provider, "getAccessToken");
      const context: DriverContext = {
        azureAccountProvider: new MockedAzureAccountProvider(),
        m365TokenProvider: provider,
        ui: undefined,
        nonInteractive: true,
        progressBar: undefined,
        logProvider: new MockLogProvider(),
        projectPath: process.cwd(),
        platform: Platform.CLI,
      };
      const outputs = new Map([
        ["titleId", "M365_TITLE_ID"],
        ["appId", "M365_APP_ID"],
      ]);
      const args = { appPackagePath: __filename };
      const result =
        operation === "extendToM365"
          ? await new M365TitleAcquireDriver().execute(args, context, outputs)
          : await new CopilotAgentPublishDriver().execute(args, context, outputs);
      if (result.result.isErr()) throw result.result.error;
      if (result.result.isOk()) {
        expect(result.result.value.get("M365_TITLE_ID")).toBe("title");
        expect(result.result.value.get("M365_APP_ID")).toBe(agentId);
      }
      expect(token).toHaveBeenCalledWith({
        scopes: ["https://graph.microsoft.com/Files.ReadWrite"],
      });
      expect(boundary.uploads).toHaveLength(3);
      expect(boundary.submitted).toHaveLength(1);
      expect(new AdmZip(boundary.submitted[0]).readFile("skills/example/SKILL.md")).toBeNull();
    }
  );
});
