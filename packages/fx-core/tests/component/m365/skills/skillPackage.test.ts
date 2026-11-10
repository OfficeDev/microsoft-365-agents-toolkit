// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { DeclarativeAgentManifestConverter, TeamsManifestConverter } from "@microsoft/teamsfx-api";
import fs from "fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import teamsManifest from "../../../plugins/resource/daTemplate/da-no-action-test-template/appPackage/manifest.json";
import agentManifest from "../../../plugins/resource/daTemplate/da-no-action-test-template/appPackage/declarativeAgent.json";
import {
  buildSkillTransportPackage,
  readSkillPackage,
  SkillPackage,
  validateSkillTransportPackage,
} from "../../../../src/component/m365/skills/skillPackage";

const agentId = "c24d7903-a3b2-49b8-8bda-a4cf639fb25d";
const skillMd = "---\nname: example\ndescription: Example skill\n---\n\nOriginal instructions.\n";

function packageZip(
  agent: Record<string, unknown> = { agent_skills: [{ folder: "skills/example" }] },
  manifest: Record<string, unknown> = {}
): AdmZip {
  const zip = new AdmZip();
  zip.addFile(
    "manifest.json",
    Buffer.from(JSON.stringify({ ...teamsManifest, id: agentId, ...manifest }))
  );
  zip.addFile("declarativeAgent.json", Buffer.from(JSON.stringify({ ...agentManifest, ...agent })));
  zip.addFile("skills/example/SKILL.md", Buffer.from(skillMd));
  zip.addFile("color.png", Buffer.from([0, 255, 128]));
  return zip;
}

function addRawName(zip: AdmZip, name: string, content = Buffer.from("unsafe")): void {
  zip.addFile("raw-entry", content);
  const entry = zip.getEntry("raw-entry");
  if (!entry) throw new Error("Fixture entry missing");
  entry.entryName = name;
}

function expectInvalid(zip: AdmZip): void {
  const result = readSkillPackage(zip.toBuffer());
  expect(result.isErr()).toBe(true);
  if (result.isErr()) {
    expect(result.error.name).toBe("InvalidSkillPackage");
    expect(result.error.source).toBe("M365");
  }
}

describe("skill package inventory", () => {
  afterEach(() => vi.restoreAllMocks());

  it("AC-03: inventories nested, binary and empty files with unchanged SKILL.md bytes", () => {
    const zip = packageZip();
    zip.addFile("skills/example/assets/binary.bin", Buffer.from([0, 255, 128, 13, 10]));
    zip.addFile("skills/example/empty", Buffer.alloc(0));
    zip.addFile("skills/example/empty-folder/", Buffer.alloc(0));
    const result = readSkillPackage(zip.toBuffer());
    expect(result.isOk()).toBe(true);
    if (result.isErr()) throw result.error;
    expect(result.value.agentId).toBe(agentId);
    expect(result.value.skills).toEqual([
      {
        folder: "skills/example",
        name: "example",
        description: "Example skill",
        files: [
          { path: "SKILL.md", content: Buffer.from(skillMd) },
          { path: "assets/binary.bin", content: Buffer.from([0, 255, 128, 13, 10]) },
          { path: "empty", content: Buffer.alloc(0) },
        ],
      },
    ]);
  });

  const remoteFolder = "11111111-2222-3333-4444-555555555555";
  const secondRemoteFolder = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";

  function inventory(zip: AdmZip): SkillPackage {
    const result = readSkillPackage(zip.toBuffer());
    if (result.isErr()) throw result.error;
    return result.value;
  }

  function readZipJson(zip: AdmZip, path: string): unknown {
    const bytes = zip.readFile(path);
    if (!bytes) throw new Error("Fixture JSON missing");
    return JSON.parse(bytes.toString("utf8"));
  }

  describe("skill transport", () => {
    afterEach(() => vi.restoreAllMocks());

    it("AC-12: mapping-independent preflight accepts remote GUID metadata but final build binds it", () => {
      const zip = packageZip();
      zip.addFile(
        "packageMetadata.json",
        Buffer.from(
          JSON.stringify({
            skills: [
              {
                folder: remoteFolder,
                name: "example",
                description: "Old description",
                extra: true,
              },
            ],
          })
        )
      );
      const source = inventory(zip);
      const before = JSON.stringify(source);
      const write = vi.spyOn(fs, "writeFileSync");
      const result = validateSkillTransportPackage(source);
      expect(result.isOk()).toBe(true);
      if (result.isErr()) throw result.error;
      expect(result.value).toBeUndefined();
      expect(JSON.stringify(source)).toBe(before);
      expect(write).not.toHaveBeenCalled();
      expect(
        buildSkillTransportPackage(
          source,
          new Map([["skills/example", secondRemoteFolder]])
        ).isErr()
      ).toBe(true);
      const built = buildSkillTransportPackage(source, new Map([["skills/example", remoteFolder]]));
      expect(built.isOk()).toBe(true);
      if (built.isErr()) throw built.error;
      expect(readZipJson(new AdmZip(built.value), "packageMetadata.json")).toEqual({
        builderInfo: { builderName: "TeamsToolKit" },
        skills: [
          { folder: remoteFolder, name: "example", description: "Example skill", extra: true },
        ],
      });
    });

    it.each([
      "{ invalid",
      "[]",
      JSON.stringify({ builderInfo: "invalid" }),
      JSON.stringify({ skills: {} }),
      JSON.stringify({ skills: [null] }),
      JSON.stringify({ skills: [{ folder: "" }] }),
      JSON.stringify({ skills: [{ folder: "other", name: "example" }] }),
      JSON.stringify({ skills: [{ folder: "skills/example", name: "wrong" }] }),
      JSON.stringify({ skills: [{ folder: remoteFolder, name: 1 }] }),
      JSON.stringify({ skills: [{ folder: remoteFolder, name: "unknown" }] }),
      JSON.stringify({ skills: [{ folder: remoteFolder, description: {} }] }),
      JSON.stringify({ skills: [{ folder: "skills/example" }, { folder: "./skills/example" }] }),
      JSON.stringify({ skills: [{ folder: remoteFolder }, { folder: remoteFolder }] }),
    ])("AC-12: preflight rejects malformed or locally unmatched metadata %#", (metadata) => {
      const zip = packageZip();
      zip.addFile("packageMetadata.json", Buffer.from(metadata));
      expect(validateSkillTransportPackage(inventory(zip)).isErr()).toBe(true);
    });

    it("AC-12: preflight rejects unsupported artifact overlap before any remote mappings exist", () => {
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example", exposure: { enabled: true } }],
        actions: [{ id: "plugin", file: "plugin.json" }],
      });
      zip.addFile(
        "plugin.json",
        Buffer.from(JSON.stringify({ file: "skills/example/openapi.yaml" }))
      );
      expect(validateSkillTransportPackage(inventory(zip)).isErr()).toBe(true);
    });

    it("AC-12: preflight checks deferred remote metadata extras for removal conflicts", () => {
      const zip = packageZip();
      zip.addFile(
        "packageMetadata.json",
        Buffer.from(
          JSON.stringify({ skills: [{ folder: remoteFolder, file: "skills/example/data.json" }] })
        )
      );
      expect(validateSkillTransportPackage(inventory(zip)).isErr()).toBe(true);
    });

    it("AC-12: preflight rejects metadata path collisions", () => {
      const zip = packageZip();
      zip.addFile("packageMetadata.json/child.txt", Buffer.from("keep"));
      expect(validateSkillTransportPackage(inventory(zip)).isErr()).toBe(true);
    });

    it("AC-11: preflight accepts local metadata and no-skill packages without modifying source", () => {
      const local = inventory(packageZip());
      const empty = inventory(packageZip({}));
      expect(validateSkillTransportPackage(local).isOk()).toBe(true);
      expect(validateSkillTransportPackage(empty).isOk()).toBe(true);
    });

    it("AC-11 SCN-ISU: first and repeated transports use remote folder GUIDs and metadata only", () => {
      const readRoot = vi.spyOn(TeamsManifestConverter, "jsonToManifest");
      const readAgent = vi.spyOn(DeclarativeAgentManifestConverter, "jsonToManifest");
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example", exposure: { enabled: false } }],
        "x-agent_skills": [{ folder: "./skills/example", exposure: { audience: ["user"] } }],
        "x-extra": { retain: true },
      });
      zip.addFile("skills/example/assets/binary.bin", Buffer.from([0, 255, 128]));
      zip.addFile("skills/example/empty", Buffer.alloc(0));
      zip.addFile("skills/example/", Buffer.alloc(0));
      zip.addFile("unrelated.bin", Buffer.from([1, 254]));
      const input = inventory(zip);
      const before = JSON.stringify(input);
      expect(validateSkillTransportPackage(input).isOk()).toBe(true);
      expect(JSON.stringify(input)).toBe(before);
      const write = vi.spyOn(fs, "writeFile");
      const writeSync = vi.spyOn(fs, "writeFileSync");
      const mkdir = vi.spyOn(fs, "mkdirSync");
      for (const folder of [remoteFolder, secondRemoteFolder]) {
        const result = buildSkillTransportPackage(input, new Map([["skills/example", folder]]));
        expect(result.isOk()).toBe(true);
        if (result.isErr()) throw result.error;
        const transport = new AdmZip(result.value);
        expect(readZipJson(transport, "packageMetadata.json")).toEqual({
          builderInfo: { builderName: "TeamsToolKit" },
          skills: [{ folder, name: "example", description: "Example skill" }],
        });
        expect(readZipJson(transport, "declarativeAgent.json")).toEqual({
          ...agentManifest,
          agent_skills: [{ folder, exposure: { enabled: false } }],
          "x-agent_skills": [{ folder, exposure: { audience: ["user"] } }],
          "x-extra": { retain: true },
        });
        expect(
          transport.getEntries().some((entry) => entry.entryName.startsWith("skills/example"))
        ).toBe(false);
        expect(transport.readFile("manifest.json")).toEqual(zip.readFile("manifest.json"));
        expect(transport.readFile("color.png")).toEqual(zip.readFile("color.png"));
        expect(transport.readFile("unrelated.bin")).toEqual(Buffer.from([1, 254]));
        expect(result.value.toString("utf8")).not.toContain("Original instructions");
        expect(JSON.stringify(input)).toBe(before);
      }
      expect(write).not.toHaveBeenCalled();
      expect(writeSync).not.toHaveBeenCalled();
      expect(mkdir).not.toHaveBeenCalled();
      expect(readRoot).toHaveBeenCalledTimes(1);
      expect(readAgent).toHaveBeenCalledTimes(1);
    });

    it("AC-11: transforms every current and legacy agent sharing the same local folder", () => {
      const zip = packageZip(
        {},
        {
          copilotExtensions: {
            declarativeCopilots: [{ id: "legacy", file: "legacy.json" }],
          },
        }
      );
      zip.addFile(
        "legacy.json",
        Buffer.from(
          JSON.stringify({ ...agentManifest, "x-agent_skills": [{ folder: "skills/example" }] })
        )
      );
      const input = inventory(zip);
      const result = buildSkillTransportPackage(input, new Map([["skills/example", remoteFolder]]));
      expect(result.isOk()).toBe(true);
      if (result.isErr()) throw result.error;
      const transport = new AdmZip(result.value);
      expect(readZipJson(transport, "legacy.json")).toEqual({
        ...agentManifest,
        "x-agent_skills": [{ folder: remoteFolder }],
      });
      expect(transport.readFile("declarativeAgent.json")).toEqual(
        zip.readFile("declarativeAgent.json")
      );
    });

    it.each(["./skills/example/", remoteFolder])(
      "AC-11: merges existing package metadata and retains matched entry extras (%s)",
      (folder) => {
        const zip = packageZip();
        zip.addFile(
          "packageMetadata.json",
          Buffer.from(
            JSON.stringify({
              unrelated: { retain: true },
              builderInfo: { builderName: "old", buildVersion: "1.2" },
              skills: [
                { folder, name: "example", description: "stale", exposure: { internal: true } },
              ],
            })
          )
        );
        const result = buildSkillTransportPackage(
          inventory(zip),
          new Map([["skills/example", remoteFolder]])
        );
        expect(result.isOk()).toBe(true);
        if (result.isErr()) throw result.error;
        expect(readZipJson(new AdmZip(result.value), "packageMetadata.json")).toEqual({
          unrelated: { retain: true },
          builderInfo: { builderName: "TeamsToolKit", buildVersion: "1.2" },
          skills: [
            {
              folder: remoteFolder,
              name: "example",
              description: "Example skill",
              exposure: { internal: true },
            },
          ],
        });
      }
    );

    it("AC-11: metadata contains the entire inventory even if existing metadata covers only one skill", () => {
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example" }, { folder: "skills/second" }],
      });
      zip.addFile(
        "skills/second/SKILL.md",
        Buffer.from("---\nname: second\ndescription: Second skill\n---\n")
      );
      zip.addFile(
        "packageMetadata.json",
        Buffer.from(
          JSON.stringify({ skills: [{ folder: "skills/example", name: "example", extra: true }] })
        )
      );
      const result = buildSkillTransportPackage(
        inventory(zip),
        new Map([
          ["skills/example", remoteFolder],
          ["skills/second", secondRemoteFolder],
        ])
      );
      expect(result.isOk()).toBe(true);
      if (result.isErr()) throw result.error;
      expect(readZipJson(new AdmZip(result.value), "packageMetadata.json")).toEqual({
        builderInfo: { builderName: "TeamsToolKit" },
        skills: [
          { folder: remoteFolder, name: "example", description: "Example skill", extra: true },
          { folder: secondRemoteFolder, name: "second", description: "Second skill" },
        ],
      });
    });

    it("AC-11: no skills preserves exact original ZIP bytes without adding metadata", () => {
      const input = inventory(packageZip({}));
      const result = buildSkillTransportPackage(input, new Map());
      expect(result.isOk()).toBe(true);
      if (result.isErr()) throw result.error;
      expect(result.value).toEqual(input.data);
      expect(new AdmZip(result.value).getEntry("packageMetadata.json")).toBeNull();
    });

    it.each(
      [
        [],
        [["skills/example", "not-a-guid"]],
        [["skills/example", `{${remoteFolder}}`]],
        [["skills/example", ` ${remoteFolder}`]],
        [["skills/example", `https://example.com/${remoteFolder}`]],
        [["other/example", remoteFolder]],
        [
          ["skills/example", remoteFolder],
          ["extra", secondRemoteFolder],
        ],
      ].map((entries) => ({ entries }))
    )("AC-12: rejects incomplete, extraneous or non-GUID mappings %#", ({ entries }) => {
      const folders = new Map<string, string>();
      for (const entry of entries) folders.set(entry[0], entry[1]);
      expect(buildSkillTransportPackage(inventory(packageZip()), folders).isErr()).toBe(true);
    });

    it("AC-12: rejects duplicate remote GUID mappings case-insensitively", () => {
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example" }, { folder: "skills/second" }],
      });
      zip.addFile(
        "skills/second/SKILL.md",
        Buffer.from("---\nname: second\ndescription: Second skill\n---\n")
      );
      expect(
        buildSkillTransportPackage(
          inventory(zip),
          new Map([
            ["skills/example", secondRemoteFolder],
            ["skills/second", secondRemoteFolder.toUpperCase()],
          ])
        ).isErr()
      ).toBe(true);
    });

    it.each([
      "{ invalid",
      "[]",
      JSON.stringify({ builderInfo: "invalid" }),
      JSON.stringify({ skills: {} }),
      JSON.stringify({ skills: [null] }),
      JSON.stringify({ skills: [{ folder: "other", name: "example" }] }),
      JSON.stringify({ skills: [{ folder: "skills/example", name: "wrong" }] }),
      JSON.stringify({ skills: [{ folder: "skills/example" }, { folder: "./skills/example" }] }),
      JSON.stringify({ skills: [{ folder: "skills/example" }, { folder: remoteFolder }] }),
    ])("AC-12: rejects malformed, ambiguous or unmatched existing metadata %#", (metadata) => {
      const zip = packageZip();
      zip.addFile("packageMetadata.json", Buffer.from(metadata));
      expect(
        buildSkillTransportPackage(
          inventory(zip),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
    });

    it.each([
      { actions: [{ id: "plugin", file: "skills/example/plugin.json" }] },
      { instructions: "skills/example/instructions.txt" },
      { "x-connectors": [{ file: "./skills/example/connector.json" }] },
      {
        capabilities: [
          { name: "OneDriveAndSharePoint", items_by_url: [{ url: "skills/example/data.txt" }] },
        ],
      },
    ])("AC-12: rejects other agent artifact references into removed roots %#", (extra) => {
      const zip = packageZip({ agent_skills: [{ folder: "skills/example" }], ...extra });
      expect(
        buildSkillTransportPackage(
          inventory(zip),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
    });

    it("AC-12: rejects manifest icon references and declarations overlapping removed skill roots", () => {
      const icons = packageZip(undefined, {
        icons: { color: "skills/example/color.png", outline: "outline.png" },
      });
      expect(
        buildSkillTransportPackage(
          inventory(icons),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
      const declaration = packageZip(undefined, {
        copilotAgents: { declarativeAgents: [{ id: "nested", file: "skills/example/agent.json" }] },
      });
      declaration.addFile(
        "skills/example/agent.json",
        Buffer.from(
          JSON.stringify({ ...agentManifest, agent_skills: [{ folder: "skills/example" }] })
        )
      );
      expect(
        buildSkillTransportPackage(
          inventory(declaration),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
    });

    it("AC-12: rejects transitive plugin references into removed skill roots", () => {
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example" }],
        actions: [{ id: "plugin", file: "plugins/plugin.json" }],
      });
      zip.addFile(
        "plugins/plugin.json",
        Buffer.from(
          JSON.stringify({ runtimes: [{ spec: { url: "../skills/example/openapi.yaml" } }] })
        )
      );
      expect(
        buildSkillTransportPackage(
          inventory(zip),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
    });

    it.each([
      "skills/example/100%invalid.txt",
      "/skills/example/absolute.txt",
      "skills%2fexample%2fencoded.txt",
      "skills\\example\\windows.txt",
      "$[file('skills/example/instructions.txt')]",
    ])("AC-12: rejects normalized non-skill overlap references: %s", (reference) => {
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example" }],
        "x-file-reference": reference,
      });
      expect(
        buildSkillTransportPackage(
          inventory(zip),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
    });

    it("AC-11: preserves unreferenced malformed JSON and does not follow remote URLs", () => {
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example" }],
        "x-remote": "https://example.com/skills/example/openapi.json",
      });
      zip.addFile("unrelated.json", Buffer.from("not JSON; preserve these bytes"));
      const result = buildSkillTransportPackage(
        inventory(zip),
        new Map([["skills/example", remoteFolder.toUpperCase()]])
      );
      expect(result.isOk()).toBe(true);
      if (result.isErr()) throw result.error;
      expect(new AdmZip(result.value).readFile("unrelated.json")).toEqual(
        zip.readFile("unrelated.json")
      );
    });

    it("AC-12: rejects transitive OpenAPI YAML references into removed skill roots", () => {
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example" }],
        actions: [{ id: "plugin", file: "plugin.json" }],
      });
      zip.addFile(
        "plugin.json",
        Buffer.from(JSON.stringify({ runtimes: [{ spec: { url: "spec/openapi.yaml" } }] }))
      );
      zip.addFile(
        "spec/openapi.yaml",
        Buffer.from(
          "components:\n  schemas:\n    input:\n      $ref: ../skills/example/schema.json\n"
        )
      );
      expect(
        buildSkillTransportPackage(
          inventory(zip),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
    });

    it("AC-12: rejects extraneous mappings even when the package has no skills", () => {
      expect(
        buildSkillTransportPackage(
          inventory(packageZip({})),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
    });

    it.each(["PackageMetadata.json", "packageMetadata.json/"])(
      "AC-12: rejects ambiguous reserved metadata entry %s",
      (path) => {
        const zip = packageZip();
        zip.addFile(path, Buffer.from("{}"));
        expect(
          buildSkillTransportPackage(
            inventory(zip),
            new Map([["skills/example", remoteFolder]])
          ).isErr()
        ).toBe(true);
      }
    );

    it("AC-12: rejects an implicit directory occupying the metadata file path", () => {
      const zip = packageZip();
      zip.addFile("packageMetadata.json/child.txt", Buffer.from("keep"));
      expect(
        buildSkillTransportPackage(
          inventory(zip),
          new Map([["skills/example", remoteFolder]])
        ).isErr()
      ).toBe(true);
    });

    it("AC-11: follows referenced cyclic YAML safely without changing its bytes", () => {
      const zip = packageZip({
        agent_skills: [{ folder: "skills/example" }],
        "x-spec": "openapi.yaml",
      });
      zip.addFile("openapi.yaml", Buffer.from("schema: &schema\n  recursive: *schema\n"));
      const result = buildSkillTransportPackage(
        inventory(zip),
        new Map([["skills/example", remoteFolder]])
      );
      expect(result.isOk()).toBe(true);
      if (result.isErr()) throw result.error;
      expect(new AdmZip(result.value).readFile("openapi.yaml")).toEqual(
        zip.readFile("openapi.yaml")
      );
    });
  });

  it("AC-11: preserves source bytes and unrelated fields without filesystem writes or extraction", () => {
    const zip = packageZip({
      agent_skills: [{ folder: "./skills/example/", exposure: { enabled: false } }],
      "x-unrelated": { nested: ["keep", 42] },
    });
    const bytes = zip.toBuffer();
    const original = Buffer.from(bytes);
    const writeSync = vi.spyOn(fs, "writeFileSync");
    const write = vi.spyOn(fs, "writeFile");
    const mkdir = vi.spyOn(fs, "mkdirSync");
    const result = readSkillPackage(bytes);
    expect(result.isOk()).toBe(true);
    if (result.isErr()) throw result.error;
    expect(bytes).toEqual(original);
    expect(result.value.data).toEqual(original);
    expect(result.value.declarativeAgents[0].manifest["x-unrelated"]).toEqual({
      nested: ["keep", 42],
    });
    expect(result.value.declarativeAgents[0].manifest.agent_skills).toEqual([
      { folder: "./skills/example/", exposure: { enabled: false } },
    ]);
    expect(result.value.references).toEqual([
      {
        manifestPath: "declarativeAgent.json",
        property: "agent_skills",
        index: 0,
        folder: "skills/example",
      },
    ]);
    expect(new AdmZip(result.value.data).readFile("color.png")).toEqual(Buffer.from([0, 255, 128]));
    expect(writeSync).not.toHaveBeenCalled();
    expect(write).not.toHaveBeenCalled();
    expect(mkdir).not.toHaveBeenCalled();
  });

  it("AC-11: visits both agent lists and both skill properties, deduplicating shared roots", () => {
    const zip = packageZip(
      {
        agent_skills: [{ folder: "skills/example" }],
        "x-agent_skills": [{ folder: "./skills/example" }],
      },
      {
        copilotAgents: {
          declarativeAgents: [
            { id: "first", file: "declarativeAgent.json" },
            { id: "second", file: "agents/second.json" },
          ],
        },
        copilotExtensions: {
          declarativeCopilots: [{ id: "legacy", file: "legacy.json" }],
        },
      }
    );
    const agent = {
      ...agentManifest,
      "x-agent_skills": [{ folder: "skills/example", exposure: { name: "retained" } }],
    };
    zip.addFile("agents/second.json", Buffer.from(JSON.stringify(agent)));
    zip.addFile("legacy.json", Buffer.from(JSON.stringify(agent)));
    const result = readSkillPackage(zip.toBuffer());
    expect(result.isOk()).toBe(true);
    if (result.isErr()) throw result.error;
    expect(result.value.skills).toHaveLength(1);
    expect(result.value.declarativeAgents.map((agent) => agent.path)).toEqual([
      "declarativeAgent.json",
      "agents/second.json",
      "legacy.json",
    ]);
    expect(result.value.references).toHaveLength(4);
  });

  it("AC-02: returns an empty inventory for an agent with no skill properties", () => {
    const result = readSkillPackage(packageZip({}).toBuffer());
    expect(result.isOk()).toBe(true);
    if (result.isErr()) throw result.error;
    expect(result.value.skills).toEqual([]);
  });

  it("AC-02: leaves nonlocal references and unsupported unrelated extensions untouched", () => {
    const reference = { remote: { provider: "future" }, exposure: { enabled: true } };
    const result = readSkillPackage(
      packageZip({ agent_skills: [reference], "x-custom-extension": { opaque: true } }).toBuffer()
    );
    expect(result.isOk()).toBe(true);
    if (result.isErr()) throw result.error;
    expect(result.value.skills).toEqual([]);
    expect(result.value.references).toEqual([]);
    expect(result.value.declarativeAgents[0].manifest.agent_skills).toEqual([reference]);
  });

  it("AC-11: preserves BOM/CRLF frontmatter, folded descriptions and optional exposure fields", () => {
    const zip = packageZip();
    const source =
      "\ufeff---\r\nname: example\r\ndescription: >-\r\n  Example\r\n  skill\r\n" +
      "exposure:\r\n  internal: true\r\n---\r\nbody\r\n";
    zip.addFile("skills/example/SKILL.md", Buffer.from(source));
    const result = readSkillPackage(zip.toBuffer());
    expect(result.isOk()).toBe(true);
    if (result.isErr()) throw result.error;
    expect(result.value.skills[0].description).toBe("Example skill");
    expect(result.value.skills[0].files[0].content).toEqual(Buffer.from(source));
  });

  it("AC-12: accepts maximum skill name and description lengths", () => {
    const name = "a".repeat(64);
    const zip = packageZip({ agent_skills: [{ folder: `skills/${name}` }] });
    zip.addFile(
      `skills/${name}/SKILL.md`,
      Buffer.from(`---\nname: ${name}\ndescription: ${"d".repeat(1024)}\n---\n`)
    );
    expect(readSkillPackage(zip.toBuffer()).isOk()).toBe(true);
  });

  it.each(["manifest.json", "declarativeAgent.json", "skills/example/SKILL.md"])(
    "AC-12: rejects invalid UTF-8 metadata bytes in %s",
    (file) => {
      const zip = packageZip();
      zip.addFile(file, Buffer.from([0xff]));
      expectInvalid(zip);
    }
  );

  it("AC-12: rejects invalid ZIP bytes", () => {
    expect(readSkillPackage(Buffer.from("not a ZIP")).isErr()).toBe(true);
  });

  it.each([undefined, "", "   ", 42, null])(
    "AC-12: requires nonempty Teams app version for unknown manifest schema: %j",
    (version) => {
      expectInvalid(packageZip({}, { manifestVersion: "future-schema", version }));
    }
  );

  it("AC-12: accepts a nonempty Teams app version for unknown manifest schema", () => {
    expect(
      readSkillPackage(
        packageZip({}, { manifestVersion: "future-schema", version: "2.1.0" }).toBuffer()
      ).isOk()
    ).toBe(true);
  });

  it.each(["manifest.json", "declarativeAgent.json"])(
    "AC-12: rejects missing referenced manifest %s",
    (file) => {
      const zip = packageZip();
      zip.deleteFile(file);
      expectInvalid(zip);
    }
  );

  it.each(["manifest.json", "declarativeAgent.json"])(
    "AC-12: rejects malformed JSON in %s",
    (file) => {
      const zip = packageZip();
      zip.addFile(file, Buffer.from("{ invalid"));
      expectInvalid(zip);
    }
  );

  it.each([null, [], "text", 42])("AC-12: rejects a non-object agent manifest: %j", (value) => {
    const zip = packageZip();
    zip.addFile("declarativeAgent.json", Buffer.from(JSON.stringify(value)));
    expectInvalid(zip);
  });

  it("AC-12: runs validating converters for known manifest versions", () => {
    expectInvalid(packageZip({}, { icons: "not-an-icons-object" }));
    expectInvalid(packageZip({ capabilities: { name: "WebSearch" } }));
  });

  it.each([undefined, "", 42])("AC-12: rejects invalid agent id: %j", (id) => {
    expectInvalid(packageZip({}, { id }));
  });

  it.each([
    "../escape",
    "skills/../escape",
    "/absolute",
    "C:/absolute",
    "\\\\server\\share",
    "skills\\example",
    "skills/%2e%2e/escape",
    "skills/example/evil:stream",
    "skills/example/evil\u0000",
    "skills/example/trailing.",
    "skills/example/trailing ",
    "skills/example/CON",
  ])("AC-12: rejects unsafe ZIP path %j", (name) => {
    const zip = packageZip();
    addRawName(zip, name);
    expectInvalid(zip);
  });

  it.each(["./skills/example/SKILL.md", "skills/example/skill.md", "skills//example/SKILL.md"])(
    "AC-12: rejects normalized or case-colliding ZIP names: %s",
    (name) => {
      const zip = packageZip();
      addRawName(zip, name);
      expectInvalid(zip);
    }
  );

  it("AC-12: rejects exact duplicate ZIP entries", () => {
    const zip = packageZip();
    addRawName(zip, "skills/example/SKILL.md");
    expectInvalid(zip);
  });

  it("AC-12: rejects case-ambiguous directory segments and file-directory collisions", () => {
    const caseZip = packageZip();
    caseZip.addFile("skills/EXAMPLE/other.txt", Buffer.from("other"));
    expectInvalid(caseZip);
    const fileZip = packageZip();
    fileZip.addFile("skills/example/assets", Buffer.from("file"));
    fileZip.addFile("skills/example/assets/nested", Buffer.from("nested"));
    expectInvalid(fileZip);
  });

  it("AC-12: rejects symlink entries without following them", () => {
    const zip = packageZip();
    zip.addFile("skills/example/link", Buffer.from("../outside"));
    const entry = zip.getEntry("skills/example/link");
    if (!entry) throw new Error("Fixture entry missing");
    entry.attr = (0o120777 << 16) >>> 0;
    expectInvalid(zip);
  });

  it.each([
    "../skills/example",
    "/skills/example",
    "C:\\skills\\example",
    ".",
    "https://host/skill",
  ])("AC-12: rejects unsafe referenced folder %s", (folder) => {
    expectInvalid(packageZip({ agent_skills: [{ folder }] }));
  });

  it("AC-12: rejects unsafe agent manifest references", () => {
    expectInvalid(
      packageZip(
        {},
        { copilotAgents: { declarativeAgents: [{ id: "one", file: "../agent.json" }] } }
      )
    );
  });

  it.each([
    { copilotAgents: null },
    { copilotAgents: { declarativeAgents: {} } },
    { copilotAgents: { declarativeAgents: [null] } },
    { copilotAgents: { declarativeAgents: [{ file: "declarativeAgent.json" }] } },
    { copilotExtensions: { declarativeCopilots: "not-a-list" } },
  ])("AC-12: rejects malformed current or legacy agent lists: %j", (manifest) => {
    expectInvalid(packageZip({}, manifest));
  });

  it.each([null, {}, "skills/example", [{ folder: 42 }], [null], [{ folder: "" }]])(
    "AC-12: rejects malformed skill references: %j",
    (agent_skills) => {
      expectInvalid(packageZip({ agent_skills }));
    }
  );

  it("AC-12: rejects missing SKILL.md or mismatched filename casing", () => {
    const zip = packageZip();
    zip.deleteFile("skills/example/SKILL.md");
    expectInvalid(zip);
    zip.addFile("skills/example/skill.md", Buffer.from(skillMd));
    expectInvalid(zip);
  });

  it.each([
    "body without frontmatter",
    "---\nname: [invalid\n---\n",
    "---\n- example\n---\n",
    "---\nname: example\nname: duplicate\ndescription: test\n---\n",
    "---\nname: different\ndescription: test\n---\n",
    "---\nname: example\n---\n",
    "---\nname: example\ndescription: '   '\n---\n",
    "---\nname: example\ndescription: 12\n---\n",
    `---\nname: example\ndescription: ${"d".repeat(1025)}\n---\n`,
  ])("AC-12: rejects invalid YAML/frontmatter %#", (content) => {
    const zip = packageZip();
    zip.addFile("skills/example/SKILL.md", Buffer.from(content));
    expectInvalid(zip);
  });

  it.each(["UPPER", "-leading", "trailing-", "double--hyphen", "under_score", "a".repeat(65)])(
    "AC-12: rejects invalid skill name %s",
    (name) => {
      const zip = packageZip({ agent_skills: [{ folder: `skills/${name}` }] });
      zip.addFile(
        `skills/${name}/SKILL.md`,
        Buffer.from(`---\nname: ${name}\ndescription: valid\n---\n`)
      );
      expectInvalid(zip);
    }
  );

  it("AC-12: rejects overlapping roots and colliding skill names", () => {
    const overlap = packageZip({
      agent_skills: [{ folder: "skills/example" }, { folder: "skills/example/nested" }],
    });
    overlap.addFile(
      "skills/example/nested/SKILL.md",
      Buffer.from("---\nname: nested\ndescription: valid\n---\n")
    );
    expectInvalid(overlap);
    const collision = packageZip({
      agent_skills: [{ folder: "skills/example" }, { folder: "other/example" }],
    });
    collision.addFile("other/example/SKILL.md", Buffer.from(skillMd));
    expectInvalid(collision);
  });

  it("AC-12: does not mistake similarly prefixed sibling folders for overlapping roots", () => {
    const zip = packageZip({
      agent_skills: [{ folder: "skills/example" }, { folder: "skills/example-two" }],
    });
    zip.addFile(
      "skills/example-two/SKILL.md",
      Buffer.from("---\nname: example-two\ndescription: valid\n---\n")
    );
    const result = readSkillPackage(zip.toBuffer());
    expect(result.isOk()).toBe(true);
    if (result.isErr()) throw result.error;
    expect(result.value.skills).toHaveLength(2);
  });
});
