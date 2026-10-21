// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  DeclarativeAgentManifestConverter,
  err,
  FxError,
  ok,
  Result,
  TeamsManifestConverter,
  UserError,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import yaml from "js-yaml";
import { posix } from "path";
import { TextDecoder } from "util";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";

export interface SkillPackageFile {
  /** Confined, slash-separated path relative to the skill folder. */
  path: string;
  content: Buffer;
}

export interface LocalSkill {
  folder: string;
  name: string;
  description: string;
  files: SkillPackageFile[];
}

export interface SkillReference {
  manifestPath: string;
  property: "agentSkills" | "agent_skills" | "x-agent_skills";
  index: number;
  folder: string;
}

export interface SkillAgentManifest {
  path: string;
  /** Original parsed fields, including extensions and exposure metadata. */
  manifest: Record<string, unknown>;
}

/** Validated inventory from readSkillPackage; consumers must not mutate its source data. */
export interface SkillPackage {
  /** Unmodified source ZIP, retained for a separately defined transport transformation. */
  data: Buffer;
  agentId: string;
  manifest: Record<string, unknown>;
  declarativeAgents: SkillAgentManifest[];
  references: SkillReference[];
  skills: LocalSkill[];
}

/** Build the verified MOS reference package without changing the portable source package. */
export function buildSkillTransportPackage(
  source: SkillPackage,
  folders: ReadonlyMap<string, string>
): Result<Buffer, FxError> {
  return processSkillTransportPackage(source, folders);
}

/** Validate before remote mutations; remote-GUID associations require the final builder mapping. */
export function validateSkillTransportPackage(source: SkillPackage): Result<void, FxError> {
  return processSkillTransportPackage(source).map(() => undefined);
}

function processSkillTransportPackage(
  input: SkillPackage,
  folders?: ReadonlyMap<string, string>
): Result<Buffer, FxError> {
  try {
    if (folders) {
      if (folders.size !== input.skills.length) return invalidPackage();
      const remoteIds = new Set<string>();
      for (const skill of input.skills) {
        const remote = folders.get(skill.folder);
        if (!remote || !isGuid(remote) || remoteIds.has(remote.toLowerCase())) {
          return invalidPackage();
        }
        remoteIds.add(remote.toLowerCase());
      }
    }
    if (input.skills.length === 0) return ok(Buffer.from(input.data));

    const zip = new AdmZip(input.data);
    const entries = new Map<string, AdmZip.IZipEntry>();
    for (const entry of zip.getEntries()) {
      const path = confinedPath(entry.entryName);
      if (!path || path.toLowerCase().startsWith("packagemetadata.json/")) {
        return invalidPackage();
      }
      entries.set(path, entry);
    }
    // Do not create a second, differently cased metadata entry or replace a directory.
    const metadataEntry = [...entries].find(
      ([path]) => path.toLowerCase() === "packagemetadata.json"
    );
    let metadata: Record<string, unknown> = {};
    if (metadataEntry) {
      if (metadataEntry[0] !== "packageMetadata.json" || metadataEntry[1].isDirectory) {
        return invalidPackage();
      }
      const parsed: unknown = JSON.parse(decodeText(metadataEntry[1].getData()));
      if (!isRecord(parsed)) return invalidPackage();
      metadata = parsed;
    }
    if (metadata.builderInfo !== undefined && !isRecord(metadata.builderInfo)) {
      return invalidPackage();
    }
    if (metadata.skills !== undefined && !Array.isArray(metadata.skills)) {
      return invalidPackage();
    }
    const existingSkills = new Map<string, Record<string, unknown>>();
    const deferredSkills: Record<string, unknown>[] = [];
    const metadataFolders = new Set<string>();
    const metadataNames = new Set<string>();
    if (Array.isArray(metadata.skills)) {
      if (metadata.skills.length > input.skills.length) return invalidPackage();
      for (const existing of metadata.skills) {
        if (!isRecord(existing) || typeof existing.folder !== "string") return invalidPackage();
        const existingFolder = existing.folder;
        const folder = confinedPath(existingFolder);
        if (
          !folder ||
          metadataFolders.has(folder.toLowerCase()) ||
          (existing.name !== undefined &&
            (typeof existing.name !== "string" ||
              !input.skills.some((skill) => skill.name === existing.name) ||
              metadataNames.has(existing.name))) ||
          (existing.description !== undefined &&
            (typeof existing.description !== "string" ||
              !existing.description.trim() ||
              existing.description.length > 1024))
        ) {
          return invalidPackage();
        }
        metadataFolders.add(folder.toLowerCase());
        if (typeof existing.name === "string") metadataNames.add(existing.name);
        const matches = input.skills.filter(
          (skill) =>
            skill.folder === folder ||
            folders?.get(skill.folder)?.toLowerCase() === existingFolder.toLowerCase()
        );
        if (!folders && matches.length === 0 && isGuid(existingFolder)) {
          deferredSkills.push(existing);
          continue;
        }
        const match = matches[0];
        if (
          matches.length !== 1 ||
          !match ||
          existingSkills.has(match.folder) ||
          (existing.name !== undefined && existing.name !== match.name)
        ) {
          return invalidPackage();
        }
        existingSkills.set(match.folder, existing);
      }
    }
    const transportMetadata = {
      ...metadata,
      builderInfo: {
        ...(isRecord(metadata.builderInfo) ? metadata.builderInfo : {}),
        builderName: "TeamsToolKit",
      },
      skills: [
        ...input.skills.map((skill) => ({
          ...existingSkills.get(skill.folder),
          folder: folders?.get(skill.folder),
          name: skill.name,
          description: skill.description,
        })),
        ...deferredSkills,
      ],
    };
    const changedManifests = new Map<string, Record<string, unknown>>();
    for (const reference of input.references) {
      const manifest =
        changedManifests.get(reference.manifestPath) ??
        (reference.manifestPath === "manifest.json"
          ? input.manifest
          : input.declarativeAgents.find((candidate) => candidate.path === reference.manifestPath)
              ?.manifest);
      if (!manifest) return invalidPackage();
      const skills = manifest[reference.property];
      if (!Array.isArray(skills) || !isRecord(skills[reference.index])) return invalidPackage();
      const updatedSkills = [...skills];
      updatedSkills[reference.index] = {
        ...skills[reference.index],
        folder: folders?.get(reference.folder),
      };
      changedManifests.set(reference.manifestPath, {
        ...manifest,
        [reference.property]: updatedSkills,
      });
    }
    const roots = input.skills.map((skill) => skill.folder.toLowerCase());
    const documents = new Map<string, unknown>([
      ["manifest.json", changedManifests.get("manifest.json") ?? input.manifest],
      ["packageMetadata.json", transportMetadata],
      ...input.declarativeAgents.map((agent): [string, unknown] => [
        agent.path,
        changedManifests.get(agent.path) ?? agent.manifest,
      ]),
    ]);
    if (!safeRemainingReferences(documents, entries, roots)) return invalidPackage();
    if (!folders) return ok(input.data);

    for (const [path, manifest] of changedManifests) {
      const entry = entries.get(path);
      if (!entry) return invalidPackage();
      zip.updateFile(entry, Buffer.from(JSON.stringify(manifest, undefined, 2)));
    }
    if (metadataEntry) {
      zip.updateFile(
        metadataEntry[1],
        Buffer.from(JSON.stringify(transportMetadata, undefined, 2))
      );
    } else {
      zip.addFile(
        "packageMetadata.json",
        Buffer.from(JSON.stringify(transportMetadata, undefined, 2))
      );
    }
    for (const [path, entry] of entries) {
      if (insideSkillRoot(path, roots)) zip.deleteFile(entry);
    }
    return ok(zip.toBuffer());
  } catch {
    return invalidPackage();
  }
}

function isGuid(value: string): boolean {
  return /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(value);
}

function insideSkillRoot(path: string, roots: string[]): boolean {
  const lower = path.toLowerCase();
  return roots.some((root) => lower === root || lower.startsWith(`${root}/`));
}

/** Follow local JSON/YAML artifact references, not remote URLs or arbitrary unreferenced files. */
function safeRemainingReferences(
  documents: Map<string, unknown>,
  entries: Map<string, AdmZip.IZipEntry>,
  roots: string[]
): boolean {
  const byLowerPath = new Map([...entries].map(([path, entry]) => [path.toLowerCase(), entry]));
  for (const [documentPath, document] of documents) {
    if (insideSkillRoot(documentPath, roots)) return false;
    const pending: unknown[] = [document];
    const visited = new Set<object>();
    while (pending.length) {
      const value = pending.pop();
      if (typeof value === "object" && value !== null) {
        if (visited.has(value)) continue;
        visited.add(value);
      }
      if (Array.isArray(value)) {
        pending.push(...value);
      } else if (isRecord(value)) {
        pending.push(...Object.values(value));
      } else if (typeof value === "string") {
        const reference = value.match(/^\$\[file\(['"](.+)['"]\)\]$/)?.[1] ?? value;
        if (/^[a-z][a-z0-9+.-]*:/i.test(reference)) continue;
        let decoded: string;
        try {
          decoded = decodeURIComponent(reference);
        } catch {
          decoded = reference;
        }
        const path = decoded.replace(/\\/g, "/").split(/[?#]/)[0];
        const candidates = [
          posix.normalize(path).replace(/^\.\//, ""),
          posix.normalize(posix.join(posix.dirname(documentPath), path)),
        ];
        for (const candidate of candidates) {
          if (insideSkillRoot(candidate, roots)) return false;
          const entry = byLowerPath.get(candidate.toLowerCase());
          if (!entry || entry.isDirectory || !/\.(json|ya?ml)$/i.test(candidate)) continue;
          const entryPath = confinedPath(entry.entryName);
          if (!entryPath || documents.has(entryPath)) continue;
          const text = decodeText(entry.getData());
          const parsed: unknown = /\.json$/i.test(candidate) ? JSON.parse(text) : yaml.load(text);
          documents.set(entryPath, parsed);
        }
      }
    }
  }
  return true;
}

/** Inventory and validate only; this function never extracts files or performs remote I/O. */
export function readSkillPackage(data: Buffer): Result<SkillPackage, FxError> {
  try {
    const entries = readPackageEntries(new AdmZip(data));
    if (entries.isErr()) return err(entries.error);
    const root = readRootManifest(entries.value);
    if (root.isErr()) return err(root.error);
    const agents = readAgentManifests(entries.value, root.value.manifest);
    if (agents.isErr()) return err(agents.error);

    const rootReferences = readSkillReferences(root.value.manifest, "manifest.json", "agentSkills");
    if (rootReferences.isErr()) return err(rootReferences.error);
    const references = rootReferences.value;
    const properties: SkillReference["property"][] = ["agent_skills", "x-agent_skills"];
    for (const agent of agents.value) {
      for (const property of properties) {
        const agentReferences = readSkillReferences(agent.manifest, agent.path, property);
        if (agentReferences.isErr()) return err(agentReferences.error);
        references.push(...agentReferences.value);
      }
    }
    const skills = readLocalSkills(entries.value, references);
    if (skills.isErr()) return err(skills.error);
    return ok({
      data: Buffer.from(data),
      ...root.value,
      declarativeAgents: agents.value,
      references,
      skills: skills.value,
    });
  } catch {
    // Parser errors may contain entire manifest/frontmatter values. Do not expose those values.
    return invalidPackage();
  }
}

function readPackageEntries(zip: AdmZip): Result<Map<string, AdmZip.IZipEntry>, FxError> {
  const entries = new Map<string, AdmZip.IZipEntry>();
  const paths = new Map<string, { path: string; directory: boolean }>();
  for (const entry of zip.getEntries()) {
    const path = confinedPath(entry.entryName);
    const mode = (entry.attr >>> 16) & 0xf000;
    if (!path || entries.has(path) || (mode !== 0 && mode !== 0x8000 && mode !== 0x4000)) {
      return invalidPackage();
    }
    const segments = path.split("/");
    for (let index = 0; index < segments.length; index++) {
      const prefix = segments.slice(0, index + 1).join("/");
      const directory = index < segments.length - 1 || entry.isDirectory;
      const existing = paths.get(prefix.toLowerCase());
      if (existing && (existing.path !== prefix || existing.directory !== directory)) {
        return invalidPackage();
      }
      paths.set(prefix.toLowerCase(), { path: prefix, directory });
    }
    entries.set(path, entry);
  }
  return ok(entries);
}

function readRootManifest(
  entries: ReadonlyMap<string, AdmZip.IZipEntry>
): Result<Pick<SkillPackage, "manifest" | "agentId">, FxError> {
  const entry = entries.get("manifest.json");
  if (!entry || entry.isDirectory) return invalidPackage();
  const manifest: unknown = JSON.parse(decodeText(entry.getData()));
  if (
    !isRecord(manifest) ||
    typeof manifest.id !== "string" ||
    !isGuid(manifest.id) ||
    typeof manifest.version !== "string" ||
    !manifest.version.trim() ||
    typeof manifest.manifestVersion !== "string" ||
    !manifest.manifestVersion.trim()
  ) {
    return invalidPackage();
  }
  // Guards cover future schemas; validate the two agent-list spellings separately below.
  TeamsManifestConverter.jsonToManifest(
    JSON.stringify({ ...manifest, copilotAgents: undefined, copilotExtensions: undefined })
  );
  return ok({ manifest, agentId: manifest.id });
}

function readAgentManifests(
  entries: ReadonlyMap<string, AdmZip.IZipEntry>,
  manifest: Record<string, unknown>
): Result<SkillAgentManifest[], FxError> {
  const result: SkillAgentManifest[] = [];
  const agentPaths = new Set<string>();
  const agentLists = [
    { container: manifest.copilotAgents, property: "declarativeAgents" },
    { container: manifest.copilotExtensions, property: "declarativeCopilots" },
  ];
  for (const { container, property } of agentLists) {
    if (container === undefined) continue;
    if (!isRecord(container)) return invalidPackage();
    const agents = container[property];
    if (agents === undefined) continue;
    if (!Array.isArray(agents)) return invalidPackage();
    for (const agent of agents) {
      if (
        !isRecord(agent) ||
        typeof agent.id !== "string" ||
        !agent.id.trim() ||
        typeof agent.file !== "string"
      ) {
        return invalidPackage();
      }
      const path = confinedPath(agent.file);
      if (!path) return invalidPackage();
      const entry = entries.get(path);
      if (!entry || entry.isDirectory) return invalidPackage();
      if (agentPaths.has(path)) continue;
      agentPaths.add(path);
      const text = decodeText(entry.getData());
      const agentManifest: unknown = JSON.parse(text);
      if (
        !isRecord(agentManifest) ||
        typeof agentManifest.version !== "string" ||
        !agentManifest.version.trim() ||
        typeof agentManifest.name !== "string" ||
        !agentManifest.name.trim()
      ) {
        return invalidPackage();
      }
      DeclarativeAgentManifestConverter.jsonToManifest(text);
      result.push({ path, manifest: agentManifest });
    }
  }
  return ok(result);
}

function readLocalSkills(
  entries: ReadonlyMap<string, AdmZip.IZipEntry>,
  references: SkillReference[]
): Result<LocalSkill[], FxError> {
  const skills: LocalSkill[] = [];
  const folders = new Set<string>();
  const names = new Set<string>();
  for (const { folder } of references) {
    if (folders.has(folder)) continue;
    const name = folder.split("/").pop();
    if (
      !name ||
      name.length > 64 ||
      !/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(name) ||
      names.has(name) ||
      [...folders].some(
        (existing) =>
          existing.toLowerCase() === folder.toLowerCase() ||
          existing.toLowerCase().startsWith(`${folder.toLowerCase()}/`) ||
          folder.toLowerCase().startsWith(`${existing.toLowerCase()}/`)
      )
    ) {
      return invalidPackage();
    }
    const skillMd = entries.get(`${folder}/SKILL.md`);
    if (!skillMd || skillMd.isDirectory) return invalidPackage();
    const match = /^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/.exec(decodeText(skillMd.getData()));
    if (!match) return invalidPackage();
    const metadata: unknown = yaml.load(match[1]);
    if (
      !isRecord(metadata) ||
      metadata.name !== name ||
      typeof metadata.description !== "string" ||
      !metadata.description.trim() ||
      metadata.description.length > 1024
    ) {
      return invalidPackage();
    }
    const files: SkillPackageFile[] = [];
    for (const [path, entry] of entries) {
      if (path.startsWith(`${folder}/`) && !entry.isDirectory) {
        files.push({ path: path.slice(folder.length + 1), content: entry.getData() });
      }
    }
    files.sort((left, right) => (left.path < right.path ? -1 : left.path > right.path ? 1 : 0));
    skills.push({ folder, name, description: metadata.description, files });
    folders.add(folder);
    names.add(name);
  }
  return ok(skills);
}

function readSkillReferences(
  manifest: Record<string, unknown>,
  manifestPath: string,
  property: SkillReference["property"]
): Result<SkillReference[], FxError> {
  const skills = manifest[property];
  if (skills === undefined) return ok([]);
  if (!Array.isArray(skills)) return invalidPackage();
  const references: SkillReference[] = [];
  for (const [index, skill] of skills.entries()) {
    if (!isRecord(skill)) return invalidPackage();
    // DA references without a local folder belong to the existing package submission path.
    if (property !== "agentSkills" && !Object.prototype.hasOwnProperty.call(skill, "folder")) {
      continue;
    }
    if (typeof skill.folder !== "string") return invalidPackage();
    const folder = confinedPath(skill.folder);
    if (!folder) return invalidPackage();
    references.push({ manifestPath, property, index, folder });
  }
  return ok(references);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function decodeText(content: Buffer): string {
  return new TextDecoder("utf-8", { fatal: true }).decode(content);
}

function confinedPath(value: string): string | undefined {
  if (
    !value ||
    value.startsWith("/") ||
    /[\\:\u0000-\u001f\u007f<>"|?*]/.test(value) ||
    /%(?:2e|2f|5c)/i.test(value)
  ) {
    return undefined;
  }
  const segments = value.normalize("NFC").split("/");
  if (
    segments.some(
      (segment) =>
        segment === ".." ||
        (segment !== "." && /[. ]$/.test(segment)) ||
        /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(segment)
    )
  ) {
    return undefined;
  }
  return segments.filter((segment) => segment !== "" && segment !== ".").join("/") || undefined;
}

function invalidPackage(): Result<never, FxError> {
  const key = "error.m365.skills.invalidPackage";
  return err(
    new UserError({
      source: "M365",
      name: "InvalidSkillPackage",
      message: getDefaultString(key),
      displayMessage: getLocalizedString(key),
    })
  );
}
