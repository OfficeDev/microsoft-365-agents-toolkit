// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Result, UserError, err, ok } from "@microsoft/teamsfx-api";
import { createHash } from "crypto";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";

export interface LocalSkillFile {
  path: string;
  content: Buffer;
}

export interface RemoteSkillItem {
  id: string;
  name: string;
  kind: "file" | "folder";
  eTag: string;
  size?: number;
  hashes?: { sha1Hash?: string; quickXorHash?: string };
}

/** The port owns pagination, authentication, safe retries, and eTag preconditions. */
export interface SkillFileStore {
  listChildren(folderId: string): Promise<Result<RemoteSkillItem[], FxError>>;
  createFolder(parentId: string, name: string): Promise<Result<RemoteSkillItem, FxError>>;
  uploadFile(
    parentId: string,
    name: string,
    content: Buffer,
    existing?: RemoteSkillItem
  ): Promise<Result<RemoteSkillItem, FxError>>;
  deleteItem(item: RemoteSkillItem): Promise<Result<void, FxError>>;
}

type LocalEntry = { kind: "file"; content: Buffer } | { kind: "folder" };

function invalidInventory(): UserError {
  return new UserError({
    source: "M365",
    name: "InvalidSkillInventory",
    message: getDefaultString("error.m365.skills.invalidInventory"),
    displayMessage: getLocalizedString("error.m365.skills.invalidInventory"),
  });
}

function validName(name: unknown): name is string {
  return (
    typeof name === "string" &&
    name.length > 0 &&
    name !== "." &&
    name !== ".." &&
    name.trim() === name &&
    !name.endsWith(".") &&
    !/[<>:"/\\|?*\u0000-\u001f\u007f]/.test(name)
  );
}

function validId(id: unknown): id is string {
  return typeof id === "string" && id.trim().length > 0;
}

function validItem(item: RemoteSkillItem): boolean {
  return (
    item !== null &&
    typeof item === "object" &&
    validId(item.id) &&
    validName(item.name) &&
    (item.kind === "file" || item.kind === "folder") &&
    validId(item.eTag) &&
    (item.size === undefined || (Number.isSafeInteger(item.size) && item.size >= 0))
  );
}

function collisionKey(path: string): string {
  return path.normalize("NFC").toLowerCase();
}

function localInventory(files: LocalSkillFile[]): Result<Map<string, LocalEntry>, FxError> {
  const entries = new Map<string, LocalEntry>();
  const paths = new Map<string, string>();
  for (const file of files) {
    if (!file || typeof file.path !== "string" || !Buffer.isBuffer(file.content)) {
      return err(invalidInventory());
    }
    const parts = file.path.split("/");
    if (!parts.every(validName)) return err(invalidInventory());
    let path = "";
    for (let i = 0; i < parts.length; i++) {
      path = path ? `${path}/${parts[i]}` : parts[i];
      const kind = i === parts.length - 1 ? "file" : "folder";
      const existingPath = paths.get(collisionKey(path));
      if (existingPath !== undefined) {
        if (existingPath !== path || kind === "file" || entries.get(path)?.kind !== "folder") {
          return err(invalidInventory());
        }
      } else {
        paths.set(collisionKey(path), path);
        entries.set(path, kind === "file" ? { kind, content: file.content } : { kind });
      }
    }
  }
  return ok(entries);
}

async function remoteInventory(
  store: SkillFileStore,
  rootId: string
): Promise<Result<Map<string, RemoteSkillItem>, FxError>> {
  const entries = new Map<string, RemoteSkillItem>();
  const ids = new Set([rootId]);
  const paths = new Set<string>();
  const folders = [{ id: rootId, path: "" }];
  // Iterative traversal bounds the call stack even for deeply nested drafts.
  for (let index = 0; index < folders.length; index++) {
    const parent = folders[index];
    const children = await store.listChildren(parent.id);
    if (children.isErr()) return err(children.error);
    if (!Array.isArray(children.value)) return err(invalidInventory());
    for (const item of children.value) {
      if (!validItem(item) || ids.has(item.id)) return err(invalidInventory());
      const path = parent.path ? `${parent.path}/${item.name}` : item.name;
      const key = collisionKey(path);
      if (paths.has(key)) return err(invalidInventory());
      paths.add(key);
      ids.add(item.id);
      entries.set(path, item);
      if (item.kind === "folder") folders.push({ id: item.id, path });
    }
  }
  return ok(entries);
}

function quickXorHash(content: Buffer): string {
  // Matches Microsoft's C# reference: circular 160-bit XOR, 11-bit stride,
  // followed by little-endian length XOR into bytes 12..19.
  // https://learn.microsoft.com/onedrive/developer/code-snippets/quickxorhash
  const hash = Buffer.alloc(20);
  for (let i = 0; i < Math.min(content.length, 160); i++) {
    let byte = 0;
    for (let j = i; j < content.length; j += 160) byte ^= content[j];
    const shift = (i * 11) % 160;
    const index = Math.floor(shift / 8);
    const offset = shift % 8;
    hash[index] ^= byte << offset;
    hash[(index + 1) % hash.length] ^= byte >>> (8 - offset);
  }
  const length = Buffer.alloc(8);
  length.writeBigUInt64LE(BigInt(content.length));
  for (let i = 0; i < length.length; i++) hash[12 + i] ^= length[i];
  return hash.toString("base64");
}

function sameContent(item: RemoteSkillItem, content: Buffer): boolean {
  if (item.size !== content.length) return false;
  // These are change detectors, never credential/security primitives. Prefer
  // SHA-1 when usable; a mismatch must not be overridden by a weaker hash.
  const sha1 = item.hashes?.sha1Hash;
  if (typeof sha1 === "string" && /^[a-fA-F0-9]{40}$/.test(sha1)) {
    return createHash("sha1").update(content).digest("hex") === sha1.toLowerCase();
  }
  const quickXor = item.hashes?.quickXorHash;
  if (
    typeof quickXor !== "string" ||
    !/^[A-Za-z0-9+/]{27}=$/.test(quickXor) ||
    Buffer.from(quickXor, "base64").toString("base64") !== quickXor
  ) {
    return false;
  }
  return quickXorHash(content) === quickXor;
}

async function refreshEmptyFolder(
  store: SkillFileStore,
  rootId: string,
  path: string,
  item: RemoteSkillItem,
  entries: Map<string, RemoteSkillItem>
): Promise<Result<RemoteSkillItem, FxError>> {
  const parentPath = path.substring(0, Math.max(0, path.lastIndexOf("/")));
  const parentId = parentPath ? entries.get(parentPath)?.id : rootId;
  if (!parentId) return err(invalidInventory());
  const siblings = await store.listChildren(parentId);
  if (siblings.isErr()) return err(siblings.error);
  if (!Array.isArray(siblings.value)) return err(invalidInventory());
  const expected = new Map(
    Array.from(entries).filter(
      ([entryPath]) =>
        entryPath.substring(0, Math.max(0, entryPath.lastIndexOf("/"))) === parentPath
    )
  );
  if (siblings.value.length !== expected.size) return err(invalidInventory());
  const seen = new Set<string>();
  let refreshed: RemoteSkillItem | undefined;
  for (const sibling of siblings.value) {
    if (!validItem(sibling) || seen.has(sibling.id)) return err(invalidInventory());
    seen.add(sibling.id);
    const siblingPath = parentPath ? `${parentPath}/${sibling.name}` : sibling.name;
    const original = expected.get(siblingPath);
    if (original?.id !== sibling.id || original.kind !== sibling.kind) {
      return err(invalidInventory());
    }
    if (sibling.id === item.id) refreshed = sibling;
  }
  if (!refreshed) return err(invalidInventory());
  // Capture the new version BEFORE checking emptiness. A child added before the
  // check is detected here; a later change must fail the port's If-Match delete.
  const children = await store.listChildren(item.id);
  if (children.isErr()) return err(children.error);
  if (!Array.isArray(children.value) || children.value.length !== 0) {
    return err(invalidInventory());
  }
  return ok(refreshed);
}

/**
 * Reconciles only descendants enumerated from the caller's draft root. Every
 * inventory read completes before mutation; errors retain partial progress for
 * a subsequent invocation to reconcile against fresh metadata. Deleted counts
 * include files and folders; uploaded and skipped counts include only files.
 */
export async function synchronizeSkillFiles(
  store: SkillFileStore,
  rootId: string,
  files: LocalSkillFile[]
): Promise<Result<{ uploaded: number; skipped: number; deleted: number }, FxError>> {
  if (!validId(rootId) || !Array.isArray(files)) return err(invalidInventory());
  const local = localInventory(files);
  if (local.isErr()) return err(local.error);
  const remote = await remoteInventory(store, rootId);
  if (remote.isErr()) return err(remote.error);
  const counts = { uploaded: 0, skipped: 0, deleted: 0 };
  const entries = remote.value;
  const ids = new Set([rootId, ...Array.from(entries.values(), (item) => item.id)]);

  // Removing stale children first also makes file/folder and case replacements
  // possible without recursive deletes or a name-conflict overwrite.
  const stale = Array.from(entries).filter(
    ([path, item]) => local.value.get(path)?.kind !== item.kind
  );
  stale.sort(([a], [b]) => b.split("/").length - a.split("/").length);
  for (const [path, item] of stale) {
    let versionedItem = item;
    if (item.kind === "folder") {
      const refreshed = await refreshEmptyFolder(store, rootId, path, item, entries);
      if (refreshed.isErr()) return err(refreshed.error);
      versionedItem = refreshed.value;
    }
    const deleted = await store.deleteItem(versionedItem);
    if (deleted.isErr()) return err(deleted.error);
    entries.delete(path);
    counts.deleted++;
  }

  // Local inventory inserts each parent before its children.
  for (const [path, entry] of local.value) {
    const separator = path.lastIndexOf("/");
    const parentPath = separator < 0 ? "" : path.substring(0, separator);
    const parentId = parentPath ? entries.get(parentPath)?.id : rootId;
    if (!parentId) return err(invalidInventory());
    const name = path.substring(separator + 1);
    const existing = entries.get(path);
    if (entry.kind === "folder") {
      if (existing) continue;
      const created = await store.createFolder(parentId, name);
      if (created.isErr()) return err(created.error);
      if (
        !validItem(created.value) ||
        created.value.kind !== "folder" ||
        created.value.name !== name ||
        ids.has(created.value.id)
      ) {
        return err(invalidInventory());
      }
      ids.add(created.value.id);
      entries.set(path, created.value);
    } else {
      if (existing && sameContent(existing, entry.content)) {
        counts.skipped++;
        continue;
      }
      const uploaded = await store.uploadFile(parentId, name, entry.content, existing);
      if (uploaded.isErr()) return err(uploaded.error);
      if (
        !validItem(uploaded.value) ||
        uploaded.value.kind !== "file" ||
        uploaded.value.name !== name ||
        (existing ? uploaded.value.id !== existing.id : ids.has(uploaded.value.id))
      ) {
        return err(invalidInventory());
      }
      ids.add(uploaded.value.id);
      entries.set(path, uploaded.value);
      counts.uploaded++;
    }
  }
  return ok(counts);
}
