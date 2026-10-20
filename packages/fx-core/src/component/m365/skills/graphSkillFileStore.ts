// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, SystemError } from "@microsoft/teamsfx-api";
import { AxiosInstance } from "axios";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { WrappedAxiosClient } from "../../../common/wrappedAxiosClient";
import type { RemoteSkillItem, SkillFileStore } from "./skillFileSync";
import { requestSkillOperation } from "./skillRequest";

export class GraphSkillFileStore implements SkillFileStore {
  private readonly client: AxiosInstance;
  private readonly baseUrl: string;
  private readonly drivePath: string;

  constructor(graphEndpoint: string, token: string, driveId: string) {
    this.baseUrl = `${graphEndpoint.replace(/\/$/, "")}/v1.0`;
    this.drivePath = `/drives/${encodeURIComponent(driveId)}/items`;
    this.client = WrappedAxiosClient.create({
      baseURL: this.baseUrl,
      headers: { Authorization: `Bearer ${token}` },
      timeout: 30000,
      maxRedirects: 0,
    });
  }

  async listChildren(folderId: string): Promise<Result<RemoteSkillItem[], FxError>> {
    const path = `${this.itemPath(folderId)}/children`;
    let next: string | undefined =
      `${path}?$select=id,name,eTag,size,file,folder,remoteItem,parentReference`;
    const visited = new Set<string>();
    const items: RemoteSkillItem[] = [];
    while (next !== undefined) {
      let absolute: URL;
      try {
        absolute = new URL(next.startsWith("/") ? `${this.baseUrl}${next}` : next);
      } catch {
        return err(invalidResponse());
      }
      const expected = new URL(`${this.baseUrl}${path}`);
      if (
        absolute.origin !== expected.origin ||
        absolute.pathname !== expected.pathname ||
        absolute.username ||
        absolute.password ||
        absolute.hash ||
        visited.has(absolute.href)
      ) {
        return err(invalidResponse());
      }
      visited.add(absolute.href);
      const page = next;
      const response = await requestSkillOperation(() => this.client.get<unknown>(page));
      if (response.isErr()) return err(response.error);
      const data = response.value.data;
      if (response.value.status !== 200 || !isRecord(data) || !Array.isArray(data.value)) {
        return err(invalidResponse());
      }
      for (const value of data.value) {
        const item = parseItem(value);
        if (!item) return err(invalidResponse());
        if (
          isRecord(value) &&
          isRecord(value.parentReference) &&
          value.parentReference.id !== folderId
        ) {
          return err(invalidResponse());
        }
        items.push(item);
      }
      const link: unknown = data["@odata.nextLink"];
      if (link !== undefined && (typeof link !== "string" || !link.startsWith("https://"))) {
        return err(invalidResponse());
      }
      next = link;
    }
    return ok(items);
  }

  async createFolder(parentId: string, name: string): Promise<Result<RemoteSkillItem, FxError>> {
    const response = await requestSkillOperation(
      () =>
        this.client.post<unknown>(`${this.itemPath(parentId)}/children`, {
          name,
          folder: {},
          "@microsoft.graph.conflictBehavior": "fail",
        }),
      false
    );
    if (response.isErr()) return err(response.error);
    const item = parseItem(response.value.data);
    return response.value.status === 201 && item?.kind === "folder" && item.name === name
      ? ok(item)
      : err(invalidResponse());
  }

  async uploadFile(
    parentId: string,
    name: string,
    content: Buffer,
    existing?: RemoteSkillItem
  ): Promise<Result<RemoteSkillItem, FxError>> {
    const path = existing
      ? `${this.itemPath(existing.id)}/content`
      : `${this.itemPath(parentId)}:/${encodeURIComponent(name)}:/content`;
    const response = await requestSkillOperation(
      () =>
        this.client.put<unknown>(path, content, {
          headers: {
            "Content-Type": "application/octet-stream",
            ...(existing ? { "If-Match": existing.eTag } : { "If-None-Match": "*" }),
          },
          params: { "@microsoft.graph.conflictBehavior": "fail" },
          maxBodyLength: Infinity,
        }),
      false
    );
    if (response.isErr()) return err(response.error);
    const item = parseItem(response.value.data);
    return (response.value.status === 200 || response.value.status === 201) &&
      item?.kind === "file" &&
      item.name === name &&
      (!existing || item.id === existing.id)
      ? ok(item)
      : err(invalidResponse());
  }

  async deleteItem(item: RemoteSkillItem): Promise<Result<void, FxError>> {
    const response = await requestSkillOperation(
      () =>
        this.client.delete(this.itemPath(item.id), {
          headers: { "If-Match": item.eTag },
        }),
      false
    );
    if (response.isErr()) return err(response.error);
    return response.value.status === 204 ? ok(undefined) : err(invalidResponse());
  }

  private itemPath(id: string): string {
    return `${this.drivePath}/${encodeURIComponent(id)}`;
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseItem(value: unknown): RemoteSkillItem | undefined {
  if (
    !isRecord(value) ||
    typeof value.id !== "string" ||
    !value.id ||
    typeof value.name !== "string" ||
    !value.name ||
    typeof value.eTag !== "string" ||
    !value.eTag ||
    value.remoteItem !== undefined ||
    isRecord(value.file) === isRecord(value.folder)
  ) {
    return undefined;
  }
  if (isRecord(value.folder)) {
    return { id: value.id, name: value.name, eTag: value.eTag, kind: "folder" };
  }
  if (!isRecord(value.file) || !Number.isSafeInteger(value.size) || Number(value.size) < 0) {
    return undefined;
  }
  const hashes = isRecord(value.file.hashes) ? value.file.hashes : {};
  return {
    id: value.id,
    name: value.name,
    eTag: value.eTag,
    kind: "file",
    size: Number(value.size),
    hashes: {
      ...(typeof hashes.sha1Hash === "string" ? { sha1Hash: hashes.sha1Hash } : {}),
      ...(typeof hashes.quickXorHash === "string" ? { quickXorHash: hashes.quickXorHash } : {}),
    },
  };
}

function invalidResponse(): SystemError {
  return new SystemError({
    source: "M365",
    name: "InvalidSkillServiceResponse",
    message: getDefaultString("error.m365.skills.invalidResponse"),
    displayMessage: getLocalizedString("error.m365.skills.invalidResponse"),
  });
}
