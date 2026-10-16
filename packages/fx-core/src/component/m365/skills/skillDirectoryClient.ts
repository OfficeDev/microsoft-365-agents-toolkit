// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, SystemError, UserError } from "@microsoft/teamsfx-api";
import { AxiosInstance } from "axios";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { WrappedAxiosClient } from "../../../common/wrappedAxiosClient";
import { requestSkillOperation } from "./skillRequest";

export interface ReadySkillDirectory {
  folder: string;
  driveId: string;
  itemId: string;
}

const httpTimeoutMs = 30000;
const sharePointSuffixes = [
  ".sharepoint.com",
  ".sharepoint.us",
  ".sharepoint-mil.us",
  ".sharepoint.cn",
  ".sharepoint.de",
];

export class SkillDirectoryClient {
  private readonly client: AxiosInstance;

  constructor(mosEndpoint: string, mosToken: string) {
    this.client = WrappedAxiosClient.create({
      baseURL: mosEndpoint.replace(/\/$/, ""),
      headers: { Authorization: `Bearer ${mosToken}` },
      timeout: httpTimeoutMs,
      maxRedirects: 0,
    });
  }

  async createReadyDirectory(
    agentId: string,
    name: string
  ): Promise<Result<ReadySkillDirectory, FxError>> {
    const response = await requestSkillOperation(
      () =>
        this.client.post<unknown>("/builder/v1/directories", {
          agentId,
          directoryType: "skill",
          name,
          skipPublishedContentCopy: true,
        }),
      "never"
    );
    if (response.isErr()) return err(response.error);
    if (response.value.status !== 200) return err(failure("directoryFailed"));
    const parsed = parseDirectory(response.value.data);
    if (!parsed) return err(failure("invalidResponse"));

    // Exclusive ATK skip-copy provenance is a prerequisite, not inferred from this response.
    if (parsed.monitors.length > 0) return err(failure("copyNotSupported"));
    return ok(parsed.directory);
  }
}

function parseDirectory(
  value: unknown
): { directory: ReadySkillDirectory; monitors: string[] } | undefined {
  if (!isRecord(value) || !guidHex(value.folder) || !isRecord(value.sharePointIds)) {
    return undefined;
  }
  const ids = value.sharePointIds;
  const site = guidHex(ids.siteId);
  const web = guidHex(ids.webId);
  const list = guidHex(ids.listId);
  if (
    typeof value.folder !== "string" ||
    !site ||
    !web ||
    !list ||
    typeof ids.uniqueId !== "string" ||
    !guidHex(ids.uniqueId) ||
    typeof ids.siteUrl !== "string"
  ) {
    return undefined;
  }
  const siteUrl = safeUrl(ids.siteUrl);
  if (
    !siteUrl ||
    siteUrl.port ||
    !sharePointSuffixes.some((suffix) => siteUrl.hostname.endsWith(suffix))
  ) {
    return undefined;
  }
  const monitors: string[] = [];
  if (value.copyStatusUrls != null) {
    if (!Array.isArray(value.copyStatusUrls)) return undefined;
    for (const entry of value.copyStatusUrls) {
      if (typeof entry !== "string") return undefined;
      const monitor = safeUrl(entry);
      if (!monitor || monitor.origin !== siteUrl.origin) return undefined;
      monitors.push(entry);
    }
  }
  return {
    directory: {
      folder: value.folder,
      driveId: `b!${Buffer.concat([guidBytes(site), guidBytes(web), guidBytes(list)])
        .toString("base64")
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "")}`,
      itemId: ids.uniqueId,
    },
    monitors,
  };
}

function guidHex(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const text = value.startsWith("{") && value.endsWith("}") ? value.slice(1, -1) : value;
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(text)
    ? text.replace(/-/g, "")
    : undefined;
}

function guidBytes(hex: string): Buffer {
  // .NET Guid.ToByteArray reverses the first 4/2/2-byte fields, not the final eight bytes.
  const bytes = Buffer.from(hex, "hex");
  bytes.subarray(0, 4).reverse();
  bytes.subarray(4, 6).reverse();
  bytes.subarray(6, 8).reverse();
  return bytes;
}

function safeUrl(value: string): URL | undefined {
  if (/[\s\\#]/.test(value) || /^https:\/\/[^/]*@/.test(value)) return undefined;
  try {
    const url = new URL(value);
    return value.startsWith("https://") && !url.username && !url.password && !url.hash
      ? url
      : undefined;
  } catch {
    return undefined;
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function failure(kind: "directoryFailed" | "copyNotSupported" | "invalidResponse"): FxError {
  const names = {
    directoryFailed: "SkillDirectoryFailed",
    copyNotSupported: "SkillUploadCopyNotSupported",
    invalidResponse: "SkillInvalidResponse",
  };
  const key = `error.m365.skills.${kind}`;
  const options = {
    source: "M365",
    name: names[kind],
    message: getDefaultString(key),
    displayMessage: getLocalizedString(key),
  };
  return kind === "copyNotSupported" ? new UserError(options) : new SystemError(options);
}
