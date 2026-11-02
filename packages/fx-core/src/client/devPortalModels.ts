// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { TeamsManifest } from "@microsoft/teamsfx-api";
import { validate as isUuid } from "uuid";

/** Developer Portal drafts can contain null or incomplete manifest sections. */
type Draft<T> = T extends (infer Item)[]
  ? Draft<Item>[]
  : T extends object
    ? { [Key in keyof T]?: Draft<T[Key]> | null }
    : T;

export interface AppUserPayload {
  aadId: string;
  tenantId: string;
  role: string;
}

/** Native /v1.0/apps detail response; timestamps are only available on list summaries. */
export interface AppResponse {
  appId: string;
  appProfile: {
    appAccessControl?: { appUsers?: AppUserPayload[] | null } | null;
    appMetadata?: {
      ownerAadId?: string | null;
      updatedByAadId?: string | null;
      source?: string | number | null;
      appType?: string | number | null;
      appPublishInfos?: unknown[] | null;
    } | null;
    appDetails?: { applicationManifest?: Draft<TeamsManifest> | null } | null;
  } | null;
}

export interface AppSummary {
  appId: string;
  appName?: string | null;
  appVersion?: string | null;
  appExternalId?: string | null;
  updatedAt?: string | null;
  updatedByAadId?: string | null;
  colorIconPath?: string | null;
  outlineIconPath?: string | null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isNullishString(value: unknown): value is string | null | undefined {
  return value === undefined || value === null || typeof value === "string";
}

// Validate only the envelope; consumers enforce their required nested fields.
export function hasValidAppResponse(value: unknown): boolean {
  return (
    isRecord(value) &&
    typeof value.appId === "string" &&
    isUuid(value.appId) &&
    (value.appProfile === null || isRecord(value.appProfile))
  );
}

export function isAppSummary(value: unknown): value is AppSummary {
  return (
    isRecord(value) &&
    typeof value.appId === "string" &&
    isUuid(value.appId) &&
    [
      "appName",
      "appVersion",
      "appExternalId",
      "updatedAt",
      "updatedByAadId",
      "colorIconPath",
      "outlineIconPath",
    ].every((key) => isNullishString(value[key]))
  );
}
