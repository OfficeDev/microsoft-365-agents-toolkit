// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createHash } from "crypto";
import { err, ok } from "neverthrow";
import type { Result } from "neverthrow";
import {
  assembleApplicationPackage,
  SharePointApplicationPackageError,
  toApplicationPackageError,
} from "./component/driver/deploy/spfx/utility/applicationPackage";
import { inspectSharePointPackage } from "./component/driver/deploy/spfx/utility/packageFile";

export interface AssembleSharePointApplicationPackageInput {
  applicationPackage: Buffer;
  sharePointPackage: Buffer;
}

export interface SharePointApplicationPackage {
  packageBytes: Buffer;
  packageSha256: string;
  sourceApplicationSha256: string;
  sourceSharePointSha256: string;
  pluginSha256: string;
  appId: string;
  appVersion: string;
  productId: string;
  managedReference: {
    kind: "declarativeAgentAction" | "agentConnector";
    id: string;
    /** Path stored in the manifest reference, relative to its containing manifest. */
    path: string;
    /** ZIP entry path containing the managed plugin bytes. */
    entryPath: string;
  };
  validation: {
    level: "local-structural";
    tenantNeutral: true;
  };
}

function sha256(bytes: Buffer): string {
  return createHash("sha256").update(bytes).digest("hex");
}

export async function assembleSharePointApplicationPackage(
  input: AssembleSharePointApplicationPackageInput
): Promise<Result<SharePointApplicationPackage, SharePointApplicationPackageError>> {
  const inspection = await inspectSharePointPackage(input.sharePointPackage);
  if (inspection.isErr()) {
    return err(toApplicationPackageError(inspection.error));
  }
  const assembled = assembleApplicationPackage(
    input.applicationPackage,
    inspection.value.pluginManifest,
    inspection.value.pluginManifest,
    inspection.value.productId
  );
  if (assembled.isErr()) {
    return err(assembled.error);
  }
  return ok({
    packageBytes: assembled.value.bytes,
    packageSha256: sha256(assembled.value.bytes),
    sourceApplicationSha256: sha256(input.applicationPackage),
    sourceSharePointSha256: inspection.value.sourceSha256,
    pluginSha256: assembled.value.pluginSha256,
    appId: assembled.value.appId,
    appVersion: assembled.value.appVersion,
    productId: inspection.value.productId,
    managedReference: assembled.value.managedReference,
    validation: {
      level: "local-structural",
      tenantNeutral: true,
    },
  });
}

export {
  inspectSharePointPackage,
  inspectSharePointPackageIdentity,
  readSharePointPackage,
  SHAREPOINT_PACKAGE_LIMITS,
  SharePointPackageInspectionError,
} from "./component/driver/deploy/spfx/utility/packageFile";
export type {
  SharePointPackageIdentity,
  SharePointPackageInspection,
} from "./component/driver/deploy/spfx/utility/packageFile";
export { SharePointApplicationPackageError };
