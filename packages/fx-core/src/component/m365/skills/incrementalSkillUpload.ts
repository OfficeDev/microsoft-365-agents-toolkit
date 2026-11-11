// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  err,
  FxError,
  LogProvider,
  M365TokenProvider,
  ok,
  Result,
  UserError,
} from "@microsoft/teamsfx-api";
import { getResourceServiceEndpoint, ResourceServiceType } from "../../../common/constants";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { Component, sendTelemetryEvent, TelemetryEvent } from "../../../common/telemetry";
import { GraphSkillFileStore } from "./graphSkillFileStore";
import { ReadySkillDirectory, SkillDirectoryClient } from "./skillDirectoryClient";
import { synchronizeSkillFiles } from "./skillFileSync";
import {
  buildSkillTransportPackage,
  LocalSkill,
  readSkillPackage,
  validateSkillTransportPackage,
} from "./skillPackage";

export async function prepareIncrementalSkillUpload(
  data: Buffer,
  serviceUrl: string,
  mosToken: string,
  tokenProvider?: M365TokenProvider,
  logger?: LogProvider
): Promise<Result<{ data: Buffer; hasSkills: boolean }, FxError>> {
  const inventory = readSkillPackage(data);
  if (inventory.isErr()) return err(inventory.error);
  const source = inventory.value;
  if (source.skills.length === 0) return ok({ data, hasSkills: false });

  const preflight = validateSkillTransportPackage(source);
  if (preflight.isErr()) return err(preflight.error);

  if (!tokenProvider) {
    const key = "error.m365.skills.authorization";
    return err(
      new UserError({
        source: "M365",
        name: "SkillUploadAuthorizationFailed",
        message: getDefaultString(key),
        displayMessage: getLocalizedString(key),
      })
    );
  }
  const graphEndpoint = getResourceServiceEndpoint(ResourceServiceType.Graph);
  const graphToken = await tokenProvider.getAccessToken({
    scopes: [`${graphEndpoint}/Files.ReadWrite`],
  });
  if (graphToken.isErr()) return err(graphToken.error);

  logger?.warning(getLocalizedString("info.m365.skills.preview"));
  const directories = new SkillDirectoryClient(serviceUrl, mosToken);
  const folders = new Map<string, string>();
  const readySkills: { skill: LocalSkill; directory: ReadySkillDirectory }[] = [];
  for (const [index, skill] of source.skills.entries()) {
    logger?.info(getLocalizedString("info.m365.skills.preparing", index + 1, source.skills.length));
    const directory = await directories.createReadyDirectory(source.agentId, skill.name);
    if (directory.isErr()) return err(directory.error);
    folders.set(skill.folder, directory.value.folder);
    readySkills.push({ skill, directory: directory.value });
  }
  // Metadata may refer to existing remote GUIDs, which only MOS can resolve.
  const transport = buildSkillTransportPackage(source, folders);
  if (transport.isErr()) return err(transport.error);

  const counts = { uploaded: 0, skipped: 0, deleted: 0 };
  for (const [index, { skill, directory }] of readySkills.entries()) {
    logger?.info(
      getLocalizedString("info.m365.skills.synchronizing", index + 1, source.skills.length)
    );
    const store = new GraphSkillFileStore(graphEndpoint, graphToken.value, directory.driveId);
    const synchronized = await synchronizeSkillFiles(store, directory.itemId, skill.files);
    if (synchronized.isErr()) return err(synchronized.error);
    counts.uploaded += synchronized.value.uploaded;
    counts.skipped += synchronized.value.skipped;
    counts.deleted += synchronized.value.deleted;
  }
  logger?.info(
    getLocalizedString("info.m365.skills.summary", counts.uploaded, counts.skipped, counts.deleted)
  );
  sendTelemetryEvent(Component.core, TelemetryEvent.MosSkillSync, undefined, counts);
  return ok({ data: transport.value, hasSkills: true });
}
