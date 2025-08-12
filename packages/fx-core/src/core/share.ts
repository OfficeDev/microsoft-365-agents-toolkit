// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { FxError, Result, err, ok } from "@microsoft/teamsfx-api";
import "reflect-metadata";
import { TOOLS } from "../common/globalVars";
import { getLocalizedString } from "../common/localizeUtils";
import "../component/driver/index";
import "../component/driver/script/scriptDriver";
import "../component/feature/sso";
import { M365AppEntity, M365EntityType } from "../component/m365/interface";
import { PackageService } from "../component/m365/packageService";
import { InputValidationError } from "../error/common";
import { CollaborationUtil } from "./collaborator";

export async function shareWithTenant(
  mosToken: string,
  sharedTitleId: string
): Promise<Result<undefined, FxError>> {
  // share with tenant users
  const res = await PackageService.GetSharedInstance().shareWithTenant(mosToken, sharedTitleId);
  if (res.isErr()) {
    return err(res.error);
  }
  const msg = getLocalizedString("core.common.shareWithTenant.success");
  TOOLS.ui?.showMessage("info", msg, false);
  return ok(undefined);
}

export async function addSharedUsers(
  mosToken: string,
  sharedTitleId: string,
  emails: string[]
): Promise<Result<undefined, FxError>> {
  const tokenProvider = TOOLS.tokenProvider.m365TokenProvider;
  const entities: Set<M365AppEntity> = new Set();
  const existingEntities = await PackageService.GetSharedInstance().getSharedUsers(
    mosToken,
    sharedTitleId
  );
  if (existingEntities.isErr()) {
    return err(existingEntities.error);
  }
  // merge existing entities with new entities
  for (const entity of existingEntities.value) {
    entities.add({
      entityId: entity.entityId,
      entityType: entity.entityType,
    });
  }

  for (const email of emails) {
    const userInfo = await CollaborationUtil.getUserInfo(tokenProvider, email);
    if (!userInfo) {
      return err(new InputValidationError("shareWithUser", `Invalid user: ${email}`));
    } else {
      entities.add({
        entityId: userInfo.aadId,
        entityType: M365EntityType.User,
      });
    }
  }

  // Call Builder API to change shared scope
  const res = await PackageService.GetSharedInstance().shareWithUsers(
    mosToken,
    Array.from(entities),
    sharedTitleId
  );
  if (res.isErr()) {
    return err(res.error);
  }
  const msg = getLocalizedString("core.common.shareWithUser.success", emails);
  TOOLS.ui?.showMessage("info", msg, false);
  return ok(undefined);
}

export async function removeShareAccess(
  mosToken: string,
  sharedTitleId: string,
  emails: string[]
): Promise<Result<undefined, FxError>> {
  const tokenProvider = TOOLS.tokenProvider.m365TokenProvider;
  const entities: Set<string> = new Set();
  const existingEntities = await PackageService.GetSharedInstance().getSharedUsers(
    mosToken,
    sharedTitleId
  );
  if (existingEntities.isErr()) {
    return err(existingEntities.error);
  }

  // remove users from shared access
  for (const email of emails) {
    const userInfo = await CollaborationUtil.getUserInfo(tokenProvider, email);
    if (!userInfo) {
      return err(new InputValidationError("shareWithUser", `Invalid user: ${email}`));
    } else {
      entities.add(userInfo.aadId);
    }
  }
  const remainingEntities = Array.from(existingEntities.value).filter(
    (entity) => !entities.has(entity.entityId)
  );

  if (remainingEntities.length === 0) {
    // no users left, call Builder API to unshare
    const res = await PackageService.GetSharedInstance().unshare(mosToken, sharedTitleId);
    if (res.isErr()) {
      return err(res.error);
    }
  } else {
    // Call Builder API to change shared scope
    const res = await PackageService.GetSharedInstance().shareWithUsers(
      mosToken,
      Array.from(remainingEntities),
      sharedTitleId
    );
    if (res.isErr()) {
      return err(res.error);
    }
  }

  const msg = getLocalizedString("core.common.removeShareAccess.success", emails);
  TOOLS.ui?.showMessage("info", msg, false);
  return ok(undefined);
}
