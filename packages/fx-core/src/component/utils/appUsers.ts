// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, M365TokenProvider, ok, Result } from "@microsoft/teamsfx-api";
import { GraphClient } from "../../client/graphClient";
import { teamsDevPortalClient } from "../../client/teamsDevPortalClientProvider";
import { AppUser } from "../driver/teamsApp/interfaces/appdefinitions/appUser";
import { appIdentityError } from "./appIdentity";

export async function getAppUsers(
  token: string,
  resourceId: string,
  tokenProvider: M365TokenProvider
): Promise<Result<AppUser[], FxError>> {
  const users = await teamsDevPortalClient.getUserList(token, resourceId);
  const result: AppUser[] = [];
  const graph = new GraphClient(tokenProvider);
  for (const user of users ?? []) {
    if ("isAdministrator" in user) {
      result.push(user);
      continue;
    }
    const info = await graph.getUserInfoFromId(encodeURIComponent(user.aadId));
    if (!info?.userPrincipalName || !info.displayName) {
      return err(appIdentityError("appAccessControl.appUsers"));
    }
    result.push({
      ...user,
      isAdministrator: user.role === "Owner",
      displayName: info.displayName,
      userPrincipalName: info.userPrincipalName,
    });
  }
  return ok(result);
}
