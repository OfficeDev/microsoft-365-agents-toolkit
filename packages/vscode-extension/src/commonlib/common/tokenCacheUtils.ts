// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { AccountInfo } from "@azure/msal-node";

export function getAccountByHomeId(
  homeAccountId: string,
  allAccounts: AccountInfo[]
): AccountInfo | null {
  if (homeAccountId && allAccounts && allAccounts.length) {
    return (
      allAccounts.filter((accountObj) => accountObj.homeAccountId === homeAccountId)[0] || null
    );
  } else {
    return null;
  }
}
