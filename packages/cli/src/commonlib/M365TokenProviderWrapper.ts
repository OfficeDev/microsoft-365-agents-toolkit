// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import {
  FxError,
  LoginStatus,
  M365TokenProvider,
  Result,
  TokenRequest,
} from "@microsoft/teamsfx-api";
import M365TokenProviderUserPassword from "./m365LoginUserPassword";
import { M365Login } from "./m365Login";
import ui from "../userInteraction";

/**
 * this class is a wrapper for M365TokenProvider that will use M365Login if interactive is true, otherwise use M365TokenProviderUserPassword
 */
class MM365TokenProviderWrapper implements M365TokenProvider {
  getProvider(): M365TokenProvider {
    // if interactive is false and system environment variables (M365_ACCOUNT_NAME, M365_ACCOUNT_PASSWORD) are set, then use M365TokenProviderUserPassword
    const m365Login =
      !ui.interactive && process.env.M365_ACCOUNT_NAME && process.env.M365_ACCOUNT_PASSWORD
        ? M365TokenProviderUserPassword
        : M365Login.getInstance();
    return m365Login;
  }
  getAccessToken(tokenRequest: TokenRequest): Promise<Result<string, FxError>> {
    return this.getProvider().getAccessToken(tokenRequest);
  }
  getJsonObject(
    tokenRequest: TokenRequest,
    tenantId?: string
  ): Promise<Result<Record<string, unknown>, FxError>> {
    return this.getProvider().getJsonObject(tokenRequest, tenantId);
  }
  getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>> {
    return this.getProvider().getStatus(tokenRequest);
  }
  setStatusChangeMap(
    name: string,
    tokenRequest: TokenRequest,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    immediateCall?: boolean
  ): Promise<Result<boolean, FxError>> {
    return this.getProvider().setStatusChangeMap(name, tokenRequest, statusChange, immediateCall);
  }
  removeStatusChangeMap(name: string): Promise<Result<boolean, FxError>> {
    return this.getProvider().removeStatusChangeMap(name);
  }
  async signout(): Promise<boolean> {
    return await (this.getProvider() as any).signout();
  }
  async switchTenant(tenantId: string): Promise<Result<string, FxError>> {
    return await this.getProvider().switchTenant(tenantId);
  }
  async getTenant(): Promise<string | undefined> {
    return await (this.getProvider() as any).getTenant();
  }
}

export default new MM365TokenProviderWrapper();
