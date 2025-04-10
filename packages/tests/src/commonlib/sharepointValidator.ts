// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosInstance } from "axios";
import * as chai from "chai";

import MockM365TokenProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365LoginUserPassword";
import { M365TokenProvider } from "@microsoft/teamsfx-api";
import {
  getSPFxTenant,
  GraphReadUserScopes,
  SPFxScopes,
} from "@microsoft/teamsfx-core";

export class SharepointValidator {
  public static provider: M365TokenProvider;

  public static init(provider?: M365TokenProvider) {
    SharepointValidator.provider = provider || MockM365TokenProvider;
  }

  public static async validateDeploy(appId: string) {
    const graphToken = await this.provider.getAccessToken({
      scopes: GraphReadUserScopes,
    });
    let spfxToken = undefined;
    if (graphToken.isOk()) {
      const tenant = await getSPFxTenant(graphToken.value);
      const spfxTokenRes = await this.provider.getAccessToken({
        scopes: SPFxScopes(tenant),
      });
      spfxToken = spfxTokenRes.isOk() ? spfxTokenRes.value : undefined;
    }
    chai.assert.isNotEmpty(spfxToken);

    const requester = this.createRequesterWithToken(spfxToken!);
    const response = await requester.get(
      `/_api/web/tenantappcatalog/AvailableApps/GetById('${appId}')`
    );
    chai.assert.isTrue(response.data.Deployed);
  }

  public static async deleteApp(appId: string) {
    const graphToken = await this.provider.getAccessToken({
      scopes: GraphReadUserScopes,
    });
    let spfxToken = undefined;
    if (graphToken.isOk()) {
      const tenant = await getSPFxTenant(graphToken.value);
      const spfxTokenRes = await this.provider.getAccessToken({
        scopes: SPFxScopes(tenant),
      });
      spfxToken = spfxTokenRes.isOk() ? spfxTokenRes.value : undefined;
    }
    chai.assert.isNotEmpty(spfxToken);

    const requester = this.createRequesterWithToken(spfxToken!);
    await requester.post(
      `/_api/web/tenantappcatalog/AvailableApps/GetById('${appId}')/Remove`
    );
  }

  private static createRequesterWithToken(
    sharepointToken: string
  ): AxiosInstance {
    const instance = axios.create({
      baseURL: "https://utest0.sharepoint.com/sites/test",
    });
    instance.defaults.headers.common[
      "Authorization"
    ] = `Bearer ${sharepointToken}`;
    return instance;
  }
}
