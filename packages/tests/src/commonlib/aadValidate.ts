// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import axios from "axios";
import * as chai from "chai";

import { M365TokenProvider } from "@microsoft/teamsfx-api";

import MockM365TokenProvider from "@microsoft/m365agentstoolkit-cli/src/commonlib/m365LoginUserPassword";
import { GraphScopes } from "@microsoft/teamsfx-core";
import { EnvConstants } from "../commonlib/constants";
import {
  IAADDefinition,
  IAadObject,
  IAadObjectLocal,
} from "./interfaces/IAADDefinition";

const baseUrl = "https://graph.microsoft.com/v1.0";

function delay(ms: number) {
  // tslint:disable-next-line no-string-based-set-timeout
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export class AadValidator {
  public static provider: M365TokenProvider;

  public static init(
    ctx: any,
    isLocalDebug = false,
    provider?: M365TokenProvider
  ): IAadObject {
    console.log("Start to init validator for Microsoft Entra app.");

    AadValidator.provider = provider || MockM365TokenProvider;

    const aadObject: IAadObject | undefined = AadValidator.parseConfig(
      ctx,
      isLocalDebug
    );
    chai.assert.exists(aadObject);

    console.log("Successfully init validator for Microsoft Entra app.");
    return aadObject!;
  }

  public static async validate(
    aadObject: IAadObject,
    expectedPermission?: string,
    expectAadName?: string,
    expectApplicationIdUri?: string
  ) {
    console.log("Start to validate Microsoft Entra app.");

    const groundTruth = await AadValidator.getAadApp(aadObject.objectId);
    chai.assert.exists(groundTruth);

    chai.assert(aadObject.clientId, groundTruth?.appId);
    chai.assert(
      aadObject.oauth2PermissionScopeId,
      groundTruth?.api?.oauth2PermissionScopes![0].id
    );
    if (expectApplicationIdUri) {
      chai.assert(expectApplicationIdUri, groundTruth?.identifierUris![0]);
    } else {
      chai.assert(aadObject.applicationIdUris, groundTruth?.identifierUris![0]);
    }

    if (expectedPermission) {
      console.log("Start to validate permission for Microsoft Entra app.");
      chai.assert(
        expectedPermission,
        JSON.stringify(groundTruth?.requiredResourceAccess)
      );
    }

    console.log("Successfully validate Microsoft Entra app.");
  }

  private static parseConfig(
    ctx: any,
    isLocalDebug: boolean
  ): IAadObject | undefined {
    return AadValidator.objectTransformV3(ctx);
  }

  private static async getAadApp(objectId: string) {
    const appStudioTokenRes = await this.provider.getAccessToken({
      scopes: GraphScopes,
    });
    const appStudioToken = appStudioTokenRes.isOk()
      ? appStudioTokenRes.value
      : undefined;

    let retries = 10;
    while (retries > 0) {
      try {
        retries = retries - 1;
        axios.defaults.headers.common[
          "Authorization"
        ] = `Bearer ${appStudioToken}`;
        const aadGetResponse = await axios.get(
          `${baseUrl}/applications/${objectId}`
        );
        if (
          aadGetResponse &&
          aadGetResponse.data &&
          aadGetResponse.data["identifierUris"][0]
        ) {
          return <IAADDefinition>aadGetResponse.data;
        }
      } catch (error) {
        console.log("Microsoft Entra app get failed. Retry.");
      }

      await delay(10000);
    }

    return undefined;
  }

  private static objectTransform(localObject: IAadObjectLocal): IAadObject {
    return {
      clientId: localObject.local_clientId,
      clientSecret: localObject.local_clientSecret,
      objectId: localObject.local_objectId,
      oauth2PermissionScopeId: localObject.local_oauth2PermissionScopeId,
      applicationIdUris: localObject.local_applicationIdUris,
      oauthAuthority: localObject.oauthAuthority,
      teamsMobileDesktopAppId: localObject.teamsMobileDesktopAppId,
      teamsWebAppId: localObject.teamsWebAppId,
    } as IAadObject;
  }

  private static objectTransformV3(ctxObj: Record<string, string>): IAadObject {
    return {
      clientId: ctxObj[EnvConstants.AAD_APP_CLIENT_ID],
      clientSecret: ctxObj[EnvConstants.AAD_APP_CLIENT_SECRETS],
      objectId: ctxObj[EnvConstants.AAD_APP_OBJECT_ID],
      oauth2PermissionScopeId:
        ctxObj[EnvConstants.AAD_APP_ACCESS_AS_USER_PERMISSION_ID],
      applicationIdUris: ctxObj[EnvConstants.AAD_APP_OAUTH_AUTHORITY_HOST],
      oauthAuthority: ctxObj[EnvConstants.AAD_APP_OAUTH_AUTHORITY],
      teamsMobileDesktopAppId: "test",
      teamsWebAppId: "test",
    };
  }
}
