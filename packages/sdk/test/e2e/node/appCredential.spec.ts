 
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert, expect, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import { AppCredential, AppCredentialAuthConfig } from "../../../src";
import { ErrorCode, ErrorWithCode } from "../../../src/core/errors";
import { jwtDecode } from "jwt-decode";
import {
  AADJwtPayLoad,
  convertCertificateContent,
  extractIntegrationEnvVariables,
} from "../helper";

chaiUse(chaiPromises);
extractIntegrationEnvVariables();

describe("AppCredential Tests with Auth Config - Node", () => {
  const fake_client_secret = "fake_client_secret";
  const defaultGraphScope = ["https://graph.microsoft.com/.default"];

  it("create AppCredential instance should success with valid configuration", function () {
    const authConfig: AppCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential: any = new AppCredential(authConfig);

    assert.strictEqual(credential.msalClient.config.auth.clientId, authConfig.clientId);
    assert.strictEqual(
      credential.msalClient.config.auth.authority,
      authConfig.authorityHost + "/" + authConfig.tenantId
    );
    assert.strictEqual(credential.msalClient.config.auth.clientSecret, authConfig.clientSecret);
  });

  it("getToken should success with .default scope when authority host has tailing slash", async function () {
    const authConfig: AppCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST! + "/",
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential = new AppCredential(authConfig);
    const token = await credential.getToken(defaultGraphScope);

    const decodedToken = jwtDecode<AADJwtPayLoad>(token!.token);
    assert.strictEqual(decodedToken.aud, "https://graph.microsoft.com");
    assert.strictEqual(decodedToken.appid, authConfig.clientId);
    assert.strictEqual(decodedToken.idtyp, "app");
  });

  it("getToken should success with .default scope for Client Secret", async function () {
    const authConfig: AppCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential = new AppCredential(authConfig);
    const token = await credential.getToken(defaultGraphScope);

    const decodedToken = jwtDecode<AADJwtPayLoad>(token!.token);
    assert.strictEqual(decodedToken.aud, "https://graph.microsoft.com");
    assert.strictEqual(decodedToken.appid, authConfig.clientId);
    assert.strictEqual(decodedToken.idtyp, "app");
  });

  it("getToken should success with .default scope for Client Certificate", async function () {
    const authConfig: AppCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      certificateContent: convertCertificateContent(
        process.env.SDK_INTEGRATION_TEST_M365_AAD_CERTIFICATE_CONTENT!
      ),
    };
    const credential = new AppCredential(authConfig);
    const token = await credential.getToken(defaultGraphScope);

    const decodedToken = jwtDecode<AADJwtPayLoad>(token!.token);
    assert.strictEqual(decodedToken.aud, "https://graph.microsoft.com");
    assert.strictEqual(decodedToken.appid, authConfig.clientId);
    assert.strictEqual(decodedToken.idtyp, "app");
  });

  it("getToken should throw ServiceError with invalid secret", async function () {
    const authConfig: AppCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: fake_client_secret,
    };
    const credential = new AppCredential(authConfig);

    const errorResult = await expect(
      credential.getToken(defaultGraphScope)
    ).to.eventually.be.rejectedWith(ErrorWithCode);
    assert.strictEqual(errorResult.code, ErrorCode.ServiceError);
    assert.include(
      errorResult.message,
      "Get M365 tenant credential failed with error: invalid_client: Error(s): 7000215"
    );
  });
});
