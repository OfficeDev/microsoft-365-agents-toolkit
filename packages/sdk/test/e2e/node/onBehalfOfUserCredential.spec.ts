 
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert, expect, use as chaiUse } from "chai";

import * as chaiPromises from "chai-as-promised";

import {
  ErrorCode,
  ErrorWithCode,
  OnBehalfOfCredentialAuthConfig,
  OnBehalfOfUserCredential,
} from "../../../src";
import { SSOTokenV2Info } from "../../../src/models/ssoTokenInfo";
import { parseJwt } from "../../../src/util/utils";
import {
  convertCertificateContent,
  extractIntegrationEnvVariables,
  getSsoTokenFromTeams,
} from "../helper";

chaiUse(chaiPromises);
extractIntegrationEnvVariables();

let ssoToken: string;

describe("OnBehalfOfUserCredential Tests with obo auth config - Node", () => {
  const defaultScope = "https://graph.microsoft.com/User.Read";
  const expiredSsoToken =
    // eslint-disable-next-line no-secrets/no-secrets
    "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Imk2bEdrM0ZaenhSY1ViMkMzbkVRN3N5SEpsWSIsImtpZCI6Imk2bEdrM0ZaenhSY1ViMkMzbkVRN3N5SEpsWSJ9.eyJhdWQiOiJlZjFkYTlkNC1mZjc3LTRjM2UtYTAwNS04NDBjM2Y4MzA3NDUiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9mYTE1ZDY5Mi1lOWM3LTQ0NjAtYTc0My0yOWYyOTUyMjIyOS8iLCJpYXQiOjE1MzcyMzMxMDYsIm5iZiI6MTUzNzIzMzEwNiwiZXhwIjoxNTM3MjM3MDA2LCJhY3IiOiIxIiwiYWlvIjoiQVhRQWkvOElBQUFBRm0rRS9RVEcrZ0ZuVnhMaldkdzhLKzYxQUdyU091TU1GNmViYU1qN1hPM0libUQzZkdtck95RCtOdlp5R24yVmFUL2tES1h3NE1JaHJnR1ZxNkJuOHdMWG9UMUxrSVorRnpRVmtKUFBMUU9WNEtjWHFTbENWUERTL0RpQ0RnRTIyMlRJbU12V05hRU1hVU9Uc0lHdlRRPT0iLCJhbXIiOlsid2lhIl0sImFwcGlkIjoiNzVkYmU3N2YtMTBhMy00ZTU5LTg1ZmQtOGMxMjc1NDRmMTdjIiwiYXBwaWRhY3IiOiIwIiwiZW1haWwiOiJBYmVMaUBtaWNyb3NvZnQuY29tIiwiZmFtaWx5X25hbWUiOiJMaW5jb2xuIiwiZ2l2ZW5fbmFtZSI6IkFiZSAoTVNGVCkiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMjIyNDcvIiwiaXBhZGRyIjoiMjIyLjIyMi4yMjIuMjIiLCJuYW1lIjoiYWJlbGkiLCJvaWQiOiIwMjIyM2I2Yi1hYTFkLTQyZDQtOWVjMC0xYjJiYjkxOTQ0MzgiLCJyaCI6IkkiLCJzY3AiOiJ1c2VyX2ltcGVyc29uYXRpb24iLCJzdWIiOiJsM19yb0lTUVUyMjJiVUxTOXlpMmswWHBxcE9pTXo1SDNaQUNvMUdlWEEiLCJ0aWQiOiJmYTE1ZDY5Mi1lOWM3LTQ0NjAtYTc0My0yOWYyOTU2ZmQ0MjkiLCJ1bmlxdWVfbmFtZSI6ImFiZWxpQG1pY3Jvc29mdC5jb20iLCJ1dGkiOiJGVnNHeFlYSTMwLVR1aWt1dVVvRkFBIiwidmVyIjoiMS4wIn0.D3H6pMUtQnoJAGq6AHd";

  before(async () => {
    ssoToken = await getSsoTokenFromTeams();
  });

  it("getToken should success with valid config", async function () {
    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    let ssoTokenFromCredential = await credential.getToken([]);
    assert.strictEqual(ssoTokenFromCredential!.token, ssoToken);

    ssoTokenFromCredential = await credential.getToken("");
    assert.strictEqual(ssoTokenFromCredential!.token, ssoToken);
  });

  it("get sso token should throw TokenExpiredError when sso token is expired", async function () {
    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential = new OnBehalfOfUserCredential(expiredSsoToken, authConfig);
    let err = await expect(credential.getToken([])).to.eventually.be.rejectedWith(ErrorWithCode);
    assert.strictEqual(err.code, ErrorCode.TokenExpiredError);

    err = await expect(credential.getToken("")).to.eventually.be.rejectedWith(ErrorWithCode);
    assert.strictEqual(err.code, ErrorCode.TokenExpiredError);
  });

  it("getUserInfo should success with valid config", async function () {
    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const userInfo = await credential.getUserInfo();
    const tokenObject = parseJwt(ssoToken) as SSOTokenV2Info;
    assert.strictEqual(userInfo.preferredUserName, tokenObject.preferred_username);
    assert.strictEqual(userInfo.tenantId, tokenObject.tid);
    assert.strictEqual(userInfo.objectId, tokenObject.oid);
    assert.strictEqual(userInfo.displayName, tokenObject.name);
  });

  it("get graph access token should success with valid config for Client Secret", async function () {
    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const graphToken = await credential.getToken(defaultScope);
    const tokenObject = parseJwt(graphToken!.token);
    const userInfo = await credential.getUserInfo();
    assert.strictEqual(tokenObject.oid, userInfo.objectId);
    assert.strictEqual(tokenObject.aud, "https://graph.microsoft.com");
    assert.include(tokenObject.scp, "User.Read");
  });

  it("get graph access token should success with valid config for Client Certificate", async function () {
    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      certificateContent: convertCertificateContent(
        process.env.SDK_INTEGRATION_TEST_M365_AAD_CERTIFICATE_CONTENT!
      ),
    };
    const credential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const graphToken = await credential.getToken(defaultScope);

    const tokenObject = parseJwt(graphToken!.token);
    const userInfo = await credential.getUserInfo();
    assert.strictEqual(tokenObject.oid, userInfo.objectId);
    assert.strictEqual(tokenObject.aud, "https://graph.microsoft.com");
    assert.include(tokenObject.scp, "User.Read");
  });

  it("get graph access token should success when authority host has tailing slash", async function () {
    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST! + "/",
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      certificateContent: convertCertificateContent(
        process.env.SDK_INTEGRATION_TEST_M365_AAD_CERTIFICATE_CONTENT!
      ),
    };

    const credential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const graphToken = await credential.getToken(defaultScope);
    const tokenObject = parseJwt(graphToken!.token);
    const userInfo = await credential.getUserInfo();
    assert.strictEqual(tokenObject.oid, userInfo.objectId);
    assert.strictEqual(tokenObject.aud, "https://graph.microsoft.com");
    assert.include(tokenObject.scp, "User.Read");
  });

  it("get graph access token should throw UiRequiredError without permission", async function () {
    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    await expect(credential.getToken("https://graph.microsoft.com/Calendars.Read"))
      .to.eventually.be.rejectedWith(ErrorWithCode)
      .and.property("code", ErrorCode.UiRequiredError);
  });

  it("get graph access token should throw TokenExpiredError when sso token is expired", async function () {
    const authConfig: OnBehalfOfCredentialAuthConfig = {
      authorityHost: process.env.SDK_INTEGRATION_TEST_AAD_AUTHORITY_HOST!,
      clientId: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID!,
      tenantId: process.env.SDK_INTEGRATION_TEST_AAD_TENANT_ID!,
      clientSecret: process.env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_SECRET!,
    };
    const credential = new OnBehalfOfUserCredential(expiredSsoToken, authConfig);
    const err = await expect(credential.getToken(defaultScope)).to.eventually.be.rejectedWith(
      ErrorWithCode
    );

    assert.strictEqual(err.code, ErrorCode.TokenExpiredError);
  });
});
