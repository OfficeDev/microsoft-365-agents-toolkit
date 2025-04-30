// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert, expect, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import {
  ErrorCode,
  ErrorWithCode,
  OnBehalfOfCredentialAuthConfig,
  OnBehalfOfUserCredential,
  UserInfo,
} from "../../../../src";
import * as sinon from "sinon";
import { AuthenticationResult, ConfidentialClientApplication, AuthError } from "@azure/msal-node";

chaiUse(chaiPromises);
const jwtBuilder = require("jwt-builder");

describe("OnBehalfOfUserCredential Tests - Node", () => {
  const scope = "fake_scope";
  const clientId = "fake_client_id";
  const clientSecret = "fake_client_secret";
  const certificateContent = `-----BEGIN PRIVATE KEY-----
fakeKey
-----END PRIVATE KEY-----
-----BEGIN CERTIFICATE-----
fakeCert
-----END CERTIFICATE-----`;
  const authorityHost = "fake_authority_host";
  const tenantId = "fake_tenant_id";
  const accessToken = "fake_access_token";
  const accessTokenExpDate = new Date("2021-04-14T02:02:23.742Z");
  const accessTokenExpNumber = accessTokenExpDate.getTime();

  // Error code
  const InvalidConfiguration = "InvalidConfiguration";
  const InternalError = "InternalError";
  const ServiceError = "ServiceError";

  const now = Math.floor(Date.now() / 1000);
  const timeInterval = 4000;
  const ssoTokenExp = now + timeInterval;
  const testDisplayName = "Teams Framework Unit Test";
  const testObjectId = "11111111-2222-3333-4444-555555555555";
  const testTenantId = "11111111-2222-3333-4444-555555555555";
  const testPreferredUserName = "test@microsoft.com";
  const ssoToken = jwtBuilder({
    algorithm: "HS256",
    secret: "super-secret",
    aud: "test_audience",
    iss: "https://login.microsoftonline.com/test_aad_id/v2.0",
    iat: now,
    nbf: now,
    exp: timeInterval,
    aio: "test_aio",
    name: testDisplayName,
    oid: testObjectId,
    preferred_username: testPreferredUserName,
    rh: "test_rh",
    scp: "access_as_user",
    sub: "test_sub",
    tid: testTenantId,
    uti: "test_uti",
    ver: "2.0",
  });
  const authConfig: OnBehalfOfCredentialAuthConfig = {
    clientId: clientId,
    clientSecret: clientSecret,
    authorityHost: authorityHost,
    tenantId: tenantId,
  };

  const sandbox = sinon.createSandbox();

  beforeEach(function () {
    // Mock ConfidentialClientApplication implementation
    sandbox
      .stub(ConfidentialClientApplication.prototype, "acquireTokenOnBehalfOf")
      .callsFake((): Promise<AuthenticationResult | null> => {
        const authResult: AuthenticationResult = {
          authority: "fake_authority",
          uniqueId: "fake_uniqueId",
          tenantId: "fake_tenant_id",
          scopes: [],
          account: null,
          idToken: "fake_id_token",
          idTokenClaims: new Object(),
          accessToken: accessToken,
          fromCache: false,
          tokenType: "fake_tokenType",
          correlationId: "fake_correlation_id",
          expiresOn: accessTokenExpDate,
        };
        return new Promise<AuthenticationResult>((resolve) => {
          resolve(authResult);
        });
      });
  });

  afterEach(function () {
    sandbox.restore();
  });

  it("create OnBehalfOfUserCredential instance should not throw InvalidConfiguration Error when clientSecret not found", async function () {
    expect(() => {
      new OnBehalfOfUserCredential(ssoToken, {
        clientId: clientId,
        certificateContent: certificateContent,
        authorityHost: authorityHost,
        tenantId: tenantId,
      });
    }).to.not.throw();
  });

  it("create OnBehalfOfUserCredential instance should not throw InvalidConfiguration Error when certificateContent not found", async function () {
    expect(() => {
      new OnBehalfOfUserCredential(ssoToken, {
        clientId: clientId,
        clientSecret: clientSecret,
        authorityHost: authorityHost,
        tenantId: tenantId,
      });
    }).to.not.throw();
  });

  it("create OnBehalfOfUserCredential instance should not throw InvalidConfiguration Error and respect certificateContent when clientSecret and certificateContent are both set", async function () {
    const oboCredential: any = new OnBehalfOfUserCredential(ssoToken, {
      clientId: clientId,
      clientSecret: clientSecret,
      certificateContent: certificateContent,
      authorityHost: authorityHost,
      tenantId: tenantId,
    } as unknown as OnBehalfOfCredentialAuthConfig);

    // certificateContent has higher priority than clientSecret
    assert.strictEqual(
      oboCredential.msalClient.config.auth.clientCertificate.thumbprintSha256,
      "90AF5A3B906DCC32226BCCD6D369165CFB9F1E0FE123F0D18B7CC48261995A6C" // thumbprint is calculated from certificate content "fakeCert"
    );
    assert.strictEqual(oboCredential.msalClient.config.auth.clientSecret, "");
  });

  it("create OnBehalfOfUserCredential instance should throw InvalidConfiguration Error when clientId not found", async function () {
    expect(() => {
      new OnBehalfOfUserCredential(ssoToken, {
        clientSecret: clientSecret,
        authorityHost: authorityHost,
        tenantId: tenantId,
      } as unknown as OnBehalfOfCredentialAuthConfig);
    })
      .to.throw(ErrorWithCode, "clientId in configuration is invalid: undefined")
      .with.property("code", InvalidConfiguration);
  });

  it("create OnBehalfOfUserCredential instance should throw InvalidConfiguration Error when authorityHost not found", async function () {
    expect(() => {
      new OnBehalfOfUserCredential(ssoToken, {
        clientSecret: clientSecret,
        clientId: clientId,
        tenantId: tenantId,
      } as unknown as OnBehalfOfCredentialAuthConfig);
    })
      .to.throw(ErrorWithCode, "authorityHost in configuration is invalid: undefined")
      .with.property("code", InvalidConfiguration);
  });

  it("create OnBehalfOfUserCredential instance should throw InvalidConfiguration Error when clientSecret, certificateContent not found", async function () {
    expect(() => {
      new OnBehalfOfUserCredential(ssoToken, {
        clientId: clientId,
        authorityHost: authorityHost,
        tenantId: tenantId,
      } as unknown as OnBehalfOfCredentialAuthConfig);
    })
      .to.throw(
        ErrorWithCode,
        "clientSecret or certificateContent in configuration is invalid: undefined"
      )
      .with.property("code", InvalidConfiguration);
  });

  it("create OnBehalfOfUserCredential instance should throw InvalidConfiguration Error when tenantId not found", async function () {
    expect(() => {
      new OnBehalfOfUserCredential(ssoToken, {
        clientId: clientId,
        clientSecret: clientSecret,
        authorityHost: authorityHost,
      } as unknown as OnBehalfOfCredentialAuthConfig);
    })
      .to.throw(ErrorWithCode, "tenantId in configuration is invalid: undefined")
      .with.property("code", InvalidConfiguration);
  });

  it("create OnBehalfOfUserCredential instance should throw InvalidConfiguration Error when clientId, clientSecret, certificateContent, authorityHost, tenantId not found", async function () {
    expect(() => {
      new OnBehalfOfUserCredential(ssoToken, {} as unknown as OnBehalfOfCredentialAuthConfig);
    })
      .to.throw(
        ErrorWithCode,
        "clientId, authorityHost, clientSecret or certificateContent, tenantId in configuration is invalid: undefined"
      )
      .with.property("code", InvalidConfiguration);
  });

  it("create OnBehalfOfUserCredential instance should throw InternalError with invalid sso token", async function () {
    const invalidSsoToken = "invalid_sso_token";

    expect(() => {
      new OnBehalfOfUserCredential(invalidSsoToken, authConfig);
    })
      .to.throw(ErrorWithCode, "Parse jwt token failed in node env with error: ")
      .with.property("code", InternalError);
  });

  it("create OnBehalfOfUserCredential instance should throw InvalidCertificate with invalid certificate", async function () {
    expect(() => {
      new OnBehalfOfUserCredential(ssoToken, {
        clientId: clientId,
        certificateContent: "invalid_certificate_content",
        authorityHost: authorityHost,
        tenantId: tenantId,
      });
    })
      .to.throw(
        ErrorWithCode,
        "The certificate content does not contain a PEM-encoded certificate."
      )
      .with.property("code", ErrorCode.InvalidCertificate);
  });

  it("getToken should success for client certificate", async function () {
    const oboCredential = new OnBehalfOfUserCredential(ssoToken, {
      clientId: clientId,
      certificateContent: certificateContent,
      authorityHost: authorityHost,
      tenantId: tenantId,
    });
    const token = await oboCredential.getToken(scope);
    assert.strictEqual(token?.token, accessToken);
    assert.strictEqual(token?.expiresOnTimestamp, accessTokenExpNumber);
  });

  it("getToken should success when scopes is empty string", async function () {
    const oboCredential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const token = await oboCredential.getToken("");
    assert.strictEqual(token?.token, ssoToken);
    assert.strictEqual(token?.expiresOnTimestamp, ssoTokenExp);
  });

  it("getToken should success when scopes is empty array", async function () {
    const oboCredential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const token = await oboCredential.getToken([]);
    assert.strictEqual(token?.token, ssoToken);
    assert.strictEqual(token?.expiresOnTimestamp, ssoTokenExp);
  });

  it("getToken should success when scopes is string", async function () {
    const oboCredential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const token = await oboCredential.getToken(scope);
    assert.strictEqual(token?.token, accessToken);
    assert.strictEqual(token?.expiresOnTimestamp, accessTokenExpNumber);
  });

  it("getToken should success when scopes is string array", async function () {
    const oboCredential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const scopesArray: string[] = [scope, "fake_scope_2"];
    const token = await oboCredential.getToken(scopesArray);
    assert.strictEqual(token?.token, accessToken);
    assert.strictEqual(token?.expiresOnTimestamp, accessTokenExpNumber);
  });

  it("getToken should throw TokenExpiredError when get SSO token with sso token expired", async function () {
    sandbox.restore();
    sandbox
      .stub(ConfidentialClientApplication.prototype, "acquireTokenOnBehalfOf")
      .callsFake((): Promise<AuthenticationResult | null> => {
        return new Promise<AuthenticationResult>(() => {
          throw new AuthError(
            "ServerError: invalid_grant: 50013",
            "AADSTS50013: Assertionfailed signature validation"
          );
        });
      });

    const expiredSsoToken =
      // eslint-disable-next-line no-secrets/no-secrets
      "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiJjZWVkYTJjNi00MDBmLTQyYjMtYjE4ZC1jY2NmYzk5NjM4NmYiLCJpc3MiOiJodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGluZS5jb20vNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3L3YyLjAiLCJpYXQiOjE2MTk0OTI3MzEsIm5iZiI6MTYxOTQ5MjczMSwiZXhwIjoxNjE5NDk2NjMxLCJhaW8iOiJBVFFBeS84VEFBQUFFWDZLU0prRjlOaEFDL1NXV1hWTXFPVDNnNGZXR2dqS0ZEWjRramlEb25OVlY2cDlZTVFMaTFqVXdHWEZaclpaIiwiYXpwIjoiYjBjNDdmMjktM2M1Ny00MDQyLTkzM2YtYTdkNTQ2YmFlMzg3IiwiYXpwYWNyIjoiMCIsIm5hbWUiOiJNZXRhIE9TIHNlcnZpY2UgYWNjb3VudCBmb3IgZGV2ZWxvcG1lbnQiLCJvaWQiOiIyYTYxYzRjMy1lY2Y5LTQ5ZWItYjcxNy02NjczZmZmZDg5MmQiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiJtZXRhZGV2QG1pY3Jvc29mdC5jb20iLCJyaCI6IjAuQVFFQXY0ajVjdkdHcjBHUnF5MTgwQkhiUnlsX3hMQlhQRUpBa3otbjFVYTY0NGNhQUpRLiIsInNjcCI6ImFjY2Vzc19hc191c2VyIiwic3ViIjoiNEhUVXFCbWVBQVFWa2ZrbU0wcFRtVHh3QjRkcDdITGtxSjRSYXFvb3dUTSIsInRpZCI6IjcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0NyIsInV0aSI6ImFVQkxZSENBWmsyZE9LNW1wR2ctQUEiLCJ2ZXIiOiIyLjAifQ.QCkyqat72TS85vQ6h-jqAj-pnAOOkeOy3-WxgEQ1DJbW6fsoXmVGgso-ncMmeiYIoA1r9jy1cBfnEMBI1tBKcq4TOHseyde2uM-pxCGHNhFC_WiWy9KXKiou5bvgXdVqqCT7CQejpiNdm3wL-EFhXWBRj6OlLMLcUtnlcnKfOSmx8IIOuQrCjWtuE_wjpfo2AwkguuJ5defyOkYqlCfcJ9FyUrqhqsONMdh0lJiVY94PZ00UTjH3zPaC2tnKrGeXn-qrr9dccEUx2HqyAfdzPwymBLWMCrirVRKCZV3DtfKuozKkIxIPZz0891QZcFO8VgfBJaLmr6J7EL8lPtFKnw";
    const credential = new OnBehalfOfUserCredential(expiredSsoToken, authConfig);
    let err = await expect(credential.getToken([])).to.eventually.be.rejectedWith(ErrorWithCode);
    assert.strictEqual(err.code, ErrorCode.TokenExpiredError);
    assert.strictEqual(err.message, "Sso token has already expired.");

    err = await expect(credential.getToken("")).to.eventually.be.rejectedWith(ErrorWithCode);
    assert.strictEqual(err.code, ErrorCode.TokenExpiredError);
    assert.strictEqual(err.message, "Sso token has already expired.");

    err = await expect(credential.getToken(scope)).to.eventually.be.rejectedWith(ErrorWithCode);
    assert.strictEqual(err.code, ErrorCode.TokenExpiredError);
    assert.isTrue(
      err.message?.indexOf(
        "Failed to get access token from AAD server, assertion is invalid because of various reasons: "
      ) >= 0
    );
  });

  it("getToken should throw ServiceError when fail to get access token due to AAD outage", async function () {
    // Mock AAD outage
    sandbox.restore();
    sandbox
      .stub(ConfidentialClientApplication.prototype, "acquireTokenOnBehalfOf")
      .callsFake((): Promise<AuthenticationResult | null> => {
        return new Promise<AuthenticationResult>(() => {
          throw new AuthError("AAD Outage", "AAD Outage");
        });
      });
    const oboCredential = new OnBehalfOfUserCredential(ssoToken, authConfig);

    const errorResult = await expect(oboCredential.getToken(scope)).to.eventually.be.rejectedWith(
      ErrorWithCode
    );
    assert.strictEqual(errorResult.code, ServiceError);
    assert.isTrue(
      errorResult.message?.indexOf("Failed to acquire access token on behalf of user: ") >= 0
    );
  });

  it("getToken should throw FailToAcquireTokenOnBehalfOfUser when access token is null", async function () {
    sandbox.restore();
    sandbox
      .stub(ConfidentialClientApplication.prototype, "acquireTokenOnBehalfOf")
      .callsFake((): Promise<AuthenticationResult | null> => {
        return Promise.resolve(null);
      });

    const credential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const err = await expect(credential.getToken(scope)).to.eventually.be.rejectedWith(
      ErrorWithCode
    );
    assert.strictEqual(err.code, ErrorCode.InternalError);
    assert.strictEqual(
      err.message,
      "Failed to acquire access token on behalf of user: Access token is null"
    );
  });

  it("getUserInfo should succeed", async function () {
    const oboCredential = new OnBehalfOfUserCredential(ssoToken, authConfig);
    const userinfo: UserInfo = oboCredential.getUserInfo();
    assert.strictEqual(userinfo.displayName, testDisplayName);
    assert.strictEqual(userinfo.objectId, testObjectId);
    assert.strictEqual(userinfo.tenantId, testTenantId);
    assert.strictEqual(userinfo.preferredUserName, testPreferredUserName);
  });
});
