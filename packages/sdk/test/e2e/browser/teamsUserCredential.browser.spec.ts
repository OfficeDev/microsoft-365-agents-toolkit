// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as chai from "chai";
import * as chaiPromises from "chai-as-promised";
import { AccessToken } from "@azure/core-auth";
import * as sinon from "sinon";
import { TeamsUserCredential, ErrorWithCode, UserInfo } from "../../../src/index.browser";
import {
  extractIntegrationEnvVariables,
  getSSOToken,
  AADJwtPayLoad,
  SSOToken,
  getGraphToken,
} from "../helper.browser";
import { jwtDecode } from "jwt-decode";
import { AccountInfo, AuthenticationResult, PublicClientApplication } from "@azure/msal-browser";

chai.use(chaiPromises);
const env = (window as any).__env__;
extractIntegrationEnvVariables();

describe("TeamsUserCredential Tests - Browser", () => {
  const TEST_USER_ACCOUNT_NAME = env.SDK_INTEGRATION_TEST_ACCOUNT_NAME;
  const TEST_USER_OBJECT_ID = env.SDK_INTEGRATION_TEST_USER_OBJECT_ID;
  const TEST_AAD_TENANT_ID = env.SDK_INTEGRATION_TEST_AAD_TENANT_ID;
  const UIREQUIREDERROR = "UiRequiredError";
  const FAKE_LOGIN_ENDPOINT = "FakeLoginEndpoint";
  let ssoToken: SSOToken;

  beforeEach(async () => {
    ssoToken = await getSSOToken();
    sinon
      .stub(TeamsUserCredential.prototype, <any>"getSSOToken")
      .callsFake((): Promise<AccessToken | null> => {
        return new Promise((resolve) => {
          resolve({
            token: ssoToken.token!,
            expiresOnTimestamp: ssoToken.expire_time!,
          });
        });
      });
  });

  afterEach(() => {
    sinon.restore();
  });

  it("GetUserInfo should success with SSOToken", async function () {
    const credential: TeamsUserCredential = new TeamsUserCredential({
      initiateLoginEndpoint: FAKE_LOGIN_ENDPOINT,
      clientId: env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID,
    });
    const info: UserInfo = await credential.getUserInfo();
    chai.assert.strictEqual(info.preferredUserName, TEST_USER_ACCOUNT_NAME);
    chai.assert.strictEqual(info.objectId, TEST_USER_OBJECT_ID);
    chai.assert.strictEqual(info.tenantId, TEST_AAD_TENANT_ID);
  });

  it("GetToken should success with consent scope", async function () {
    const credential: TeamsUserCredential = new TeamsUserCredential({
      initiateLoginEndpoint: FAKE_LOGIN_ENDPOINT,
      clientId: env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID,
    });
    sinon
      .stub(PublicClientApplication.prototype, <any>"acquireTokenSilent")
      .callsFake(async (): Promise<AuthenticationResult> => {
        const graphToken = await getGraphToken(ssoToken, "User.Read");
        return new Promise((resolve) => {
          resolve({
            authority: "authority",
            uniqueId: "uniqueId",
            tenantId: "tenantId",
            scopes: ["User.Read"],
            account: {} as unknown as AccountInfo,
            idToken: "idToken",
            idTokenClaims: {},
            accessToken: graphToken,
            fromCache: false,
            expiresOn: null,
            tokenType: "tokenType",
            correlationId: "correlationId",
          });
        });
      });

    // await expect(credential.getToken(["User.Read"])).to.be.eventually.have.property("token");
    const accessToken = await credential.getToken(["User.Read"]);
    const decodedToken = jwtDecode<AADJwtPayLoad>(accessToken!.token);
    chai.assert.strictEqual(decodedToken.aud, "00000003-0000-0000-c000-000000000000");
    chai.assert.isTrue(decodedToken.scp!.includes("User.Read"));
  });

  it("GetToken should throw UiRequiredError with unconsent scope", async function () {
    const credential: TeamsUserCredential = new TeamsUserCredential({
      initiateLoginEndpoint: FAKE_LOGIN_ENDPOINT,
      clientId: env.SDK_INTEGRATION_TEST_M365_AAD_CLIENT_ID,
    });
    await chai
      .expect(credential.getToken(["Calendars.Read"]))
      .to.eventually.be.rejectedWith(ErrorWithCode)
      .and.property("code", UIREQUIREDERROR);
  });
});
