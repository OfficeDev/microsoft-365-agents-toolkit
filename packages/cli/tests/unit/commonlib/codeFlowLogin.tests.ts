// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import sinon from "sinon";
import { AccountInfo } from "@azure/msal-node";
import { expect } from "../utils";
import { CodeFlowLogin } from "../../../src/commonlib/codeFlowLogin";
import CliTelemetry from "../../../src/telemetry/cliTelemetry";
import * as cacheAccess from "../../../src/commonlib/cacheAccess";

describe("CodeFlowLogin.loginWithBroker", function () {
  const sandbox = sinon.createSandbox();

  // A minimal JWT-like token: header.payload.signature
  // payload = base64({"oid":"fake-oid","upn":"test@test.com"})
  const fakeAccessToken =
    "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9." +
    Buffer.from(JSON.stringify({ oid: "fake-oid", upn: "test@test.com" })).toString("base64") +
    ".fake-signature";

  const fakeResponse = {
    account: {
      homeAccountId: "fake-id",
      environment: "login.microsoftonline.com",
      tenantId: "fake-tenant",
      username: "test@test.com",
      localAccountId: "fake-local-id",
    },
    accessToken: fakeAccessToken,
  };

  const config = {
    auth: {
      clientId: "fake-client-id",
      authority: "https://login.microsoftonline.com/common",
    },
  };

  afterEach(() => {
    sandbox.restore();
  });

  function setupLogin(accountName: string) {
    sandbox.stub(CliTelemetry, "sendTelemetryEvent");

    const codeFlowLogin = new CodeFlowLogin([], config, 0, accountName);
    let capturedRequest: any;
    sandbox.stub(codeFlowLogin.pca, "acquireTokenInteractive").callsFake(async (request: any) => {
      capturedRequest = request;
      return fakeResponse as any;
    });
    sandbox.stub(codeFlowLogin as any, "mutex").value({
      runExclusive: async (fn: any) => fn(),
    });

    return { codeFlowLogin, getCapturedRequest: () => capturedRequest };
  }

  it("should replace accountName placeholder with M365 in loopback template for m365 account", async () => {
    const { codeFlowLogin, getCapturedRequest } = setupLogin("appStudio");

    await codeFlowLogin.loginWithBroker(["scope1"]);
    const req = getCapturedRequest();

    expect(req.successTemplate).to.include("M365 - Sign In");
    expect(req.successTemplate).to.not.include("$" + "{accountName}");
    expect(req.errorTemplate).to.include("M365 - Sign In");
    expect(req.errorTemplate).to.not.include("$" + "{accountName}");
  });

  it("should replace accountName placeholder with Azure in loopback template for azure account", async () => {
    const { codeFlowLogin, getCapturedRequest } = setupLogin("azure");

    await codeFlowLogin.loginWithBroker(["scope1"]);
    const req = getCapturedRequest();

    expect(req.successTemplate).to.include("Azure - Sign In");
    expect(req.successTemplate).to.not.include("$" + "{accountName}");
    expect(req.errorTemplate).to.include("Azure - Sign In");
    expect(req.errorTemplate).to.not.include("$" + "{accountName}");
  });
});

describe("CodeFlowLogin.logout", function () {
  const sandbox = sinon.createSandbox();

  const config = {
    auth: {
      clientId: "fake-client-id",
      authority: "https://login.microsoftonline.com/common",
    },
  };

  afterEach(() => {
    sandbox.restore();
  });

  function createMockAccount(homeAccountId: string): Pick<AccountInfo, "homeAccountId"> {
    return { homeAccountId };
  }

  it("should only sign out cached account when broker is available", async () => {
    const codeFlowLogin = new CodeFlowLogin([], config, 0, "appStudio");
    codeFlowLogin.isBrokerAvailable = true;

    const accountA = createMockAccount("account-a");
    const accountB = createMockAccount("account-b");

    sandbox.stub(cacheAccess, "loadAccountId").resolves("account-b");
    sandbox.stub(cacheAccess, "clearCache").resolves();
    sandbox.stub(cacheAccess, "saveAccountId").resolves();
    sandbox.stub(cacheAccess, "saveTenantId").resolves();
    sandbox
      .stub(codeFlowLogin.pca, "getAllAccounts")
      .resolves([accountA, accountB] as AccountInfo[]);
    const signOutStub = sandbox.stub(codeFlowLogin.pca, "signOut").resolves();

    const result = await codeFlowLogin.logout();

    expect(result).to.equal(true);
    expect(signOutStub.calledOnce).to.equal(true);
    expect(signOutStub.firstCall.firstArg.account.homeAccountId).to.equal(accountB.homeAccountId);
  });

  it("should sign out all accounts when broker is not available", async () => {
    const codeFlowLogin = new CodeFlowLogin([], config, 0, "appStudio");
    codeFlowLogin.isBrokerAvailable = false;

    const accountA = createMockAccount("account-a");
    const accountB = createMockAccount("account-b");

    sandbox.stub(cacheAccess, "clearCache").resolves();
    sandbox.stub(cacheAccess, "saveAccountId").resolves();
    sandbox.stub(cacheAccess, "saveTenantId").resolves();
    sandbox
      .stub(codeFlowLogin.pca, "getAllAccounts")
      .resolves([accountA, accountB] as AccountInfo[]);
    const signOutStub = sandbox.stub(codeFlowLogin.pca, "signOut").resolves();

    const result = await codeFlowLogin.logout();

    expect(result).to.equal(true);
    expect(signOutStub.callCount).to.equal(2);
    expect(signOutStub.firstCall.firstArg.account.homeAccountId).to.equal(accountA.homeAccountId);
    expect(signOutStub.secondCall.firstArg.account.homeAccountId).to.equal(accountB.homeAccountId);
  });
});
