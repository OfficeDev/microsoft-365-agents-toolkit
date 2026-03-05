// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import * as sinon from "sinon";
import * as daSpecParser from "../../../../src/common/daSpecParser";
import { getAuthInfo, OauthInfo } from "../../../../src/component/driver/oauth/utility/utility";
import { CreateOauthArgs } from "../../../../src/component/driver/oauth/interface/createOauthArgs";
import { UpdateOauthArgs } from "../../../../src/component/driver/oauth/interface/updateOauthArgs";
import { OauthAuthInfoInvalid } from "../../../../src/component/driver/oauth/error/oauthAuthInfoInvalid";
import { OauthAuthMissingInSpec } from "../../../../src/component/driver/oauth/error/oauthAuthMissingInSpec";

chai.use(chaiAsPromised);
const expect = chai.expect;

function makeOAuth2Operation(
  server: string,
  authName: string,
  flows: Record<string, any>
): any {
  return {
    api: "api",
    server,
    operationId: "get",
    auth: {
      name: authName,
      authScheme: {
        type: "oauth2",
        flows,
      },
    },
    isValid: true,
    reason: [],
  };
}

function makeListAPIResult(operations: any[]): any {
  return {
    APIs: operations,
    allAPICount: operations.length,
    validAPICount: operations.length,
  };
}

describe("getAuthInfo - getandValidateOauthInfoFromSpec flow handling", () => {
  const mockedDriverContext: any = {
    projectPath: "/mock/project",
  };
  const actionName = "testAction";

  afterEach(() => {
    sinon.restore();
  });

  it("should extract authorizationCode flow info", async () => {
    const flows = {
      authorizationCode: {
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        refreshUrl: "https://example.com/refresh",
        scopes: {
          "read:data": "Read data",
          "write:data": "Write data",
        },
      },
    };
    const operation = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([operation]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "authorizationCode",
    };

    const result = await getAuthInfo(args, mockedDriverContext, actionName);

    expect(result.domain).to.deep.equal(["https://api.example.com"]);
    expect(result.authorizationEndpoint).to.equal("https://example.com/authorize");
    expect(result.tokenExchangeEndpoint).to.equal("https://example.com/token");
    expect(result.tokenRefreshEndpoint).to.equal("https://example.com/refresh");
    expect(result.scopes).to.deep.equal(["read:data", "write:data"]);
  });

  it("should use authorizationCode flow as default when flow is unknown", async () => {
    const flows = {
      authorizationCode: {
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        scopes: {
          "scope1": "desc",
        },
      },
    };
    const operation = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([operation]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "unknownFlow",
    };

    const result = await getAuthInfo(args, mockedDriverContext, actionName);

    expect(result.authorizationEndpoint).to.equal("https://example.com/authorize");
    expect(result.tokenExchangeEndpoint).to.equal("https://example.com/token");
    expect(result.scopes).to.deep.equal(["scope1"]);
  });

  it("should throw OauthAuthInfoInvalid when authorizationCode flow is not present", async () => {
    const flows = {
      implicit: {
        authorizationUrl: "https://example.com/authorize",
        scopes: { "scope1": "desc" },
      },
    };
    const operation = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([operation]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "authorizationCode",
    };

    await expect(
      getAuthInfo(args, mockedDriverContext, actionName)
    ).to.be.rejectedWith(OauthAuthInfoInvalid);
  });

  it("should handle refreshUrl being undefined", async () => {
    const flows = {
      authorizationCode: {
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        scopes: { "scope1": "desc" },
      },
    };
    const operation = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([operation]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "authorizationCode",
    };

    const result = await getAuthInfo(args, mockedDriverContext, actionName);

    expect(result.tokenRefreshEndpoint).to.be.undefined;
  });

  it("should extract scopes as keys from the scopes object", async () => {
    const flows = {
      authorizationCode: {
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        scopes: {
          "User.Read": "Read user profile",
          "Mail.Send": "Send mail",
          "Files.ReadWrite": "Read and write files",
        },
      },
    };
    const operation = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([operation]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "authorizationCode",
    };

    const result = await getAuthInfo(args, mockedDriverContext, actionName);

    expect(result.scopes).to.deep.equal(["User.Read", "Mail.Send", "Files.ReadWrite"]);
  });

  it("should throw OauthAuthInfoInvalid when multiple operations have different auth info", async () => {
    const flows1 = {
      authorizationCode: {
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        scopes: { "scope1": "desc" },
      },
    };
    const flows2 = {
      authorizationCode: {
        authorizationUrl: "https://other.com/authorize",
        tokenUrl: "https://other.com/token",
        scopes: { "scope2": "desc" },
      },
    };
    const op1 = makeOAuth2Operation("https://api.example.com", "oauthTest", flows1);
    const op2 = makeOAuth2Operation("https://api.example.com", "oauthTest", flows2);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([op1, op2]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "authorizationCode",
    };

    await expect(
      getAuthInfo(args, mockedDriverContext, actionName)
    ).to.be.rejectedWith(OauthAuthInfoInvalid);
  });

  it("should deduplicate operations with identical auth info", async () => {
    const flows = {
      authorizationCode: {
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        refreshUrl: "https://example.com/refresh",
        scopes: { "scope1": "desc" },
      },
    };
    const op1 = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    const op2 = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([op1, op2]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "authorizationCode",
    };

    const result = await getAuthInfo(args, mockedDriverContext, actionName);

    expect(result.authorizationEndpoint).to.equal("https://example.com/authorize");
    expect(result.tokenExchangeEndpoint).to.equal("https://example.com/token");
  });

  it("should throw OauthAuthMissingInSpec when no matching oauth operations found", async () => {
    const flows = {
      authorizationCode: {
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        scopes: { "scope1": "desc" },
      },
    };
    const operation = makeOAuth2Operation("https://api.example.com", "differentName", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([operation]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "authorizationCode",
    };

    await expect(
      getAuthInfo(args, mockedDriverContext, actionName)
    ).to.be.rejectedWith(OauthAuthMissingInSpec);
  });

  it("should default flow to authorizationCode when flow property is not in args", async () => {
    const flows = {
      authorizationCode: {
        authorizationUrl: "https://example.com/authorize",
        tokenUrl: "https://example.com/token",
        scopes: { "scope1": "desc" },
      },
    };
    const operation = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([operation]));

    // UpdateOauthArgs does not have 'flow' property
    const args: UpdateOauthArgs = {
      name: "oauthTest",
      apiSpecPath: "spec.yaml",
      configurationId: "configId",
    };

    const result = await getAuthInfo(args, mockedDriverContext, actionName);

    expect(result.authorizationEndpoint).to.equal("https://example.com/authorize");
    expect(result.tokenExchangeEndpoint).to.equal("https://example.com/token");
    expect(result.scopes).to.deep.equal(["scope1"]);
  });

  it("should throw OauthAuthInfoInvalid when flows object has no matching flow type", async () => {
    const flows = {
      clientCredentials: {
        tokenUrl: "https://example.com/token",
        scopes: { "scope1": "desc" },
      },
    };
    const operation = makeOAuth2Operation("https://api.example.com", "oauthTest", flows);
    sinon.stub(daSpecParser, "listAPIInfo").resolves(makeListAPIResult([operation]));

    const args: CreateOauthArgs = {
      name: "oauthTest",
      appId: "appId",
      apiSpecPath: "spec.yaml",
      flow: "authorizationCode",
    };

    await expect(
      getAuthInfo(args, mockedDriverContext, actionName)
    ).to.be.rejectedWith(OauthAuthInfoInvalid);
  });
});
