// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SpecParser } from "@microsoft/m365-spec-parser";
import { SystemError, err } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClient";
import { setTools } from "../../../../src/common/globalVars";
import { CreateApiKeyDriver } from "../../../../src/component/driver/apiKey/create";
import {
  ApiSecretRegistrationAppType,
  ApiSecretRegistrationTargetAudience,
} from "../../../../src/component/driver/teamsApp/interfaces/ApiSecretRegistration";
import { UserCancelError } from "../../../../src/error";
import * as visitor from "../../../../src/ui/visitor";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";

chai.use(chaiAsPromised);
const expect = chai.expect;

const outputKeys = {
  registrationId: "REGISTRATION_ID",
};

const outputEnvVarNames = new Map<string, string>(Object.entries(outputKeys));

describe("CreateApiKeyDriver", () => {
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    ui: new MockedUserInteraction(),
  };
  const createApiKeyDriver = new CreateApiKeyDriver();

  let envRestore: RestoreFn | undefined;

  beforeEach(() => {
    setTools({
      ui: new MockedUserInteraction(),
      logProvider: new MockedLogProvider(),
      tokenProvider: {
        azureAccountProvider: new MockedAzureAccountProvider(),
        m365TokenProvider: new MockedM365Provider(),
      },
    });
  });

  afterEach(() => {
    sinon.restore();
    if (envRestore) {
      envRestore();
      envRestore = undefined;
    }
  });

  it("happy path: create registraionid, read domain from api spec, clientSecret from input", async () => {
    sinon.stub(teamsDevPortalClient, "createApiKeyRegistration").resolves({
      id: "mockedRegistrationId",
      clientSecrets: [],
      targetUrlsShouldStartWith: [],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
    });
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "apiKey",
              in: "header",
              name: "test",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.get(outputKeys.registrationId)).to.equal("mockedRegistrationId");
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: create registraionid, read domain from api spec, clientSecret and secondaryClientSecret from input", async () => {
    sinon.stub(teamsDevPortalClient, "createApiKeyRegistration").resolves({
      id: "mockedRegistrationId",
      clientSecrets: [],
      targetUrlsShouldStartWith: [],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
    });

    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      secondaryClientSecret: "mockedSecondaryClientSecret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.get(outputKeys.registrationId)).to.equal("mockedRegistrationId");
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: create registraionid and read domain from env and secret from env", async () => {
    sinon.stub(teamsDevPortalClient, "createApiKeyRegistration").resolves({
      id: "mockedRegistrationId",
      clientSecrets: [],
      targetUrlsShouldStartWith: [],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
    });

    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    envRestore = mockedEnv({
      ["api-key"]: "existingvalue",
    });
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.get(outputKeys.registrationId)).to.equal("mockedRegistrationId");
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: registration id exists in env", async () => {
    sinon.stub(teamsDevPortalClient, "getApiKeyRegistrationById").resolves({
      id: "mockedRegistrationId",
      clientSecrets: [],
      targetUrlsShouldStartWith: [],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
    });

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    envRestore = mockedEnv({
      [outputKeys.registrationId]: "existing value",
    });
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(0);
    }
  });

  it("happy path: create registrationid, read applicableToApps and targetAudience from input", async () => {
    sinon
      .stub(teamsDevPortalClient, "createApiKeyRegistration")
      .callsFake(async (token, apiKey) => {
        expect(apiKey.targetAudience).equals(ApiSecretRegistrationTargetAudience.HomeTenant);
        expect(apiKey.specificAppId).equals("mockedAppId");
        expect(apiKey.applicableToApps).equals(ApiSecretRegistrationAppType.SpecificApp);
        return {
          id: "mockedRegistrationId",
          clientSecrets: [],
          targetUrlsShouldStartWith: [],
          applicableToApps: ApiSecretRegistrationAppType.AnyApp,
          targetAudience: ApiSecretRegistrationTargetAudience.AnyTenant,
        };
      });
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
      applicableToApps: "SpecificApp",
      targetAudience: "HomeTenant",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.get(outputKeys.registrationId)).to.equal("mockedRegistrationId");
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: create registraionid, read domain from api spec, clientSecret from input with invalid api", async () => {
    sinon.stub(teamsDevPortalClient, "createApiKeyRegistration").resolves({
      id: "mockedRegistrationId",
      clientSecrets: [],
      targetUrlsShouldStartWith: [],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
    });
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: false,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.get(outputKeys.registrationId)).to.equal("mockedRegistrationId");
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("should throw error when empty outputEnvVarNames", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, undefined);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("OutputEnvironmentVariableUndefined");
    }
  });

  it("should throw error when failed to get app studio token", async () => {
    sinon
      .stub(MockedM365Provider.prototype, "getAccessToken")
      .resolves(err(new SystemError("source", "name", "message")));
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("name");
    }
  });

  it("should show warning if registration id exists and failed to get API key", async () => {
    sinon
      .stub(teamsDevPortalClient, "getApiKeyRegistrationById")
      .throws(new SystemError("source", "name", "message"));

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    envRestore = mockedEnv({
      [outputKeys.registrationId]: "existing value",
    });
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isOk()).to.be.true;
  });

  it("should throw error if missing name", async () => {
    const args: any = {
      name: "",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if name is too long", async () => {
    const args: any = {
      name: "a".repeat(513),
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyNameTooLong");
    }
  });

  it("should throw error if missing appId", async () => {
    const args: any = {
      name: "test",
      appId: "",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if invalid clientSecret", async () => {
    const args: any = {
      name: "test",
      appId: "",
      primaryClientSecret: "secret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyClientSecretInvalid");
    }
  });

  it("should throw error if clientSecret equals space when from scratch", async () => {
    const args: any = {
      name: "test",
      appId: "",
      primaryClientSecret: " ",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("apiKeyFromScratchClientSecretInvalid");
    }
  });

  it("should throw error if invalid secondaryClientSecret", async () => {
    const args: any = {
      name: "test",
      appId: "",
      primaryClientSecret: "mockedClientSecret",
      secondaryClientSecret: "secret",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyClientSecretInvalid");
    }
  });

  it("should throw error if missing apiSpecPath", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if domain > 1", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedSecret",
      apiSpecPath: "mockedPath",
    };

    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          api: "api",
          server: "https://test2",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 2,
      validAPICount: 2,
    });

    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyDomainInvalid");
    }
  });

  it("should throw error if list api is empty and domain = 0", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedSecret",
      apiSpecPath: "mockedPath",
    };
    sinon
      .stub(SpecParser.prototype, "list")
      .resolves({ APIs: [], validAPICount: 0, allAPICount: 1 });
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyAuthMissingInSpec");
    }
  });

  it("should throw error if list api contains no auth and domain = 0", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedSecret",
      apiSpecPath: "mockedPath",
    };
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 1,
      allAPICount: 1,
    });
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyAuthMissingInSpec");
    }
  });

  it("should throw error if list api contains unsupported auth and domain = 0", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedSecret",
      apiSpecPath: "mockedPath",
    };
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api1",
          server: "https://test",
          operationId: "get1",
          auth: {
            name: "test1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          api: "api2",
          server: "https://test",
          operationId: "get2",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "basic",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          api: "api3",
          server: "https://test",
          operationId: "get3",
          auth: {
            name: "test1",
            authScheme: {
              type: "apiKey",
              in: "header",
              name: "test1",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 3,
      allAPICount: 3,
    });
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyAuthMissingInSpec");
    }
  });

  it("should throw error if list api contains auth but server info is null", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedSecret",
      apiSpecPath: "mockedPath",
    };
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api1",
          server: "https://test",
          operationId: "get1",
          auth: {
            name: "test1",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          api: "api2",
          server: "",
          operationId: "get2",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
        {
          api: "api3",
          server: "https://test",
          operationId: "get3",
          auth: {
            name: "test1",
            authScheme: {
              type: "apiKey",
              in: "header",
              name: "test1",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      validAPICount: 3,
      allAPICount: 3,
    });
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyFailedToGetDomain");
    }
  });

  it("should throw error if failed to create API key", async () => {
    sinon
      .stub(teamsDevPortalClient, "createApiKeyRegistration")
      .throws(new SystemError("source", "name", "message"));

    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret, mockedClientSecret2",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("name");
    }
  });

  it("should throw unhandled error if error is not SystemError or UserError", async () => {
    sinon.stub(MockedM365Provider.prototype, "getAccessToken").throws(new Error("unhandled error"));
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret, mockedClientSecret2",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.source).to.equal("apiKeyRegister");
    }
  });

  it("should throw error if invalid applicableToApps and targetAudience", async () => {
    sinon.stub(teamsDevPortalClient, "createApiKeyRegistration").resolves({
      id: "mockedRegistrationId",
      clientSecrets: [],
      targetUrlsShouldStartWith: [],
      applicableToApps: ApiSecretRegistrationAppType.AnyApp,
      targetAudience: ApiSecretRegistrationTargetAudience.AnyTenant,
    });
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      primaryClientSecret: "mockedClientSecret",
      apiSpecPath: "mockedPath",
      applicableToApps: "specificapp",
      targetAudience: "hometenant",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
      expect(result.result.error.message.includes("applicableToApps")).to.be.true;
      expect(result.result.error.message.includes("targetAudience")).to.be.true;
    }
  });

  it("should throw error if user cancel", async () => {
    sinon.stub(teamsDevPortalClient, "createApiKeyRegistration").resolves({
      id: "mockedRegistrationId",
      clientSecrets: [],
      targetUrlsShouldStartWith: [],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
    });
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "http",
              scheme: "bearer",
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    sinon.stub(visitor, "traverse").resolves(err(new UserCancelError("apikey")));

    const args: any = {
      name: "test",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
    };
    const result = await createApiKeyDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.source).to.equal("apikey");
    }
  });
});
