// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SpecParser } from "@microsoft/m365-spec-parser";
import { ConfirmConfig, UserError, err, ok } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import "mocha";
import { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClient";
import { setTools } from "../../../../src/common/globalVars";
import { UpdateApiKeyArgs } from "../../../../src/component/driver/apiKey/interface/updateApiKeyArgs";
import { UpdateApiKeyDriver } from "../../../../src/component/driver/apiKey/update";
import {
  ApiSecretRegistrationAppType,
  ApiSecretRegistrationTargetAudience,
} from "../../../../src/component/driver/teamsApp/interfaces/ApiSecretRegistration";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";

chai.use(chaiAsPromised);
const expect = chai.expect;

describe("UpdateApiKeyDriver", () => {
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    ui: new MockedUserInteraction(),
  };
  const updateApiKeyDriver = new UpdateApiKeyDriver();

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

  it("happy path: update all fields", async () => {
    sinon.stub(teamsDevPortalClient, "updateApiKeyRegistration").resolves({
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test2"],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
      targetAudience: ApiSecretRegistrationTargetAudience.HomeTenant,
      specificAppId: "mockedAppId",
    });
    sinon.stub(teamsDevPortalClient, "getApiKeyRegistrationById").resolves({
      id: "mockedRegistrationId",
      description: "mockedDescription",
      clientSecrets: [],
      targetUrlsShouldStartWith: ["https://test"],
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
        {
          api: "api2",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test2",
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

    sinon.stub(mockedDriverContext.ui, "confirm").callsFake(async (config) => {
      expect((config as ConfirmConfig).title.includes("description")).to.be.true;
      expect((config as ConfirmConfig).title.includes("applicableToApps")).to.be.true;
      expect((config as ConfirmConfig).title.includes("specificAppId")).to.be.true;
      expect((config as ConfirmConfig).title.includes("targetAudience")).to.be.true;
      return ok({ type: "success", value: true });
    });

    const args: UpdateApiKeyArgs = {
      name: "test2",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      registrationId: "mockedRegistrationId",
    };

    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: update all fields with baseURL", async () => {
    sinon.stub(teamsDevPortalClient, "updateApiKeyRegistration").resolves({
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test2"],
      applicableToApps: ApiSecretRegistrationAppType.SpecificApp,
      targetAudience: ApiSecretRegistrationTargetAudience.HomeTenant,
      specificAppId: "mockedAppId",
    });
    sinon.stub(teamsDevPortalClient, "getApiKeyRegistrationById").resolves({
      id: "mockedRegistrationId",
      description: "mockedDescription",
      clientSecrets: [],
      targetUrlsShouldStartWith: ["https://test"],
      applicableToApps: ApiSecretRegistrationAppType.AnyApp,
      targetAudience: ApiSecretRegistrationTargetAudience.AnyTenant,
    });

    sinon.stub(mockedDriverContext.ui, "confirm").callsFake(async (config) => {
      expect((config as ConfirmConfig).title.includes("description")).to.be.true;
      expect((config as ConfirmConfig).title.includes("applicableToApps")).to.be.true;
      expect((config as ConfirmConfig).title.includes("specificAppId")).to.be.true;
      expect((config as ConfirmConfig).title.includes("targetAudience")).to.be.true;
      return ok({ type: "success", value: true });
    });

    const args: UpdateApiKeyArgs = {
      name: "test2",
      appId: "mockedAppId",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      registrationId: "mockedRegistrationId",
      baseUrl: "https://test",
    };

    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: does not update when no changes", async () => {
    sinon.stub(teamsDevPortalClient, "getApiKeyRegistrationById").resolves({
      id: "test",
      description: "test",
      clientSecrets: [],
      targetUrlsShouldStartWith: ["https://test"],
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
        {
          api: "api2",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test2",
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
    const args: UpdateApiKeyArgs = {
      name: "test",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "AnyTenant",
      applicableToApps: "AnyApp",
      registrationId: "mockedRegistrationId",
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: should not show confirm when only devtunnel url is different", async () => {
    sinon.stub(teamsDevPortalClient, "updateApiKeyRegistration").resolves({
      description: "test",
      targetUrlsShouldStartWith: ["https://test2.asse.devtunnels.ms"],
      applicableToApps: ApiSecretRegistrationAppType.AnyApp,
      targetAudience: ApiSecretRegistrationTargetAudience.AnyTenant,
    });
    sinon.stub(teamsDevPortalClient, "getApiKeyRegistrationById").resolves({
      id: "test",
      description: "test",
      clientSecrets: [],
      targetUrlsShouldStartWith: ["https://test.asse.devtunnels.ms"],
      applicableToApps: ApiSecretRegistrationAppType.AnyApp,
      targetAudience: ApiSecretRegistrationTargetAudience.AnyTenant,
    });
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test2.asse.devtunnels.ms",
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

    const confirmStub = sinon
      .stub(mockedDriverContext.ui, "confirm")
      .resolves(ok({ type: "success", value: true }));

    const args: UpdateApiKeyArgs = {
      name: "test",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "AnyTenant",
      applicableToApps: "AnyApp",
      registrationId: "mockedRegistrationId",
    };

    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
    expect(confirmStub.notCalled).to.be.true;
  });

  it("should throw error when user canel", async () => {
    sinon.stub(teamsDevPortalClient, "getApiKeyRegistrationById").resolves({
      id: "mockedRegistrationId",
      description: "mockedDescription",
      clientSecrets: [],
      targetUrlsShouldStartWith: ["https://test"],
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
        {
          api: "api2",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test2",
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

    sinon
      .stub(mockedDriverContext.ui, "confirm")
      .returns(err(new UserError("source", "userCancelled", "Cancel by user")));

    const args: UpdateApiKeyArgs = {
      name: "test2",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      registrationId: "mockedRegistrationId",
    };

    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("userCancelled");
    }
  });

  it("should throw error if missing name", async () => {
    const args: any = {
      name: "",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      registrationId: "mockedRegistrationId",
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if name is too long", async () => {
    const args: any = {
      name: "a".repeat(129),
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      registrationId: "mockedRegistrationId",
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("ApiKeyNameTooLong");
    }
  });

  it("should throw error if missing registrationId", async () => {
    const args: any = {
      name: "name",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if missing both baseUrl and apiSpecPath", async () => {
    const args: any = {
      name: "name",
      appId: "mockedAppId",
      regirstrationid: "mockedRegistrationId",
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
      expect(result.result.error.message).to.include("baseUrl");
      expect(result.result.error.message).to.include("apiSpecPath");
    }
  });

  it("should throw error if apiSpecPath is not string", async () => {
    const args: any = {
      name: "name",
      appId: "mockedAppId",
      regirstrationid: "mockedRegistrationId",
      apiSpecPath: ["invalidPathType"],
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
      expect(result.result.error.message).to.include("apiSpecPath");
    }
  });

  it("should throw error if baseURL is not valid https url", async () => {
    const args: any = {
      name: "name",
      appId: "mockedAppId",
      regirstrationid: "mockedRegistrationId",
      baseUrl: "http://test",
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
      expect(result.result.error.message).to.include("baseUrl");
    }
  });

  it("should throw error if invalid applicableToApps", async () => {
    const args: any = {
      name: "name",
      appId: "mockedAppId",
      regirstrationid: "mockedRegistrationId",
      apiSpecPath: "mockedPath",
      applicableToApps: "test",
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if invalid targetAudience", async () => {
    const args: any = {
      name: "name",
      appId: "mockedAppId",
      regirstrationid: "mockedRegistrationId",
      apiSpecPath: "mockedPath",
      targetAudience: "test",
    };
    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error when unhandled error", async () => {
    sinon.stub(MockedM365Provider.prototype, "getAccessToken").throws(new Error("unhandled error"));
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
          api: "api2",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test2",
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
    const args: UpdateApiKeyArgs = {
      name: "test2",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      registrationId: "mockedRegistrationId",
    };

    const result = await updateApiKeyDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.source).to.equal("apiKeyUpdate");
    }
  });
});
