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
import { UpdateOauthArgs } from "../../../../src/component/driver/oauth/interface/updateOauthArgs";
import { UpdateOauthDriver } from "../../../../src/component/driver/oauth/update";
import {
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
} from "../../../../src/component/driver/teamsApp/interfaces/OauthRegistration";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";
import * as utiltiy from "../../../../src/component/driver/oauth/utility/utility";

chai.use(chaiAsPromised);
const expect = chai.expect;

describe("UpdateOauthDriver", () => {
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    ui: new MockedUserInteraction(),
  };
  const updateOauthDriver = new UpdateOauthDriver();

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
    sinon.stub(teamsDevPortalClient, "updateOauthRegistration").resolves({
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test2"],
      applicableToApps: OauthRegistrationAppType.SpecificApp,
      targetAudience: OauthRegistrationTargetAudience.HomeTenant,
      m365AppId: "mockedAppId",
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      scopes: ["mockedScope"],
      isPKCEEnabled: true,
    });
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves({
      oAuthConfigId: "mockedRegistrationId",
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test"],
      applicableToApps: OauthRegistrationAppType.AnyApp,
      targetAudience: OauthRegistrationTargetAudience.AnyTenant,
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      scopes: ["mockedScope"],
      isPKCEEnabled: false,
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  refreshUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
      expect((config as ConfirmConfig).title.includes("m365AppId")).to.be.true;
      expect((config as ConfirmConfig).title.includes("targetAudience")).to.be.true;
      expect((config as ConfirmConfig).title.includes("isPKCEEnabled")).to.be.true;
      expect((config as ConfirmConfig).title.includes("authorizationEndpoint")).to.be.true;
      expect((config as ConfirmConfig).title.includes("tokenExchangeEndpoint")).to.be.true;
      expect((config as ConfirmConfig).title.includes("tokenRefreshEndpoint")).to.be.true;
      expect((config as ConfirmConfig).title.includes("scopes")).to.be.true;
      return ok({ type: "success", value: true });
    });

    const args: UpdateOauthArgs = {
      name: "test2",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      configurationId: "mockedRegistrationId",
      isPKCEEnabled: true,
    };

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: update all fields for Entra SSO", async () => {
    sinon.stub(teamsDevPortalClient, "updateOauthRegistration").resolves({
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test2"],
      applicableToApps: OauthRegistrationAppType.SpecificApp,
      targetAudience: OauthRegistrationTargetAudience.HomeTenant,
      m365AppId: "mockedAppId",
      clientId: "mockedClientId2",
      identityProvider: "MicrosoftEntra",
      isPKCEEnabled: false,
    } as any);
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves({
      oAuthConfigId: "mockedRegistrationId",
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test"],
      applicableToApps: OauthRegistrationAppType.AnyApp,
      targetAudience: OauthRegistrationTargetAudience.AnyTenant,
      clientId: "mockedClientId",
      identityProvider: "MicrosoftEntra",
      isPKCEEnabled: false,
    } as any);
    sinon.stub(SpecParser.prototype, "list").resolves({
      APIs: [
        {
          api: "api",
          server: "https://test",
          operationId: "get",
          auth: {
            name: "test",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  refreshUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
      expect((config as ConfirmConfig).title.includes("m365AppId")).to.be.true;
      expect((config as ConfirmConfig).title.includes("targetAudience")).to.be.true;

      expect((config as ConfirmConfig).title.includes("clientId")).to.be.true;
      return ok({ type: "success", value: true });
    });

    const args: UpdateOauthArgs = {
      name: "test2",
      appId: "mockedAppId",
      clientId: "mockedClientId2",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      configurationId: "mockedRegistrationId",
    };

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("should throw error if try to disable PKCE", async () => {
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves({
      oAuthConfigId: "mockedRegistrationId",
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test"],
      applicableToApps: OauthRegistrationAppType.AnyApp,
      targetAudience: OauthRegistrationTargetAudience.AnyTenant,
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      scopes: ["mockedScope"],
      isPKCEEnabled: true,
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const args: UpdateOauthArgs = {
      name: "test2",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      configurationId: "mockedRegistrationId",
      isPKCEEnabled: false,
    };

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("OauthDisablePKCEError");
    }
  });

  it("happy path: does not update when no changes", async () => {
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves({
      oAuthConfigId: "mockedRegistrationId",
      description: "test",
      targetUrlsShouldStartWith: ["https://test"],
      applicableToApps: OauthRegistrationAppType.AnyApp,
      targetAudience: OauthRegistrationTargetAudience.AnyTenant,
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      scopes: ["mockedScopes"],
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationEndpoint",
                  tokenUrl: "mockedTokenExchangeEndpoint",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScope",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });

    const args: UpdateOauthArgs = {
      name: "test",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "AnyTenant",
      applicableToApps: "AnyApp",
      configurationId: "mockedRegistrationId",
    };

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
  });

  it("happy path: should not show confirm when only devtunnel url is different", async () => {
    sinon.stub(teamsDevPortalClient, "updateOauthRegistration").resolves({
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test2.asse.devtunnels.ms"],
      applicableToApps: OauthRegistrationAppType.SpecificApp,
      targetAudience: OauthRegistrationTargetAudience.HomeTenant,
      m365AppId: "mockedAppId",
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      scopes: ["mockedScope"],
    });
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves({
      oAuthConfigId: "mockedRegistrationId",
      description: "test",
      targetUrlsShouldStartWith: ["https://test.asse.devtunnels.ms"],
      applicableToApps: OauthRegistrationAppType.AnyApp,
      targetAudience: OauthRegistrationTargetAudience.AnyTenant,
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      scopes: ["mockedScopes"],
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "mockedAuthorizationEndpoint",
                  tokenUrl: "mockedTokenExchangeEndpoint",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
        {
          api: "api2",
          server: "https://test2.asse.devtunnels.ms",
          operationId: "get",
          auth: {
            name: "test2",
            authScheme: {
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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

    const args: UpdateOauthArgs = {
      name: "test",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "AnyTenant",
      applicableToApps: "AnyApp",
      configurationId: "mockedRegistrationId",
    };

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
    expect(confirmStub.notCalled).to.be.true;
  });

  it("should throw error when user canel", async () => {
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves({
      oAuthConfigId: "mockedRegistrationId",
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test"],
      applicableToApps: OauthRegistrationAppType.AnyApp,
      targetAudience: OauthRegistrationTargetAudience.AnyTenant,
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      scopes: ["mockedScope"],
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
    const args: UpdateOauthArgs = {
      name: "test2",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      configurationId: "mockedRegistrationId",
    };

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
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
      configurationId: "mockedRegistrationId",
    };
    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if isPKCEEnabled is not boolean", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      clientId: "mockedClientId",
      flow: "authorizationCode",
      refreshUrl: "mockedRefreshUrl",
      isPKCEEnabled: "invalid",
      identityProvider: "Custom",
      configurationId: "mockedRegistrationId",
    };
    sinon.stub(utiltiy, "getandValidateOauthInfoFromSpec").resolves({} as any);
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves(
      ok({
        identityProvider: "Custom",
      }) as any
    );

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
      expect(result.result.error.message).to.include("isPKCEEnabled");
    }
  });

  it("should throw error if secret is not string", async () => {
    const args: any = {
      name: "test",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      clientId: "mockedClientId",
      flow: "authorizationCode",
      refreshUrl: "mockedRefreshUrl",
      isPKCEEnabled: false,
      identityProvider: "Custom",
      clientSecret: 123,
      configurationId: "mockedRegistrationId",
    };
    sinon.stub(utiltiy, "getandValidateOauthInfoFromSpec").resolves({} as any);
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves(
      ok({
        identityProvider: "Custom",
      }) as any
    );
    sinon.stub(MockedM365Provider.prototype, "getAccessToken").resolves(ok({}) as any);

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
      expect(result.result.error.message).to.include("clientSecret");
    }
  });

  it("should throw error if name is too long", async () => {
    const args: any = {
      name: "a".repeat(129),
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      configurationId: "mockedRegistrationId",
    };
    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("OauthNameTooLong");
    }
  });

  it("should throw error if missing appId", async () => {
    const args: any = {
      name: "",
      apiSpecPath: "mockedPath",
      configurationId: "mockedRegistrationId",
    };
    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if missing apiSpecPath", async () => {
    const args: any = {
      name: "",
      appId: "mockedAppId",
      configurationId: "mockedRegistrationId",
    };
    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if missing registrationId", async () => {
    const args: any = {
      name: "",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
    };
    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if invalid applicableToApps", async () => {
    const args: any = {
      name: "name",
      appId: "mockedAppId",
      configurationId: "mockedRegistrationId",
      apiSpecPath: "mockedPath",
      applicableToApps: "test",
    };
    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.equal("InvalidActionInputError");
    }
  });

  it("should throw error if invalid targetAudience", async () => {
    const args: any = {
      name: "name",
      appId: "mockedAppId",
      configurationId: "mockedRegistrationId",
      apiSpecPath: "mockedPath",
      targetAudience: "test",
    };
    const result = await updateOauthDriver.execute(args, mockedDriverContext);
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
            },
          },
          isValid: true,
          reason: [],
        },
      ],
      allAPICount: 1,
      validAPICount: 1,
    });
    const args: UpdateOauthArgs = {
      name: "test2",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      configurationId: "mockedRegistrationId",
    };

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.source).to.equal("oauthUpdate");
    }
  });

  it("should update if tokenRefreshEndpoint and scopes are undefined", async () => {
    sinon.stub(teamsDevPortalClient, "updateOauthRegistration").resolves({
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test2"],
      applicableToApps: OauthRegistrationAppType.SpecificApp,
      targetAudience: OauthRegistrationTargetAudience.HomeTenant,
      m365AppId: "mockedAppId",
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      scopes: ["mockedScope"],
      isPKCEEnabled: true,
    });
    sinon.stub(teamsDevPortalClient, "getOauthRegistrationById").resolves({
      oAuthConfigId: "mockedRegistrationId",
      description: "mockedDescription",
      targetUrlsShouldStartWith: ["https://test"],
      applicableToApps: OauthRegistrationAppType.AnyApp,
      targetAudience: OauthRegistrationTargetAudience.AnyTenant,
      clientId: "mockedClientId",
      clientSecret: "mockedClientSecret",
      authorizationEndpoint: "mockedAuthorizationEndpoint",
      tokenExchangeEndpoint: "mockedTokenExchangeEndpoint",
      tokenRefreshEndpoint: "mockedTokenRefreshEndpoint",
      scopes: ["mockedScope"],
      isPKCEEnabled: false,
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {
                    mockedScopes: "mockedScopes",
                  },
                },
              },
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
              type: "oauth2",
              flows: {
                authorizationCode: {
                  authorizationUrl: "https://test",
                  tokenUrl: "https://test",
                  scopes: {},
                },
              },
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
      expect((config as ConfirmConfig).title.includes("m365AppId")).to.be.true;
      expect((config as ConfirmConfig).title.includes("targetAudience")).to.be.true;
      expect((config as ConfirmConfig).title.includes("isPKCEEnabled")).to.be.true;
      expect((config as ConfirmConfig).title.includes("authorizationEndpoint")).to.be.true;
      expect((config as ConfirmConfig).title.includes("tokenExchangeEndpoint")).to.be.true;
      expect((config as ConfirmConfig).title.includes("tokenRefreshEndpoint")).to.be.true;
      expect((config as ConfirmConfig).title.includes("scopes")).to.be.true;
      return ok({ type: "success", value: true });
    });

    const args: UpdateOauthArgs = {
      name: "test2",
      appId: "mockedAppId",
      apiSpecPath: "mockedPath",
      targetAudience: "HomeTenant",
      applicableToApps: "SpecificApp",
      configurationId: "mockedRegistrationId",
      isPKCEEnabled: true,
    };

    const result = await updateOauthDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(0);
      expect(result.summaries.length).to.equal(1);
    }
  });
});
