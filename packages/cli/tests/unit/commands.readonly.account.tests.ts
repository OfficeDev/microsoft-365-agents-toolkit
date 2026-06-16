// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLIContext, err, ok, signedIn, signedOut } from "@microsoft/teamsfx-api";
import { UserCancelError } from "@microsoft/teamsfx-core";
import * as tools from "@microsoft/teamsfx-core/build/common/tools";
import { assert } from "chai";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { accountLogoutCommand, accountShowCommand, accountUtils } from "../../src/commands/models";
import AzureTokenProvider from "../../src/commonlib/azureLogin";
import AzureTokenCIProvider from "../../src/commonlib/azureLoginCI";
import { AzureSpCrypto } from "../../src/commonlib/cacheAccess";
import { logger } from "../../src/commonlib/logger";
import M365TokenProvider from "../../src/commonlib/M365TokenProviderWrapper";

describe("CLI read-only commands account", () => {
  const sandbox = sinon.createSandbox();
  let messages: string[] = [];

  beforeEach(() => {
    sandbox.stub(process.stdout, "write").returns(true as any);
    sandbox.stub(process.stderr, "write").returns(true as any);
    sandbox.stub(logger, "info").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
    sandbox.stub(logger, "error").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
    sandbox.stub(logger, "outputInfo").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
    sandbox.stub(logger, "outputError").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
  });

  afterEach(() => {
    messages = [];
    sandbox.restore();
  });

  describe("AccountUtils", async () => {
    it("outputAccountInfoOffline", async () => {
      const res = accountUtils.outputAccountInfoOffline("m365", "xxx");
      assert.isTrue(res);
    });
    it("outputM365Info login success", async () => {
      sandbox.stub(M365TokenProvider, "getJsonObject").resolves(ok({ upn: "fakename" }));
      sandbox.stub(M365TokenProvider, "getTenant").resolves(undefined);
      const res = await accountUtils.outputM365Info("login");
      assert.isTrue(res);
    });
    context("outputM365Info login under hosting tenant", () => {
      let mocks: RestoreFn;
      beforeEach(() => {
        mocks = mockedEnv({ TEAMSFX_MULTI_TENANT: "true" });
        sandbox.stub(M365TokenProvider, "getJsonObject").resolves(ok({ unique_name: "fakename" }));
        sandbox.stub(M365TokenProvider, "getTenant").resolves("faked_tenant_id");
      });

      afterEach(() => {
        mocks();
      });

      it("specified tenant name displayed", async () => {
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
        sandbox
          .stub(tools, "listAllTenants")
          .resolves([
            { tenantId: "faked_tid_1" },
            { tenantId: "faked_tenant_id", displayName: "Test tenant" },
          ]);
        const res = await accountUtils.outputM365Info("login", "faked_tenant_id");
        assert.isTrue(res);
      });

      it("specified tenant not match", async () => {
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
        sandbox
          .stub(tools, "listAllTenants")
          .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tid_2" }]);
        const res = await accountUtils.outputM365Info("login", "faked_tenant_id");
        assert.isTrue(res);
      });

      it("failed to retrieve access token", async () => {
        sandbox
          .stub(M365TokenProvider, "getAccessToken")
          .resolves(err("failed to get access token" as any));
        const res = await accountUtils.outputM365Info("login", "faked_tenant_id");
        assert.isTrue(res);
      });
    });
    it("outputM365Info login fail", async () => {
      sandbox.stub(M365TokenProvider, "getJsonObject").resolves(err(new UserCancelError()));
      const res = await accountUtils.outputM365Info("login");
      assert.isFalse(res);
    });
    it("outputM365Info show success", async () => {
      sandbox.stub(M365TokenProvider, "getJsonObject").resolves(ok({ upn: "fakename" }));
      sandbox.stub(M365TokenProvider, "getTenant").resolves("faked_tenant_id");
      sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
      sandbox
        .stub(tools, "listAllTenants")
        .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tenant_id" }]);
      const res = await accountUtils.outputM365Info("show");
      assert.isTrue(res);
    });
    it("outputM365Info show fail", async () => {
      sandbox.stub(M365TokenProvider, "getJsonObject").resolves(err(new UserCancelError()));
      const res = await accountUtils.outputM365Info("show");
      assert.isFalse(res);
    });
    it("outputAzureInfo login", async () => {
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ upn: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      const res = await accountUtils.outputAzureInfo("login", undefined, true);
      assert.isTrue(res);
    });
    it("outputAzureInfo login with tenant parameter", async () => {
      const mockedEnvRestore = mockedEnv({ TEAMSFX_MULTI_TENANT: "true" });
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "switchTenant").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ unique_name: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      sandbox.stub(AzureTokenCIProvider, "getTenant").resolves("faked_tenant_id");
      sandbox.stub(AzureTokenCIProvider, "getIdentityCredentialAsync").resolves({
        getToken: async () => {
          return Promise.resolve({ token: "faked_token" });
        },
      } as any);
      sandbox
        .stub(tools, "listAllTenants")
        .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tenant_id" }]);
      const res = await accountUtils.outputAzureInfo("login", "faked_tenant_id", true);
      assert.isTrue(res);
      mockedEnvRestore();
    });
    it("outputAzureInfo login fail with tenant parameter - invalid token", async () => {
      const mockedEnvRestore = mockedEnv({ TEAMSFX_MULTI_TENANT: "true" });
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "switchTenant").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ unique_name: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      sandbox.stub(AzureTokenCIProvider, "getTenant").resolves("faked_tenant_id");
      sandbox.stub(AzureTokenCIProvider, "getIdentityCredentialAsync").resolves(undefined);
      const res = await accountUtils.outputAzureInfo("login", "faked_tenant_id", true);
      assert.isTrue(res);
      mockedEnvRestore();
    });
    it("outputAzureInfo login fail with tenant parameter - tenant mismatch", async () => {
      const mockedEnvRestore = mockedEnv({ TEAMSFX_MULTI_TENANT: "true" });
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "switchTenant").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ unique_name: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      sandbox.stub(AzureTokenCIProvider, "getTenant").resolves("faked_tenant_id");
      sandbox.stub(AzureTokenCIProvider, "getIdentityCredentialAsync").resolves({
        getToken: async () => {
          return Promise.resolve({ token: "faked_token" });
        },
      } as any);
      sandbox
        .stub(tools, "listAllTenants")
        .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tid_2" }]);
      const res = await accountUtils.outputAzureInfo("login", "faked_tenant_id", true);
      assert.isTrue(res);
      mockedEnvRestore();
    });
    it("outputAzureInfo login fail", async () => {
      sandbox.stub(AzureTokenProvider, "getJsonObject").resolves(undefined);
      const res = await accountUtils.outputAzureInfo("login");
      assert.isFalse(res);
    });
    it("outputAzureInfo show", async () => {
      sandbox.stub(AzureTokenProvider, "getJsonObject").resolves({ upn: "test" });
      sandbox.stub(AzureTokenProvider, "listSubscriptions").resolves([]);
      const res = await accountUtils.outputAzureInfo("show");
      assert.isTrue(res);
    });
    it("outputAzureInfo show fail", async () => {
      sandbox.stub(AzureTokenProvider, "getJsonObject").resolves(undefined);
      const res = await accountUtils.outputAzureInfo("show");
      assert.isFalse(res);
    });
    it("outputAzureInfo show with sp login", async () => {
      sandbox.stub(AzureSpCrypto, "checkAzureSPFile").resolves(true);
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "switchTenant").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ unique_name: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      sandbox.stub(AzureTokenCIProvider, "getTenant").resolves("faked_tenant_id");
      const getTokenFake = {
        getToken: async (scope: string) => {
          return Promise.resolve({ token: "faked_token" });
        },
      };
      const getTokenSpy = sandbox.spy(getTokenFake, "getToken");
      sandbox
        .stub(AzureTokenCIProvider, "getIdentityCredentialAsync")
        .resolves(getTokenFake as any);
      sandbox
        .stub(tools, "listAllTenants")
        .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tid_2" }]);
      const res = await accountUtils.outputAzureInfo("login", "faked_tenant_id", true);
      assert.isTrue(res);
      assert.isTrue(getTokenSpy.calledOnceWith("https://management.core.windows.net/.default"));
    });
  });
  describe("accountShowCommand", async () => {
    it("both signedOut", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(ok({ status: signedOut }));
      sandbox.stub(AzureTokenProvider, "getStatus").resolves({ status: signedOut });
      messages = [];
      const ctx: CLIContext = {
        command: {
          ...accountShowCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth list`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("both signedIn and checkIsOnline = true", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(ok({ status: signedIn }));
      sandbox.stub(AzureTokenProvider, "getStatus").resolves({ status: signedIn });
      sandbox.stub(accountUtils, "checkIsOnline").resolves(true);
      const outputM365Info = sandbox.stub(accountUtils, "outputM365Info").resolves();
      const outputAzureInfo = sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      messages = [];
      const ctx: CLIContext = {
        command: {
          ...accountShowCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth list`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(outputM365Info.calledOnce);
      assert.isTrue(outputAzureInfo.calledOnce);
    });
    it("both signedIn and checkIsOnline = false", async () => {
      sandbox
        .stub(M365TokenProvider, "getStatus")
        .resolves(ok({ status: signedIn, accountInfo: { upn: "xxx" } }));
      sandbox
        .stub(AzureTokenProvider, "getStatus")
        .resolves({ status: signedIn, accountInfo: { upn: "xxx" } });
      sandbox.stub(accountUtils, "checkIsOnline").resolves(false);
      const outputAccountInfoOffline = sandbox.stub(accountUtils, "outputAccountInfoOffline");
      messages = [];
      const ctx: CLIContext = {
        command: {
          ...accountShowCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth list`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(outputAccountInfoOffline.calledTwice);
    });
    it("M365TokenProvider.getStatus() returns error", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(err(new UserCancelError()));
      messages = [];
      const ctx: CLIContext = {
        command: {
          ...accountShowCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth list`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("accountLogoutCommand", async () => {
    it("azure success", async () => {
      sandbox.stub(AzureTokenProvider, "signout").resolves(true);
      const ctx: CLIContext = {
        command: {
          ...accountLogoutCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth logout`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["azure"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("azure fail", async () => {
      sandbox.stub(AzureTokenProvider, "signout").resolves(false);
      const ctx: CLIContext = {
        command: {
          ...accountLogoutCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth logout`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["azure"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("m365 success", async () => {
      sandbox.stub(M365TokenProvider, "signout").resolves(true);
      const ctx: CLIContext = {
        command: {
          ...accountLogoutCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth logout`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["m365"],
        telemetryProperties: {},
      };
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("m365 fail", async () => {
      sandbox.stub(M365TokenProvider, "signout").resolves(false);
      const ctx: CLIContext = {
        command: {
          ...accountLogoutCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth logout`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["m365"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
});
