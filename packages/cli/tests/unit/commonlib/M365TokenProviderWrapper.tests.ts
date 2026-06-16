// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import mockedEnv, { RestoreFn } from "mocked-env";
import sinon from "sinon";
import { expect } from "../utils";
import { ok, err, FxError, UserError } from "@microsoft/teamsfx-api";
import ui from "../../../src/userInteraction";
import { M365Login } from "../../../src/commonlib/m365Login";
import M365TokenProviderUserPassword from "../../../src/commonlib/m365LoginUserPassword";
import M365TokenProviderWrapper from "../../../src/commonlib/M365TokenProviderWrapper";

describe("M365TokenProviderWrapper Tests", function () {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn = () => {};

  afterEach(() => {
    sandbox.restore();
    mockedEnvRestore();
  });

  describe("getProvider", () => {
    it("should return M365Login when interactive is true", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockM365Login = {
        getAccessToken: sandbox.stub().resolves(ok("token")),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const provider = M365TokenProviderWrapper.getProvider();

      expect(provider).to.equal(mockM365Login);
    });

    it("should return M365Login when interactive is false but env vars not set", async () => {
      sandbox.stub(ui, "interactive").value(false);
      // Explicitly ensure env vars are not set
      mockedEnvRestore = mockedEnv({
        M365_ACCOUNT_NAME: undefined,
        M365_ACCOUNT_PASSWORD: undefined,
      });
      const mockM365Login = {
        getAccessToken: sandbox.stub().resolves(ok("token")),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const provider = M365TokenProviderWrapper.getProvider();

      expect(provider).to.equal(mockM365Login);
    });

    it("should return M365TokenProviderUserPassword when interactive is false and env vars are set", async () => {
      sandbox.stub(ui, "interactive").value(false);
      mockedEnvRestore = mockedEnv({
        M365_ACCOUNT_NAME: "test@test.com",
        M365_ACCOUNT_PASSWORD: "password",
      });

      const provider = M365TokenProviderWrapper.getProvider();

      expect(provider).to.equal(M365TokenProviderUserPassword);
    });
  });

  describe("getAccessToken", () => {
    it("should delegate to the provider's getAccessToken", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockToken = "test-token";
      const mockM365Login = {
        getAccessToken: sandbox.stub().resolves(ok(mockToken)),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const result = await M365TokenProviderWrapper.getAccessToken({ scopes: ["scope1"] });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal(mockToken);
      expect(mockM365Login.getAccessToken.calledOnce).to.be.true;
    });
  });

  describe("getJsonObject", () => {
    it("should delegate to the provider's getJsonObject", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockJson = { name: "test" };
      const mockM365Login = {
        getJsonObject: sandbox.stub().resolves(ok(mockJson)),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const result = await M365TokenProviderWrapper.getJsonObject(
        { scopes: ["scope1"] },
        "tenantId"
      );

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.deep.equal(mockJson);
      expect(mockM365Login.getJsonObject.calledOnce).to.be.true;
    });
  });

  describe("getStatus", () => {
    it("should delegate to the provider's getStatus", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockStatus = { status: "signedIn", accountInfo: { upn: "test@test.com" } };
      const mockM365Login = {
        getStatus: sandbox.stub().resolves(ok(mockStatus)),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const result = await M365TokenProviderWrapper.getStatus({ scopes: ["scope1"] });

      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.deep.equal(mockStatus);
      expect(mockM365Login.getStatus.calledOnce).to.be.true;
    });
  });

  describe("setStatusChangeMap", () => {
    it("should delegate to the provider's setStatusChangeMap", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockM365Login = {
        setStatusChangeMap: sandbox.stub().resolves(ok(true)),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const statusChange = async () => {};
      const result = await M365TokenProviderWrapper.setStatusChangeMap(
        "test",
        { scopes: ["scope1"] },
        statusChange,
        true
      );

      expect(result.isOk()).to.be.true;
      expect(mockM365Login.setStatusChangeMap.calledOnce).to.be.true;
    });
  });

  describe("removeStatusChangeMap", () => {
    it("should delegate to the provider's removeStatusChangeMap", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockM365Login = {
        removeStatusChangeMap: sandbox.stub().resolves(ok(true)),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const result = await M365TokenProviderWrapper.removeStatusChangeMap("test");

      expect(result.isOk()).to.be.true;
      expect(mockM365Login.removeStatusChangeMap.calledOnce).to.be.true;
    });
  });

  describe("signout", () => {
    it("should delegate to the provider's signout", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockM365Login = {
        signout: sandbox.stub().resolves(true),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const result = await M365TokenProviderWrapper.signout();

      expect(result).to.be.true;
      expect(mockM365Login.signout.calledOnce).to.be.true;
    });
  });

  describe("switchTenant", () => {
    it("should delegate to the provider's switchTenant", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockM365Login = {
        switchTenant: sandbox.stub().resolves(ok("newTenantId")),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const result = await M365TokenProviderWrapper.switchTenant("newTenantId");

      expect(result.isOk()).to.be.true;
      expect(mockM365Login.switchTenant.calledOnce).to.be.true;
    });
  });

  describe("getTenant", () => {
    it("should delegate to the provider's getTenant", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockTenantId = "test-tenant-id";
      const mockM365Login = {
        getTenant: sandbox.stub().resolves(mockTenantId),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const result = await M365TokenProviderWrapper.getTenant();

      expect(result).to.equal(mockTenantId);
      expect(mockM365Login.getTenant.calledOnce).to.be.true;
    });

    it("should return undefined when provider's getTenant returns undefined", async () => {
      sandbox.stub(ui, "interactive").value(true);
      const mockM365Login = {
        getTenant: sandbox.stub().resolves(undefined),
      };
      sandbox.stub(M365Login, "getInstance").returns(mockM365Login as any);

      const result = await M365TokenProviderWrapper.getTenant();

      expect(result).to.be.undefined;
    });
  });
});
