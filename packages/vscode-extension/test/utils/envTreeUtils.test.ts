import * as chai from "chai";
import * as sinon from "sinon";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import { Uri } from "vscode";
import { envUtil, metadataUtil, pathUtils } from "@microsoft/teamsfx-core";
import * as envTreeUtils from "../../src/utils/envTreeUtils";
import { ok } from "@microsoft/teamsfx-api";
import * as fileSystemUtils from "../../src/utils/fileSystemUtils";

describe("EnvTreeUtils", () => {
   
  describe("getSubscriptionInfoFromEnv", () => {
    const sandbox = sinon.createSandbox();
    const subscriptionInfo = {
      subscriptionName: "subscriptionName",
      subscriptionId: "subscriptionId",
      tenantId: "tenantId",
    };
    const provisionResult: Record<string, any> = {
      solution: subscriptionInfo,
    };

    afterEach(() => {
      sandbox.restore();
    });

    it("returns subscription info successfully", async () => {
      sandbox.stub(fileSystemUtils, "getProvisionResultJson").resolves(provisionResult);
      const result = await envTreeUtils.getSubscriptionInfoFromEnv("test");
      chai.expect(result).deep.equals(subscriptionInfo);
    });

    it("returns undefined if get provision result throws error", async () => {
      sandbox.stub(fileSystemUtils, "getProvisionResultJson").rejects(new Error());
      const result = await envTreeUtils.getSubscriptionInfoFromEnv("test");
      chai.expect(result).is.undefined;
    });

    it("returns undefined if get provision result is undefined", async () => {
      sandbox.stub(fileSystemUtils, "getProvisionResultJson").resolves(undefined);
      const result = await envTreeUtils.getSubscriptionInfoFromEnv("test");
      chai.expect(result).is.undefined;
    });

    it("returns undefined if get provision result does not contain subscriptionId", async () => {
      sandbox.stub(fileSystemUtils, "getProvisionResultJson").resolves({ solution: {} } as any);
      const result = await envTreeUtils.getSubscriptionInfoFromEnv("test");
      chai.expect(result).is.undefined;
    });
  });

  describe("getM365TenantFromEnv", () => {
    const sandbox = sinon.createSandbox();
    const m365TenantId = {
      teamsAppTenantId: "fakeTenantId",
    };

    beforeEach(() => {
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/test" });
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("returns m365 tenantId successfully", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFileSync").returns("TEAMS_APP_TENANT_ID=fakeTenantId\n");
      const result = await envTreeUtils.getM365TenantFromEnv("test");
      chai.expect(result).equal("fakeTenantId");
    });

    it("returns undefined if env file doesn't exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const result = await envTreeUtils.getM365TenantFromEnv("test");
      chai.expect(result).equal(undefined);
    });

    it("returns undefined if tenant id doesn't exist in env file", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFileSync").returns("");
      const result = await envTreeUtils.getM365TenantFromEnv("test");
      chai.expect(result).equal(undefined);
    });
  });

  describe("getResourceGroupNameFromEnv", () => {
    const sandbox = sinon.createSandbox();
    const resourceGroupName = {
      resourceGroupName: "fakeResourceGroupName",
    };
    const provisionResult: Record<string, any> = {
      solution: resourceGroupName,
    };

    afterEach(() => {
      sandbox.restore();
    });

    it("returns resource group name successfully", async () => {
      sandbox.stub(fileSystemUtils, "getProvisionResultJson").resolves(provisionResult);
      const result = await envTreeUtils.getResourceGroupNameFromEnv("test");
      chai.expect(result).equal("fakeResourceGroupName");
    });

    it("returns undefined if get provision result throws error", async () => {
      sandbox.stub(fileSystemUtils, "getProvisionResultJson").rejects(new Error());
      const result = await envTreeUtils.getResourceGroupNameFromEnv("test");
      chai.expect(result).is.undefined;
    });

    it("returns undefined if get provision result returns undefined", async () => {
      sandbox.stub(fileSystemUtils, "getProvisionResultJson").resolves(undefined);
      const result = await envTreeUtils.getResourceGroupNameFromEnv("test");
      chai.expect(result).is.undefined;
    });

    it("returns undefined if get provision result does not contain solution", async () => {
      sandbox.stub(fileSystemUtils, "getProvisionResultJson").resolves({});
      const result = await envTreeUtils.getResourceGroupNameFromEnv("test");
      chai.expect(result).is.undefined;
    });
  });

  describe("getProvisionSucceedFromEnv", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("returns false if teamsAppId is empty", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(Uri.file("test"));
      sandbox.stub(envUtil, "readEnv").resolves(
        ok({
          TEAMS_APP_ID: "",
        })
      );

      const result = await envTreeUtils.getProvisionSucceedFromEnv("test");

      chai.expect(result).equals(false);
    });

    it("returns true if teamsAppId is not empty", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(Uri.file("test"));
      sandbox.stub(envUtil, "readEnv").resolves(
        ok({
          TEAMS_APP_ID: "xxx",
        })
      );
      sandbox.stub(globalVariables, "workspaceUri").value(Uri.file("test"));
      sandbox.stub(pathUtils, "getYmlFilePath");
      sandbox.stub(metadataUtil, "parse").resolves(ok({} as any));

      const result = await envTreeUtils.getProvisionSucceedFromEnv("test");

      chai.expect(result).equals(true);
    });

    it("returns false if teamsAppId has error", async () => {
      sandbox.stub(globalVariables, "workspaceUri").value(Uri.file("test"));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));

      const result = await envTreeUtils.getProvisionSucceedFromEnv("test");

      chai.expect(result).equals(false);
    });
  });
});
