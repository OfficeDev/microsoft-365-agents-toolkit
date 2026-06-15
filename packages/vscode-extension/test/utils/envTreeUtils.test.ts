import * as chai from "chai";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import { Uri } from "vscode";
import { envUtil, metadataUtil, pathUtils } from "@microsoft/teamsfx-core";
import * as envTreeUtils from "../../src/utils/envTreeUtils";
import { envTreeUtilsDeps } from "../../src/utils/envTreeUtils";
import { ok } from "@microsoft/teamsfx-api";
import * as fileSystemUtils from "../../src/utils/fileSystemUtils";
import * as appDefinitionUtils from "../../src/utils/appDefinitionUtils";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";

describe("EnvTreeUtils", () => {
  // eslint-disable-next-line no-secrets/no-secrets
  describe("getSubscriptionInfoFromEnv", () => {
    const subscriptionInfo = {
      subscriptionName: "subscriptionName",
      subscriptionId: "subscriptionId",
      tenantId: "tenantId",
    };
    const provisionResult: Record<string, any> = {
      solution: subscriptionInfo,
    };

    it("returns subscription info successfully", async () => {
      vi.spyOn(envTreeUtilsDeps, "getProvisionResultJson").mockResolvedValue(provisionResult);
      const result = await envTreeUtils.getSubscriptionInfoFromEnv("test");
      chai.expect(result).deep.equals(subscriptionInfo);
    });

    it("returns undefined if get provision result throws error", async () => {
      vi.spyOn(envTreeUtilsDeps, "getProvisionResultJson").mockRejectedValue(new Error());
      const result = await envTreeUtils.getSubscriptionInfoFromEnv("test");
      chai.expect(result).is.undefined;
    });

    it("returns undefined if get provision result is undefined", async () => {
      vi.spyOn(envTreeUtilsDeps, "getProvisionResultJson").mockResolvedValue(undefined);
      const result = await envTreeUtils.getSubscriptionInfoFromEnv("test");
      chai.expect(result).is.undefined;
    });

    it("returns undefined if get provision result does not contain subscriptionId", async () => {
      vi.spyOn(envTreeUtilsDeps, "getProvisionResultJson").mockResolvedValue({
        solution: {},
      } as any);
      const result = await envTreeUtils.getSubscriptionInfoFromEnv("test");
      chai.expect(result).is.undefined;
    });
  });

  describe("getM365TenantFromEnv", () => {
    const m365TenantId = {
      teamsAppTenantId: "fakeTenantId",
    };

    beforeEach(() => {
      mockValue(globalVariables, "workspaceUri", { fsPath: "/test" });
      vi.spyOn(envTreeUtilsDeps, "getWorkspacePath").mockReturnValue("/test");
    });

    it("returns m365 tenantId successfully", async () => {
      vi.spyOn(envTreeUtilsDeps, "pathExists").mockResolvedValue(true);
      vi.spyOn(envTreeUtilsDeps, "readFileSync").mockReturnValue(
        "TEAMS_APP_TENANT_ID=fakeTenantId\n"
      );
      const result = await envTreeUtils.getM365TenantFromEnv("test");
      chai.expect(result).equal("fakeTenantId");
    });

    it("returns undefined if env file doesn't exist", async () => {
      vi.spyOn(envTreeUtilsDeps, "pathExists").mockResolvedValue(false);
      const result = await envTreeUtils.getM365TenantFromEnv("test");
      chai.expect(result).equal(undefined);
    });

    it("returns undefined if tenant id doesn't exist in env file", async () => {
      vi.spyOn(envTreeUtilsDeps, "pathExists").mockResolvedValue(true);
      vi.spyOn(envTreeUtilsDeps, "readFileSync").mockReturnValue("");
      const result = await envTreeUtils.getM365TenantFromEnv("test");
      chai.expect(result).equal(undefined);
    });
  });

  describe("getResourceGroupNameFromEnv", () => {
    const resourceGroupName = {
      resourceGroupName: "fakeResourceGroupName",
    };
    const provisionResult: Record<string, any> = {
      solution: resourceGroupName,
    };

    it("returns resource group name successfully", async () => {
      vi.spyOn(envTreeUtilsDeps, "getProvisionResultJson").mockResolvedValue(provisionResult);
      const result = await envTreeUtils.getResourceGroupNameFromEnv("test");
      chai.expect(result).equal("fakeResourceGroupName");
    });

    it("returns undefined if get provision result throws error", async () => {
      vi.spyOn(envTreeUtilsDeps, "getProvisionResultJson").mockRejectedValue(new Error());
      const result = await envTreeUtils.getResourceGroupNameFromEnv("test");
      chai.expect(result).is.undefined;
    });

    it("returns undefined if get provision result returns undefined", async () => {
      vi.spyOn(envTreeUtilsDeps, "getProvisionResultJson").mockResolvedValue(undefined);
      const result = await envTreeUtils.getResourceGroupNameFromEnv("test");
      chai.expect(result).is.undefined;
    });

    it("returns undefined if get provision result does not contain solution", async () => {
      vi.spyOn(envTreeUtilsDeps, "getProvisionResultJson").mockResolvedValue({});
      const result = await envTreeUtils.getResourceGroupNameFromEnv("test");
      chai.expect(result).is.undefined;
    });
  });

  describe("getProvisionSucceedFromEnv", () => {
    it("returns false if teamsAppId is empty", async () => {
      mockValue(globalVariables, "workspaceUri", Uri.file("test"));
      vi.spyOn(envTreeUtilsDeps, "getWorkspacePath").mockReturnValue("test");
      vi.spyOn(envTreeUtilsDeps, "getV3TeamsAppId").mockResolvedValue("");

      const result = await envTreeUtils.getProvisionSucceedFromEnv("test");

      chai.expect(result).equals(false);
    });

    it("returns true if teamsAppId is not empty", async () => {
      mockValue(globalVariables, "workspaceUri", Uri.file("test"));
      vi.spyOn(envTreeUtilsDeps, "getWorkspacePath").mockReturnValue("test");
      vi.spyOn(envTreeUtilsDeps, "getV3TeamsAppId").mockResolvedValue("xxx");

      const result = await envTreeUtils.getProvisionSucceedFromEnv("test");

      chai.expect(result).equals(true);
    });

    it("returns false if teamsAppId has error", async () => {
      mockValue(globalVariables, "workspaceUri", Uri.file("test"));
      vi.spyOn(envTreeUtilsDeps, "getWorkspacePath").mockReturnValue("test");
      vi.spyOn(envTreeUtilsDeps, "getV3TeamsAppId").mockRejectedValue(new Error("test"));

      const result = await envTreeUtils.getProvisionSucceedFromEnv("test");

      chai.expect(result).equals(false);
    });
  });

  describe("envTreeUtilsDeps delegation", () => {
    it("delegates getWorkspacePath", () => {
      mockValue(globalVariables, "workspaceUri", Uri.file("C:\\test"));
      const workspacePath = envTreeUtilsDeps.getWorkspacePath();
      chai.expect(workspacePath?.toLowerCase()).equals("c:\\test");
    });

    it("delegates pathExists and readFileSync", async () => {
      vi.spyOn(fs, "pathExists").mockResolvedValue(true as never);
      vi.spyOn(fs, "readFileSync").mockReturnValue("A=B" as never);

      const exists = await envTreeUtilsDeps.pathExists("C:\\test\\.env.dev");
      const content = envTreeUtilsDeps.readFileSync("C:\\test\\.env.dev", "utf8");

      chai.expect(exists).to.be.true;
      chai.expect(content).equals("A=B");
    });

    it("delegates getV3TeamsAppId", async () => {
      vi.spyOn(appDefinitionUtils, "getV3TeamsAppId").mockResolvedValue("app-id");
      const appId = await envTreeUtilsDeps.getV3TeamsAppId("C:\\test", "dev");
      chai.expect(appId).equals("app-id");
    });
  });
});
