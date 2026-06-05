// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Siglud <siglud@gmail.com>
 */
import * as appService from "@azure/arm-appservice";
import { expect } from "chai";
import * as sinon from "sinon";
import {
  AzureStaticWebAppGetDeploymentTokenDriver,
  azureStaticWebAppGetTokenDeps,
} from "../../../../../src/component/driver/deploy/azure/azureStaticWebAppGetDeploymentTokenDriver";

describe("AzureStaticWebAppGetDeploymentTokenDriver", () => {
  let driver: AzureStaticWebAppGetDeploymentTokenDriver;
  let clientStub: sinon.SinonStubbedInstance<appService.WebSiteManagementClient>;

  beforeEach(() => {
    driver = new AzureStaticWebAppGetDeploymentTokenDriver();
    clientStub = sinon.createStubInstance(appService.WebSiteManagementClient);
    sinon.stub(azureStaticWebAppGetTokenDeps, "createWebSiteManagementClient").returns(clientStub);
    clientStub.staticSites = {
      listStaticSiteSecrets: () => {},
    } as any;
    sinon.stub(azureStaticWebAppGetTokenDeps, "getAzureAccountCredential").resolves({} as any);
  });

  afterEach(() => {
    sinon.restore();
  });

  it("should get deployment token", async () => {
    const secrets = { properties: { apiKey: "testKey" } };
    sinon.stub(clientStub.staticSites, "listStaticSiteSecrets").resolves(secrets);

    const result = await driver.execute(
      {
        resourceId:
          "/subscriptions/aaa-bbb-ccc/resourceGroups/fff-rg/providers/Microsoft.Web/staticSites/aaabbbbccc",
      },
      { azureAccountProvider: {} } as any,
      new Map([["deploymentToken", "SECRET_TAB_SWA_DEPLOYMENT_TOKEN"]])
    );

    expect(result.result.isOk()).to.be.true;
    expect(result.result.unwrapOr(new Map()).get("SECRET_TAB_SWA_DEPLOYMENT_TOKEN")).to.equal(
      "testKey"
    );
  });

  it("azureStaticWebAppGetTokenDeps should create management client", async () => {
    (azureStaticWebAppGetTokenDeps.createWebSiteManagementClient as sinon.SinonStub).restore();
    const client = azureStaticWebAppGetTokenDeps.createWebSiteManagementClient({} as any, "sub-id");
    expect(client).to.be.instanceOf(appService.WebSiteManagementClient);
  });

  it("should get deployment token use default settings", async () => {
    const secrets = { properties: { apiKey: "testKey" } };
    sinon.stub(clientStub.staticSites, "listStaticSiteSecrets").resolves(secrets);

    const result = await driver.execute(
      {
        resourceId:
          "/subscriptions/aaa-bbb-ccc/resourceGroups/fff-rg/providers/Microsoft.Web/staticSites/aaabbbbcccdd",
      },
      { azureAccountProvider: {} } as any,
      new Map([["deploymentToken", ""]])
    );

    expect(result.result.isOk()).to.be.true;
    expect(result.result.unwrapOr(new Map()).get("SECRET_TAB_SWA_DEPLOYMENT_TOKEN")).to.equal(
      undefined
    );
  });

  it("should handle error when getting deployment token", async () => {
    sinon.stub(clientStub.staticSites, "listStaticSiteSecrets").throws(new Error("test error"));

    const result = await driver.execute(
      { resourceId: "testResourceId" },
      { azureAccountProvider: {} } as any,
      new Map([["deploymentToken", "SECRET_TAB_SWA_DEPLOYMENT_TOKEN"]])
    );

    expect(result.result.isErr()).to.be.true;
  });

  it("should handle error when no output env var name", async () => {
    const result = await driver.execute({ resourceId: "testResourceId" }, {
      azureAccountProvider: {},
    } as any);

    expect(result.result.isErr()).to.be.true;
  });

  it("should handle error when no deployment token output env var name", async () => {
    const result = await driver.execute(
      { resourceId: "testResourceId" },
      { azureAccountProvider: {} } as any,
      new Map([["test", "test"]])
    );

    expect(result.result.isErr()).to.be.true;
  });
});
