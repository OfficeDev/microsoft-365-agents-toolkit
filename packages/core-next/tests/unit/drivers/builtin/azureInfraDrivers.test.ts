/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import axios from "axios";
import * as retryModule from "../../../../src/http/retry";
import { createMockContext } from "../../testHelper";
import { builtinDrivers } from "../../../../src/drivers/builtin";
import { armDeployDriver } from "../../../../src/drivers/builtin/arm/deploy";
import { azureAppServiceZipDeployDriver } from "../../../../src/drivers/builtin/azureAppService/zipDeploy";
import { azureFunctionsZipDeployDriver } from "../../../../src/drivers/builtin/azureFunctions/zipDeploy";
import { AzureArmClient } from "../../../../src/clients/azure/client";
import {
  ARM_BASE_URL,
  azureManagementScopes,
  isTerminalState,
} from "../../../../src/clients/azure/types";

const FAKE_SUB = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const FAKE_RG = "test-rg";
const FAKE_TOKEN = "fake-azure-token";

describe("Azure Infrastructure Drivers", () => {
  let sandbox: sinon.SinonSandbox;
  let mockAxios: {
    put: sinon.SinonStub;
    post: sinon.SinonStub;
    get: sinon.SinonStub;
    defaults: { headers: { common: Record<string, string> } };
    interceptors: {
      request: { use: sinon.SinonStub };
      response: { use: sinon.SinonStub };
    };
  };

  beforeEach(() => {
    sandbox = sinon.createSandbox();
    mockAxios = {
      put: sandbox.stub(),
      post: sandbox.stub(),
      get: sandbox.stub(),
      defaults: { headers: { common: {} } },
      interceptors: {
        request: { use: sandbox.stub() },
        response: { use: sandbox.stub() },
      },
    };
    sandbox.stub(axios, "create").returns(mockAxios as any);
    sandbox.stub(retryModule, "sendWithRetry").callsFake(async (fn: any) => fn());
  });

  afterEach(() => {
    sandbox.restore();
  });

  // ── Registration ──────────────────────────────────────────────────────

  describe("registration", () => {
    it("all 22 builtin drivers are registered", () => {
      expect(builtinDrivers).to.have.lengthOf(22);
    });

    it("arm/deploy is in builtinDrivers", () => {
      const ids = builtinDrivers.map((d) => d.id);
      expect(ids).to.include("arm/deploy");
      expect(ids).to.include("azureAppService/zipDeploy");
      expect(ids).to.include("azureFunctions/zipDeploy");
    });
  });

  // ── Azure ARM Client ──────────────────────────────────────────────────

  describe("AzureArmClient", () => {
    it("sets Authorization header", () => {
      const ctx = createMockContext();
      const _client = new AzureArmClient(ctx, FAKE_TOKEN);
      expect(mockAxios.defaults.headers.common["Authorization"]).to.equal(`Bearer ${FAKE_TOKEN}`);
    });

    it("deployTemplate succeeds on Succeeded state", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.put.resolves({
        status: 200,
        data: {
          properties: {
            provisioningState: "Succeeded",
            outputs: { endpoint: { type: "String", value: "https://example.com" } },
          },
        },
      });

      const result = await client.deployTemplate(FAKE_SUB, FAKE_RG, "test-deploy", {
        properties: { template: {}, parameters: null, mode: "Incremental" },
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.properties?.outputs?.endpoint?.value).to.equal("https://example.com");
      }
    });

    it("deployTemplate returns error on Failed state", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.put.resolves({
        status: 200,
        data: {
          properties: {
            provisioningState: "Failed",
            error: { code: "DeploymentFailed", message: "Something went wrong" },
          },
        },
      });

      const result = await client.deployTemplate(FAKE_SUB, FAKE_RG, "test-deploy", {
        properties: { template: {}, parameters: null, mode: "Incremental" },
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("DeployArmError");
        expect(result.error.message).to.include("Failed");
      }
    });

    it("deployTemplate handles network error", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.put.rejects(new Error("Network error"));

      const result = await client.deployTemplate(FAKE_SUB, FAKE_RG, "test-deploy", {
        properties: { template: {}, parameters: null, mode: "Incremental" },
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.kind).to.equal("system");
      }
    });

    it("getScmEndpoint returns SCM hostname", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.get.resolves({
        status: 200,
        data: {
          properties: {
            enabledHostNames: ["myapp.azurewebsites.net", "myapp.scm.azurewebsites.net"],
          },
        },
      });

      const result = await client.getScmEndpoint(FAKE_SUB, FAKE_RG, "myapp");
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal("https://myapp.scm.azurewebsites.net");
      }
    });

    it("getScmEndpoint returns error when no SCM host", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.get.resolves({
        status: 200,
        data: {
          properties: {
            enabledHostNames: ["myapp.azurewebsites.net"],
          },
        },
      });

      const result = await client.getScmEndpoint(FAKE_SUB, FAKE_RG, "myapp");
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("ScmEndpointNotFound");
      }
    });

    it("restartSite succeeds", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.post.resolves({ status: 200 });

      const result = await client.restartSite(FAKE_SUB, FAKE_RG, "myapp");
      expect(result.isOk()).to.be.true;
    });
  });

  // ── Azure Types ─────────────────────────────────────────────────────

  describe("Azure types", () => {
    it("azureManagementScopes returns correct scope", () => {
      expect(azureManagementScopes()).to.deep.equal([`${ARM_BASE_URL}/.default`]);
    });

    it("isTerminalState identifies terminal states", () => {
      expect(isTerminalState("Succeeded")).to.be.true;
      expect(isTerminalState("Failed")).to.be.true;
      expect(isTerminalState("Canceled")).to.be.true;
      expect(isTerminalState("Running")).to.be.false;
      expect(isTerminalState("Accepted")).to.be.false;
    });
  });

  // ── arm/deploy driver ─────────────────────────────────────────────────

  describe("arm/deploy driver", () => {
    it("has correct metadata", () => {
      expect(armDeployDriver.id).to.equal("arm/deploy");
      expect(armDeployDriver.name).to.equal("Deploy ARM Templates");
    });

    it("validates config schema", () => {
      const result = armDeployDriver.validateFn!({
        subscriptionId: "not-a-uuid",
        resourceGroupName: FAKE_RG,
        templates: [{ path: "main.bicep", deploymentName: "deploy1" }],
      });
      expect(result.isErr()).to.be.true;
    });

    it("validates templates must not be empty", () => {
      const result = armDeployDriver.validateFn!({
        subscriptionId: FAKE_SUB,
        resourceGroupName: FAKE_RG,
        templates: [],
      });
      expect(result.isErr()).to.be.true;
    });

    it("accepts valid config", () => {
      const result = armDeployDriver.validateFn!({
        subscriptionId: FAKE_SUB,
        resourceGroupName: FAKE_RG,
        templates: [{ path: "main.json", deploymentName: "deploy1" }],
      });
      expect(result.isOk()).to.be.true;
    });

    it("requires projectPath", async () => {
      const ctx = createMockContext({ projectPath: undefined });
      const result = await armDeployDriver.executeFn(ctx, {
        subscriptionId: FAKE_SUB,
        resourceGroupName: FAKE_RG,
        templates: [{ path: "main.json", deploymentName: "deploy1" }],
      });
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("MissingProjectPath");
      }
    });

    it("returns error when Azure credential fails", async () => {
      const ctx = createMockContext();
      (ctx.auth.azureAccountProvider as any).getIdentityCredentialAsync = sandbox
        .stub()
        .resolves(undefined);

      const result = await armDeployDriver.executeFn(ctx, {
        subscriptionId: FAKE_SUB,
        resourceGroupName: FAKE_RG,
        templates: [{ path: "main.json", deploymentName: "deploy1" }],
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("AzureCredentialError");
      }
    });

    it("deploys JSON template and returns outputs", async () => {
      const tmpDir = path.join(process.env.TEMP || "/tmp", "arm-test-" + Date.now());
      await fs.mkdir(tmpDir, { recursive: true });
      const templatePath = path.join(tmpDir, "main.json");
      await fs.writeFile(
        templatePath,
        JSON.stringify({
          $schema:
            "https://schema.management.azure.com/schemas/2019-04-01/deploymentTemplate.json#",
          contentVersion: "1.0.0.0",
          resources: [],
          outputs: {},
        })
      );

      try {
        const ctx = createMockContext({ projectPath: tmpDir });
        (ctx.auth.azureAccountProvider as any).getIdentityCredentialAsync = sandbox
          .stub()
          .resolves({ getToken: sandbox.stub().resolves({ token: FAKE_TOKEN }) });

        mockAxios.put.resolves({
          status: 200,
          data: {
            properties: {
              provisioningState: "Succeeded",
              outputs: {
                endpoint: { type: "String", value: "https://example.com" },
              },
            },
          },
        });

        const result = await armDeployDriver.executeFn(ctx, {
          subscriptionId: FAKE_SUB,
          resourceGroupName: FAKE_RG,
          templates: [{ path: "main.json", deploymentName: "deploy1" }],
        });

        expect(result.isOk()).to.be.true;
        if (result.isOk()) {
          expect(result.value.outputs["ENDPOINT"]).to.equal("https://example.com");
        }
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true });
      }
    });

    it("loads parameters with env resolution", async () => {
      const tmpDir = path.join(process.env.TEMP || "/tmp", "arm-params-" + Date.now());
      await fs.mkdir(tmpDir, { recursive: true });
      const templatePath = path.join(tmpDir, "main.json");
      const paramsPath = path.join(tmpDir, "params.json");
      await fs.writeFile(templatePath, JSON.stringify({ resources: [], outputs: {} }));
      process.env.TEST_ARM_VALUE = "resolved-value";
      await fs.writeFile(
        paramsPath,
        JSON.stringify({ parameters: { appName: { value: "${{TEST_ARM_VALUE}}" } } })
      );

      try {
        const ctx = createMockContext({ projectPath: tmpDir });
        (ctx.auth.azureAccountProvider as any).getIdentityCredentialAsync = sandbox
          .stub()
          .resolves({ getToken: sandbox.stub().resolves({ token: FAKE_TOKEN }) });

        mockAxios.put.resolves({
          status: 200,
          data: { properties: { provisioningState: "Succeeded", outputs: {} } },
        });

        const result = await armDeployDriver.executeFn(ctx, {
          subscriptionId: FAKE_SUB,
          resourceGroupName: FAKE_RG,
          templates: [{ path: "main.json", parameters: "params.json", deploymentName: "deploy1" }],
        });

        expect(result.isOk()).to.be.true;
        // Verify PUT was called with resolved parameters
        const putCall = mockAxios.put.getCall(0);
        const body = putCall.args[1];
        expect(body.properties.parameters.appName.value).to.equal("resolved-value");
      } finally {
        delete process.env.TEST_ARM_VALUE;
        await fs.rm(tmpDir, { recursive: true, force: true });
      }
    });
  });

  // ── azureAppService/zipDeploy driver ──────────────────────────────────

  describe("azureAppService/zipDeploy driver", () => {
    it("has correct metadata", () => {
      expect(azureAppServiceZipDeployDriver.id).to.equal("azureAppService/zipDeploy");
      expect(azureAppServiceZipDeployDriver.name).to.equal("Deploy to Azure App Service");
    });

    it("validates resource ID is required", () => {
      const result = azureAppServiceZipDeployDriver.validateFn!({
        resourceId: "",
        artifactFolder: "dist",
      });
      expect(result.isErr()).to.be.true;
    });

    it("accepts valid config", () => {
      const result = azureAppServiceZipDeployDriver.validateFn!({
        resourceId: `/subscriptions/${FAKE_SUB}/resourceGroups/${FAKE_RG}/providers/Microsoft.Web/sites/myapp`,
        artifactFolder: "dist",
      });
      expect(result.isOk()).to.be.true;
    });

    it("requires projectPath", async () => {
      const ctx = createMockContext({ projectPath: undefined });
      const result = await azureAppServiceZipDeployDriver.executeFn(ctx, {
        resourceId: `/subscriptions/${FAKE_SUB}/resourceGroups/${FAKE_RG}/providers/Microsoft.Web/sites/myapp`,
        artifactFolder: "dist",
      });
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("MissingProjectPath");
      }
    });

    it("returns error on invalid resource ID", async () => {
      const ctx = createMockContext();
      const result = await azureAppServiceZipDeployDriver.executeFn(ctx, {
        resourceId: "invalid-resource-id",
        artifactFolder: "dist",
      });
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("InvalidResourceId");
      }
    });

    it("returns error when artifact folder missing", async () => {
      const tmpDir = path.join(process.env.TEMP || "/tmp", "zip-test-" + Date.now());
      await fs.mkdir(tmpDir, { recursive: true });
      try {
        const ctx = createMockContext({ projectPath: tmpDir });
        const result = await azureAppServiceZipDeployDriver.executeFn(ctx, {
          resourceId: `/subscriptions/${FAKE_SUB}/resourceGroups/${FAKE_RG}/providers/Microsoft.Web/sites/myapp`,
          artifactFolder: "nonexistent-dist",
        });
        expect(result.isErr()).to.be.true;
        if (result.isErr()) {
          expect(result.error.code).to.equal("ArtifactFolderNotFound");
        }
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true });
      }
    });

    it("creates ZIP and deploys", async () => {
      const tmpDir = path.join(process.env.TEMP || "/tmp", "zip-deploy-" + Date.now());
      const distDir = path.join(tmpDir, "dist");
      await fs.mkdir(distDir, { recursive: true });
      await fs.writeFile(path.join(distDir, "index.js"), "module.exports = {}");

      try {
        const ctx = createMockContext({ projectPath: tmpDir });
        (ctx.auth.azureAccountProvider as any).getIdentityCredentialAsync = sandbox
          .stub()
          .resolves({ getToken: sandbox.stub().resolves({ token: FAKE_TOKEN }) });

        // Mock SCM endpoint lookup
        mockAxios.get.resolves({
          status: 200,
          data: {
            properties: {
              enabledHostNames: ["myapp.azurewebsites.net", "myapp.scm.azurewebsites.net"],
            },
          },
        });

        // Mock zip deploy upload
        const axiosPostStub = sandbox.stub(axios, "post").resolves({
          status: 202,
          headers: { location: "https://myapp.scm.azurewebsites.net/api/deployments/latest" },
        });

        // Mock deploy status check
        const _axiosGetStub = sandbox.stub(axios, "get").resolves({
          status: 200,
          data: { status: 4 }, // DeployStatus.Success
        });

        // Mock restart
        mockAxios.post.resolves({ status: 200 });

        const result = await azureAppServiceZipDeployDriver.executeFn(ctx, {
          resourceId: `/subscriptions/${FAKE_SUB}/resourceGroups/${FAKE_RG}/providers/Microsoft.Web/sites/myapp`,
          artifactFolder: "dist",
        });

        expect(result.isOk()).to.be.true;
        expect(axiosPostStub.calledOnce).to.be.true;
      } finally {
        await fs.rm(tmpDir, { recursive: true, force: true });
      }
    });
  });

  // ── azureFunctions/zipDeploy driver ───────────────────────────────────

  describe("azureFunctions/zipDeploy driver", () => {
    it("has correct metadata", () => {
      expect(azureFunctionsZipDeployDriver.id).to.equal("azureFunctions/zipDeploy");
      expect(azureFunctionsZipDeployDriver.name).to.equal("Deploy to Azure Functions");
    });

    it("accepts valid config", () => {
      const result = azureFunctionsZipDeployDriver.validateFn!({
        resourceId: `/subscriptions/${FAKE_SUB}/resourceGroups/${FAKE_RG}/providers/Microsoft.Web/sites/myfunc`,
        artifactFolder: "dist",
      });
      expect(result.isOk()).to.be.true;
    });

    it("requires projectPath", async () => {
      const ctx = createMockContext({ projectPath: undefined });
      const result = await azureFunctionsZipDeployDriver.executeFn(ctx, {
        resourceId: `/subscriptions/${FAKE_SUB}/resourceGroups/${FAKE_RG}/providers/Microsoft.Web/sites/myfunc`,
        artifactFolder: "dist",
      });
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("MissingProjectPath");
      }
    });
  });
});
