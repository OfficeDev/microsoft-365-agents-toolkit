/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import axios from "axios";
import * as retryModule from "../../../src/http/retry";
import { createMockContext } from "../testHelper";
import { AzureArmClient } from "../../../src/clients/azure/client";

const FAKE_TOKEN = "fake-azure-token";
const FAKE_SUB = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const FAKE_RG = "test-rg";

describe("AzureArmClient", () => {
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

  it("constructor sets Authorization header", () => {
    const ctx = createMockContext();
    new AzureArmClient(ctx, FAKE_TOKEN);
    expect(mockAxios.defaults.headers.common["Authorization"]).to.equal(`Bearer ${FAKE_TOKEN}`);
  });

  it("constructor sets Content-Type header", () => {
    const ctx = createMockContext();
    new AzureArmClient(ctx, FAKE_TOKEN);
    expect(mockAxios.defaults.headers.common["Content-Type"]).to.equal("application/json");
  });

  describe("deployTemplate", () => {
    it("returns outputs on success", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.put.resolves({
        status: 200,
        data: {
          properties: {
            provisioningState: "Succeeded",
            outputs: {
              botEndpoint: { type: "String", value: "https://bot.example.com" },
            },
          },
        },
      });

      const result = await client.deployTemplate(FAKE_SUB, FAKE_RG, "my-deployment", {
        properties: { template: {}, parameters: null, mode: "Incremental" },
      });

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.properties?.provisioningState).to.equal("Succeeded");
        expect(result.value.properties?.outputs?.botEndpoint.value).to.equal(
          "https://bot.example.com"
        );
      }
    });

    it("returns error on Canceled state", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.put.resolves({
        status: 200,
        data: { properties: { provisioningState: "Canceled" } },
      });

      const result = await client.deployTemplate(FAKE_SUB, FAKE_RG, "my-deployment", {
        properties: { template: {}, parameters: null, mode: "Incremental" },
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.message).to.include("Canceled");
      }
    });

    it("includes ARM error details in error message", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.put.resolves({
        status: 200,
        data: {
          properties: {
            provisioningState: "Failed",
            error: { code: "InvalidTemplate", message: "The template is invalid" },
          },
        },
      });

      const result = await client.deployTemplate(FAKE_SUB, FAKE_RG, "my-deployment", {
        properties: { template: {}, parameters: null, mode: "Incremental" },
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.message).to.include("InvalidTemplate");
        expect(result.error.message).to.include("The template is invalid");
      }
    });

    it("classifies 4xx API errors as user errors", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      const axiosError = {
        response: {
          status: 404,
          data: { error: { code: "ResourceGroupNotFound", message: "RG not found" } },
        },
        message: "Request failed with status code 404",
      };
      mockAxios.put.rejects(axiosError);

      const result = await client.deployTemplate(FAKE_SUB, FAKE_RG, "my-deployment", {
        properties: { template: {}, parameters: null, mode: "Incremental" },
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.kind).to.equal("user");
        expect(result.error.code).to.equal("ResourceGroupNotFound");
      }
    });

    it("classifies 5xx API errors as system errors", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      const axiosError = {
        response: {
          status: 502,
          data: { error: { code: "BadGateway", message: "Service unavailable" } },
        },
        message: "Request failed with status code 502",
      };
      mockAxios.put.rejects(axiosError);

      const result = await client.deployTemplate(FAKE_SUB, FAKE_RG, "my-deployment", {
        properties: { template: {}, parameters: null, mode: "Incremental" },
      });

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.kind).to.equal("system");
      }
    });
  });

  describe("getScmEndpoint", () => {
    it("extracts SCM hostname from enabled host names", async () => {
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

    it("returns error when no SCM host found", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.get.resolves({
        status: 200,
        data: { properties: { enabledHostNames: [] } },
      });

      const result = await client.getScmEndpoint(FAKE_SUB, FAKE_RG, "myapp");
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.code).to.equal("ScmEndpointNotFound");
      }
    });
  });

  describe("restartSite", () => {
    it("succeeds on 200", async () => {
      const ctx = createMockContext();
      const client = new AzureArmClient(ctx, FAKE_TOKEN);
      mockAxios.post.resolves({ status: 200, data: {} });

      const result = await client.restartSite(FAKE_SUB, FAKE_RG, "myapp");
      expect(result.isOk()).to.be.true;
      expect(mockAxios.post.calledOnce).to.be.true;
    });
  });
});
