// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import axios, { AxiosRequestConfig, AxiosResponse } from "axios";
import chai from "chai";
import chaiAsPromised from "chai-as-promised";
import sinon from "sinon";
import { SPOClient } from "../../../../../src/component/driver/deploy/spfx/utility/spoClient";

chai.use(chaiAsPromised);
const expect = chai.expect;

describe("SPFx SPO Client", async () => {
  afterEach(() => {
    sinon.restore();
  });

  it("get app catalog site", async () => {
    sinon.stub(axios, "create").returns({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: sinon.stub(),
        },
      },
      get: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { CorporateCatalogUrl: "fakeUrl" } } as any;
      },
    } as any);
    await expect(SPOClient.getAppCatalogSite("")).to.eventually.equal("fakeUrl");
  });

  it("get app catalog site - undefined", async () => {
    sinon.stub(axios, "create").returns({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: sinon.stub(),
        },
      },
      get: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: {} } as any;
      },
    } as any);
    await expect(SPOClient.getAppCatalogSite("")).to.eventually.equal(undefined);
  });

  it("upload app package", async () => {
    sinon.stub(axios, "create").returns({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: sinon.stub(),
        },
      },
      post: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { CorporateCatalogUrl: "fakeUrl" } } as any;
      },
    } as any);
    await expect(SPOClient.uploadAppPackage("", "", Buffer.from(""))).not.rejected;
  });

  it("deploy app package", async () => {
    sinon.stub(axios, "create").returns({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: sinon.stub(),
        },
      },
      post: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { CorporateCatalogUrl: "fakeUrl" } } as any;
      },
    } as any);
    await expect(SPOClient.deployAppPackage("", "")).not.rejected;
  });

  it("create app catelog", async () => {
    sinon.stub(axios, "create").returns({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: sinon.stub(),
        },
      },
      post: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { CorporateCatalogUrl: "fakeUrl" } } as any;
      },
    } as any);
    await expect(SPOClient.createAppCatalog("")).not.rejected;
  });
});
