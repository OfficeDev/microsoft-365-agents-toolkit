// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { M365TokenProvider, ok, Platform } from "@microsoft/teamsfx-api";
import axios, { AxiosRequestConfig, AxiosResponse } from "axios";
import chai from "chai";
import chaiAsPromised from "chai-as-promised";
import faker from "faker";
import fs from "fs-extra";
import "mocha";
import sinon from "sinon";

import path from "path";
import * as Tools from "../../../../../src/common/tools";
import { SPFxDeployDriver } from "../../../../../src/component/driver/deploy/spfx/deployDriver";
import { CreateAppCatalogFailedError } from "../../../../../src/component/driver/deploy/spfx/error/createAppCatalogFailedError";
import { GetGraphTokenFailedError } from "../../../../../src/component/driver/deploy/spfx/error/getGraphTokenFailedError";
import { GetSPOTokenFailedError } from "../../../../../src/component/driver/deploy/spfx/error/getSPOTokenFailedError";
import { GetTenantFailedError } from "../../../../../src/component/driver/deploy/spfx/error/getTenantFailedError";
import { InsufficientPermissionError } from "../../../../../src/component/driver/deploy/spfx/error/insufficientPermissionError";
import { NoSPPackageError } from "../../../../../src/component/driver/deploy/spfx/error/noSPPackageError";
import { NoValidAppCatelog } from "../../../../../src/component/driver/deploy/spfx/error/noValidAppCatelogError";
import { UploadAppPackageFailedError } from "../../../../../src/component/driver/deploy/spfx/error/uploadAppPackageFailedError";
import { DeploySPFxArgs } from "../../../../../src/component/driver/deploy/spfx/interface/deployArgs";
import { Constants } from "../../../../../src/component/driver/deploy/spfx/utility/constants";
import { SPOClient } from "../../../../../src/component/driver/deploy/spfx/utility/spoClient";
import { FileNotFoundError } from "../../../../../src/error/common";
import { MockedM365Provider } from "../../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../../plugins/solution/util";

chai.use(chaiAsPromised);
const expect = chai.expect;

describe("SPFx Deploy Driver", async () => {
  const args: DeploySPFxArgs = {
    createAppCatalogIfNotExist: true,
    packageSolutionPath: "./SPFx/config/package-solution.json",
  };
  const deployDriver = new SPFxDeployDriver();
  const mockedDriverContext: any = {
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    m365TokenProvider: new MockedM365Provider(),
    platform: Platform.VSCode,
    projectPath: "C://TeamsApp",
  };

  afterEach(() => {
    sinon.restore();
  });

  it("should successfully deploy if app catelog exists - VSCode", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return path === "C://fake/appPackage/a.zip";
    });
    sinon.stub(fs, "readFile").resolves(Buffer.from("content"));

    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    sinon.stub(SPFxDeployDriver.prototype, "getPackagePath").resolves("C://fake/appPackage/a.zip");
    sinon.stub(SPFxDeployDriver.prototype, "getAppID").resolves("fakeAppID");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves("fakeAppCatelogSite");
    sinon.stub(SPOClient, "uploadAppPackage");
    sinon.stub(SPOClient, "deployAppPackage");

    const result = await deployDriver.run(args, mockedDriverContext);
    expect(result.isOk()).to.be.true;
  });

  it("should successfully deploy if app catelog exists - VSCode - execute", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return path === "C://fake/appPackage/a.zip";
    });
    sinon.stub(fs, "readFile").resolves(Buffer.from("content"));

    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    sinon.stub(SPFxDeployDriver.prototype, "getPackagePath").resolves("C://fake/appPackage/a.zip");
    sinon.stub(SPFxDeployDriver.prototype, "getAppID").resolves("fakeAppID");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves("fakeAppCatelogSite");
    sinon.stub(SPOClient, "uploadAppPackage");
    sinon.stub(SPOClient, "deployAppPackage");

    const result = await deployDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    expect(result.summaries.length).to.eq(3);
  });

  it("should successfully deploy if app catelog exists - CLI", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return path === "C://fake/appPackage/a.zip";
    });
    sinon.stub(fs, "readFile").resolves(Buffer.from("content"));

    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    sinon.stub(SPFxDeployDriver.prototype, "getPackagePath").resolves("C://fake/appPackage/a.zip");
    sinon.stub(SPFxDeployDriver.prototype, "getAppID").resolves("fakeAppID");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves("fakeAppCatelogSite");
    sinon.stub(SPOClient, "uploadAppPackage");
    sinon.stub(SPOClient, "deployAppPackage");

    const result = await deployDriver.run(args, {
      logProvider: new MockedLogProvider(),
      m365TokenProvider: new MockedM365Provider(),
      platform: Platform.CLI,
      projectPath: "C://TeamsApp",
    } as any);
    expect(result.isOk()).to.be.true;
  });

  it("should successfully deploy if app catelog not exist", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return path === "C://fake/appPackage/a.zip";
    });
    sinon.stub(fs, "readFile").resolves(Buffer.from("content"));
    (Constants as any).APP_CATALOG_REFRESH_TIME = 0;
    (Constants as any).APP_CATALOG_ACTIVE_TIME = 0;

    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    sinon.stub(SPFxDeployDriver.prototype, "getPackagePath").resolves("C://fake/appPackage/a.zip");
    sinon.stub(SPFxDeployDriver.prototype, "getAppID").resolves("fakeAppID");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon
      .stub(SPOClient, "getAppCatalogSite")
      .onFirstCall()
      .resolves(undefined)
      .onSecondCall()
      .resolves(undefined)
      .onThirdCall()
      .resolves("fakeAppCatelogSite");
    sinon.stub(SPOClient, "createAppCatalog");
    sinon.stub(SPOClient, "uploadAppPackage");
    sinon.stub(SPOClient, "deployAppPackage");

    const result = await deployDriver.run(args, mockedDriverContext);
    expect(result.isOk()).to.be.true;
  });

  it("fail to get SPFx token - GetSPOTokenFailed", async () => {
    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    sinon.stub(Tools, "getSPFxToken").resolves(undefined);

    const result = await deployDriver.run(args, mockedDriverContext);
    expect(result.isErr()).to.be.true;
    expect((result as any).error).instanceOf(GetSPOTokenFailedError);
  });

  it("fail to create app catelog - CreateAppCatalogFailed", async () => {
    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves(undefined);

    const result = await deployDriver.run(args, mockedDriverContext);
    expect(result.isErr()).to.be.true;
    expect((result as any).error).instanceOf(CreateAppCatalogFailedError);
  });

  it("fail to create app catelog - NoValidAppCatelog", async () => {
    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves(undefined);

    const result = await deployDriver.run(
      {
        createAppCatalogIfNotExist: false,
        packageSolutionPath: "./SPFx/config/package-solution.json",
      },
      mockedDriverContext
    );
    expect(result.isErr()).to.be.true;
    expect((result as any).error).instanceOf(NoValidAppCatelog);
  });

  it("fail to get app catelog - CreateAppCatalogFailedError", async () => {
    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    (Constants as any).APP_CATALOG_REFRESH_TIME = 0;

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves(undefined);
    sinon.stub(SPOClient, "createAppCatalog");

    const result = await deployDriver.run(args, mockedDriverContext);
    expect(result.isErr()).to.be.true;
    expect((result as any).error).instanceOf(CreateAppCatalogFailedError);
  });

  it("fail to get package path - NoSPPackageError", async () => {
    sinon.stub(fs, "pathExists").resolves(false);

    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    sinon.stub(SPFxDeployDriver.prototype, "getPackagePath").resolves("C://fake/appPackage/a.zip");
    sinon.stub(SPFxDeployDriver.prototype, "getAppID").resolves("fakeAppID");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves("fakeAppCatelogSite");
    sinon.stub(SPOClient, "createAppCatalog");

    const result = await deployDriver.run(args, mockedDriverContext);
    expect(result.isErr()).to.be.true;
    expect((result as any).error).instanceOf(NoSPPackageError);
  });

  it("fail to upload app package - 403", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return path === "C://fake/appPackage/a.zip";
    });
    sinon.stub(fs, "readFile").resolves(Buffer.from("content"));

    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    sinon.stub(SPFxDeployDriver.prototype, "getPackagePath").resolves("C://fake/appPackage/a.zip");
    sinon.stub(SPFxDeployDriver.prototype, "getAppID").resolves("fakeAppID");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves("fakeAppCatelogSite");
    sinon.stub(SPOClient, "uploadAppPackage").throws({
      response: {
        status: 403,
      },
    });

    const result = await deployDriver.run(args, mockedDriverContext);
    expect(result.isErr()).to.be.true;
    expect((result as any).error).instanceOf(InsufficientPermissionError);
  });

  it("fail to upload app package - UploadAppPackageFailedError", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
      return path === "C://fake/appPackage/a.zip";
    });
    sinon.stub(fs, "readFile").resolves(Buffer.from("content"));

    sinon.stub(SPFxDeployDriver.prototype, "getTenant").resolves("fakeTenant");
    sinon.stub(SPFxDeployDriver.prototype, "getPackagePath").resolves("C://fake/appPackage/a.zip");
    sinon.stub(SPFxDeployDriver.prototype, "getAppID").resolves("fakeAppID");

    sinon.stub(Tools, "getSPFxToken").resolves("fakeSPFxToken");
    sinon.stub(SPOClient, "getAppCatalogSite").resolves("fakeAppCatelogSite");
    sinon.stub(SPOClient, "uploadAppPackage").throws(new Error("fakeError"));

    const result = await deployDriver.run(args, mockedDriverContext);
    expect(result.isErr()).to.be.true;
    expect((result as any).error).instanceOf(UploadAppPackageFailedError);
  });

  it("get tenant from M365TokenProvider", async () => {
    sinon.stub(axios, "create").returns({
      defaults: { headers: { common: {} } },
      get: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { webUrl: "fakeWebUrl" } } as any;
      },
    } as any);
    await expect(deployDriver.getTenant(mockM365TokenProvider())).to.eventually.equal("fakeWebUrl");
  });

  it("fail to tenant from M365TokenProvider - GetGraphTokenFailedError", async () => {
    const mockedM365TokenProvider = mockM365TokenProvider();
    (mockedM365TokenProvider as any).getAccessToken = sinon.stub().returns(ok(undefined));
    await expect(deployDriver.getTenant(mockedM365TokenProvider)).to.be.rejectedWith(
      GetGraphTokenFailedError
    );
  });

  it("fail to get tenant from M365TokenProvider - GetTenantFailedError", async () => {
    sinon.stub(axios, "create").returns({
      defaults: { headers: { common: {} } },
      get: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return undefined as any;
      },
    } as any);
    await expect(deployDriver.getTenant(mockM365TokenProvider())).to.be.rejectedWith(
      GetTenantFailedError
    );
  });

  it("fail to get tenant from M365TokenProvider - GetTenantFailedError", async () => {
    sinon.stub(axios, "create").returns({
      defaults: { headers: { common: {} } },
      get: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        throw new Error();
      },
    } as any);
    await expect(deployDriver.getTenant(mockM365TokenProvider())).to.be.rejectedWith(
      GetTenantFailedError
    );
  });

  it("get package path from solutionConfigPath", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readJson").resolves({ paths: { zippedPackage: "solution/a.zip" } });
    await expect(
      deployDriver.getPackagePath("C://test/config/package-solution.json")
    ).to.eventually.equal(path.resolve("C://test/sharepoint/solution/a.zip"));
  });

  it("fail to get package path from solutionConfigPath - PathNotExistsError", async () => {
    sinon.stub(fs, "pathExists").resolves(false);
    await expect(
      deployDriver.getPackagePath("C://test/config/package-solution.json")
    ).to.be.rejectedWith(FileNotFoundError);
  });

  it("get app id from solutionConfigPath", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readJson").resolves({ solution: { id: "fakeID" } });
    await expect(
      deployDriver.getAppID("C://test/config/package-solution.json")
    ).to.eventually.equal("fakeID");
  });

  it("fail to get app id from solutionConfigPath - PathNotExistsError", async () => {
    sinon.stub(fs, "pathExists").resolves(false);
    await expect(deployDriver.getAppID("C://test/config/package-solution.json")).to.be.rejectedWith(
      FileNotFoundError
    );
  });
});

export function mockM365TokenProvider(): M365TokenProvider {
  const provider = {} as M365TokenProvider;
  const mockTokenObject = {
    tid: faker.datatype.uuid(),
  };

  provider.getAccessToken = sinon.stub().returns(ok("token"));
  provider.getJsonObject = sinon.stub().returns(ok(mockTokenObject));
  return provider;
}
