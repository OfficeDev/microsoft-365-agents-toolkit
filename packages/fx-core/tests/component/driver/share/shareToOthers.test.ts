// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import chai from "chai";
import fs from "fs-extra";
import { PackageService } from "../../../../src/component/m365/packageService";
import { ShareToOthersDriver } from "../../../../src/component/driver/share/shareToOthers";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import {
  FileNotFoundError,
  InvalidActionInputError,
  UnhandledError,
} from "../../../../src/error/common";
import { MockedM365Provider } from "../../../core/utils";

describe("teamsApp/shareToOthers", async () => {
  const shareDriver = new ShareToOthersDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  const mockedDriverContextWithoutLogProvider: any = {
    m365TokenProvider: new MockedM365Provider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  afterEach(() => {
    sinon.restore();
  });

  it("run: happy path", async () => {
    const args = {
      appPackagePath: "fakePath",
    };

    const result = await shareDriver.run(args, mockedDriverContext);
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidActionInputError);
      chai.assert.isTrue(result.error.message.includes("writeToEnvironmentFile"));
    }
  });

  it("execute: invalid param error", async () => {
    const args = {
      appPackagePath: false,
    } as any;
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof InvalidActionInputError);
      chai.assert.isTrue(result.result.error.message.includes("appPackagePath"));
    }
  });

  it("execute: appPackagePath is undefined", async () => {
    const args = {
      // appPackagePath is missing
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof InvalidActionInputError);
      chai.assert.isTrue(result.result.error.message.includes("appPackagePath"));
    }
  });

  it("execute: writeToEnvironmentFile undefined", async () => {
    const args = {
      appPackagePath: "fakePath",
    };

    const result = await shareDriver.execute(args, mockedDriverContext, undefined);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof InvalidActionInputError);
      chai.assert.isTrue(result.result.error.message.includes("writeToEnvironmentFile"));
    }
  });

  it("execute: missing titleId", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof InvalidActionInputError);
      chai.assert.isTrue(result.result.error.message.includes("writeToEnvironmentFile"));
    }
  });

  it("execute: missing appId", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
    ]);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof InvalidActionInputError);
      chai.assert.isTrue(result.result.error.message.includes("writeToEnvironmentFile"));
    }
  });

  it("execute: missing shareLink", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof InvalidActionInputError);
      chai.assert.isTrue(result.result.error.message.includes("writeToEnvironmentFile"));
    }
  });

  it("execute: should throw error if file not exists", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof FileNotFoundError);
    }
  });

  it("execute: unhandled error", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    sinon.stub(PackageService.prototype, "sideLoading").throws(new Error("test error"));
    sinon.stub(fs, "pathExists").resolves(true);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof UnhandledError);
    }
  });

  it("execute: token provider returns error", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(mockedDriverContext.m365TokenProvider, "getAccessToken").resolves({
      isErr: () => true,
      error: new Error("Failed to get token"),
    });
    try {
      await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    } catch (error) {
      chai.assert.isTrue(error instanceof Error);
      chai.assert.isTrue(error.message.include("Failed to get token"));
    }
  });

  it("execute: token provider returns error without log provider", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(mockedDriverContext.m365TokenProvider, "getAccessToken").resolves({
      isErr: () => true,
      error: new Error("Failed to get token"),
    });

    try {
      await shareDriver.execute(args, mockedDriverContextWithoutLogProvider, outputEnvVarNames);
    } catch (error) {
      chai.assert.isTrue(error instanceof Error);
      chai.assert.isTrue(error.message.include("Failed to get token"));
    }
  });

  it("execute: happy path", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    sinon
      .stub(PackageService.prototype, "sideLoading")
      .resolves(["test-title-id", "test-app-id", "www.fake.shareLink.com"]);
    sinon.stub(fs, "pathExists").resolves(true);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
    chai.assert.equal((result.result as any).value.get("MY_TITLE_ID"), "test-title-id");
    chai.assert.equal((result.result as any).value.get("MY_APP_ID"), "test-app-id");
    chai.assert.equal((result.result as any).value.get("MY_SHARE_LINK"), "www.fake.shareLink.com");
  });

  it("execute: happy path with custom sideloading env", async () => {
    process.env.SIDELOADING_SERVICE_ENDPOINT = "customEndpoint";
    process.env.SIDELOADING_SERVICE_SCOPE = "customScope";

    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
      ["shareLink", "MY_SHARE_LINK"],
    ]);

    sinon.stub(fs, "pathExists").resolves(true);

    sinon.stub(mockedDriverContext.m365TokenProvider, "getAccessToken").resolves({
      isErr: () => false,
      value: "fakeToken",
    });

    // Stub PackageService.sideLoading to return custom values
    sinon
      .stub(PackageService.prototype, "sideLoading")
      .resolves(["custom-title-id", "custom-app-id", "www.custom.shareLink.com"]);

    const result = await shareDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
    chai.assert.equal((result.result as any).value.get("MY_TITLE_ID"), "custom-title-id");
    chai.assert.equal((result.result as any).value.get("MY_APP_ID"), "custom-app-id");
    chai.assert.equal(
      (result.result as any).value.get("MY_SHARE_LINK"),
      "www.custom.shareLink.com"
    );

    // Restore environment variables
    delete process.env.SIDELOADING_SERVICE_ENDPOINT;
    delete process.env.SIDELOADING_SERVICE_SCOPE;
  });
});
