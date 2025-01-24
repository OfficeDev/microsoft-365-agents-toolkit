// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import chai from "chai";
import fs from "fs-extra";
import { PackageService } from "../../../../src/component/m365/packageService";
import { M365TitleAcquireDriver } from "../../../../src/component/driver/m365/acquire";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import {
  FileNotFoundError,
  InvalidActionInputError,
  UnhandledError,
} from "../../../../src/error/common";
import { MockedM365Provider } from "../../../core/utils";

describe("teamsApp/extendToM365", async () => {
  const acquireDriver = new M365TitleAcquireDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
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

    const result = await acquireDriver.run(args, mockedDriverContext);
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
    ]);

    const result = await acquireDriver.execute(args, mockedDriverContext, outputEnvVarNames);
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

    const result = await acquireDriver.execute(args, mockedDriverContext, undefined);
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
    const outputEnvVarNames = new Map([["appId", "MY_APP_ID"]]);

    const result = await acquireDriver.execute(args, mockedDriverContext, outputEnvVarNames);
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
    const outputEnvVarNames = new Map([["titleId", "MY_TITLE_ID"]]);

    const result = await acquireDriver.execute(args, mockedDriverContext, outputEnvVarNames);
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
    ]);

    const result = await acquireDriver.execute(args, mockedDriverContext, outputEnvVarNames);
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
    ]);

    sinon.stub(PackageService.prototype, "sideLoading").throws(new Error("test error"));
    sinon.stub(fs, "pathExists").resolves(true);

    const result = await acquireDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert(result.result.isErr());
    if (result.result.isErr()) {
      chai.assert.isTrue(result.result.error instanceof UnhandledError);
    }
  });

  it("execute: happy path", async () => {
    const args = {
      appPackagePath: "fakePath",
    };
    const outputEnvVarNames = new Map([
      ["titleId", "MY_TITLE_ID"],
      ["appId", "MY_APP_ID"],
    ]);

    sinon
      .stub(PackageService.prototype, "sideLoading")
      .resolves(["test-title-id", "test-app-id", ""]);
    sinon.stub(fs, "pathExists").resolves(true);

    const result = await acquireDriver.execute(args, mockedDriverContext, outputEnvVarNames);
    chai.assert.isTrue(result.result.isOk());
    chai.assert.equal((result.result as any).value.get("MY_TITLE_ID"), "test-title-id");
    chai.assert.equal((result.result as any).value.get("MY_APP_ID"), "test-app-id");
  });
});
