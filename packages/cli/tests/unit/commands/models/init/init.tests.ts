// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-non-null-assertion */

import { CLIContext, err, ok, SystemError } from "@microsoft/teamsfx-api";
import { FxCore, UserCancelError } from "@microsoft/teamsfx-core";
import { assert } from "chai";
import "mocha";
import * as sinon from "sinon";
import * as activate from "../../../../../src/activate";
import { initCommand } from "../../../../../src/commands/models/init/init";
import { logger } from "../../../../../src/commonlib/logger";

describe("init command", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(logger, "info").resolves(true);
    sandbox.stub(logger, "error").resolves(true);
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("handler", () => {
    it("should successfully generate config files with all options", async () => {
      const mockCore = new FxCore({} as any);
      sandbox.stub(activate, "getFxCore").returns(mockCore);
      sandbox.stub(mockCore, "generateConfigFiles").resolves(ok(undefined));

      const ctx: CLIContext = {
        command: { ...initCommand, fullName: "init" },
        optionValues: {
          playground: true,
          local: true,
          remote: true,
          language: "typescript",
          "teams-manifest-file": "./appPackage/manifest.json",
          folder: "./",
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const result = await initCommand.handler!(ctx);

      assert.isTrue(result.isOk());
      assert.isTrue(
        (mockCore.generateConfigFiles as sinon.SinonStub).calledOnceWith(ctx.optionValues)
      );
    });

    it("should successfully generate config files with default options", async () => {
      const mockCore = new FxCore({} as any);
      sandbox.stub(activate, "getFxCore").returns(mockCore);
      sandbox.stub(mockCore, "generateConfigFiles").resolves(ok(undefined));

      const ctx: CLIContext = {
        command: { ...initCommand, fullName: "init" },
        optionValues: {
          playground: true,
          local: true,
          remote: false,
          language: "typescript",
          "teams-manifest-file": "./appPackage/manifest.json",
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const result = await initCommand.handler!(ctx);

      assert.isTrue(result.isOk());
      assert.isTrue(
        (mockCore.generateConfigFiles as sinon.SinonStub).calledOnceWith(ctx.optionValues)
      );
    });

    it("should return error when generateConfigFiles fails with UserCancelError", async () => {
      const mockCore = new FxCore({} as any);
      const expectedError = new UserCancelError();
      sandbox.stub(activate, "getFxCore").returns(mockCore);
      sandbox.stub(mockCore, "generateConfigFiles").resolves(err(expectedError));

      const ctx: CLIContext = {
        command: { ...initCommand, fullName: "init" },
        optionValues: {
          playground: true,
          local: true,
          remote: false,
          language: "typescript",
          "teams-manifest-file": "./appPackage/manifest.json",
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const result = await initCommand.handler!(ctx);

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error, expectedError);
      }
    });

    it("should return error when generateConfigFiles fails with SystemError", async () => {
      const mockCore = new FxCore({} as any);
      const expectedError = new SystemError("TestSource", "TestError", "Test error message");
      sandbox.stub(activate, "getFxCore").returns(mockCore);
      sandbox.stub(mockCore, "generateConfigFiles").resolves(err(expectedError));

      const ctx: CLIContext = {
        command: { ...initCommand, fullName: "init" },
        optionValues: {
          playground: false,
          local: true,
          remote: true,
          language: "typescript",
          "teams-manifest-file": "./manifest.json",
          folder: "./test",
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const result = await initCommand.handler!(ctx);

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error, expectedError);
      }
    });

    it("should pass correct inputs to generateConfigFiles", async () => {
      const mockCore = new FxCore({} as any);
      sandbox.stub(activate, "getFxCore").returns(mockCore);
      const generateConfigFilesStub = sandbox
        .stub(mockCore, "generateConfigFiles")
        .resolves(ok(undefined));

      const expectedInputs = {
        playground: false,
        local: false,
        remote: true,
        language: "typescript",
        "teams-manifest-file": "./custom/path/manifest.json",
        folder: "./custom/folder",
      };

      const ctx: CLIContext = {
        command: { ...initCommand, fullName: "init" },
        optionValues: expectedInputs,
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const result = await initCommand.handler!(ctx);

      assert.isTrue(result.isOk());
      assert.isTrue(generateConfigFilesStub.calledOnce);
      const actualInputs = generateConfigFilesStub.firstCall.args[0];
      assert.equal(actualInputs.playground, expectedInputs.playground);
      assert.equal(actualInputs.local, expectedInputs.local);
      assert.equal(actualInputs.remote, expectedInputs.remote);
      assert.equal(actualInputs.language, expectedInputs.language);
      assert.equal(actualInputs["teams-manifest-file"], expectedInputs["teams-manifest-file"]);
      assert.equal(actualInputs.folder, expectedInputs.folder);
    });

    it("should handle empty option values", async () => {
      const mockCore = new FxCore({} as any);
      sandbox.stub(activate, "getFxCore").returns(mockCore);
      sandbox.stub(mockCore, "generateConfigFiles").resolves(ok(undefined));

      const ctx: CLIContext = {
        command: { ...initCommand, fullName: "init" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const result = await initCommand.handler!(ctx);

      assert.isTrue(result.isOk());
    });
  });
});
