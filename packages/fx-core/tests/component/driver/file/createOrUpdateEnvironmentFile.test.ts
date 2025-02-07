// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";

import * as chai from "chai";
import fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import * as sinon from "sinon";
import * as util from "util";

import * as localizeUtils from "../../../../src/common/localizeUtils";
import { CreateOrUpdateEnvironmentFileDriver } from "../../../../src/component/driver/file/createOrUpdateEnvironmentFile";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import {
  InvalidActionInputError,
  UnhandledError,
  UserCancelError,
} from "../../../../src/error/common";
import { pathUtils } from "../../../../src/component/utils/pathUtils";
import { err, ok } from "@microsoft/teamsfx-api";

describe("CreateOrUpdateEnvironmentFileDriver", () => {
  const mockedDriverContexts = [
    {
      logProvider: new MockedLogProvider(),
      projectPath: "/path/to/project",
      ui: new MockedUserInteraction(),
    } as any,
    {
      projectPath: "/path/to/project",
    } as any,
  ];
  const driver = new CreateOrUpdateEnvironmentFileDriver();

  beforeEach(() => {
    sinon.stub(localizeUtils, "getDefaultString").callsFake((key, ...params) => {
      if (key === "error.yaml.InvalidActionInputError") {
        return util.format("error.yaml.InvalidActionInputError. %s. %s.", ...params);
      } else if (key === "error.common.UnhandledError") {
        return util.format("error.common.UnhandledError. %s. %s", ...params);
      } else if (key === "driver.file.createOrUpdateEnvironmentFile.description") {
        return "driver.file.createOrUpdateEnvironmentFile.description";
      } else if (key === "driver.file.createOrUpdateEnvironmentFile.summary") {
        return util.format("driver.file.createOrUpdateEnvironmentFile.summary. %s.", ...params);
      }
      return "";
    });
    sinon
      .stub(localizeUtils, "getLocalizedString")
      .callsFake((key, ...params) => localizeUtils.getDefaultString(key, ...params));
  });

  afterEach(() => {
    sinon.restore();
  });

  describe("run", () => {
    for (const mockedDriverContext of mockedDriverContexts) {
      it("invalid args: empty target", async () => {
        const args: any = {
          target: null,
          envs: {
            key: "value",
          },
        };
        const result = await driver.run(args, mockedDriverContext);
        chai.assert(result.isErr());
        if (result.isErr()) {
          chai.assert(result.error instanceof InvalidActionInputError);
        }
      });

      it("invalid args: envs is not object", async () => {
        const args: any = {
          target: ".env.teamsfx.local",
          envs: "value",
        };
        const result = await driver.run(args, mockedDriverContext);
        chai.assert(result.isErr());
        if (result.isErr()) {
          chai.assert(result.error instanceof InvalidActionInputError);
        }
      });

      it("invalid args: envs is not key value pairs", async () => {
        const args: any = {
          target: ".env.teamsfx.local",
          envs: {
            key1: "value1",
            key2: {
              key3: "value3",
            },
          },
        };
        const result = await driver.run(args, mockedDriverContext);
        chai.assert(result.isErr());
        if (result.isErr()) {
          chai.assert(result.error instanceof InvalidActionInputError);
        }
      });

      it("exception", async () => {
        sinon.stub(fs, "ensureFile").throws(new Error("exception"));
        const args: any = {
          target: "path",
          envs: {
            key1: "value1",
            key2: "value2",
          },
        };
        const result = await driver.run(args, mockedDriverContext);
        chai.assert(result.isErr());
        if (result.isErr()) {
          chai.assert(result.error instanceof UnhandledError);
          const message =
            "error.common.UnhandledError. file/createOrUpdateEnvironmentFile. exception.";
          chai.assert(result.error.message, message);
        }
      });

      it("happy path: output to target", async () => {
        const target = path.join(mockedDriverContext.projectPath, ".env.local");
        const existingEnvs = {
          existing1: "value1",
          existing2: "value2",
        };
        let content = Object.entries(existingEnvs)
          .map(([key, value]) => `${key}=${value}`)
          .join(os.EOL);
        sinon.stub(fs, "ensureFile").resolves();
        sinon.stub(fs, "readFile").callsFake(async (path) => {
          return Buffer.from(content);
        });
        sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
          content = data;
        });
        const args: any = {
          target: ".env.local",
          envs: {
            key1: 10,
            key2: true,
            key3: "value3",
          },
        };
        sinon.stub(pathUtils, "getEnvFilePath").resolves(ok(target));
        const result = await driver.run(args, mockedDriverContext);
        chai.assert(result.isOk());
        if (result.isOk()) {
          chai.assert.equal(result.value.size, 3);
          const expectedEnvs = { ...existingEnvs, ...args.envs };
          const expectedContent = Object.entries(expectedEnvs)
            .map(([key, value]) => `${key}=${value}`)
            .join(os.EOL);
          chai.assert.equal(content, expectedContent);
        }
      });
      it("happy path: not env file", async () => {
        const existingEnvs = {
          existing1: "value1",
          existing2: "value2",
        };
        let content = Object.entries(existingEnvs)
          .map(([key, value]) => `${key}=${value}`)
          .join(os.EOL);
        sinon.stub(fs, "ensureFile").resolves();
        sinon.stub(fs, "readFile").callsFake(async (path) => {
          return Buffer.from(content);
        });
        sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
          content = data;
        });
        const args: any = {
          target: "E:\\home\\test\\.env.local",
          envs: {
            key1: 10,
            key2: true,
            key3: "value3",
          },
        };
        sinon.stub(pathUtils, "getEnvFilePath").resolves(ok("fake-path"));
        const result = await driver.run(args, mockedDriverContext);
        chai.assert(result.isOk());
        if (result.isOk()) {
          chai.assert.equal(result.value.size, 0);
          const expectedEnvs = { ...existingEnvs, ...args.envs };
          const expectedContent = Object.entries(expectedEnvs)
            .map(([key, value]) => `${key}=${value}`)
            .join(os.EOL);
          chai.assert.equal(content, expectedContent);
        }
      });
      it("happy path: getEnvFilePath error", async () => {
        const existingEnvs = {
          existing1: "value1",
          existing2: "value2",
        };
        let content = Object.entries(existingEnvs)
          .map(([key, value]) => `${key}=${value}`)
          .join(os.EOL);
        sinon.stub(fs, "ensureFile").resolves();
        sinon.stub(fs, "readFile").callsFake(async (path) => {
          return Buffer.from(content);
        });
        sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
          content = data;
        });
        const args: any = {
          target: "E:\\home\\test\\.env.local",
          envs: {
            key1: 10,
            key2: true,
            key3: "value3",
          },
        };
        sinon.stub(pathUtils, "getEnvFilePath").resolves(err(new UserCancelError()));
        const result = await driver.run(args, mockedDriverContext);
        chai.assert(result.isOk());
        if (result.isOk()) {
          chai.assert.equal(result.value.size, 0);
          const expectedEnvs = { ...existingEnvs, ...args.envs };
          const expectedContent = Object.entries(expectedEnvs)
            .map(([key, value]) => `${key}=${value}`)
            .join(os.EOL);
          chai.assert.equal(content, expectedContent);
        }
      });
    }
  });

  describe("execute", () => {
    beforeEach(() => {
      process.env.TEAMSFX_ENV = "local";
    });

    afterEach(() => {
      delete process.env.TEAMSFX_ENV;
    });

    for (const mockedDriverContext of mockedDriverContexts) {
      it("happy path: output to target", async () => {
        const target = path.join(mockedDriverContext.projectPath, ".env.teamsfx.local");
        const existingEnvs = {
          existing1: "value1",
          existing2: "value2",
        };
        let content = Object.entries(existingEnvs)
          .map(([key, value]) => `${key}=${value}`)
          .join(os.EOL);
        sinon.stub(fs, "ensureFile").callsFake(async (path) => {
          if (path !== target) {
            content = "";
          }
        });
        sinon.stub(fs, "readFile").callsFake(async (path) => {
          if (path === target) {
            return Buffer.from(content);
          }
          return Buffer.from("");
        });
        sinon.stub(fs, "writeFile").callsFake(async (path, data) => {
          if (path === target) {
            content = data;
          }
        });
        const args: any = {
          target: ".env.teamsfx.local",
          envs: {
            key1: 10,
            key2: true,
            key3: "value3",
          },
        };
        sinon.stub(pathUtils, "getEnvFilePath").resolves(ok(target));
        const executionResult = await driver.execute(args, mockedDriverContext);
        chai.assert(executionResult.result.isOk());
        if (executionResult.result.isOk()) {
          chai.assert.equal(executionResult.result.value.size, 3);
          const expectedEnvs = { ...existingEnvs, ...args.envs };
          const expectedContent = Object.entries(expectedEnvs)
            .map(([key, value]) => `${key}=${value}`)
            .join(os.EOL);
          chai.assert.equal(content, expectedContent);
        }
        chai.assert.equal(executionResult.summaries.length, 1);
        chai.assert.equal(
          executionResult.summaries[0],
          `driver.file.createOrUpdateEnvironmentFile.summary. ${path.normalize(target)}.`
        );
      });
    }
  });

  describe("askForOpenAIEnvironmentVariables", () => {
    let envOutput: Map<string, string>;
    const mockedDriverContext = {
      logProvider: new MockedLogProvider(),
      projectPath: "/path/to/project",
      ui: new MockedUserInteraction(),
    } as any;

    beforeEach(() => {
      envOutput = new Map<string, string>();
    });

    afterEach(() => {
      sinon.restore();
    });

    it("should prompt for AZURE_OPENAI_API_KEY and update envOutput", async () => {
      const args = {
        envs: {
          AZURE_OPENAI_API_KEY: "${{ AZURE_OPENAI_API_KEY }}",
        },
      };
      // (mockedDriverContext.ui!.inputText as sinon.SinonStub).resolves(ok({ result: "fakeApiKey" }));
      sinon.stub(mockedDriverContext.ui!, "inputText").resolves(ok({ result: "fakeApiKey" }));

      const result = await driver.askForOpenAIEnvironmentVariables(
        mockedDriverContext,
        args,
        envOutput
      );

      chai.assert(result.isOk());
      chai.assert.equal(envOutput.get("AZURE_OPENAI_API_KEY"), "fakeApiKey");
      chai.assert.equal(args.envs["AZURE_OPENAI_API_KEY"], "fakeApiKey");
    });

    it("should prompt for AZURE_OPENAI_ENDPOINT and update envOutput", async () => {
      const args = {
        envs: {
          AZURE_OPENAI_ENDPOINT: "${{ AZURE_OPENAI_ENDPOINT }}",
        },
      };
      sinon
        .stub(mockedDriverContext.ui!, "inputText")
        .resolves(ok({ result: "https://fakeEndpoint" }));

      const result = await driver.askForOpenAIEnvironmentVariables(
        mockedDriverContext,
        args,
        envOutput
      );

      chai.assert(result.isOk());
      chai.assert.equal(envOutput.get("AZURE_OPENAI_ENDPOINT"), "https://fakeEndpoint");
      chai.assert.equal(args.envs["AZURE_OPENAI_ENDPOINT"], "https://fakeEndpoint");
    });

    it("should prompt for AZURE_OPENAI_DEPLOYMENT_NAME and update envOutput", async () => {
      const args = {
        envs: {
          AZURE_OPENAI_DEPLOYMENT_NAME: "${{ AZURE_OPENAI_DEPLOYMENT_NAME }}",
        },
      };
      sinon
        .stub(mockedDriverContext.ui!, "inputText")
        .resolves(ok({ result: "fakeDeploymentName" }));

      const result = await driver.askForOpenAIEnvironmentVariables(
        mockedDriverContext,
        args,
        envOutput
      );

      chai.assert(result.isOk());
      chai.assert.equal(envOutput.get("AZURE_OPENAI_DEPLOYMENT_NAME"), "fakeDeploymentName");
      chai.assert.equal(args.envs["AZURE_OPENAI_DEPLOYMENT_NAME"], "fakeDeploymentName");
    });

    it("should prompt for OPENAI_API_KEY and update envOutput", async () => {
      const args = {
        envs: {
          OPENAI_API_KEY: "${{ OPENAI_API_KEY }}",
        },
      };
      sinon.stub(mockedDriverContext.ui!, "inputText").resolves(ok({ result: "fakeOpenAIKey" }));

      const result = await driver.askForOpenAIEnvironmentVariables(
        mockedDriverContext,
        args,
        envOutput
      );

      chai.assert(result.isOk());
      chai.assert.equal(envOutput.get("OPENAI_API_KEY"), "fakeOpenAIKey");
      chai.assert.equal(args.envs["OPENAI_API_KEY"], "fakeOpenAIKey");
    });

    it("should return error if inputText fails", async () => {
      const args = {
        envs: {
          AZURE_OPENAI_API_KEY: "${{ AZURE_OPENAI_API_KEY }}",
        },
      };
      sinon.stub(mockedDriverContext.ui!, "inputText").resolves(err(new UserCancelError()));

      const result = await driver.askForOpenAIEnvironmentVariables(
        mockedDriverContext,
        args,
        envOutput
      );

      chai.assert(result.isErr());
    });

    /*
    it("should prompt for AZURE_OPENAI_ENDPOINT and update envOutput", async () => {
      (ctx.ui!.inputText as sinon.SinonStub).resolves(ok({ result: "https://fakeEndpoint" }));
  
      const result = await driver.askForOpenAIEnvironmentVariables(ctx, args, envOutput);
  
      expect(result.isOk()).to.be.true;
      expect(envOutput.get("AZURE_OPENAI_ENDPOINT")).to.equal("https://fakeEndpoint");
      expect(args.envs[OpenAIEnvironmentVariables.AZURE_OPENAI_ENDPOINT]).to.equal("https://fakeEndpoint");
    });
  
    it("should prompt for AZURE_OPENAI_DEPLOYMENT_NAME and update envOutput", async () => {
      (ctx.ui!.inputText as sinon.SinonStub).resolves(ok({ result: "fakeDeploymentName" }));
  
      const result = await driver.askForOpenAIEnvironmentVariables(ctx, args, envOutput);
  
      expect(result.isOk()).to.be.true;
      expect(envOutput.get("AZURE_OPENAI_DEPLOYMENT_NAME")).to.equal("fakeDeploymentName");
      expect(args.envs[OpenAIEnvironmentVariables.AZURE_OPENAI_DEPLOYMENT_NAME]).to.equal("fakeDeploymentName");
    });
  
    it("should prompt for OPENAI_API_KEY and update envOutput", async () => {
      (ctx.ui!.inputText as sinon.SinonStub).resolves(ok({ result: "fakeOpenAIKey" }));
  
      const result = await driver.askForOpenAIEnvironmentVariables(ctx, args, envOutput);
  
      expect(result.isOk()).to.be.true;
      expect(envOutput.get("OPENAI_API_KEY")).to.equal("fakeOpenAIKey");
      expect(args.envs[OpenAIEnvironmentVariables.OPENAI_API_KEY]).to.equal("fakeOpenAIKey");
    });
  
    it("should return error if inputText fails", async () => {
      (ctx.ui!.inputText as sinon.SinonStub).resolves(err(new FxError("UserCancel", "User canceled the input")));
  
      const result = await driver.askForOpenAIEnvironmentVariables(ctx, args, envOutput);
  
      expect(result.isErr()).to.be.true;
    });
    */
  });
});
