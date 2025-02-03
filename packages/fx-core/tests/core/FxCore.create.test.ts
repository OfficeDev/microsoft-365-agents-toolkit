// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Context,
  CreateProjectInputs,
  err,
  FxError,
  GeneratorResult,
  IGenerator,
  Inputs,
  ok,
  Platform,
  Result,
  UserError,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import * as os from "os";
import sinon from "sinon";
import { AppDefinition, FxCore, InputValidationError, UserCancelError } from "../../src";
import { coordinator } from "../../src/component/coordinator";
import { setTools } from "../../src/common/globalVars";
import {
  CapabilityOptions,
  ProjectTypeOptions,
  QuestionNames,
  ScratchOptions,
} from "../../src/question/constants";
import { MockTools, randomAppName } from "./utils";

describe("FxCore.createProject", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  beforeEach(() => {});
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path", async () => {
    sandbox.stub(coordinator, "create").resolves(ok({ projectPath: "" }));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
      [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AppName]: randomAppName(),
    };
    const core = new FxCore(tools);
    const res = await core.createProject(inputs);
    assert.isTrue(res.isOk());
  });

  it("coordinator error", async () => {
    sandbox.stub(coordinator, "create").resolves(err(new UserError({})));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
      [QuestionNames.Capabilities]: CapabilityOptions.tab().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AppName]: randomAppName(),
    };
    const core = new FxCore(tools);
    const res = await core.createProject(inputs);
    assert.isTrue(res.isErr());
  });
});

describe("createProjectFromTdp", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  beforeEach(() => {});
  afterEach(() => {
    sandbox.restore();
  });
  it("TDP input error", async () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "mock-id",
      appId: "mock-id",
      staticTabs: [
        {
          name: "tab1",
          entityId: "tab1",
          contentUrl: "mock-contentUrl",
          websiteUrl: "mock-websiteUrl",
          context: [],
          scopes: [],
        },
      ],
      bots: [
        {
          botId: "mock-bot-id",
          isNotificationOnly: false,
          needsChannelSelector: false,
          supportsCalling: false,
          supportsFiles: false,
          supportsVideo: false,
          scopes: [],
          teamCommands: [],
          groupChatCommands: [],
          personalCommands: [],
        },
      ],
      connectors: [
        {
          name: "connector1",
          configurationUrl: "https://test.com",
          scopes: [],
        },
      ],
    };
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
      [QuestionNames.Capabilities]: CapabilityOptions.tab().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AppName]: randomAppName(),
      teamsAppFromTdp: appDefinition,
    };
    const core = new FxCore(tools);
    sandbox.stub(tools.ui, "selectOptions").resolves(ok({ type: "success", result: [] }));
    const res = await core.createProjectFromTdp(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof InputValidationError);
    }
  });

  it("happy", async () => {
    const appDefinition: AppDefinition = {
      teamsAppId: "mock-id",
      appId: "mock-id",
      staticTabs: [
        {
          name: "tab1",
          entityId: "tab1",
          contentUrl: "mock-contentUrl",
          websiteUrl: "mock-websiteUrl",
          context: [],
          scopes: [],
        },
      ],
      bots: [
        {
          botId: "mock-bot-id",
          isNotificationOnly: false,
          needsChannelSelector: false,
          supportsCalling: false,
          supportsFiles: false,
          supportsVideo: false,
          scopes: [],
          teamCommands: [],
          groupChatCommands: [],
          personalCommands: [],
        },
      ],
    };
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Scratch]: ScratchOptions.yes().id,
      [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
      [QuestionNames.Capabilities]: CapabilityOptions.tab().id,
      [QuestionNames.ProgrammingLanguage]: "javascript",
      [QuestionNames.Folder]: os.tmpdir(),
      [QuestionNames.AppName]: randomAppName(),
      teamsAppFromTdp: appDefinition,
    };
    const core = new FxCore(tools);
    sandbox.stub(tools.ui, "selectOptions").resolves(ok({ type: "success", result: [] }));
    sandbox.stub(coordinator, "create").resolves(ok({ projectPath: "." }));
    const res = await core.createProjectFromTdp(inputs);
    assert.isTrue(res.isOk());
  });
});

describe("FxCore.createProjectByCustomizedGenerator", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  beforeEach(() => {});
  afterEach(() => {
    sandbox.restore();
  });

  class MyGenerator implements IGenerator {
    componentName = "my-generator";
    async run(
      context: Context,
      inputs: Inputs,
      destinationPath: string
    ): Promise<Result<GeneratorResult, FxError>> {
      return Promise.resolve(ok({}));
    }
  }

  it("happy path", async () => {
    const myGenerator = new MyGenerator();
    sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
    sandbox.stub(fs, "pathExists").resolves(ok("mock-id"));
    const inputs: CreateProjectInputs = {
      platform: Platform.VSCode,
      folder: ".",
      "app-name": "test-app",
    };
    const core = new FxCore(tools);
    const res = await core.createProjectByCustomizedGenerator(inputs, myGenerator);
    assert.isTrue(res.isOk());
  });

  it("folder is empty", async () => {
    const myGenerator = new MyGenerator();
    const inputs: CreateProjectInputs = {
      platform: Platform.VSCode,
      folder: "",
      "app-name": "test-app",
    };
    const core = new FxCore(tools);
    const res = await core.createProjectByCustomizedGenerator(inputs, myGenerator);
    assert.isTrue(res.isErr());
  });

  it("appname is empty", async () => {
    const myGenerator = new MyGenerator();
    const inputs: CreateProjectInputs = {
      platform: Platform.VSCode,
      folder: ".",
      "app-name": "",
    };
    const core = new FxCore(tools);
    const res = await core.createProjectByCustomizedGenerator(inputs, myGenerator);
    assert.isTrue(res.isErr());
  });

  it("app is invalid", async () => {
    const myGenerator = new MyGenerator();
    const inputs: CreateProjectInputs = {
      platform: Platform.VSCode,
      folder: ".",
      "app-name": "123",
    };
    const core = new FxCore(tools);
    const res = await core.createProjectByCustomizedGenerator(inputs, myGenerator);
    assert.isTrue(res.isErr());
  });

  it("generator error", async () => {
    const myGenerator = new MyGenerator();
    sandbox.stub(myGenerator, "run").resolves(err(new UserCancelError()));
    const inputs: CreateProjectInputs = {
      platform: Platform.VSCode,
      folder: ".",
      "app-name": "test-app",
    };
    const core = new FxCore(tools);
    const res = await core.createProjectByCustomizedGenerator(inputs, myGenerator);
    assert.isTrue(res.isErr());
  });

  it("ensureTrackingId error", async () => {
    const myGenerator = new MyGenerator();
    sandbox.stub(coordinator, "ensureTrackingId").resolves(err(new UserCancelError()));
    sandbox.stub(fs, "pathExists").resolves(ok("mock-id"));
    const inputs: CreateProjectInputs = {
      platform: Platform.VSCode,
      folder: ".",
      "app-name": "test-app",
    };
    const core = new FxCore(tools);
    const res = await core.createProjectByCustomizedGenerator(inputs, myGenerator);
    assert.isTrue(res.isErr());
  });
});
