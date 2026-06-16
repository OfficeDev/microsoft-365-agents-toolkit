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
import sinon from "sinon";
import { FxCore, pathUtils, UserCancelError } from "../../src";
import { setTools } from "../../src/common/globalVars";
import { coordinator } from "../../src/component/coordinator";
import { QuestionNames } from "../../src/question/constants";
import { MockTools } from "./utils";

describe("FxCore.createProject", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  beforeEach(() => {});
  afterEach(() => {
    sandbox.restore();
  });

  it("happy path", async () => {
    const core = new FxCore(tools);
    assert.isFunction(core.createProject);
  });

  it("create teams agent with key", async () => {
    const core = new FxCore(tools);
    assert.isFunction(core.createProject);
  });

  it("create teams agent without AI key", async () => {
    const core = new FxCore(tools);
    assert.isFunction(core.createProject);
  });

  it("create teams agent without AI endpoint", async () => {
    const core = new FxCore(tools);
    assert.isFunction(core.createProject);
  });

  it("startWithGithubCopilot", async () => {
    const core = new FxCore(tools);
    assert.isFunction(core.createProject);
  });

  it("coordinator error", async () => {
    const core = new FxCore(tools);
    assert.isFunction(core.createProject);
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
    const core = new FxCore(tools);
    assert.isFunction(core.createProjectFromTdp);
  });

  it("happy", async () => {
    const core = new FxCore(tools);
    assert.isFunction(core.createProjectFromTdp);
  });
});

describe("metaOSExtendToDA", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  beforeEach(() => {});
  afterEach(() => {
    sandbox.restore();
  });

  it("happy path", async () => {
    sandbox.stub(coordinator, "create").resolves(ok({ projectPath: "path" }));
    sandbox.stub(tools, "logProvider").value(undefined);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: "path",
      [QuestionNames.AppName]: "abc",
    };
    const core = new FxCore(tools);
    const res = await core.metaOSExtendToDA(inputs, "path");
    assert.isTrue(res.isOk());
  });

  it("happy path: coordinator error", async () => {
    sandbox.stub(coordinator, "create").resolves(err(new UserError({})));
    sandbox.stub(tools, "logProvider").value(undefined);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [QuestionNames.Folder]: "path",
      [QuestionNames.AppName]: "abc",
    };
    const core = new FxCore(tools);
    const res = await core.metaOSExtendToDA(inputs, "path");
    assert.isTrue(res.isErr());
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
    sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
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
