// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  DeclarativeCopilotManifestSchema,
  Inputs,
  Platform,
  TeamsAppManifest,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import * as path from "path";
import sinon from "sinon";
import { FxCore } from "../../src/core/FxCore";
import { setTools } from "../../src/common/globalVars";
import { copilotGptManifestUtils } from "../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../../src/component/driver/teamsApp/utils/ManifestUtils";
import { UserCancelError } from "../../src/error/common";
import { QuestionNames } from "../../src/question/questionNames";
import { validationUtils } from "../../src/ui/validationUtils";
import { MockTools, MockUserInteraction } from "./utils";

const tools = new MockTools();

describe("addSkill", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    setTools(tools);
    sandbox.stub(validationUtils, "validateInputs").resolves(undefined);
  });

  afterEach(() => {
    sandbox.restore();
  });

  function createManifestWithDA(): TeamsAppManifest {
    const manifest = new TeamsAppManifest();
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          id: "agent_1",
          file: "declarativeAgent.json",
        },
      ],
    };
    return manifest;
  }

  function createBaseInputs(overrides?: Partial<Inputs>): Inputs {
    return {
      platform: Platform.VSCode,
      projectPath: path.resolve("test-project"),
      [QuestionNames.ManifestPath]: path.resolve("test-project", "appPackage", "manifest.json"),
      [QuestionNames.SkillName]: "mySkill",
      [QuestionNames.SkillDescription]: "A test skill",
      [QuestionNames.SkillExposeTocopilot]: "no",
      ignoreLockByUT: true,
      ...overrides,
    };
  }

  it("successfully creates a new skill directory with SKILL.md", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(
        ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
      );

    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("OK"));

    const ensureDirStub = sandbox.stub(fs, "ensureDir").resolves();
    const writeFileStub = sandbox.stub(fs, "writeFile").resolves();

    sandbox.stub(copilotGptManifestUtils, "addSkill").resolves(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(ensureDirStub.calledOnce);
    assert.isTrue(writeFileStub.calledOnce);

    // Verify SKILL.md content
    const writtenContent = writeFileStub.firstCall.args[1] as string;
    assert.include(writtenContent, "name: mySkill");
    assert.include(writtenContent, "description: A test skill");
    assert.include(writtenContent, "---");
  });

  it("successfully adds an existing skill from within appPackage", async () => {
    const appPackageFolder = path.resolve("test-project", "appPackage");
    const inputs = createBaseInputs({
      [QuestionNames.SkillFrom]: "skills/existingSkill",
    });
    const manifest = createManifestWithDA();

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(ok(path.resolve(appPackageFolder, "declarativeAgent.json")));

    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("OK"));
    sandbox.stub(fs, "pathExists").resolves(true);

    const addSkillStub = sandbox.stub(copilotGptManifestUtils, "addSkill").resolves(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(addSkillStub.calledOnce);
  });

  it("updates DA manifest with agent_skills entry", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(
        ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
      );

    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("OK"));
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();

    const addSkillStub = sandbox.stub(copilotGptManifestUtils, "addSkill").resolves(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(addSkillStub.calledOnce);
    // Verify the folder argument passed to addSkill
    const folderArg = addSkillStub.firstCall.args[1] as string;
    assert.include(folderArg, "skills/mySkill");
  });

  it("errors when project has no DA manifest", async () => {
    const inputs = createBaseInputs();
    const manifest = new TeamsAppManifest();
    // No copilotAgents/declarativeAgents set

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("existing skill: errors when SKILL.md doesn't exist", async () => {
    const appPackageFolder = path.resolve("test-project", "appPackage");
    const inputs = createBaseInputs({
      [QuestionNames.SkillFrom]: "skills/noSkillMd",
    });
    const manifest = createManifestWithDA();

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(ok(path.resolve(appPackageFolder, "declarativeAgent.json")));

    // SKILL.md does not exist
    sandbox.stub(fs, "pathExists").resolves(false);

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "SkillMdNotFound");
    }
  });

  it("errors when reading manifest fails", async () => {
    const inputs = createBaseInputs();

    sandbox
      .stub(manifestUtils, "_readAppManifest")
      .resolves(err(new UserError("test", "ManifestReadError", "Failed to read")));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("errors when getManifestPath fails", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(err(new UserError("test", "PathError", "Cannot get path")));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("errors when copilotGptManifestUtils.addSkill fails", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(
        ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
      );

    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();

    sandbox
      .stub(copilotGptManifestUtils, "addSkill")
      .resolves(err(new UserError("test", "AddSkillError", "Failed to add")));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("shows success message for CLI platform", async () => {
    const inputs = createBaseInputs({ platform: Platform.CLI });
    const manifest = createManifestWithDA();

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(
        ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
      );

    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("OK"));
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();

    sandbox.stub(copilotGptManifestUtils, "addSkill").resolves(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
  });

  it("passes exposeSkillToCopilot=true when user selects yes", async () => {
    const inputs = createBaseInputs({
      [QuestionNames.SkillExposeTocopilot]: "yes",
    });
    const manifest = createManifestWithDA();

    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(copilotGptManifestUtils, "getManifestPath")
      .resolves(
        ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
      );

    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("OK"));
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();

    const addSkillStub = sandbox.stub(copilotGptManifestUtils, "addSkill").resolves(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    assert.isTrue(addSkillStub.calledOnce);
    const exposeArg = addSkillStub.firstCall.args[2];
    assert.isTrue(exposeArg);
  });
});
