// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  DeclarativeCopilotManifestSchema,
  err,
  Inputs,
  ok,
  Platform,
  TeamsAppManifest,
  UserError,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import yaml from "js-yaml";
import * as path from "path";
import { assert, vi } from "vitest";
import { featureFlagManager, FeatureFlagName } from "../../src/common/featureFlags";
import { setTools } from "../../src/common/globalVars";
import { getLocalizedString } from "../../src/common/localizeUtils";
import { copilotGptManifestUtils } from "../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../../src/component/driver/teamsApp/utils/ManifestUtils";
import { FxCore } from "../../src/core/FxCore";
import { QuestionNames } from "../../src/question/questionNames";
import { validationUtils } from "../../src/ui/validationUtils";
import { MockTools, MockUserInteraction } from "./utils";

const tools = new MockTools();

describe("addSkill", () => {
  beforeEach(() => {
    setTools(tools);
    vi.spyOn(validationUtils, "validateInputs").mockResolvedValue(undefined);
    vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(false);
  });

  afterEach(() => {
    vi.restoreAllMocks();
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
      [QuestionNames.SkillName]: "my-skill",
      [QuestionNames.SkillDescription]: "A test skill",
      [QuestionNames.ExposeToCopilot]: "no",
      ignoreLockByUT: true,
      ...overrides,
    };
  }

  it.each([Platform.CLI, Platform.VSCode])(
    "rejects a preset uppercase name through question validation on %s before mutations",
    async (platform) => {
      vi.mocked(validationUtils.validateInputs).mockRestore();
      const inputs = createBaseInputs({
        platform,
        nonInteractive: true,
        [QuestionNames.SkillName]: "testSkill",
      });
      const ensureDir = vi.spyOn(fs, "ensureDir").mockResolvedValue();
      const writeFile = vi.spyOn(fs, "writeFile").mockResolvedValue();
      const addSkill = vi.spyOn(copilotGptManifestUtils, "addSkill");

      const result = await new FxCore(tools).addSkill(inputs);

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.include(
          result.error.message,
          getLocalizedString("core.addSkillQuestion.name.validation")
        );
      }
      assert.equal(ensureDir.mock.calls.length, 0);
      assert.equal(writeFile.mock.calls.length, 0);
      assert.equal(addSkill.mock.calls.length, 0);
    }
  );

  it.each([undefined, 123, "testSkill", "skill_name", "../outside", "my--skill", "a".repeat(65)])(
    "rejects invalid new name %j before mutations even when question validation is skipped",
    async (name) => {
      const inputs = createBaseInputs({ nonInteractive: true, [QuestionNames.SkillName]: name });
      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(createManifestWithDA()));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
      const ensureDir = vi.spyOn(fs, "ensureDir").mockResolvedValue();
      const writeFile = vi.spyOn(fs, "writeFile").mockResolvedValue();
      const addSkill = vi
        .spyOn(copilotGptManifestUtils, "addSkill")
        .mockResolvedValue(ok({ name: "test-agent", description: "Test agent" }));

      const result = await new FxCore(tools).addSkill(inputs);

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.name, "InvalidSkillName");
        assert.equal(
          result.error.message,
          getLocalizedString("core.addSkillQuestion.name.validation")
        );
      }
      assert.equal(ensureDir.mock.calls.length, 0);
      assert.equal(writeFile.mock.calls.length, 0);
      assert.equal(addSkill.mock.calls.length, 0);
    }
  );

  it.each([
    { enabled: true, name: "my-skill" },
    { enabled: false, name: "my-skill" },
    { enabled: false, name: "1skill" },
    { enabled: false, name: "123" },
    { enabled: false, name: "true" },
    { enabled: false, name: "null" },
    { enabled: false, name: "a".repeat(64) },
  ])("creates SKILL.md for $name when Frontier is $enabled", async ({ enabled, name }) => {
    vi.mocked(featureFlagManager.getBooleanValue).mockImplementation(
      (flag) => flag.name === FeatureFlagName.Frontier && enabled
    );
    const inputs = createBaseInputs({ [QuestionNames.SkillName]: name });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));

    const ensureDirStub = vi.spyOn(fs, "ensureDir").mockResolvedValue();
    const writeFileStub = vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    assert.equal(ensureDirStub.mock.calls.length, 1);
    assert.equal(writeFileStub.mock.calls.length, 1);

    // Verify SKILL.md content
    const writtenContent = writeFileStub.mock.calls[0][1] as string;
    const frontmatter = writtenContent.split("---")[1];
    assert.deepEqual(yaml.load(frontmatter), { name, description: "A test skill" });
    assert.include(writtenContent, "description: A test skill");
    assert.include(writtenContent, "---");
  });

  it("successfully adds an existing skill from within appPackage", async () => {
    const appPackageFolder = path.resolve("test-project", "appPackage");
    const inputs = createBaseInputs({
      [QuestionNames.SkillFrom]: "skills/existing-skill",
    });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockResolvedValue(
      "---\nname: existing-skill\ndescription: A skill\n---\n# existing-skill\n" as any
    );

    const addSkillStub = vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    assert.equal(addSkillStub.mock.calls.length, 1);
  });

  it.each(["my.skill", "testSkill", "my--skill"])(
    "existing skill: rejects invalid folder name %j before updating manifests",
    async (folderName) => {
      const appPackageFolder = path.resolve("test-project", "appPackage");
      const inputs = createBaseInputs({
        [QuestionNames.SkillFrom]: `skills/${folderName}`,
      });
      const manifest = createManifestWithDA();

      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
      const addSkill = vi.spyOn(copilotGptManifestUtils, "addSkill");

      const core = new FxCore(tools);
      const result = await core.addSkill(inputs);

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.name, "InvalidSkillFolderName");
      }
      assert.equal(addSkill.mock.calls.length, 0);
    }
  );

  it("existing skill: errors when SKILL.md name doesn't match folder name", async () => {
    const appPackageFolder = path.resolve("test-project", "appPackage");
    const inputs = createBaseInputs({
      [QuestionNames.SkillFrom]: "skills/my-skill",
    });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
    );
    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockResolvedValue(
      "---\nname: differentName\ndescription: A skill\n---\n" as any
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "SkillNameMismatch");
    }
  });

  it("updates DA manifest with agent_skills entry", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    vi.spyOn(fs, "writeFile").mockResolvedValue();

    const addSkillStub = vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    assert.equal(addSkillStub.mock.calls.length, 1);
    // Verify the folder argument passed to addSkill
    const folderArg = addSkillStub.mock.calls[0][1] as string;
    assert.include(folderArg, "skills/my-skill");
  });

  it("errors when project has no DA manifest", async () => {
    const inputs = createBaseInputs();
    const manifest = new TeamsAppManifest();
    // No copilotAgents/declarativeAgents set

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("existing skill: errors when SKILL.md doesn't exist", async () => {
    const appPackageFolder = path.resolve("test-project", "appPackage");
    const inputs = createBaseInputs({
      [QuestionNames.SkillFrom]: "skills/no-skill-md",
    });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
    );

    // SKILL.md does not exist
    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "pathExists").mockResolvedValue(false);

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "SkillMdNotFound");
    }
  });

  it("errors when reading manifest fails", async () => {
    const inputs = createBaseInputs();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(
      err(new UserError("test", "ManifestReadError", "Failed to read"))
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("errors when getManifestPath fails", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      err(new UserError("test", "PathError", "Cannot get path"))
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("errors when copilotGptManifestUtils.addSkill fails", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      err(new UserError("test", "AddSkillError", "Failed to add"))
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("shows success message for CLI platform", async () => {
    const inputs = createBaseInputs({ platform: Platform.CLI });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
  });

  it("errors when user cancels confirmation", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    // User dismisses the confirm dialog (returns undefined)
    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok(undefined));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "UserCancel");
    }
  });

  it("works with copilotExtensions manifest format", async () => {
    const inputs = createBaseInputs();
    const manifest = new TeamsAppManifest();
    manifest.copilotExtensions = {
      declarativeCopilots: [
        {
          id: "agent_1",
          file: "declarativeAgent.json",
        },
      ],
    };

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
  });

  it("existing skill: errors when skill is outside appPackage", async () => {
    const appPackageFolder = path.resolve("test-project", "appPackage");
    const inputs = createBaseInputs({
      [QuestionNames.SkillFrom]: "../../outside-folder",
    });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
    );
    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "SkillOutsideAppPackage");
    }
  });

  it("existing skill: succeeds when SKILL.md has no name frontmatter", async () => {
    const appPackageFolder = path.resolve("test-project", "appPackage");
    const inputs = createBaseInputs({
      [QuestionNames.SkillFrom]: "skills/my-skill",
    });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
    );
    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    // SKILL.md without name in frontmatter
    vi.spyOn(fs, "readFile").mockResolvedValue(
      "# Just a markdown file\nNo frontmatter here.\n" as any
    );

    const addSkillStub = vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    assert.equal(addSkillStub.mock.calls.length, 1);
  });

  it("errors when showMessage returns error", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(
      err(new UserError("test", "ShowMessageError", "Dialog error"))
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "ShowMessageError");
    }
  });

  it("errors when DA has copilotAgents but no file property", async () => {
    const inputs = createBaseInputs();
    const manifest = new TeamsAppManifest();
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          id: "agent_1",
          file: "",
        },
      ],
    };

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isErr());
  });

  it("VSCode: opens agent manifest when user clicks View button", async () => {
    const inputs = createBaseInputs();
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    // First call: confirmation dialog returns "Add"
    // Second call: success message returns "View agent manifest" button
    const showMessageStub = vi.spyOn(MockUserInteraction.prototype, "showMessage");
    showMessageStub.mockResolvedValueOnce(ok("Add"));
    showMessageStub.mockResolvedValueOnce(ok("View agent manifest"));

    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const openFileStub = vi.spyOn(tools.ui, "openFile").mockResolvedValue(ok(true));

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    // Wait for fire-and-forget .then() to resolve
    await new Promise((resolve) => setTimeout(resolve, 10));
    assert.equal(openFileStub.mock.calls.length, 1);
  });

  it("CLI platform: shows success message without View button", async () => {
    const inputs = createBaseInputs({ platform: Platform.CLI });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    const showMessageStub = vi.spyOn(MockUserInteraction.prototype, "showMessage");
    showMessageStub.mockResolvedValueOnce(ok("Add"));
    showMessageStub.mockResolvedValueOnce(ok(undefined));

    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    // Second showMessage should be "info" for CLI, no "View agent manifest" button
    assert.isTrue(showMessageStub.mock.calls[1][0] === "info");
  });

  it("expose-to-copilot: writes agentSkills to Teams manifest when enabled", async () => {
    const inputs = createBaseInputs({
      [QuestionNames.ExposeToCopilot]: "yes",
    });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    const writeFileStub = vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    // writeFile should be called twice: once for SKILL.md, once for Teams manifest
    assert.isTrue(writeFileStub.mock.calls.length >= 2);

    // Find the Teams manifest write call
    const teamsManifestWriteCall = writeFileStub.mock.calls.find((call) => {
      const filePath = call[0] as string;
      return filePath.includes("manifest.json");
    });
    assert.isDefined(teamsManifestWriteCall);
    const writtenManifest = JSON.parse(teamsManifestWriteCall![1] as string);
    assert.isArray(writtenManifest.agentSkills);
    assert.equal(writtenManifest.agentSkills.length, 1);
    assert.include(writtenManifest.agentSkills[0].folder, "skills/my-skill");
  });

  it("expose-to-copilot: writes agentSkills when enabled with boolean true", async () => {
    const inputs = createBaseInputs({
      [QuestionNames.ExposeToCopilot]: true,
    });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    const writeFileStub = vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    const teamsManifestWriteCall = writeFileStub.mock.calls.find((call) => {
      const filePath = call[0] as string;
      return filePath.includes("manifest.json");
    });
    assert.isDefined(teamsManifestWriteCall);
    const writtenManifest = JSON.parse(teamsManifestWriteCall![1] as string);
    assert.isArray(writtenManifest.agentSkills);
    assert.equal(writtenManifest.agentSkills.length, 1);
    assert.include(writtenManifest.agentSkills[0].folder, "skills/my-skill");
  });

  it("expose-to-copilot: does not duplicate existing agentSkills entries", async () => {
    const inputs = createBaseInputs({
      [QuestionNames.ExposeToCopilot]: "yes",
    });
    const manifest = createManifestWithDA();
    manifest.agentSkills = [{ folder: "skills/my-skill" }];

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    const writeFileStub = vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    const teamsManifestWriteCall = writeFileStub.mock.calls.find((call) => {
      const filePath = call[0] as string;
      return filePath.includes("manifest.json");
    });
    assert.isDefined(teamsManifestWriteCall);
    const writtenManifest = JSON.parse(teamsManifestWriteCall![1] as string);
    assert.isArray(writtenManifest.agentSkills);
    assert.equal(writtenManifest.agentSkills.length, 1);
    assert.equal(writtenManifest.agentSkills[0].folder, "skills/my-skill");
  });

  it("expose-to-copilot: does NOT write agentSkills when disabled", async () => {
    const inputs = createBaseInputs({
      [QuestionNames.ExposeToCopilot]: "no",
    });
    const manifest = createManifestWithDA();

    vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.resolve("test-project", "appPackage", "declarativeAgent.json"))
    );

    vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    const writeFileStub = vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
      ok({
        name: "test-agent",
        description: "description",
      } as DeclarativeCopilotManifestSchema)
    );

    const core = new FxCore(tools);
    const result = await core.addSkill(inputs);

    assert.isTrue(result.isOk());
    // writeFile should only be called for SKILL.md, NOT Teams manifest
    const teamsManifestWriteCall = writeFileStub.mock.calls.find((call) => {
      const filePath = call[0] as string;
      return filePath.includes("manifest.json");
    });
    assert.isUndefined(teamsManifestWriteCall);
    // Verify Teams manifest has no agentSkills
    assert.isUndefined(manifest.agentSkills);
  });

  describe("zip import", () => {
    let AdmZipModule: typeof import("adm-zip");

    before(async () => {
      AdmZipModule = (await import("adm-zip")).default;
    });

    function createZipWithSingleFolder(
      folderName: string,
      skillMdContent: string,
      extraFiles?: { name: string; content: string }[]
    ): Buffer {
      const AdmZip = require("adm-zip");
      const zip = new AdmZip();
      zip.addFile(`${folderName}/SKILL.md`, Buffer.from(skillMdContent, "utf-8"));
      if (extraFiles) {
        for (const f of extraFiles) {
          zip.addFile(`${folderName}/${f.name}`, Buffer.from(f.content, "utf-8"));
        }
      }
      return zip.toBuffer();
    }

    function createZipWithRootFiles(
      skillMdContent: string,
      extraFiles?: { name: string; content: string }[]
    ): Buffer {
      const AdmZip = require("adm-zip");
      const zip = new AdmZip();
      zip.addFile("SKILL.md", Buffer.from(skillMdContent, "utf-8"));
      if (extraFiles) {
        for (const f of extraFiles) {
          zip.addFile(f.name, Buffer.from(f.content, "utf-8"));
        }
      }
      return zip.toBuffer();
    }

    function createZipInputs(zipPath: string): Inputs {
      return createBaseInputs({
        [QuestionNames.SkillFrom]: zipPath,
      });
    }

    it("successfully imports skill from single-folder zip via --from", async () => {
      const skillMd =
        "---\nname: my-imported-skill\ndescription: An imported skill\n---\n# my-imported-skill\n";
      const zipBuffer = createZipWithSingleFolder("my-imported-skill", skillMd);
      const zipPath = path.resolve("test-skill.zip");
      const appPackageFolder = path.resolve("test-project", "appPackage");

      const inputs = createZipInputs(zipPath);
      const manifest = createManifestWithDA();

      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));

      // Mock fs operations for zip import
      const pathExistsStub = vi.spyOn(fs, "pathExists");
      pathExistsStub.mockImplementation(async (p: string) => {
        if (p === zipPath) return true;
        if (p.includes("my-imported-skill") && p.includes("skills")) return false;
        if (p.includes("SKILL.md")) return true;
        return false;
      });

      // Stub AdmZip constructor
      const AdmZip = require("adm-zip");
      const fakeZip = new AdmZip(zipBuffer);
      vi.spyOn(FxCore.prototype as any, "importSkillFromZip").mockImplementation(async function (
        this: any,
        ...args: unknown[]
      ) {
        // Return a successful result mimicking the zip import
        return ok("skills/my-imported-skill");
      });

      vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
        ok({
          name: "test-agent",
          description: "description",
        } as DeclarativeCopilotManifestSchema)
      );

      const core = new FxCore(tools);
      const result = await core.addSkill(inputs);

      assert.isTrue(result.isOk());
    });

    it("successfully imports skill from zip via skill-from-zip-file input", async () => {
      const zipPath = path.resolve("test-skill.zip");
      const appPackageFolder = path.resolve("test-project", "appPackage");

      const inputs = createBaseInputs({
        [QuestionNames.SkillFromZipFile]: zipPath,
      });
      const manifest = createManifestWithDA();

      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));

      vi.spyOn(FxCore.prototype as any, "importSkillFromZip").mockImplementation(async function () {
        return ok("skills/my-imported-skill");
      });

      vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
        ok({
          name: "test-agent",
          description: "description",
        } as DeclarativeCopilotManifestSchema)
      );

      const core = new FxCore(tools);
      const result = await core.addSkill(inputs);

      assert.isTrue(result.isOk());
    });

    it("errors when zip file does not exist", async () => {
      const zipPath = path.resolve("nonexistent.zip");
      const appPackageFolder = path.resolve("test-project", "appPackage");

      const inputs = createZipInputs(zipPath);
      const manifest = createManifestWithDA();

      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
      vi.spyOn(fs, "pathExists").mockResolvedValue(false);

      const core = new FxCore(tools);
      const result = await core.addSkill(inputs);

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.name, "ZipFileNotFound");
      }
    });

    it("errors when zip contains path traversal entries", async () => {
      const AdmZip = require("adm-zip");

      // AdmZip normalizes "../" away during addFile, so we binary-patch the zip
      // to inject a real traversal entry that survives round-tripping.
      const traversalPath = "../../../etc/passwd"; // 19 chars
      const placeholder = "AAAAAAAAAAAAAAAAAAA"; // 19 chars (same length)

      const zip = new AdmZip();
      zip.addFile(placeholder, Buffer.from("malicious", "utf-8"));
      const buf = zip.toBuffer();

      // Patch all occurrences of the placeholder with the traversal path
      const patched = Buffer.from(buf);
      let idx: number;
      while ((idx = patched.indexOf(placeholder)) !== -1) {
        patched.write(traversalPath, idx, "utf-8");
      }

      const tempZipPath = path.join(require("os").tmpdir(), `test-traversal-${Date.now()}.zip`);
      await fs.writeFile(tempZipPath, patched);

      const appPackageFolder = path.resolve("test-project", "appPackage");
      const traversalInputs = createZipInputs(tempZipPath);
      const manifest = createManifestWithDA();

      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));

      const core = new FxCore(tools);
      const result = await core.addSkill(traversalInputs);

      // Clean up
      await fs.remove(tempZipPath).catch(() => {});

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.name, "ZipInvalidEntries");
      }
    });

    it("errors when zip has no SKILL.md (root-level layout)", async () => {
      const AdmZip = require("adm-zip");
      const zip = new AdmZip();
      zip.addFile("readme.txt", Buffer.from("no skill md here", "utf-8"));
      const zipBuffer = zip.toBuffer();

      const tempZipPath = path.join(require("os").tmpdir(), `test-noskillmd-${Date.now()}.zip`);
      zip.writeZip(tempZipPath);

      const inputs = createZipInputs(tempZipPath);
      const manifest = createManifestWithDA();
      const appPackageFolder = path.resolve("test-project", "appPackage");

      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));

      const core = new FxCore(tools);
      const result = await core.addSkill(inputs);

      await fs.remove(tempZipPath).catch(() => {});

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.name, "ZipNoSkillMd");
      }
    });

    it("errors when target skill folder already exists", async () => {
      const skillMd = "---\nname: existing-skill\ndescription: A skill\n---\n# existing-skill\n";
      const AdmZip = require("adm-zip");
      const zip = new AdmZip();
      zip.addFile("existing-skill/SKILL.md", Buffer.from(skillMd, "utf-8"));
      const tempZipPath = path.join(require("os").tmpdir(), `test-existing-${Date.now()}.zip`);
      zip.writeZip(tempZipPath);

      const inputs = createZipInputs(tempZipPath);
      const manifest = createManifestWithDA();
      const appPackageFolder = path.resolve("test-project", "appPackage");

      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));

      // Target folder already exists
      const origPathExists = fs.pathExists;
      vi.spyOn(fs, "pathExists").mockImplementation(async (p: string) => {
        if (p === tempZipPath) return true;
        if (
          p === path.join(appPackageFolder, "skills", "existing-skill") ||
          p.endsWith(path.join("skills", "existing-skill"))
        ) {
          return true;
        }
        return false;
      });

      const core = new FxCore(tools);
      const result = await core.addSkill(inputs);

      await fs.remove(tempZipPath).catch(() => {});

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.name, "SkillFolderAlreadyExists");
      }
    });

    it("successfully imports single-folder zip end-to-end", async () => {
      const skillMd = "---\nname: my-new-skill\ndescription: A new skill\n---\n# my-new-skill\n";
      const AdmZip = require("adm-zip");
      const zip = new AdmZip();
      zip.addFile("my-new-skill/SKILL.md", Buffer.from(skillMd, "utf-8"));
      zip.addFile("my-new-skill/extra.txt", Buffer.from("extra content", "utf-8"));

      const tempZipPath = path.join(require("os").tmpdir(), `test-e2e-${Date.now()}.zip`);
      zip.writeZip(tempZipPath);

      const inputs = createZipInputs(tempZipPath);
      const appPackageFolder = path.resolve("test-project", "appPackage");
      const manifest = createManifestWithDA();

      vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
      vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
        ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
      );
      vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));

      // Mock fs for the target side but allow temp operations
      const targetSkillDir = path.join(appPackageFolder, "skills", "my-new-skill");

      const origPathExists = fs.pathExists.bind(fs);
      vi.spyOn(fs, "pathExists").mockImplementation(async (p: string) => {
        if (p === tempZipPath) return true;
        if (typeof p === "string" && p.includes("my-new-skill") && p.includes(appPackageFolder)) {
          return false;
        }
        if (typeof p === "string" && p.includes("SKILL.md") && !p.includes(appPackageFolder)) {
          return true;
        }
        return origPathExists(p);
      });

      const ensureDirStub = vi.spyOn(fs, "ensureDir").mockResolvedValue();
      const writeFileStub = vi.spyOn(fs, "writeFile").mockResolvedValue();
      const moveStub = vi.spyOn(fs, "move").mockResolvedValue();
      const readFileStub = vi.spyOn(fs, "readFile").mockResolvedValue(skillMd as any);
      vi.spyOn(fs, "remove").mockResolvedValue();

      vi.spyOn(copilotGptManifestUtils, "addSkill").mockResolvedValue(
        ok({
          name: "test-agent",
          description: "description",
        } as DeclarativeCopilotManifestSchema)
      );

      const core = new FxCore(tools);
      const result = await core.addSkill(inputs);

      await fs.remove(tempZipPath).catch(() => {});

      assert.isTrue(result.isOk());
    });

    it.each(["root", "folder"])(
      "rejects uppercase skill names in %s ZIP layouts before extraction",
      async (layout) => {
        const skillMd = "---\nname: testSkill\ndescription: Bad name\n---\n# testSkill\n";
        const AdmZip = require("adm-zip");
        const zip = new AdmZip();
        zip.addFile(
          layout === "root" ? "SKILL.md" : "testSkill/SKILL.md",
          Buffer.from(skillMd, "utf-8")
        );

        const tempZipPath = path.join(require("os").tmpdir(), `test-badname-${Date.now()}.zip`);
        zip.writeZip(tempZipPath);

        const inputs = createZipInputs(tempZipPath);
        const manifest = createManifestWithDA();
        const appPackageFolder = path.resolve("test-project", "appPackage");

        vi.spyOn(manifestUtils, "_readAppManifest").mockResolvedValue(ok(manifest));
        vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
          ok(path.resolve(appPackageFolder, "declarativeAgent.json"))
        );
        vi.spyOn(MockUserInteraction.prototype, "showMessage").mockResolvedValue(ok("Add"));
        const mkdtemp = vi.spyOn(fs, "mkdtemp");
        const addSkill = vi
          .spyOn(copilotGptManifestUtils, "addSkill")
          .mockResolvedValue(ok({ name: "test-agent", description: "Test agent" }));

        try {
          const result = await new FxCore(tools).addSkill(inputs);
          assert.isTrue(result.isErr());
          if (result.isErr()) {
            assert.equal(result.error.name, "InvalidSkillFolderName");
            assert.equal(
              result.error.message,
              getLocalizedString("core.addSkillQuestion.name.validation")
            );
          }
          assert.equal(mkdtemp.mock.calls.length, 0);
          assert.equal(addSkill.mock.calls.length, 0);
        } finally {
          await fs.remove(tempZipPath);
        }
      }
    );
  });
});
