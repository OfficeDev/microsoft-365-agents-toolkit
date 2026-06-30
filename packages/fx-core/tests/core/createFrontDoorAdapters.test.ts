// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Platform, UserError, err, ok } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import { assert, vi } from "vitest";

import { setTools } from "../../src/common/globalVars";
import { coordinator } from "../../src/component/coordinator";
import { pathUtils } from "../../src/component/utils/pathUtils";
import {
  applyV3PreFill,
  collectCreateFloor,
  scaffoldV4,
  scaffoldV4Deps,
} from "../../src/core/createFrontDoorAdapters";
import { QuestionNames } from "../../src/question/constants";
import { BuildTarget, TemplateSource } from "../../src/v4";
import { MockTools } from "./utils";

const TEMPLATE_SOURCE: TemplateSource = {
  origin: "bundled",
  version: "1.0.0",
  digest: "sha256:test",
  location: "test",
};

describe("createFrontDoorAdapters", () => {
  const tools = new MockTools();
  setTools(tools);

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe("applyV3PreFill", () => {
    it("is a no-op when the target carries no Q1 picks (the direct source)", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      const target: BuildTarget = { templateId: "default-bot", engine: "v3" };

      applyV3PreFill(inputs, target);

      assert.deepEqual(inputs, { platform: Platform.VSCode });
    });

    it("is a no-op when the picks carry no projectType", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      const target: BuildTarget = { templateId: "default-bot", engine: "v3", answers: {} };

      applyV3PreFill(inputs, target);

      assert.deepEqual(inputs, { platform: Platform.VSCode });
    });

    it("maps copilot-agent no-action onto the v3 declarative-agent path", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      const target: BuildTarget = {
        templateId: "copilot-gpt-basic",
        engine: "v3",
        answers: { projectType: "copilot-agent-type", daTemplate: "no-action" },
      };

      applyV3PreFill(inputs, target);

      assert.equal(inputs[QuestionNames.ProjectType], "copilot-agent-type");
      assert.equal(inputs[QuestionNames.Capabilities], "declarative-agent");
      assert.equal(inputs[QuestionNames.WithPlugin], "no");
      assert.isUndefined(inputs[QuestionNames.ActionType]);
      assert.isUndefined(inputs[QuestionNames.ApiAuth]);
    });

    it("maps copilot-agent add-action + new-api + auth onto the full v3 path", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      const target: BuildTarget = {
        templateId: "api-plugin-from-scratch",
        engine: "v3",
        answers: {
          projectType: "copilot-agent-type",
          daTemplate: "add-action",
          actionSource: "new-api",
          apiAuth: "api-key",
        },
      };

      applyV3PreFill(inputs, target);

      assert.equal(inputs[QuestionNames.WithPlugin], "yes");
      assert.equal(inputs[QuestionNames.ActionType], "new-api");
      assert.equal(inputs[QuestionNames.ApiAuth], "api-key");
    });

    it("maps the openapi action source onto the v3 api-spec id and drops api-auth", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      const target: BuildTarget = {
        templateId: "api-plugin-existing-api",
        engine: "v3",
        answers: {
          projectType: "copilot-agent-type",
          daTemplate: "add-action",
          actionSource: "openapi",
          apiAuth: "api-key",
        },
      };

      applyV3PreFill(inputs, target);

      assert.equal(inputs[QuestionNames.ActionType], "api-spec");
      // api-auth only applies to the new-api source.
      assert.isUndefined(inputs[QuestionNames.ApiAuth]);
    });

    it("maps the mcp action source onto the v3 mcp id and drops api-auth", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      const target: BuildTarget = {
        templateId: "declarative-agent-with-action-from-mcp",
        engine: "v3",
        answers: {
          projectType: "copilot-agent-type",
          daTemplate: "add-action",
          actionSource: "mcp",
        },
      };

      applyV3PreFill(inputs, target);

      assert.equal(inputs[QuestionNames.ProjectType], "copilot-agent-type");
      assert.equal(inputs[QuestionNames.Capabilities], "declarative-agent");
      assert.equal(inputs[QuestionNames.WithPlugin], "yes");
      assert.equal(inputs[QuestionNames.ActionType], "mcp");
      // api-auth only applies to the new-api source; the v3 walk re-asks the
      // mcp-server-type dimension (the selector does not collect it).
      assert.isUndefined(inputs[QuestionNames.ApiAuth]);
    });

    it("maps the da-meta-os action source onto the v3 da-meta-os id", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      const target: BuildTarget = {
        templateId: "declarative-agent-meta-os-new-project",
        engine: "v3",
        answers: {
          projectType: "copilot-agent-type",
          daTemplate: "add-action",
          actionSource: "da-meta-os",
        },
      };

      applyV3PreFill(inputs, target);

      assert.equal(inputs[QuestionNames.WithPlugin], "yes");
      assert.equal(inputs[QuestionNames.ActionType], "da-meta-os");
      assert.isUndefined(inputs[QuestionNames.ApiAuth]);
    });

    it("maps the graph-connector and typespec da templates", () => {
      const gc: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(gc, {
        templateId: "declarative-agent-with-graph-connector",
        engine: "v3",
        answers: { projectType: "copilot-agent-type", daTemplate: "graph-connector" },
      });
      assert.equal(gc[QuestionNames.WithPlugin], "gc");
      assert.isUndefined(gc[QuestionNames.ActionType]);

      const ts: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(ts, {
        templateId: "declarative-agent-typespec",
        engine: "v3",
        answers: { projectType: "copilot-agent-type", daTemplate: "typespec" },
      });
      assert.equal(ts[QuestionNames.WithPlugin], "type-spec");
    });

    it("maps the skill da template onto WithPlugin so the v3 walk does not re-ask", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(inputs, {
        templateId: "declarative-agent-with-skill",
        engine: "v3",
        answers: { projectType: "copilot-agent-type", daTemplate: "skill" },
      });
      assert.equal(inputs[QuestionNames.Capabilities], "declarative-agent");
      assert.equal(inputs[QuestionNames.WithPlugin], "skill");
      assert.isUndefined(inputs[QuestionNames.ActionType]);
    });

    it("maps the custom-engine-agent capability onto the v3 capabilities question", () => {
      const basic: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(basic, {
        templateId: "basic-custom-engine-agent",
        engine: "v3",
        answers: {
          projectType: "custom-engine-agent-type",
          customEngineAgent: "basic-custom-engine-agent",
        },
      });
      assert.equal(basic[QuestionNames.ProjectType], "custom-engine-agent-type");
      assert.equal(basic[QuestionNames.Capabilities], "basic-custom-engine-agent");

      const weather: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(weather, {
        templateId: "weather-agent",
        engine: "v3",
        answers: { projectType: "custom-engine-agent-type", customEngineAgent: "weather-agent" },
      });
      assert.equal(weather[QuestionNames.Capabilities], "weather-agent");
    });

    it("maps the teams general agent onto the renamed v3 teams-app-type id", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(inputs, {
        templateId: "custom-copilot-basic",
        engine: "v3",
        answers: { projectType: "teams-agent-and-app-type", teamsApp: "custom-copilot-basic" },
      });

      assert.equal(inputs[QuestionNames.ProjectType], "teams-agent-and-app-type");
      assert.equal(inputs[QuestionNames.TeamsAppType], "custom-copilot-basic");
      assert.isUndefined(inputs[QuestionNames.CustomCopilotRag]);
    });

    it("renames the teams rag app and maps the rag source verbatim", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(inputs, {
        templateId: "custom-copilot-rag-custom-api",
        engine: "v3",
        answers: {
          projectType: "teams-agent-and-app-type",
          teamsApp: "rag",
          customCopilotRagType: "custom-copilot-rag-custom-api",
        },
      });

      assert.equal(inputs[QuestionNames.TeamsAppType], "custom-copilot-rag");
      assert.equal(inputs[QuestionNames.CustomCopilotRag], "custom-copilot-rag-custom-api");
    });

    it("maps the teams collaborator agent", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(inputs, {
        templateId: "teams-collaborator-agent",
        engine: "v3",
        answers: { projectType: "teams-agent-and-app-type", teamsApp: "teams-collaborator-agent" },
      });

      assert.equal(inputs[QuestionNames.TeamsAppType], "teams-collaborator-agent");
    });

    it("renames the teams other app and maps the other capability verbatim", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(inputs, {
        templateId: "default-bot",
        engine: "v3",
        answers: {
          projectType: "teams-agent-and-app-type",
          teamsApp: "other",
          teamsOtherAppType: "default-bot",
        },
      });

      assert.equal(inputs[QuestionNames.TeamsAppType], "teams-other-app-type");
      assert.equal(inputs["teams-other-app-type"], "default-bot");
    });

    it("maps the office-addin taskpane and config capabilities onto the renamed v3 ids", () => {
      const taskpane: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(taskpane, {
        templateId: "office-addin-wxpo-taskpane",
        engine: "v3",
        answers: {
          projectType: "office-meta-os-type",
          officeAddinCapability: "office-addin-wxpo-taskpane",
        },
      });
      assert.equal(taskpane[QuestionNames.ProjectType], "office-meta-os-type");
      assert.equal(taskpane[QuestionNames.Capabilities], "wxp-json-taskpane");

      const config: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(config, {
        templateId: "office-addin-config",
        engine: "v3",
        answers: {
          projectType: "office-meta-os-type",
          officeAddinCapability: "office-addin-config",
        },
      });
      assert.equal(config[QuestionNames.Capabilities], "office-addin-import");
    });

    it("maps the office DA-meta-os capability and its sub-capability onto the renamed v3 ids", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      applyV3PreFill(inputs, {
        templateId: "declarative-agent-meta-os-new-project",
        engine: "v3",
        answers: {
          projectType: "office-meta-os-type",
          officeAddinCapability: "office-da-meta-os",
          daMetaOsCapability: "declarative-agent-meta-os-new-project",
        },
      });

      assert.equal(inputs[QuestionNames.Capabilities], "office-da-meta-os");
      assert.equal(inputs[QuestionNames.DAMetaOSCapability], "da-meta-os-new-project");
    });

    it("sets only projectType for graph-connector (no capability dimension — safe re-ask)", () => {
      const inputs: Inputs = { platform: Platform.VSCode };
      const target: BuildTarget = {
        templateId: "graph-connector",
        engine: "v3",
        answers: { projectType: "graph-connector-type" },
      };

      applyV3PreFill(inputs, target);

      assert.equal(inputs[QuestionNames.ProjectType], "graph-connector-type");
      assert.isUndefined(inputs[QuestionNames.Capabilities]);
      assert.isUndefined(inputs[QuestionNames.TeamsAppType]);
    });
  });

  describe("scaffoldV4", () => {
    const v4Target: BuildTarget = {
      templateId: "da/mcp-server",
      engine: "v4",
      language: "common",
    };

    it("errors when the create floor has no folder", async () => {
      const inputs: Inputs = { platform: Platform.VSCode, [QuestionNames.AppName]: "MyApp" };

      const res = await scaffoldV4(inputs, v4Target, {});

      assert.isTrue(res.isErr());
    });

    it("errors when the create floor has no app name", async () => {
      const inputs: Inputs = { platform: Platform.VSCode, [QuestionNames.Folder]: "/tmp" };

      const res = await scaffoldV4(inputs, v4Target, {});

      assert.isTrue(res.isErr());
    });

    it("errors when the app name violates the name pattern", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: "/tmp",
        [QuestionNames.AppName]: "invalid name!",
      };

      const res = await scaffoldV4(inputs, v4Target, {});

      assert.isTrue(res.isErr());
    });

    it("scaffolds the located package and returns the project path", async () => {
      const channel = vi
        .spyOn(scaffoldV4Deps, "scaffoldDeclarativeFromV4Channel")
        .mockResolvedValue(TEMPLATE_SOURCE);
      // No teamsapp.yml ⇒ ensureTrackingId is skipped.
      vi.spyOn(pathUtils, "getYmlFilePath").mockReturnValue(undefined);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: "/tmp",
        [QuestionNames.AppName]: "MyApp",
      };
      const flagReader = (name: string): boolean => name === "TEAMSFX_TEST_FLAG";

      const res = await scaffoldV4(inputs, v4Target, { mcpServerType: "remote" }, flagReader);

      assert.isTrue(res.isOk());
      assert.equal(res._unsafeUnwrap().projectPath, path.join(path.resolve("/tmp"), "MyApp"));
      const firstCall = channel.mock.calls[0];
      assert.deepEqual(firstCall[1], { kind: "create", templateId: "da/mcp-server" });
      assert.deepEqual(firstCall[2], { mcpServerType: "remote" });
      assert.deepEqual(firstCall[3], { appName: "MyApp", language: "common" });
      assert.strictEqual(firstCall[5], flagReader);
    });

    it("ensures the tracking id when the scaffold wrote a teamsapp.yml", async () => {
      vi.spyOn(scaffoldV4Deps, "scaffoldDeclarativeFromV4Channel").mockResolvedValue(
        TEMPLATE_SOURCE
      );
      vi.spyOn(pathUtils, "getYmlFilePath").mockReturnValue("/tmp/MyApp/teamsapp.yml");
      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      const ensure = vi.spyOn(coordinator, "ensureTrackingId").mockResolvedValue(ok("tracking-id"));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: "/tmp",
        [QuestionNames.AppName]: "MyApp",
      };

      const res = await scaffoldV4(inputs, v4Target, {});

      assert.isTrue(res.isOk());
      assert.equal(res._unsafeUnwrap().projectId, "tracking-id");
      assert.equal(ensure.mock.calls.length, 1);
    });

    it("defaults the caller-floor language to common when the target has none", async () => {
      const channel = vi
        .spyOn(scaffoldV4Deps, "scaffoldDeclarativeFromV4Channel")
        .mockResolvedValue(TEMPLATE_SOURCE);
      vi.spyOn(pathUtils, "getYmlFilePath").mockReturnValue(undefined);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: "/tmp",
        [QuestionNames.AppName]: "MyApp",
      };

      const res = await scaffoldV4(inputs, { templateId: "da/mcp-server", engine: "v4" }, {});

      assert.isTrue(res.isOk());
      assert.deepEqual(channel.mock.calls[0][3], { appName: "MyApp", language: "common" });
    });

    it("surfaces a channel failure as an error", async () => {
      vi.spyOn(scaffoldV4Deps, "scaffoldDeclarativeFromV4Channel").mockRejectedValue(
        new Error("channel boom")
      );
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: "/tmp",
        [QuestionNames.AppName]: "MyApp",
      };

      const res = await scaffoldV4(inputs, v4Target, {});

      assert.isTrue(res.isErr());
    });
  });

  describe("collectCreateFloor", () => {
    it("skips the floor when folder + app-name are already preset (asks no UI)", async () => {
      // a preset app-name is validated (pattern + path-not-exists) but never re-asked;
      // MockTools UI throws if prompted, so an ok proves the preset-skip path.
      vi.spyOn(fs, "pathExists").mockResolvedValue(false);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.Folder]: "/tmp",
        [QuestionNames.AppName]: "MyApp",
      };

      const res = await collectCreateFloor(inputs, tools.ui);

      assert.isTrue(res.isOk());
    });

    it("short-circuits on a preset template-name (the CLI non-interactive contract)", async () => {
      // template-name preset ⇒ traverse returns before any question, so an absent
      // folder/app-name is taken from the CLI options downstream, never prompted.
      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.TemplateName]: "da/mcp-server",
      };

      const res = await collectCreateFloor(inputs, tools.ui);

      assert.isTrue(res.isOk());
    });

    it("prompts an interactive surface and writes the answers back to inputs", async () => {
      // VS Code interactive, no preset floor ⇒ the floor questions are asked and the
      // answers land on the same inputs bag scaffoldV4 then reads (the bug this fixes:
      // without it the v4 path reached scaffoldV4 with folder undefined).
      vi.spyOn(fs, "pathExists").mockResolvedValue(false);
      vi.spyOn(tools.ui, "selectFolder").mockResolvedValue(
        ok({ type: "success", result: "/picked" })
      );
      vi.spyOn(tools.ui, "inputText").mockResolvedValue(
        ok({ type: "success", result: "PickedApp" })
      );
      const inputs: Inputs = { platform: Platform.VSCode };

      const res = await collectCreateFloor(inputs, tools.ui);

      assert.isTrue(res.isOk());
      assert.equal(inputs[QuestionNames.Folder], "/picked");
      assert.equal(inputs[QuestionNames.AppName], "PickedApp");
    });

    it("propagates a cancellation from the interactive floor prompt", async () => {
      const cancel = new UserError({ source: "Test", name: "UserCancelError", message: "cancel" });
      vi.spyOn(tools.ui, "selectFolder").mockResolvedValue(err(cancel));
      const inputs: Inputs = { platform: Platform.VSCode };

      const res = await collectCreateFloor(inputs, tools.ui);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "UserCancelError");
      }
    });
  });
});
