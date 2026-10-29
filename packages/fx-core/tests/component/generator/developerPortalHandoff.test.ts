// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, ok, Platform, TeamsAppManifest, err, UserError } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import os from "os";
import path from "path";
import { expect, vi } from "vitest";
import { FeatureFlagName, FeatureFlags } from "../../../src/common/featureFlags";
import { createContext, setTools } from "../../../src/common/globalVars";
import * as appStudio from "../../../src/component/driver/teamsApp/appStudio";
import { copilotGptManifestUtils } from "../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import {
  DeclarativeAgentGenerator,
  declarativeAgentGeneratorDeps,
} from "../../../src/component/generator/declarativeAgent/generator";
import * as apiSpec from "../../../src/component/generator/openApiSpec/common";
import {
  CustomEngineAgentWithExistingApiSpecGenerator,
  customEngineAgentGeneratorDeps,
} from "../../../src/component/generator/openApiSpec/customEngineAgentGenerator";
import { DeclarativeAgentWithExistingApiSpecGenerator } from "../../../src/component/generator/openApiSpec/declarativeAgentGenerator";
import { TdpGenerator } from "../../../src/component/generator/other/tdpGenerator";
import { CombinedProjectGenerator } from "../../../src/component/generator/combinedProject/generator";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import { dotenvUtil } from "../../../src/component/utils/envUtil";
import { QuestionNames } from "../../../src/question";
import { MockTools } from "../../core/utils";

describe("DP generator handoff persistence", () => {
  const originalDefault = FeatureFlags.NewDeveloperPortalApis.defaultValue;
  let projectPath: string;
  const resourceId = "11111111-1111-4111-8111-111111111111";
  const distinctManifestId = "22222222-2222-4222-8222-222222222222";
  const warnings = [{ type: "test", content: "Template warning" }];

  beforeEach(async () => {
    setTools(new MockTools());
    FeatureFlags.NewDeveloperPortalApis.defaultValue = "true";
    vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, undefined);
    vi.stubEnv(FeatureFlagName.SensitivityLabelEnabled, "false");
    vi.stubEnv(FeatureFlagName.GenerateConfigFiles, "false");
    projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "dp-generator-handoff-"));
    const yaml = "version: 1.9\nprojectId: test-project\nenvironmentFolderPath: ./env\n";
    await fs.outputFile(path.join(projectPath, "m365agents.yml"), yaml);
    await fs.outputFile(path.join(projectPath, "env", ".env.dev"), "APP_ID=\n");
    await fs.outputFile(
      path.join(projectPath, "appPackage", "manifest.json"),
      JSON.stringify(new TeamsAppManifest())
    );
    await fs.outputFile(
      path.join(projectPath, ".vscode", "launch.json"),
      '{"url":"https://teams.microsoft.com/l/app/${{local:APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}"}'
    );
    vi.spyOn(copilotGptManifestUtils, "getManifestPath").mockResolvedValue(
      ok(path.join(projectPath, "appPackage", "declarativeAgent.json"))
    );
    vi.spyOn(declarativeAgentGeneratorDeps, "generateForMCPForDA").mockResolvedValue(
      ok({ warnings })
    );
    vi.spyOn(declarativeAgentGeneratorDeps, "addExistingPlugin").mockResolvedValue(
      ok({
        warnings,
        destinationPluginManifestPath: path.join(projectPath, "appPackage", "plugin.json"),
      })
    );
    vi.spyOn(apiSpec, "generateFilesFromApiSpec").mockResolvedValue(ok({ warnings }));
    vi.spyOn(customEngineAgentGeneratorDeps, "generateFilesFromApiSpec").mockResolvedValue(
      ok({ warnings })
    );
  });

  afterEach(async () => {
    FeatureFlags.NewDeveloperPortalApis.defaultValue = originalDefault;
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await fs.remove(projectPath);
  });

  const cases = [
    {
      template: TemplateNames.DeclarativeAgentBasic,
      generator: () => new DeclarativeAgentGenerator(),
    },
    {
      template: TemplateNames.DeclarativeAgentWithActionFromMCP,
      generator: () => new DeclarativeAgentGenerator(),
    },
    {
      template: TemplateNames.DeclarativeAgentWithExistingAction,
      generator: () => new DeclarativeAgentGenerator(),
    },
    {
      template: TemplateNames.DeclarativeAgentWithActionFromExistingApiSpec,
      generator: () => new DeclarativeAgentWithExistingApiSpecGenerator(),
    },
    {
      template: TemplateNames.CustomCopilotRagCustomApi,
      generator: () => new CustomEngineAgentWithExistingApiSpecGenerator(),
    },
    { template: TemplateNames.Tab, generator: () => new TdpGenerator() },
    {
      template: TemplateNames.DeclarativeAgentWithGraphConnector,
      generator: () => new CombinedProjectGenerator(),
    },
  ];

  for (const local of [false, true]) {
    for (const split of [false, true]) {
      it.each(cases)(
        `AC-11/AC-13: $template persists ${split ? "split" : "equal"} IDs in ${local ? "local" : "dev"}`,
        async ({ template, generator: makeGenerator }) => {
          if (local) {
            await fs.copy(
              path.join(projectPath, "m365agents.yml"),
              path.join(projectPath, "m365agents.local.yml")
            );
            await fs.outputFile(path.join(projectPath, "env", ".env.local"), "APP_ID=\n");
          }
          const manifestId = split ? distinctManifestId : resourceId;
          const manifest = new TeamsAppManifest();
          manifest.id = manifestId;
          const download = vi
            .spyOn(appStudio, "getAppPackage")
            .mockResolvedValue(ok({ manifest: Buffer.from(JSON.stringify(manifest)) }));
          const generator = makeGenerator();
          if (generator instanceof CombinedProjectGenerator) {
            await fs.copy(
              path.join(projectPath, "appPackage"),
              path.join(projectPath, generator.temporaryFolderName, "appPackage")
            );
          }
          // Template acquisition is outside this test; post-processing and handoff use real files.
          vi.spyOn(generator, "getTemplateInfos").mockResolvedValue(ok([]));
          const inputs: Inputs = {
            platform: Platform.VSCode,
            projectPath,
            [QuestionNames.TemplateName]: template,
            teamsAppFromTdp: {
              appId: resourceId,
              appProfile: { appDetails: { applicationManifest: { id: manifestId } } },
            },
          };
          const result = await generator.run(createContext(), inputs, projectPath);
          expect(result.isOk()).toBe(true);
          const env = dotenvUtil.deserialize(
            await fs.readFile(path.join(projectPath, "env", `.env.${local ? "local" : "dev"}`))
          ).obj;
          expect(env.APP_ID).toBe(resourceId);
          expect(env.APP_MANIFEST_ID).toBe(split ? manifestId : undefined);
          expect(env.TEAMS_APP_ID).toBeUndefined();
          expect(download).toHaveBeenCalledTimes(1);
          const outputManifest = await fs.readJson(
            path.join(projectPath, "appPackage", "manifest.json")
          );
          expect(outputManifest.id).toBe(split ? "${{APP_MANIFEST_ID}}" : "${{APP_ID}}");
          const launch = await fs.readFile(
            path.join(projectPath, ".vscode", "launch.json"),
            "utf8"
          );
          expect(launch).toBe(
            '{"url":"https://teams.microsoft.com/l/app/${{local:APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}"}'
          );
          if (local) {
            expect(await fs.readFile(path.join(projectPath, "env", ".env.dev"), "utf8")).toBe(
              "APP_ID=\n"
            );
          }
          if (template === TemplateNames.DeclarativeAgentWithActionFromMCP && result.isOk()) {
            expect(result.value.warnings).toEqual(warnings);
          }
        }
      );
    }
  }

  it("AC-13: failed template post-processing does not download or persist handoff IDs", async () => {
    const generator = new DeclarativeAgentGenerator();
    vi.spyOn(generator, "getTemplateInfos").mockResolvedValue(ok([]));
    const failure = new UserError("test", "McpFailure", "MCP failed");
    vi.mocked(declarativeAgentGeneratorDeps.generateForMCPForDA).mockResolvedValue(err(failure));
    const download = vi.spyOn(appStudio, "getAppPackage");
    const result = await generator.run(
      createContext(),
      {
        platform: Platform.VSCode,
        projectPath,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        teamsAppFromTdp: {
          appId: resourceId,
          appProfile: { appDetails: { applicationManifest: { id: distinctManifestId } } },
        },
      },
      projectPath
    );
    expect(result.isErr() && result.error).toBe(failure);
    expect(download).not.toHaveBeenCalled();
    expect(await fs.readFile(path.join(projectPath, "env", ".env.dev"), "utf8")).toBe("APP_ID=\n");
  });

  it("AC-11: legacy handoff retains a single TEAMS_APP_ID with the new API disabled", async () => {
    vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, "false");
    const generator = new TdpGenerator();
    vi.spyOn(generator, "getTemplateInfos").mockResolvedValue(ok([]));
    await fs.outputFile(path.join(projectPath, "env", ".env.dev"), "TEAMS_APP_ID=\n");
    await fs.copy(
      path.join(projectPath, "m365agents.yml"),
      path.join(projectPath, "m365agents.local.yml")
    );
    const manifest = new TeamsAppManifest();
    manifest.id = resourceId;
    vi.spyOn(appStudio, "getAppPackage").mockResolvedValue(
      ok({ manifest: Buffer.from(JSON.stringify(manifest)) })
    );
    const result = await generator.run(
      createContext(),
      {
        platform: Platform.VSCode,
        projectPath,
        teamsAppFromTdp: { teamsAppId: resourceId, appId: resourceId },
      },
      projectPath
    );
    expect(result.isOk()).toBe(true);
    const env = dotenvUtil.deserialize(
      await fs.readFile(path.join(projectPath, "env", ".env.dev"))
    ).obj;
    expect(env.TEAMS_APP_ID).toBe(resourceId);
    expect(env.APP_ID).toBeUndefined();
    expect(env.APP_MANIFEST_ID).toBeUndefined();
    const manifestOutput = await fs.readJson(path.join(projectPath, "appPackage", "manifest.json"));
    expect(manifestOutput.id).toBe("${{TEAMS_APP_ID}}");
  });

  it.each(cases)(
    "AC-20: flag-off $template retains its original handoff behavior",
    async ({ template, generator: makeGenerator }) => {
      vi.stubEnv(FeatureFlagName.NewDeveloperPortalApis, "false");
      await fs.copy(
        path.join(projectPath, "m365agents.yml"),
        path.join(projectPath, "m365agents.local.yml")
      );
      const manifest = new TeamsAppManifest();
      manifest.id = resourceId;
      const download = vi
        .spyOn(appStudio, "getAppPackage")
        .mockResolvedValue(ok({ manifest: Buffer.from(JSON.stringify(manifest)) }));
      const generator = makeGenerator();
      if (generator instanceof CombinedProjectGenerator) {
        await fs.copy(
          path.join(projectPath, "appPackage"),
          path.join(projectPath, generator.temporaryFolderName, "appPackage")
        );
      }
      vi.spyOn(generator, "getTemplateInfos").mockResolvedValue(ok([]));
      const result = await generator.run(
        createContext(),
        {
          platform: Platform.VSCode,
          projectPath,
          [QuestionNames.TemplateName]: template,
          teamsAppFromTdp: { teamsAppId: resourceId, appId: resourceId },
        },
        projectPath
      );
      expect(result.isOk()).toBe(true);
      const expected =
        generator instanceof TdpGenerator ||
        generator instanceof CombinedProjectGenerator ||
        template === TemplateNames.DeclarativeAgentBasic;
      expect(download).toHaveBeenCalledTimes(expected ? 1 : 0);
      const env = dotenvUtil.deserialize(
        await fs.readFile(path.join(projectPath, "env", ".env.dev"))
      ).obj;
      expect(env.TEAMS_APP_ID).toBe(expected ? resourceId : undefined);
      expect(env.APP_MANIFEST_ID).toBeUndefined();
    }
  );
});
