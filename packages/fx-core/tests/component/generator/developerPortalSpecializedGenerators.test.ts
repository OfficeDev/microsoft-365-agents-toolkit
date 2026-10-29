// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Platform } from "@microsoft/teamsfx-api";
import { assert, vi } from "vitest";
import { createContext, setTools } from "../../../src/common/globalVars";
import { CombinedProjectGenerator } from "../../../src/component/generator/combinedProject/generator";
import { DeclarativeAgentGenerator } from "../../../src/component/generator/declarativeAgent/generator";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import { ProgrammingLanguage, QuestionNames } from "../../../src/question";
import { MockTools } from "../../core/utils";

const resourceId = "11111111-1111-4111-8111-111111111111";
const manifestId = "22222222-2222-4222-8222-222222222222";
const cases = [
  {
    name: "TypeSpec declarative agent",
    generator: new DeclarativeAgentGenerator(),
    template: TemplateNames.DeclarativeAgentWithTypeSpec,
    language: ProgrammingLanguage.TS,
    appName: "IdentityApp",
    safeProjectName: "IdentityApp",
    solutionName: "IdentityApp",
  },
  {
    name: "C# declarative agent",
    generator: new DeclarativeAgentGenerator(),
    template: TemplateNames.DeclarativeAgentBasic,
    language: ProgrammingLanguage.CSharp,
    appName: "Identity App",
    safeProjectName: "SafeIdentity",
    solutionName: "IdentitySolution",
  },
  {
    name: "combined connector project",
    generator: new CombinedProjectGenerator(),
    template: TemplateNames.DeclarativeAgentWithGraphConnector,
    language: ProgrammingLanguage.CSharp,
    appName: "Identity App",
    safeProjectName: "SafeIdentity",
    solutionName: "IdentitySolution",
  },
];

describe("SCN-DP-MIGRATION specialized generator identity maps", () => {
  beforeEach(() => setTools(new MockTools()));
  afterEach(() => vi.unstubAllEnvs());

  for (const testCase of cases) {
    for (const enabled of [false, true]) {
      for (const split of [false, true]) {
        it(`AC-07/AC-11/AC-13: ${testCase.name}, enabled=${enabled}, split=${split}`, async () => {
          vi.stubEnv("TEAMSFX_NEW_DP_APIS", String(enabled));
          const inputs: Inputs = {
            platform: Platform.VSCode,
            [QuestionNames.AppName]: "Identity App",
            [QuestionNames.TemplateName]: testCase.template,
            [QuestionNames.ProgrammingLanguage]: testCase.language,
            [QuestionNames.SafeProjectName]: "SafeIdentity",
            [QuestionNames.SolutionName]: "IdentitySolution",
            targetFramework: "net9.0",
            placeProjectFileInSolutionDir: "true",
          };
          const context = createContext();
          const ordinary = (
            await testCase.generator.getTemplateInfos(context, inputs, ".")
          )._unsafeUnwrap();
          inputs.teamsAppFromTdp = {
            appId: resourceId,
            appProfile: {
              appDetails: { applicationManifest: { id: split ? manifestId : resourceId } },
            },
          };
          const imported = (
            await testCase.generator.getTemplateInfos(context, inputs, ".")
          )._unsafeUnwrap();
          assert.lengthOf(imported, ordinary.length);
          for (let index = 0; index < imported.length; index++) {
            const map = imported[index].replaceMap;
            assert.isDefined(map);
            assert.strictEqual(map?.NewDeveloperPortalApis, enabled ? "true" : "");
            assert.strictEqual(map?.appName, testCase.appName);
            assert.strictEqual(map?.SafeProjectName, testCase.safeProjectName);
            assert.strictEqual(map?.SolutionName, testCase.solutionName);
            assert.strictEqual(map?.TargetFramework, "net9.0");
            assert.strictEqual(map?.PlaceProjectFileInSolutionDir, "true");
            assert.deepEqual(imported[index], ordinary[index]);
          }
        });
      }
    }
  }
});
