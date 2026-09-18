// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Platform } from "@microsoft/teamsfx-api";
import fse from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createContext, setTools } from "../../../src/common/globalVars";
import { Generator } from "../../../src/component/generator/generator";
import { ScaffoldLocalTemplateAction } from "../../../src/component/generator/generatorAction";
import {
  OfficeAddinGeneratorNew,
  officeAddinGeneratorDeps,
} from "../../../src/component/generator/officeAddin/generator";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import { ProgrammingLanguage, QuestionNames } from "../../../src/question";
import { MockTools } from "../../core/utils";

describe("Office Add-in import project naming (v3)", () => {
  let root: string;
  let source: string;
  let target: string;
  const manifest = {
    manifestVersion: "devPreview",
    version: "1.0.0",
    id: "00000000-0000-4000-8000-000000000001",
    developer: {
      name: "Contoso",
      websiteUrl: "https://contoso.com",
      privacyUrl: "https://contoso.com/privacy",
      termsOfUseUrl: "https://contoso.com/terms",
    },
    name: { short: "Existing Add-in", full: "Existing Add-in Full Name" },
    description: { short: "Existing description", full: "Existing full description" },
    icons: { outline: "outline.png", color: "color.png" },
    accentColor: "#FFFFFF",
    extensions: [{ requirements: { scopes: ["workbook"] } }],
  };
  const packageJson = {
    name: "existing-add-in",
    version: "2.0.0",
    scripts: { start: "office-addin-debugging start manifest.json" },
    devDependencies: { "office-addin-debugging": "6.0.6" },
  };

  beforeEach(async () => {
    vi.stubEnv("TEAMSFX_V4_ENABLED", "false");
    setTools(new MockTools());
    root = await fse.mkdtemp(path.join(os.tmpdir(), "office-import-naming-"));
    source = path.join(root, "source");
    target = path.join(root, "target");
    await fse.ensureDir(target);
    await fse.outputJson(path.join(source, "package.json"), packageJson);
    // Render the bundled template instead of downloading it.
    vi.spyOn(Generator, "generate").mockImplementation(ScaffoldLocalTemplateAction.run);
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await fse.remove(root);
  });

  function inputs(manifestPath: string): Inputs {
    return {
      platform: Platform.VSCode,
      [QuestionNames.TemplateName]: TemplateNames.OfficeAddinCommon,
      [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
      [QuestionNames.AppName]: "My Renamed Add-in 42",
      [QuestionNames.OfficeAddinFolder]: source,
      [QuestionNames.OfficeAddinManifest]: manifestPath,
    };
  }

  it.each(["manifest.json", path.join("appPackage", "manifest.json"), "manifest.xml"])(
    "SCN-CREATE-OFFICE-CONFIG-05: applies the entered name when importing %s",
    async (relativeManifest) => {
      const sourceManifest = path.join(source, relativeManifest);
      const xml = relativeManifest.endsWith(".xml");
      const sourceContent = xml ? "<OfficeApp />" : JSON.stringify(manifest);
      await fse.outputFile(sourceManifest, sourceContent);
      if (xml) {
        vi.spyOn(officeAddinGeneratorDeps, "convertProject").mockImplementation(
          async (manifestPath) => {
            expect(manifestPath).toBe(path.join(target, relativeManifest));
            await fse.outputJson(path.join(target, "manifest.json"), manifest);
          }
        );
      }

      const generator = new OfficeAddinGeneratorNew();
      const context = createContext();
      const answers = inputs(sourceManifest);
      expect(generator.activate(context, answers)).toBe(true);
      const result = await generator.run(context, answers, target);
      expect(result.isOk(), JSON.stringify(result)).toBe(true);
      expect(Generator.generate).toHaveBeenCalledOnce();

      const copiedManifest = await fse.readJson(
        path.join(target, relativeManifest.replace(/\.xml$/, ".json"))
      );
      expect(copiedManifest).toEqual({
        ...manifest,
        name: {
          short: "My Renamed Add-in 42",
          full: "Full name for My Renamed Add-in 42",
        },
      });
      expect(await fse.readJson(path.join(target, "package.json"))).toEqual({
        ...packageJson,
        name: "myrenamedaddin42",
      });
      expect(await fse.readFile(sourceManifest, "utf8")).toBe(sourceContent);
      expect(await fse.readJson(path.join(source, "package.json"))).toEqual(packageJson);
    }
  );

  it("SCN-CREATE-OFFICE-CONFIG-05: fails import when the copied package cannot be parsed", async () => {
    const sourceManifest = path.join(source, "manifest.json");
    await fse.outputJson(sourceManifest, manifest);
    await fse.writeFile(path.join(source, "package.json"), "{invalid");

    const result = await new OfficeAddinGeneratorNew().run(
      createContext(),
      inputs(sourceManifest),
      target
    );

    expect(result.isErr()).toBe(true);
    expect(Generator.generate).not.toHaveBeenCalled();
  });

  it("SCN-CREATE-OFFICE-CONFIG-05: propagates a package write failure and restores the working directory", async () => {
    const sourceManifest = path.join(source, "manifest.json");
    await fse.outputJson(sourceManifest, manifest);
    const writeError = new Error("Cannot write imported package");
    vi.spyOn(fse, "writeJson").mockRejectedValue(writeError);
    const workingDirectory = process.cwd();

    const result = await new OfficeAddinGeneratorNew().run(
      createContext(),
      inputs(sourceManifest),
      target
    );

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.message).toContain(writeError.message);
    }
    expect(Generator.generate).not.toHaveBeenCalled();
    expect(process.cwd()).toBe(workingDirectory);
    expect(await fse.readJson(path.join(source, "package.json"))).toEqual(packageJson);
  });
});
