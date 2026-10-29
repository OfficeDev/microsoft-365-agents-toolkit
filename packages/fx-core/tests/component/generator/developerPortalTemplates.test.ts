// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Platform } from "@microsoft/teamsfx-api";
import * as fs from "fs";
import * as path from "path";
import { load } from "js-yaml";
import { assert, vi } from "vitest";
import { Validator } from "../../../src/component/configManager/validator";
import { getTemplateReplaceMap } from "../../../src/component/generator/templates/templateReplaceMap";
import { Generator } from "../../../src/component/generator/generator";
import { renderTemplateFileData } from "../../../src/component/generator/utils";
import { buildRenderContext } from "../../../src/v4/renderContext/buildRenderContext";
import { parseDeclaredKeys, parseReplaceMap } from "../../../src/v4/runtime/packageParse";
import { renderMustache } from "../../../src/v4/runtime/renderMustache";
import { createExpressionPort } from "../../../src/v4/runtime/whitelist";
import { createContext, setTools } from "../../../src/common/globalVars";
import { TdpGenerator } from "../../../src/component/generator/other/tdpGenerator";
import { DefaultTemplateGenerator } from "../../../src/component/generator/defaultGenerator";
import { MockTools } from "../../core/utils";

const root = path.resolve(__dirname, "../../../../../templates");
const resourceId = "11111111-1111-4111-8111-111111111111";
const manifestId = "22222222-2222-4222-8222-222222222222";

function files(directory: string): string[] {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(directory, entry.name);
    return entry.isDirectory() ? files(fullPath) : [fullPath];
  });
}

function variables(enabled: boolean, importedManifestId?: string) {
  vi.stubEnv("TEAMSFX_NEW_DP_APIS", String(enabled));
  const inputs: Inputs = { platform: Platform.VSCode, "app-name": "IdentityApp" };
  if (importedManifestId) {
    inputs.teamsAppFromTdp = {
      appId: resourceId,
      appProfile: { appDetails: { applicationManifest: { id: importedManifestId } } },
    };
  }
  return getTemplateReplaceMap(inputs);
}

const legacyFiles = ["vsc", path.join("vs", "csharp")]
  .flatMap((family) => files(path.join(root, family)))
  .filter((file) =>
    /TEAMS_APP_ID|AppIdEnvName|AppManifestIdEnvReference/.test(fs.readFileSync(file, "utf8"))
  );

describe("SCN-DP-MIGRATION template identity bindings", () => {
  afterEach(() => {
    vi.unstubAllEnvs();
    vi.restoreAllMocks();
  });

  it("AC-07: every active legacy family renders canonical names and preserves flag-off bytes", () => {
    const legacy = variables(false);
    const canonical = variables(true);
    assert.isAbove(legacyFiles.length, 300);
    for (const file of legacyFiles) {
      const source = fs.readFileSync(file);
      const oldText = renderTemplateFileData(file, source, legacy).toString();
      const newText = renderTemplateFileData(file, source, canonical).toString();
      assert.notInclude(newText, "TEAMS_APP_ID", file);
      assert.notInclude(newText, "APP_MANIFEST_ID", file);
      assert.notMatch(newText, /\{\{(?:AppId|AppManifestId|LocalAppManifestId)/, file);
      assert.strictEqual(
        newText
          .replace(/\bAPP_ID\b/g, "TEAMS_APP_ID")
          .replace(/appId: TEAMS_APP_ID/g, "teamsAppId: TEAMS_APP_ID"),
        oldText,
        file
      );
    }
  });

  it("AC-07: rendered create outputs validate under their existing declared YAML schema", () => {
    const validator = new Validator();
    for (const enabled of [false, true]) {
      const vars = variables(enabled);
      for (const file of legacyFiles.filter((entry) => entry.endsWith(".yml.tpl"))) {
        const rendered = renderTemplateFileData(file, fs.readFileSync(file), vars).toString();
        assert.notMatch(rendered, /[ \t]+\r?$/m, file);
        assert.isTrue(rendered.endsWith("\n"), file);
        const document = load(rendered);
        assert.isTrue(validator.validate(document, document.version), file);
      }
    }
  });

  it("AC-07: launch and task URLs retain their exact local and remote patterns", () => {
    const debugRoot = path.join(root, "vsc", "ts", "default-bot", ".vscode");
    for (const enabled of [false, true]) {
      const vars = variables(enabled);
      const id = enabled ? "APP_ID" : "TEAMS_APP_ID";
      for (const file of ["launch.json.tpl", "tasks.json.tpl"]) {
        const rendered = renderTemplateFileData(
          file,
          fs.readFileSync(path.join(debugRoot, file)),
          vars
        ).toString();
        for (const scope of ["local:", ""]) {
          const url = `teams.microsoft.com/l/app/\${{${scope}${id}}}?installAppPackage=true`;
          assert.include(
            rendered,
            file === "launch.json.tpl"
              ? `"https://${url}&webjoin=true&\${account-hint}"`
              : `"${url}"`
          );
        }
        assert.notInclude(rendered, "NewDeveloperPortalApis");
      }
    }
  });

  it("AC-11: equal-ID handoffs render exactly the ordinary one-ID project", () => {
    assert.deepEqual(variables(true, resourceId), variables(true));
    assert.strictEqual(variables(true).NewDeveloperPortalApis, "true");
    assert.strictEqual(variables(false).NewDeveloperPortalApis, "");
    assert.isFalse(
      Object.keys(variables(true)).some((key) =>
        /^(AppId|AppManifestId|LocalAppManifestId)/.test(key)
      )
    );
  });

  it("AC-07: standalone legacy generators receive canonical naming defaults", () => {
    vi.stubEnv("TEAMSFX_NEW_DP_APIS", "true");
    assert.strictEqual(Generator.getDefaultVariables("IdentityApp").NewDeveloperPortalApis, "true");
  });

  it("AC-07: rendering env templates does not opt ordinary projects into sandbox files", async () => {
    vi.stubEnv("TEAMSFX_V4_ENABLED", "false");
    vi.stubEnv("TEAMSFX_SANDBOXED_TEAM", "false");
    setTools(new MockTools());
    const generate = vi.spyOn(Generator, "generate").mockResolvedValue();
    const result = await new DefaultTemplateGenerator().run(
      createContext(),
      {
        platform: Platform.VSCode,
        "app-name": "IdentityApp",
        "template-name": "default-bot",
        "programming-language": "typescript",
      },
      "."
    );
    assert.isTrue(result.isOk());
    const filter = generate.mock.calls[0][0].filterFn;
    assert.isDefined(filter);
    assert.isFalse(filter?.("default-bot/env/.env.sandbox.tpl"));
  });

  it("AC-13: split-ID handoffs use ordinary toggle-only templates before handoff processing", () => {
    const split = variables(true, manifestId);
    assert.deepEqual(split, variables(true));
    for (const file of legacyFiles) {
      const source = fs.readFileSync(file);
      const rendered = renderTemplateFileData(file, source, split).toString();
      assert.notInclude(rendered, "APP_MANIFEST_ID", file);
      assert.notMatch(source.toString(), /\{\{(?:AppId|AppManifestId|LocalAppManifestId)/, file);
    }
  });

  it("AC-13: TDP generator passes the API toggle without split-ID variables", async () => {
    vi.stubEnv("TEAMSFX_NEW_DP_APIS", "true");
    setTools(new MockTools());
    const result = await new TdpGenerator().getTemplateInfos(
      createContext(),
      {
        platform: Platform.VSCode,
        "app-name": "IdentityApp",
        "template-name": "default-bot",
        "programming-language": "typescript",
        teamsAppFromTdp: {
          appId: resourceId,
          appProfile: { appDetails: { applicationManifest: { id: manifestId } } },
        },
      },
      "."
    );
    assert.strictEqual(result._unsafeUnwrap()[0].replaceMap?.NewDeveloperPortalApis, "true");
  });

  it("AC-07: every v4 identity-bearing package declares flag-based names without engine branches", () => {
    const descriptors = files(path.join(root, "v4")).filter(
      (file) => path.basename(file) === "descriptor.json"
    );
    let checked = 0;
    for (const file of descriptors) {
      const directory = path.dirname(file);
      const content = files(directory).filter((entry) =>
        entry.includes(`${path.sep}content${path.sep}`)
      );
      const identityFiles = content.filter((entry) =>
        /TEAMS_APP_ID|AppIdEnvName|AppManifestIdEnvReference/.test(fs.readFileSync(entry, "utf8"))
      );
      if (identityFiles.length === 0) continue;
      checked++;
      const descriptor: unknown = JSON.parse(fs.readFileSync(file, "utf8"));
      const mapping = parseReplaceMap(descriptor)
        ._unsafeUnwrap()
        .filter((entry) => entry.var === "NewDeveloperPortalApis");
      assert.lengthOf(mapping, 1, file);
      assert.notMatch(
        JSON.stringify(descriptor),
        /"(?:AppId|AppManifestId|LocalAppManifestId)\w*"/,
        file
      );
      for (const enabled of [false, true]) {
        const vars = buildRenderContext(
          mapping,
          {},
          { appName: "IdentityApp", language: "typescript" },
          createExpressionPort((name) => name === "TEAMSFX_NEW_DP_APIS" && enabled),
          parseDeclaredKeys(descriptor)
        )._unsafeUnwrap();
        assert.strictEqual(vars.NewDeveloperPortalApis, enabled ? "true" : undefined, file);
        for (const entry of identityFiles) {
          assert.isTrue(entry.endsWith(".tpl"), entry);
          const rendered = renderMustache(fs.readFileSync(entry, "utf8"), vars)._unsafeUnwrap();
          if (entry.endsWith(".yml.tpl")) {
            assert.notMatch(rendered, /[ \t]+\r?$/m, entry);
            assert.isTrue(rendered.endsWith("\n"), entry);
          }
          assert.notInclude(rendered, "APP_MANIFEST_ID", entry);
          assert.notMatch(rendered, /\{\{(?:AppId|AppManifestId|LocalAppManifestId)/, entry);
          if (enabled) assert.notInclude(rendered, "TEAMS_APP_ID", entry);
          else assert.notMatch(rendered, /\bAPP_ID\b/, entry);
        }
      }
    }
    assert.isAbove(checked, 20);
  });
});
