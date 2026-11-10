// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "vitest";
import { createInMemoryRuntime } from "../../../src/v4/runtime/inMemoryRuntime";
import { scaffold } from "../../../src/v4/runtime/scaffold";
import { loadV4Package, text } from "./helpers/scenarioHarness";

const template = loadV4Package("create", "default-bot");

describe("SCN-DP-MIGRATION v4 scaffold identity", () => {
  for (const enabled of [false, true]) {
    for (const language of ["typescript", "javascript", "python"]) {
      it(`AC-07/AC-11: ${language}, new DP API ${enabled}, scaffolds one consistent identity`, async () => {
        const runtime = createInMemoryRuntime((flag) => enabled && flag === "TEAMSFX_NEW_DP_APIS");
        const result = await scaffold(
          {
            descriptor: template.descriptor,
            pipeline: template.pipeline,
            content: template.content,
            answers: {},
            callerFloor: { appName: "IdentityApp", language },
            targetDir: { path: "/out", existing: [] },
          },
          runtime
        );
        assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "");
        const id = enabled ? "APP_ID" : "TEAMS_APP_ID";
        const output = enabled ? "appId" : "teamsAppId";
        assert.include(text(runtime.files, "m365agents.yml"), `${output}: ${id}`);
        assert.include(text(runtime.files, "appPackage/manifest.json"), `"\${{${id}}}"`);
        assert.include(text(runtime.files, "env/.env.dev"), `${id}=`);
        assert.include(text(runtime.files, ".vscode/launch.json"), `\${{local:${id}}}`);
        assert.include(text(runtime.files, ".vscode/tasks.json"), `\${{local:${id}}}`);
        assert.isFalse(runtime.files.has("env/.env.sandbox"));
        for (const [name, bytes] of runtime.files) {
          assert.notInclude(name, ".tpl");
          assert.notInclude(bytes.toString(), "APP_MANIFEST_ID", name);
          assert.notMatch(bytes.toString(), /\{\{(?:AppId|AppManifestId|LocalAppManifestId)/, name);
          if (enabled) assert.notInclude(bytes.toString(), "TEAMS_APP_ID", name);
        }
      });
    }
  }
});
