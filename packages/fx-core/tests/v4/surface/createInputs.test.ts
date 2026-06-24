// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  InputTextConfig,
  InputTextResult,
  MultiSelectConfig,
  MultiSelectResult,
  OptionItem as SurfaceOptionItem,
  SingleSelectConfig,
  SingleSelectResult,
  SystemError,
  UserError,
  UserInteraction,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import AdmZip from "adm-zip";
import path from "path";
import { Result, err, ok } from "neverthrow";
import {
  INPUT_VALIDATION_FAILED,
  OptionsProvider,
} from "../../../src/v4/collectInputs/collectInputs";
import { openCreateQuestions } from "../../../src/v4/distribution/createQuestions";
import { DeclarativeLocator } from "../../../src/v4/model/dataModel";
import { createUiPromptUI } from "../../../src/v4/surface/uiPromptUI";
import { gateLanguagesBySurface, runCreateInputs } from "../../../src/v4/surface/createInputs";

/**
 * Tests for docs/03-specs/operations/scaffolding/collect-create-inputs.md.
 * One `it` per CCI-* acceptance-criteria row. v4-isolated (no v3 import).
 *
 * The floor is built in-memory from the loose `templates/v4` source — the same
 * `addLocalFolder(templates/v4, "v4")` layout `generateV4Zip.js` ships — so the
 * real shipped `da/mcp-server` `questions.json` + `descriptor.json` are exercised
 * with no built `templates.zip` artifact (CI-clean).
 */

const TEMPLATES_V4_DIR = path.resolve(__dirname, "../../../../../templates/v4");
const MCP_DA: DeclarativeLocator = { kind: "create", templateId: "da/mcp-server" };
const OPENAPI_DA: DeclarativeLocator = {
  kind: "create",
  templateId: "da/api-plugin-from-existing-api",
};
const OPENAPI_SPEC = path.resolve(__dirname, "../scenarios/fixtures/repairs-openapi.yaml");

function buildFloor(): Buffer {
  const zip = new AdmZip();
  zip.addLocalFolder(TEMPLATES_V4_DIR, "v4");
  return zip.toBuffer();
}

/** A `mcp.serverTypes` provider yielding both server types (the local-available case). */
const twoServerTypes: OptionsProvider = {
  fetch() {
    return {
      options: [
        { id: "remote", label: "Remote" },
        { id: "local", label: "Local" },
      ],
    };
  },
};

interface Script {
  select?: Record<string, string>;
  text?: Record<string, string>;
  multi?: Record<string, string[]>;
  back?: string[];
}

function noAnswer(name: string): FxError {
  return new UserError({ source: "Test", name: "NoScriptedAnswer", message: name });
}

/**
 * A scripted host `UserInteraction`: answers `selectOption` / `inputText` /
 * `selectOptions` from a per-name script and records every config it saw. Only
 * the three faces the create bridge drives are implemented; the cast in `asUI`
 * is test-only (the src no-`as` rule does not apply to tests).
 */
class ScriptedUserInteraction {
  selectNames: string[] = [];
  textNames: string[] = [];
  multiNames: string[] = [];
  lastSelectConfig?: SingleSelectConfig;
  lastInputConfig?: InputTextConfig;
  lastMultiConfig?: MultiSelectConfig;
  constructor(private readonly script: Script) {}

  selectOption(config: SingleSelectConfig): Promise<Result<SingleSelectResult, FxError>> {
    this.selectNames.push(config.name);
    this.lastSelectConfig = config;
    if (this.script.back?.includes(config.name) === true) {
      return Promise.resolve(ok({ type: "back" }));
    }
    const answer = this.script.select?.[config.name];
    if (answer === undefined) {
      return Promise.resolve(err(noAnswer(config.name)));
    }
    const result: SingleSelectResult = { type: "success", result: answer };
    return Promise.resolve(ok(result));
  }

  inputText(config: InputTextConfig): Promise<Result<InputTextResult, FxError>> {
    this.textNames.push(config.name);
    this.lastInputConfig = config;
    if (this.script.back?.includes(config.name) === true) {
      return Promise.resolve(ok({ type: "back" }));
    }
    const answer = this.script.text?.[config.name];
    if (answer === undefined) {
      return Promise.resolve(err(noAnswer(config.name)));
    }
    const result: InputTextResult = { type: "success", result: answer };
    return Promise.resolve(ok(result));
  }

  selectOptions(config: MultiSelectConfig): Promise<Result<MultiSelectResult, FxError>> {
    this.multiNames.push(config.name);
    this.lastMultiConfig = config;
    if (this.script.back?.includes(config.name) === true) {
      return Promise.resolve(ok({ type: "back" }));
    }
    const answer = this.script.multi?.[config.name];
    if (answer === undefined) {
      return Promise.resolve(err(noAnswer(config.name)));
    }
    const result: MultiSelectResult = { type: "success", result: answer };
    return Promise.resolve(ok(result));
  }
}

function asUI(scripted: ScriptedUserInteraction): UserInteraction {
  return scripted as unknown as UserInteraction;
}

describe("runCreateInputs (collect-create-inputs)", () => {
  it("CCI-01: remote-only provider auto-skips mcpServerType, asks url + authType=none", async () => {
    const ui = new ScriptedUserInteraction({
      text: { mcpServerUrl: "https://api.example.com/mcp" },
      select: { authType: "none" },
    });

    const res = await runCreateInputs(buildFloor(), MCP_DA, {}, asUI(ui), {
      flagReader: () => false,
    });

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {
        mcpServerType: "remote",
        mcpServerUrl: "https://api.example.com/mcp",
        authType: "none",
      });
    }
    // mcpServerType has a single option (remote-only) + skipSingleOption -> never prompted.
    assert.notInclude(ui.selectNames, "mcpServerType");
    assert.deepEqual(ui.selectNames, ["authType"]);
    assert.deepEqual(ui.textNames, ["mcpServerUrl"]);
  });

  it("CCI-17: openapi.operations provider lists operations from the selected OpenAPI document", async () => {
    const ui = new ScriptedUserInteraction({
      multi: { apiOperations: ["GET /repairs"] },
    });

    const res = await runCreateInputs(
      buildFloor(),
      OPENAPI_DA,
      { apiSpecLocation: OPENAPI_SPEC },
      asUI(ui),
      { flagReader: () => false }
    );

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, {
        apiSpecLocation: OPENAPI_SPEC,
        apiOperations: ["GET /repairs"],
      });
    }
    assert.deepEqual(ui.textNames, []);
    assert.deepEqual(ui.multiNames, ["apiOperations"]);
    assert.strictEqual(ui.lastMultiConfig?.options[0].id, "GET /repairs");
  });

  it("CCI-02: provider [remote,local] prompts mcpServerType; local pick skips url, asks authType", async () => {
    const ui = new ScriptedUserInteraction({
      select: { mcpServerType: "local", authType: "none" },
    });

    const res = await runCreateInputs(buildFloor(), MCP_DA, {}, asUI(ui), {
      optionsProvider: { "mcp.serverTypes": twoServerTypes },
      flagReader: () => false,
    });

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value.mcpServerType, "local");
      assert.notProperty(res.value, "mcpServerUrl");
      assert.equal(res.value.authType, "none");
    }
    // mcpServerType prompted (two options); mcpServerUrl's `== 'remote'` condition is false.
    assert.deepEqual(ui.selectNames, ["mcpServerType", "authType"]);
    assert.deepEqual(ui.textNames, []);
  });

  it("CCI-03: an entryParams mcpServerUrl is used as-is (not prompted); authType=oauth", async () => {
    const ui = new ScriptedUserInteraction({ select: { authType: "oauth" } });

    const res = await runCreateInputs(
      buildFloor(),
      MCP_DA,
      { mcpServerUrl: "https://seed.example.com/mcp" },
      asUI(ui),
      { flagReader: () => false }
    );

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value.mcpServerUrl, "https://seed.example.com/mcp");
      assert.equal(res.value.mcpServerType, "remote");
      assert.equal(res.value.authType, "oauth");
    }
    // The pre-filled url is used as-is (INPUT-12) -> the text prompt never runs.
    assert.deepEqual(ui.textNames, []);
  });

  it("CCI-04: an invalid uri for mcpServerUrl -> UserError INPUT_VALIDATION_FAILED", async () => {
    const ui = new ScriptedUserInteraction({ text: { mcpServerUrl: "not a uri" } });

    const res = await runCreateInputs(buildFloor(), MCP_DA, {}, asUI(ui), {
      flagReader: () => false,
    });

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.instanceOf(res.error, UserError);
      assert.equal(res.error.name, INPUT_VALIDATION_FAILED);
    }
  });

  it("CCI-05: da/mcp-server languages=['common'] -> no language axis asked", async () => {
    const ui = new ScriptedUserInteraction({
      text: { mcpServerUrl: "https://api.example.com/mcp" },
      select: { authType: "none" },
    });

    const res = await runCreateInputs(buildFloor(), MCP_DA, {}, asUI(ui), {
      flagReader: () => false,
    });

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.notProperty(res.value, "language");
    }
    assert.notInclude(ui.selectNames, "language");
  });
});

describe("gateLanguagesBySurface (csharp surface/flag gate)", () => {
  // The .NET gate reads v3's `FeatureFlags.CLIDotNet` name ("TEAMSFX_CLI_DOTNET").
  const dotnetOn = (name: string): boolean => name === "TEAMSFX_CLI_DOTNET";
  const dotnetOff = (): boolean => false;

  it("CCI-14: drops csharp on the VS Code surface regardless of the .NET flag", () => {
    assert.deepEqual(
      gateLanguagesBySurface(["typescript", "csharp", "javascript"], "vscode", dotnetOn),
      ["typescript", "javascript"]
    );
    assert.deepEqual(gateLanguagesBySurface(["typescript", "csharp"], "vscode", dotnetOff), [
      "typescript",
    ]);
  });

  it("CCI-15: keeps csharp on the CLI / VS surfaces only when TEAMSFX_CLI_DOTNET is on", () => {
    assert.deepEqual(gateLanguagesBySurface(["typescript", "csharp"], "cli", dotnetOn), [
      "typescript",
      "csharp",
    ]);
    assert.deepEqual(gateLanguagesBySurface(["typescript", "csharp"], "vs", dotnetOn), [
      "typescript",
      "csharp",
    ]);
    assert.deepEqual(gateLanguagesBySurface(["typescript", "csharp"], "cli", dotnetOff), [
      "typescript",
    ]);
  });

  it("CCI-16: leaves non-csharp language lists untouched, order preserved", () => {
    assert.deepEqual(gateLanguagesBySurface(["typescript", "javascript"], "vscode", dotnetOff), [
      "typescript",
      "javascript",
    ]);
    assert.deepEqual(gateLanguagesBySurface(["common"], "vscode", dotnetOff), ["common"]);
  });
});

describe("createUiPromptUI (collect-create-inputs)", () => {
  it("CCI-06: ask maps a singleSelect to selectOption and returns the chosen id", async () => {
    const ui = new ScriptedUserInteraction({ select: { picker: "b" } });
    const prompt = createUiPromptUI(asUI(ui));

    const res = await prompt.ask({ name: "picker", type: "singleSelect", title: "Pick" }, [
      { id: "a", label: "A" },
      { id: "b" },
    ]);

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, { kind: "value", value: "b" });
    }
    assert.equal(ui.lastSelectConfig?.returnObject, false);
    const options = (ui.lastSelectConfig?.options ?? []) as SurfaceOptionItem[];
    assert.equal(options.length, 2);
    assert.equal(options[0].id, "a");
    assert.equal(options[0].label, "A");
    // a v4 option with no label defaults its surface label to its id.
    assert.equal(options[1].id, "b");
    assert.equal(options[1].label, "b");
  });

  it("CCI-07: ask maps a text question to inputText and returns the string", async () => {
    const ui = new ScriptedUserInteraction({ text: { freeText: "hello world" } });
    const prompt = createUiPromptUI(asUI(ui));

    const res = await prompt.ask({ name: "freeText", type: "text", title: "Enter" }, undefined);

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, { kind: "value", value: "hello world" });
    }
    assert.deepEqual(ui.textNames, ["freeText"]);
  });

  it("CCI-08: askMulti maps a multiSelect to selectOptions and returns the ids", async () => {
    const ui = new ScriptedUserInteraction({ multi: { servers: ["alpha", "beta"] } });
    const prompt = createUiPromptUI(asUI(ui));

    const res = await prompt.askMulti({ name: "servers", type: "multiSelect", title: "Servers" }, [
      { id: "alpha" },
      { id: "beta" },
      { id: "gamma" },
    ]);

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, { kind: "value", value: ["alpha", "beta"] });
    }
    assert.deepEqual(ui.multiNames, ["servers"]);
  });

  it("CCI-10: ask projects a host back on a singleSelect to { kind: 'back' }", async () => {
    const ui = new ScriptedUserInteraction({ back: ["picker"] });
    const prompt = createUiPromptUI(asUI(ui));

    const res = await prompt.ask({ name: "picker", type: "singleSelect", title: "Pick" }, [
      { id: "a" },
      { id: "b" },
    ]);

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, { kind: "back" });
    }
  });

  it("CCI-11: ask projects a host back on a text question to { kind: 'back' }", async () => {
    const ui = new ScriptedUserInteraction({ back: ["freeText"] });
    const prompt = createUiPromptUI(asUI(ui));

    const res = await prompt.ask({ name: "freeText", type: "text", title: "Enter" }, undefined);

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, { kind: "back" });
    }
  });

  it("CCI-12: askMulti projects a host back on a multiSelect to { kind: 'back' }", async () => {
    const ui = new ScriptedUserInteraction({ back: ["servers"] });
    const prompt = createUiPromptUI(asUI(ui));

    const res = await prompt.askMulti({ name: "servers", type: "multiSelect", title: "Servers" }, [
      { id: "alpha" },
      { id: "beta" },
    ]);

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value, { kind: "back" });
    }
  });

  it("CCI-13: ask threads the caller's step onto the host config (the Back-button gate)", async () => {
    const ui = new ScriptedUserInteraction({ select: { picker: "a" } });
    const prompt = createUiPromptUI(asUI(ui));

    const res = await prompt.ask(
      { name: "picker", type: "singleSelect", title: "Pick" },
      [{ id: "a" }, { id: "b" }],
      2
    );

    assert.isTrue(res.isOk());
    assert.equal(ui.lastSelectConfig?.step, 2);
  });
});

describe("openCreateQuestions (collect-create-inputs)", () => {
  it("CCI-09: reads the three authored da/mcp-server questions from the floor", () => {
    const res = openCreateQuestions(buildFloor(), MCP_DA);

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(
        res.value.map((q) => q.name),
        ["mcpServerType", "mcpServerUrl", "authType"]
      );
    }
  });

  it("CCI-09: an unknown templateId -> SystemError PackageFileMissing", () => {
    const res = openCreateQuestions(buildFloor(), {
      kind: "create",
      templateId: "da/does-not-exist",
    });

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.instanceOf(res.error, SystemError);
      assert.equal(res.error.name, "PackageFileMissing");
    }
  });
});
