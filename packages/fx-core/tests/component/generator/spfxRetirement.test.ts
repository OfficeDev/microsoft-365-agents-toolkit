// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform, Stage } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { Container } from "typedi";
import { afterEach, expect, it, vi } from "vitest";
import { createContext, setTools } from "../../../src/common/globalVars";
import { Lifecycle } from "../../../src/component/configManager/lifecycle";
import { cpUtils } from "../../../src/component/deps-checker";
import "../../../src/component/driver/index";
import { Generators } from "../../../src/component/generator/generatorProvider";
import { FxCore } from "../../../src/core/FxCore";
import { QuestionNames } from "../../../src/question/constants";
import { CreateProjectOptions } from "../../../src/question/options";
import { tabProjectTypeNode } from "../../../src/question/scaffold/vsc/teamsProjectTypeNode";
import { MockTools } from "../../core/utils";

afterEach(() => {
  vi.restoreAllMocks();
});

for (const platform of [Platform.CLI, Platform.VSCode]) {
  it(`RET-SPFX-02: rejects legacy addWebpart on ${platform} without prompts, file reads or tools`, async () => {
    const tools = new MockTools();
    setTools(tools);
    const core = new FxCore(tools);
    const read = vi.spyOn(fs, "readJson").mockRejectedValue(new Error("unexpected project read"));
    const execute = vi.spyOn(cpUtils, "executeCommand").mockResolvedValue("");
    const folder = vi.spyOn(tools.ui, "selectFolder");
    const option = vi.spyOn(tools.ui, "selectOption");

    const result = await core.addWebpart({
      platform,
      stage: Stage.addWebpart,
      projectPath: "retired-spfx-project",
      [QuestionNames.SPFxFolder]: "retired-spfx-project\\src",
      [QuestionNames.SPFxWebpartName]: "hello",
      nonInteractive: false,
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) expect(result.error.name).to.equal("SPFxGenerationNotSupported");
    expect(read).not.toHaveBeenCalled();
    expect(execute).not.toHaveBeenCalled();
    expect(folder).not.toHaveBeenCalled();
    expect(option).not.toHaveBeenCalled();
  });
}

for (const solution of ["new", "import"]) {
  it(`RET-SPFX-02: no generator activates for retired SPFx ${solution}`, () => {
    setTools(new MockTools());
    const context = createContext();
    const inputs = {
      platform: Platform.CLI,
      [QuestionNames.Capabilities]: "tab-spfx",
      [QuestionNames.TemplateName]: "spfx-tab",
      [QuestionNames.SPFxSolution]: solution,
    };

    expect(Generators.filter((generator) => generator.activate(context, inputs))).to.deep.equal([]);
  });
}

it("RET-SPFX-02: removes the add driver but preserves existing SPFx deployment and packaging", () => {
  expect(Container.has("spfx/add")).to.be.false;
  expect(Container.has("spfx/deploy")).to.be.true;
  expect(Container.has("teamsApp/copyAppPackageToSPFx")).to.be.true;
});

it("RET-SPFX-02: rejects legacy YAML spfx/add through normal lifecycle driver resolution", () => {
  const lifecycle = new Lifecycle("deploy", [{ uses: "spfx/add", with: {} }], "v1.7");
  const result = lifecycle.resolveDriverInstances(new MockTools().logProvider);
  expect(result.isErr()).to.be.true;
  if (result.isErr()) expect(result.error.name).to.equal("InvalidYmlActionNameError");
});

it("RET-SPFX-01: offers regular Teams tabs without an SPFx generation branch", () => {
  const capability = CreateProjectOptions.find((option) => option.name === "capability");
  expect(capability?.type).to.equal("string");
  if (capability?.type === "string") expect(capability.choices).not.to.contain("tab-spfx");
  const node = tabProjectTypeNode();
  expect(node.data.type).to.equal("singleSelect");
  if (node.data.type === "singleSelect") {
    expect(
      node.data.staticOptions?.map((option) => (typeof option === "string" ? option : option.id))
    ).not.to.contain("tab-spfx");
  }
  expect(node.children ?? []).to.deep.equal([]);
});
