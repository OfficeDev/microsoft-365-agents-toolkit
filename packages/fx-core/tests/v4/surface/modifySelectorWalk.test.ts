// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  SingleSelectConfig,
  SingleSelectResult,
  UserError,
  UserInteraction,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import AdmZip from "adm-zip";
import path from "path";
import { Result, err, ok } from "neverthrow";
import { runModifySelector } from "../../../src/v4/surface/modifySelectorWalk";

const TEMPLATES_V4_DIR = path.resolve(__dirname, "../../../../../templates/v4");
const DT = "TEAMSFX_MCP_FOR_DA_DT";

function buildFloor(): Buffer {
  const zip = new AdmZip();
  zip.addLocalFolder(TEMPLATES_V4_DIR, "v4");
  return zip.toBuffer();
}

function flagsOn(...names: string[]): (name: string) => boolean {
  const on = new Set(names);
  return (name) => on.has(name);
}

class ScriptedUI {
  selectNames: string[] = [];
  constructor(private readonly answers: Record<string, string>) {}

  selectOption(config: SingleSelectConfig): Promise<Result<SingleSelectResult, FxError>> {
    this.selectNames.push(config.name);
    const answer = this.answers[config.name];
    if (answer === undefined) {
      return Promise.resolve(
        err(new UserError({ source: "Test", name: "UserCancelError", message: config.name }))
      );
    }
    return Promise.resolve(ok({ type: "success", result: answer }));
  }
}

function asUI(ui: ScriptedUI): UserInteraction {
  return ui as unknown as UserInteraction;
}

const MCP_ADD_ACTION_PICKS: Record<string, string> = {
  addCapability: "add-action",
  actionSource: "mcp",
};

describe("runModifySelector", () => {
  it("WMS-01: add-action→mcp with DT on resolves the v4 add-mcp-server modify package", async () => {
    const ui = new ScriptedUI(MCP_ADD_ACTION_PICKS);

    const res = await runModifySelector(buildFloor(), asUI(ui), "vscode", {
      flagReader: flagsOn(DT),
    });

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value.templateId, "add-mcp-server");
      assert.equal(res.value.engine, "v4");
      assert.deepEqual(res.value.answers, MCP_ADD_ACTION_PICKS);
    }
    assert.deepEqual(ui.selectNames, ["addCapability", "actionSource"]);
  });

  it("WMS-02: add-action→mcp with DT off resolves the v3 addPlugin core method", async () => {
    const ui = new ScriptedUI(MCP_ADD_ACTION_PICKS);

    const res = await runModifySelector(buildFloor(), asUI(ui), "vscode", {
      flagReader: () => false,
    });

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value.templateId, "addPlugin");
      assert.equal(res.value.engine, "v3-core-method");
      assert.deepEqual(res.value.answers, MCP_ADD_ACTION_PICKS);
    }
  });
});
