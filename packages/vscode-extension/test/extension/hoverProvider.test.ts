// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import { envUtil } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import * as sinon from "sinon";
import { v4 } from "uuid";
import * as vscode from "vscode";
import { environmentVariableRegex } from "../../src/constants";
import * as globalVariables from "../../src/globalVariables";
import { ManifestTemplateHoverProvider } from "../../src/hoverProvider";
import { MockCore } from "../mocks/mockCore";

describe("Manifest template hover - V3", async () => {
  const sandbox = sinon.createSandbox();
  const text = `{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.14/MicrosoftTeams.schema.json",
    "manifestVersion": "1.14",
    "version": "1.0.0",
    "id": "\${{TEAMS_APP_ID}}"
  }`;
  const document: vscode.TextDocument = {
    fileName: "manifest.template.json",
    getText: () => {
      return text;
    },
    lineAt: (line: number) => {
      const lines = text.split("\n");
      return {
        lineNumber: line,
        text: lines[line - 1],
      };
    },
    getWordRangeAtPosition: (position: vscode.Position, regex?: RegExp) => {
      return undefined;
    },
  } as any;

  beforeEach(() => {
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(envUtil, "listEnv").resolves(ok(["local", "dev"]));
  });

  afterEach(() => {
    sandbox.restore();
    environmentVariableRegex.lastIndex = 0;
  });

  it("hover - match", async () => {
    sandbox.stub(envUtil, "readEnv").resolves(
      ok({
        ["TEAMS_APP_ID"]: v4(),
      })
    );

    const hoverProvider = new ManifestTemplateHoverProvider();
    const position = new vscode.Position(5, 15);
    const cts = new vscode.CancellationTokenSource();
    const hover = await hoverProvider.provideHover(document, position, cts.token);

    chai.assert.isTrue(hover !== undefined);
    if (hover !== undefined) {
      chai.assert.isTrue(hover.contents.length > 0);
    }
  });

  it("hover - local", async () => {
    sandbox.stub(envUtil, "readEnv").resolves(
      ok({
        ["TEAMS_APP_ID"]: v4(),
      })
    );

    const document: vscode.TextDocument = {
      fileName: "manifest.template.local.json",
      getText: () => {
        return text;
      },
      lineAt: (line: number) => {
        const lines = text.split("\n");
        return {
          lineNumber: line,
          text: lines[line - 1],
        };
      },
      getWordRangeAtPosition: (position: vscode.Position, regex?: RegExp) => {
        return undefined;
      },
    } as any;

    const hoverProvider = new ManifestTemplateHoverProvider();
    const position = new vscode.Position(5, 15);
    const cts = new vscode.CancellationTokenSource();
    const hover = await hoverProvider.provideHover(document, position, cts.token);

    chai.assert.isTrue(hover !== undefined);
    if (hover !== undefined) {
      chai.assert.isTrue(hover.contents.length > 0);
    }
  });

  it("hover-undefined", async () => {
    sandbox.stub(envUtil, "readEnv").resolves(
      ok({
        ["TEAMS_APP_ID"]: v4(),
      })
    );

    const hoverProvider = new ManifestTemplateHoverProvider();
    const position = new vscode.Position(1, 1);
    const cts = new vscode.CancellationTokenSource();
    const hover = await hoverProvider.provideHover(document, position, cts.token);

    chai.assert.isTrue(hover === undefined);
  });

  it("hover - no value", async () => {
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));

    const hoverProvider = new ManifestTemplateHoverProvider();
    const position = new vscode.Position(5, 15);
    const cts = new vscode.CancellationTokenSource();
    const hover = await hoverProvider.provideHover(document, position, cts.token);

    chai.assert.isTrue(hover !== undefined);
    if (hover !== undefined) {
      chai.assert.isTrue(hover.contents.length > 0);
    }
  });
});
