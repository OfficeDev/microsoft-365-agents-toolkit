// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import * as sinon from "sinon";
import * as chai from "chai";
import { window } from "vscode";

import { ProgressHandler } from "../../src/debug/progressHandler";
import * as vsc_ui from "@microsoft/vscode-ui";
import * as localizeUtils from "../../src/utils/localizeUtils";
import * as vscodeMocks from "../mocks/vsc";

afterEach(() => {
  sinon.restore();
});

describe("ProgressHandler", () => {
  let message: string | undefined = undefined;
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(window, "withProgress").callsFake(async (options, task) => {
      return await task(
        {
          report: (value) => {
            message = value.message;
          },
        },
        new vscodeMocks.CancellationToken()
      );
    });
    sandbox.stub(vsc_ui, "sleep").callsFake(async () => {});
    sandbox.stub(localizeUtils, "localize").callsFake((key) => {
      if (key === "teamstoolkit.progressHandler.showOutputLink") {
        return "Check [output window](%s) for details.";
      } else if (key === "teamstoolkit.progressHandler.showTerminalLink") {
        return "Check [terminal window](%s) for details.";
      } else if (key === "teamstoolkit.progressHandler.prepareTask") {
        return " Prepare task.";
      } else if (key === "teamstoolkit.progressHandler.reloadNotice") {
        return "%s%s%s (Notice: You can reload the window and retry if task spends too long time.)";
      }
      return "";
    });
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("terminal", async () => {
    const progressHandler = new ProgressHandler("test title", 1, "terminal");
    let expected =
      "test title: [0/1] Prepare task. Check [terminal window](command:workbench.action.terminal.focus) for details. (Notice: You can reload the window and retry if task spends too long time.)";
    await progressHandler.start();
    chai.assert.equal(message, expected);
    await progressHandler.next("test message.");
    expected =
      "test title: [1/1] test message. Check [terminal window](command:workbench.action.terminal.focus) for details. (Notice: You can reload the window and retry if task spends too long time.)";
    chai.assert.equal(message, expected);
  });

  it("output", async () => {
    const progressHandler = new ProgressHandler("test title", 1, "output");
    let expected =
      "test title: [0/1] Prepare task. Check [output window](command:fx-extension.showOutputChannel) for details. (Notice: You can reload the window and retry if task spends too long time.)";
    await progressHandler.start();
    chai.assert.equal(message, expected);
    await progressHandler.next("test message.");
    expected =
      "test title: [1/1] test message. Check [output window](command:fx-extension.showOutputChannel) for details. (Notice: You can reload the window and retry if task spends too long time.)";
    chai.assert.equal(message, expected);
  });

  it("not started", async () => {
    message = undefined;
    const progressHandler = new ProgressHandler("test title", 1, "output");
    await progressHandler.next("test message.");
    await progressHandler.end(true);
    chai.assert.equal(message, undefined);
  });
});
