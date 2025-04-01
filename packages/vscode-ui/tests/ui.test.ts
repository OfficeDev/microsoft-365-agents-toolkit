// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import {
  err,
  ok,
  SelectFileConfig,
  SelectFolderConfig,
  SingleFileOrInputConfig,
  SingleSelectConfig,
  UserError,
} from "@microsoft/teamsfx-api";
import { expect } from "chai";
import "mocha";
import * as sinon from "sinon";
import { stubInterface } from "ts-sinon";
import "./mocks/vscode-mock";
import {
  commands,
  Disposable,
  QuickInputButton,
  QuickPick,
  Task,
  tasks,
  Terminal,
  TextDocument,
  window,
  workspace,
} from "vscode";
import { UserCancelError } from "../src/error";
import { FxQuickPickItem, sleep, VSCodeUI } from "../src/ui";

describe("UI Unit Tests", async () => {
  const ui = new VSCodeUI("Test", (e) => {
    return new UserError({});
  });

  before(() => {
    // Mock user input.
  });

  describe("Manually", () => {
    it("Show Progress 2", async function (this: Mocha.Context) {
      const handler = ui.createProgressBar("Test Progress Bar", 3);
      await handler.start("Prepare");
      await handler.next("First step");
      await handler.next("Second step");
      await handler.next("Third step");
      await handler.text?.("Third step");
      await handler.end(true);
    });
  });

  describe("Select Folder", () => {
    it("has returns default folder", async function (this: Mocha.Context) {
      const config: SelectFolderConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        default: "default folder",
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "default" } as FxQuickPickItem];
        acceptListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      // const telemetryStub = sinon.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFolder(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("default folder");
      }
      sinon.restore();
    });

    it("has returns user cancel", async function (this: Mocha.Context) {
      const config: SelectFolderConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        default: "default folder",
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "browse" } as FxQuickPickItem];
        acceptListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      sinon.stub(window, "showOpenDialog").resolves(undefined);

      const result = await ui.selectFolder(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error instanceof UserCancelError).is.true;
      }
      sinon.restore();
    });
  });

  describe("Select File", () => {
    it("has returns default file", async function (this: Mocha.Context) {
      const config: SelectFileConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        default: "default file",
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "default" } as FxQuickPickItem];
        acceptListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      const result = await ui.selectFile(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("default file");
      }
      sinon.restore();
    });

    it("has returns user cancel", async function (this: Mocha.Context) {
      const config: SelectFileConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        default: "default folder",
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let onHideListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        onHideListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "browse" } as FxQuickPickItem];
        onHideListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      sinon.stub(window, "showOpenDialog").resolves(undefined);

      const result = await ui.selectFile(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error instanceof UserCancelError).is.true;
      }
      sinon.restore();
    });

    it("has returns item in possible files", async function (this: Mocha.Context) {
      const config: SelectFileConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        default: "default folder",
        possibleFiles: [
          {
            id: "1",
            label: "1",
          },
          {
            id: "2",
            label: "2",
          },
        ],
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "1" } as FxQuickPickItem];
        acceptListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const result = await ui.selectFile(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
      }
      sinon.restore();
    });

    it("has returns invalid input item id", async function (this: Mocha.Context) {
      const config: SelectFileConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        default: "default",
        possibleFiles: [
          {
            id: "default",
            label: "default",
          },
        ],
      };

      const result = await ui.selectFile(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("InvalidInput");
      }
      sinon.restore();
    });

    it("selects a file which pass validation", async function (this: Mocha.Context) {
      const config: SelectFileConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        default: "default file",
        validation: (input: string) => {
          if (input === "default file") {
            return undefined;
          }
          return "validation failed";
        },
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "default" } as FxQuickPickItem];
        acceptListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const res = await ui.selectFile(config);
      expect(res.isOk()).is.true;

      sinon.restore();
    });

    it("selects a file with error thrown when validating result", async function (this: Mocha.Context) {
      const config: SelectFileConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        default: "default file",
        validation: (input: string) => {
          throw new UserError("source", "name", "", "");
        },
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "default" } as FxQuickPickItem];
        acceptListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const res = await ui.selectFile(config);
      expect(res.isErr()).is.true;

      sinon.restore();
    });
  });

  describe("Open File", () => {
    it("open the preview of Markdown file", async function (this: Mocha.Context) {
      sinon.stub(workspace, "openTextDocument").resolves({} as TextDocument);
      let executedCommand = "";
      sinon.stub(commands, "executeCommand").callsFake((command: string, ...args: any[]) => {
        executedCommand = command;
        return Promise.resolve();
      });
      const showTextStub = sinon.stub(window, "showTextDocument");

      const result = await ui.openFile("test.md");

      expect(result.isOk()).is.true;
      expect(showTextStub.calledOnce).to.be.false;
      expect(executedCommand).to.equal("markdown.showPreview");
      sinon.restore();
    });
  });

  describe("runCommand", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });
    it("runs command timeout", async function (this: Mocha.Context) {
      const timer = sandbox.useFakeTimers();
      const runCmd = ui.runCommand({ cmd: "test", timeout: 200 });
      sandbox.stub(tasks, "executeTask").resolves();
      await timer.tickAsync(2000);
      const result = await runCmd;
      expect(result.isErr()).is.true;
      timer.restore();
    });
    it("runs command successfully", async function (this: Mocha.Context) {
      sandbox.stub(tasks, "executeTask").resolves();
      const disposable = { dispose: () => {} };
      const stub = sandbox.stub(tasks, "onDidEndTaskProcess").returns(disposable);
      stub.yields({
        exitCode: 0,
        execution: { task: { name: "Execute script action", source: "ms-teams-vscode-extension" } },
      });
      const runCmd = ui.runCommand({ cmd: "test" });
      const result = await runCmd;
      expect(result.isOk()).is.true;
    });
  });

  describe("single select", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("select success with validation", async function (this: Mocha.Context) {
      let hasRun = false;
      const config: SingleSelectConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        options: [{ id: "1", label: "label1" }],
        validation: (input: string) => {
          if (input === "1") {
            hasRun = true;
            return undefined;
          }
        },
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerItemButton.callsFake((listener: (e: any) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "1" } as FxQuickPickItem];
        acceptListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
      }
      sinon.restore();
    });

    it("select fail with validation", async function (this: Mocha.Context) {
      const hasRun = false;
      const config: SingleSelectConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        options: [{ id: "1", label: "label1" }],
        validation: (input: string) => {
          throw new UserError("name", "source", "msg", "msg");
        },
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerItemButton.callsFake((listener: (e: any) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "1" } as FxQuickPickItem];
        acceptListener();
      });
      sinon.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const result = await ui.selectOption(config);

      expect(result.isErr()).is.true;

      sinon.restore();
    });

    it("loads dynamic options in a short time", async function (this: Mocha.Context) {
      const config: SingleSelectConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        options: async () => {
          return Promise.resolve([{ id: "1", label: "label1" }]);
        },
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerItemButton.callsFake((listener: (e: any) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "1" } as FxQuickPickItem];
        acceptListener();
      });
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
      }
    });

    it("loads dynamic option in a short time and auto select", async function (this: Mocha.Context) {
      const config: SingleSelectConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        options: async () => {
          return Promise.resolve([{ id: "1", label: "label1" }]);
        },
        skipSingleOption: true,
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerItemButton.callsFake((listener: (e: any) => unknown) => {
        return mockDisposable;
      });
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
      }
      sandbox.restore();
    });

    it("loads dynamic options in a short time and shows", async function (this: Mocha.Context) {
      const config: SingleSelectConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        options: async () => {
          return Promise.resolve([
            { id: "1", label: "label1" },
            { id: "2", label: "label2" },
          ]);
        },
        skipSingleOption: true,
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerItemButton.callsFake((listener: (e: any) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "1" } as FxQuickPickItem];
        acceptListener();
      });
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
        expect(mockQuickPick.show.called).is.true;
      }
      sandbox.restore();
    });

    it("loads dynamic option in a long time and shows", async function (this: Mocha.Context) {
      const config: SingleSelectConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        options: async () => {
          await sleep(1000);
          return Promise.resolve([{ id: "1", label: "label1" }]);
        },
        skipSingleOption: true,
      };

      const mockQuickPick = stubInterface<QuickPick<FxQuickPickItem>>();
      const mockDisposable = stubInterface<Disposable>();
      let acceptListener: (e: void) => any;
      mockQuickPick.onDidAccept.callsFake((listener: (e: void) => unknown) => {
        acceptListener = listener;
        return mockDisposable;
      });
      mockQuickPick.onDidHide.callsFake((listener: (e: void) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerButton.callsFake((listener: (e: QuickInputButton) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.onDidTriggerItemButton.callsFake((listener: (e: any) => unknown) => {
        return mockDisposable;
      });
      mockQuickPick.show.callsFake(() => {
        mockQuickPick.selectedItems = [{ id: "1" } as FxQuickPickItem];
        acceptListener();
      });
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
        expect(mockQuickPick.show.called).is.true;
      }
      sandbox.restore();
    });
  });

  describe("Select local file or input", () => {
    it("selects local file successfully", async function (this: Mocha.Context) {
      const config: SingleFileOrInputConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        inputOptionItem: {
          id: "input",
          label: "input",
        },
        inputBoxConfig: {
          prompt: "prompt",
          title: "title",
          name: "input name",
        },
      };

      sinon.stub(ui, "selectFile").resolves(ok({ type: "success", result: "file" }));

      const result = await ui.selectFileOrInput(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("file");
      }
      sinon.restore();
    });

    it("selects local file error", async function (this: Mocha.Context) {
      const config: SingleFileOrInputConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        inputOptionItem: {
          id: "input",
          label: "input",
        },
        inputBoxConfig: {
          prompt: "prompt",
          title: "title",
          name: "input name",
        },
      };

      sinon.stub(ui, "selectFile").resolves(err(new UserError("source", "name", "msg", "msg")));

      const result = await ui.selectFileOrInput(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("name");
      }
      sinon.restore();
    });

    it("inputs a value sucessfully", async function (this: Mocha.Context) {
      const config: SingleFileOrInputConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        inputOptionItem: {
          id: "input",
          label: "input",
        },
        inputBoxConfig: {
          prompt: "prompt",
          title: "title",
          name: "input name",
        },
      };

      sinon.stub(ui, "selectFile").resolves(ok({ type: "success", result: "input" }));
      sinon.stub(ui, "inputText").resolves(ok({ type: "success", result: "testUrl" }));

      const result = await ui.selectFileOrInput(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("testUrl");
      }
      sinon.restore();
    });

    it("inputs a value error", async function (this: Mocha.Context) {
      const config: SingleFileOrInputConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        inputOptionItem: {
          id: "input",
          label: "input",
        },
        inputBoxConfig: {
          prompt: "prompt",
          title: "title",
          name: "input name",
        },
      };

      sinon.stub(ui, "selectFile").resolves(ok({ type: "success", result: "input" }));
      sinon.stub(ui, "inputText").resolves(err(new UserError("source", "name", "msg", "msg")));

      const result = await ui.selectFileOrInput(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("name");
      }
      sinon.restore();
    });

    it("inputs a value back and then sucessfully", async function (this: Mocha.Context) {
      const config: SingleFileOrInputConfig = {
        name: "name",
        title: "title",
        placeholder: "placeholder",
        inputOptionItem: {
          id: "input",
          label: "input",
        },
        inputBoxConfig: {
          prompt: "prompt",
          title: "title",
          name: "input name",
        },
      };

      sinon.stub(ui, "selectFile").resolves(ok({ type: "success", result: "input" }));
      sinon
        .stub(ui, "inputText")
        .onFirstCall()
        .resolves(ok({ type: "back" }))
        .onSecondCall()
        .resolves(ok({ type: "success", result: "testUrl" }));

      const result = await ui.selectFileOrInput(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("testUrl");
      }
      sinon.restore();
    });
  });
});
