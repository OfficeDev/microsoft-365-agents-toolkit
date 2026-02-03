// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import "./mocks/vscode-mock";

import {
  err,
  ok,
  SelectFileConfig,
  SelectFolderConfig,
  SingleFileOrInputConfig,
  SingleSelectConfig,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import { expect } from "chai";
import * as fs from "fs";
import "mocha";
import * as sinon from "sinon";
import { stubInterface } from "ts-sinon";
import {
  commands,
  Disposable,
  QuickInputButton,
  QuickPick,
  tasks,
  TextDocument,
  window,
  workspace,
} from "vscode";
import { UserCancelError } from "../src/error";
import { FxQuickPickItem, sleep, VSCodeUI } from "../src/ui";

describe("UI Unit Tests", async () => {
  const ui = new VSCodeUI("Test", (e) => {
    // If it's already an FxError, just return it; otherwise wrap it
    if (e instanceof UserError || e instanceof SystemError) {
      return e;
    }
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

  describe("runCommand", () => {
    let fsReadFileSyncStub: sinon.SinonStub;
    let fsUnlinkSyncStub: sinon.SinonStub;
    let tasksExecuteTaskStub: sinon.SinonStub;
    let tasksOnDidEndTaskProcessStub: sinon.SinonStub;
    let onDidEndTaskProcessListeners: any[] = [];
    let tempFiles: string[] = [];

    beforeEach(() => {
      fsUnlinkSyncStub = sinon.stub(fs, "unlinkSync");
      tasksExecuteTaskStub = sinon.stub(tasks, "executeTask").callsFake((task: any) => {
        // Capture the temp file path from the shell execution command
        // The commandLine uses tee/Tee-Object: cmd 2>&1 | tee "path/to/file" or cmd 2>&1 | Tee-Object -FilePath "path/to/file"
        const execution = task.execution;
        if (execution && execution.commandLine) {
          // Match both tee "path" and Tee-Object -FilePath "path" patterns
          const match = execution.commandLine.match(/(?:tee|Tee-Object -FilePath)\s+"([^"]+)"/);
          if (match) {
            const tempFilePath = match[1];
            tempFiles.push(tempFilePath);
          }
        }
        return Promise.resolve(undefined as any);
      });
      onDidEndTaskProcessListeners = [];

      // Store all listeners so we can invoke them later
      tasksOnDidEndTaskProcessStub = sinon
        .stub(tasks, "onDidEndTaskProcess")
        .callsFake((callback: any) => {
          onDidEndTaskProcessListeners.push(callback);
          return { dispose: () => {} };
        });
    });

    afterEach(() => {
      sinon.restore();
      onDidEndTaskProcessListeners = [];
      // Clean up any temporary files we created
      for (const tempFile of tempFiles) {
        try {
          if (fs.existsSync(tempFile)) {
            fs.unlinkSync(tempFile);
          }
        } catch (e) {
          // Ignore errors during cleanup
        }
      }
      tempFiles = [];
    });

    it("should execute command successfully and return output", async () => {
      const expectedOutput = "command output";

      const resultPromise = ui.runCommand({ cmd: "echo test" });

      // Wait a bit for the callback to be registered and temp file path to be captured
      await sleep(50);

      // Create the temporary file that runCommand expects
      if (tempFiles.length > 0) {
        const tempFile = tempFiles[tempFiles.length - 1];
        fs.writeFileSync(tempFile, expectedOutput, "utf-8");
      }

      // Now invoke the registered callback
      if (onDidEndTaskProcessListeners.length > 0) {
        onDidEndTaskProcessListeners[onDidEndTaskProcessListeners.length - 1]({
          execution: {
            task: {
              name: "Execute script action",
              source: "ms-teams-vscode-extension",
            },
          },
          exitCode: 0,
        });
      }

      const result = await resultPromise;

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value).to.equal(expectedOutput);
      }
    });

    it("should handle command execution error with non-zero exit code", async () => {
      const errorOutput = "error output";
      const windowShowErrorMessageStub = sinon.stub(window, "showErrorMessage");

      const resultPromise = ui.runCommand({ cmd: "failing command" });

      await sleep(50);

      // Create the temporary file with error output
      if (tempFiles.length > 0) {
        const tempFile = tempFiles[tempFiles.length - 1];
        fs.writeFileSync(tempFile, errorOutput, "utf-8");
      }

      if (onDidEndTaskProcessListeners.length > 0) {
        onDidEndTaskProcessListeners[onDidEndTaskProcessListeners.length - 1]({
          execution: {
            task: {
              name: "Execute script action",
              source: "ms-teams-vscode-extension",
            },
          },
          exitCode: 1,
        });
      }

      const result = await resultPromise;

      expect(result.isErr()).is.true;
      expect(windowShowErrorMessageStub.called).is.true;
    });

    it("should handle timeout error", async () => {
      const result = await ui.runCommand({
        cmd: "long running command",
        timeout: 50,
      });

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("ScriptTimeoutError");
      }
    });

    it("should execute command with working directory", async () => {
      const workingDir = "/test/path";
      const expectedOutput = "output";

      const resultPromise = ui.runCommand({
        cmd: "echo test",
        workingDirectory: workingDir,
      });

      await sleep(50);

      // Create the temporary file
      if (tempFiles.length > 0) {
        const tempFile = tempFiles[tempFiles.length - 1];
        fs.writeFileSync(tempFile, expectedOutput, "utf-8");
      }

      if (onDidEndTaskProcessListeners.length > 0) {
        onDidEndTaskProcessListeners[onDidEndTaskProcessListeners.length - 1]({
          execution: {
            task: {
              name: "Execute script action",
              source: "ms-teams-vscode-extension",
            },
          },
          exitCode: 0,
        });
      }

      const result = await resultPromise;

      expect(result.isOk()).is.true;
    });

    it("should execute command with environment variables", async () => {
      const customEnv = { TEST_VAR: "test_value" };
      const expectedOutput = "output";

      const resultPromise = ui.runCommand({
        cmd: "echo test",
        env: customEnv,
      });

      await sleep(50);

      // Create the temporary file
      if (tempFiles.length > 0) {
        const tempFile = tempFiles[tempFiles.length - 1];
        fs.writeFileSync(tempFile, expectedOutput, "utf-8");
      }

      if (onDidEndTaskProcessListeners.length > 0) {
        onDidEndTaskProcessListeners[onDidEndTaskProcessListeners.length - 1]({
          execution: {
            task: {
              name: "Execute script action",
              source: "ms-teams-vscode-extension",
            },
          },
          exitCode: 0,
        });
      }

      const result = await resultPromise;

      expect(result.isOk()).is.true;
    });

    it("should clean up temporary file on success", async () => {
      const expectedOutput = "output";

      const resultPromise = ui.runCommand({ cmd: "echo test" });

      await sleep(50);

      let createdTempFile = "";
      // Create the temporary file
      if (tempFiles.length > 0) {
        const tempFile = tempFiles[tempFiles.length - 1];
        createdTempFile = tempFile;
        fs.writeFileSync(tempFile, expectedOutput, "utf-8");
      }

      if (onDidEndTaskProcessListeners.length > 0) {
        onDidEndTaskProcessListeners[onDidEndTaskProcessListeners.length - 1]({
          execution: {
            task: {
              name: "Execute script action",
              source: "ms-teams-vscode-extension",
            },
          },
          exitCode: 0,
        });
      }

      const result = await resultPromise;

      expect(result.isOk()).is.true;
      // Verify that the temporary file was cleaned up
      expect(fs.existsSync(createdTempFile)).is.false;
    });

    it("should ignore non-matching task end events", async () => {
      const expectedOutput = "correct output";

      const resultPromise = ui.runCommand({ cmd: "echo test" });

      await sleep(30);

      // Trigger a non-matching task end event
      if (onDidEndTaskProcessListeners.length > 0) {
        onDidEndTaskProcessListeners[onDidEndTaskProcessListeners.length - 1]({
          execution: {
            task: {
              name: "Different task",
              source: "different-source",
            },
          },
          exitCode: 0,
        });
      }

      await sleep(30);

      // Create the temporary file for the correct task
      if (tempFiles.length > 0) {
        const tempFile = tempFiles[tempFiles.length - 1];
        fs.writeFileSync(tempFile, expectedOutput, "utf-8");
      }

      // Trigger the correct task end event
      if (onDidEndTaskProcessListeners.length > 0) {
        onDidEndTaskProcessListeners[onDidEndTaskProcessListeners.length - 1]({
          execution: {
            task: {
              name: "Execute script action",
              source: "ms-teams-vscode-extension",
            },
          },
          exitCode: 0,
        });
      }

      const result = await resultPromise;

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value).to.equal(expectedOutput);
      }
    });
  });
});
