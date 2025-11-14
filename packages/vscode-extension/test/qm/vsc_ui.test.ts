// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { expect } from "chai";
import * as sinon from "sinon";
import { stubInterface } from "ts-sinon";
import {
  commands,
  DiagnosticCollection,
  Disposable,
  ExtensionContext,
  languages,
  QuickInputButton,
  QuickPick,
  TextDocument,
  window,
  workspace,
} from "vscode";

import {
  err,
  ok,
  SelectFileConfig,
  SelectFolderConfig,
  SingleFileOrInputConfig,
  SingleSelectConfig,
  UserError,
} from "@microsoft/teamsfx-api";
import { featureFlagManager } from "@microsoft/teamsfx-core";
import { FxQuickPickItem, sleep, UserCancelError } from "@microsoft/vscode-ui";
import * as globalVariables from "../../src/globalVariables";
import { VsCodeUI } from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";

describe("UI Unit Tests", async () => {
  afterEach(() => {
    sinon.restore();
  });

  describe("Manually", () => {
    it("Show Progress 2", async function (this: Mocha.Context) {
      this.timeout(0);
      const VS_CODE_UI = new VsCodeUI({} as ExtensionContext);
      const handler = VS_CODE_UI.createProgressBar("Test Progress Bar", 3);

      await handler.start("Prepare");
      await sleep(2 * 1000);

      await handler.next("First step");
      await sleep(2 * 1000);

      await handler.next("Second step");
      await sleep(2 * 1000);

      await handler.next("Third step");
      await sleep(2 * 1000);

      await handler.end(true);
    });
  });

  describe("Select Folder", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("has returns default folder", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      // const telemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFolder(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("default folder");
      }
      // expect(
      //   telemetryStub.calledOnceWith("select-folder", {
      //     "selected-option": "default",
      //   })
      // ).is.true;
    });

    it("has returns user cancel", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      sandbox.stub(window, "showOpenDialog").resolves(undefined);
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFolder(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error instanceof UserCancelError).is.true;
      }
    });
  });

  describe("Select File", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("has returns default file", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFile(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("default file");
      }
    });

    it("has returns user cancel", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      sandbox.stub(window, "showOpenDialog").resolves(undefined);
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFile(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error instanceof UserCancelError).is.true;
      }
    });

    it("has returns item in possible files", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFile(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
      }
    });

    it("has returns invalid input item id", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
    });

    it("selects a file which pass validation", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const res = await ui.selectFile(config);
      expect(res.isOk()).is.true;
    });

    it("selects a file with error thrown when validating result", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });

      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const res = await ui.selectFile(config);
      expect(res.isErr()).is.true;
    });
  });

  describe("Open File", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("open the preview of Markdown file", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
      sandbox.stub(workspace, "openTextDocument").resolves({} as TextDocument);
      let executedCommand = "";
      sandbox.stub(commands, "executeCommand").callsFake((command: string, ...args: any[]) => {
        executedCommand = command;
        return Promise.resolve();
      });
      const showTextStub = sandbox.stub(window, "showTextDocument");

      const result = await ui.openFile("test.md");

      expect(result.isOk()).is.true;
      expect(showTextStub.calledOnce).to.be.false;
      expect(executedCommand).to.equal("markdown.showPreview");
    });
  });

  describe("single select", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("select success with validation", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
      }
    });

    it("select fail with validation", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(window, "createQuickPick").callsFake(() => {
        return mockQuickPick;
      });
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectOption(config);

      expect(result.isErr()).is.true;
    });

    it("loads dynamic options in a short time", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
      }
    });

    it("loads dynamic option in a short time and auto select", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
      }
      sandbox.restore();
    });

    it("loads dynamic options in a short time and shows", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectOption(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("1");
        expect(mockQuickPick.show.called).is.true;
      }
      sandbox.restore();
    });

    it("loads dynamic option in a long time and shows", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

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
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("selects local file successfully", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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

      sandbox
        .stub(VsCodeUI.prototype, "selectFile")
        .resolves(ok({ type: "success", result: "file" }));
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFileOrInput(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("file");
      }
    });

    it("selects local file error", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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

      sandbox
        .stub(VsCodeUI.prototype, "selectFile")
        .resolves(err(new UserError("source", "name", "msg", "msg")));
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFileOrInput(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("name");
      }
    });

    it("inputs a value sucessfully", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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

      sandbox
        .stub(VsCodeUI.prototype, "selectFile")
        .resolves(ok({ type: "success", result: "input" }));
      sandbox
        .stub(VsCodeUI.prototype, "inputText")
        .resolves(ok({ type: "success", result: "testUrl" }));
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFileOrInput(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("testUrl");
      }
    });

    it("inputs a value error", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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

      sandbox
        .stub(VsCodeUI.prototype, "selectFile")
        .resolves(ok({ type: "success", result: "input" }));
      sandbox
        .stub(VsCodeUI.prototype, "inputText")
        .resolves(err(new UserError("source", "name", "msg", "msg")));
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFileOrInput(config);

      expect(result.isErr()).is.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("name");
      }
    });

    it("inputs a value back and then sucessfully", async function (this: Mocha.Context) {
      const ui = new VsCodeUI({} as ExtensionContext);
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

      sandbox
        .stub(VsCodeUI.prototype, "selectFile")
        .resolves(ok({ type: "success", result: "input" }));
      sandbox
        .stub(VsCodeUI.prototype, "inputText")
        .onFirstCall()
        .resolves(ok({ type: "back" }))
        .onSecondCall()
        .resolves(ok({ type: "success", result: "testUrl" }));
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

      const result = await ui.selectFileOrInput(config);

      expect(result.isOk()).is.true;
      if (result.isOk()) {
        expect(result.value.result).to.equal("testUrl");
      }
    });
  });

  describe("showDiagnosticInfo", () => {
    const sandbox = sinon.createSandbox();
    let collection: DiagnosticCollection | undefined;

    afterEach(() => {
      sandbox.restore();
      globalVariables.setDiagnosticCollection(undefined as unknown as DiagnosticCollection);
    });

    it("do nothing if feature flag is disabled", () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const ui = new VsCodeUI({} as ExtensionContext);
      ui.showDiagnosticInfo([]);
    });

    it("show diagnostics first time if feature flag is enabled", () => {
      const records: [string, { message: string }][] = [];
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      collection = {
        set: (filePath: string, diag: { message: string }) => {
          records.push([filePath, diag]);
        },
      } as unknown as DiagnosticCollection;

      sandbox.stub(languages, "createDiagnosticCollection").returns(collection as any);
      const ui = new VsCodeUI({} as ExtensionContext);

      ui.showDiagnosticInfo([
        {
          startIndex: 0,
          startLine: 1,
          endIndex: 10,
          endLine: 10,
          severity: 2,
          filePath: "test",
          message: "error",
        },
      ]);

      expect(globalVariables.diagnosticCollection).not.undefined;
      expect(records.length).equals(1);
    });

    it("show diagnostics not first time if feature flag is enabled", () => {
      const records: [string, { message: string }][] = [];
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      collection = {
        clear: () => {
          return;
        },
        set: (filePath: string, diag: { message: string }) => {
          records.push([filePath, diag]);
        },
      } as unknown as DiagnosticCollection;

      globalVariables.setDiagnosticCollection(collection);
      const ui = new VsCodeUI({} as ExtensionContext);

      ui.showDiagnosticInfo([
        {
          startIndex: 0,
          startLine: 1,
          endIndex: 10,
          endLine: 10,
          severity: 2,
          filePath: "test",
          message: "error",
          code: {
            value: "test",
            link: "https://test.com",
          },
        },
      ]);

      expect(globalVariables.diagnosticCollection).not.undefined;
      expect(records.length).equals(1);
    });
  });
});
