// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
"use strict";
import chai from "chai";
import spies from "chai-spies";
import * as sinon from "sinon";
import * as vscode from "vscode";

import * as globalVariables from "../../src/globalVariables";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { ReleaseNote } from "../../src/utils/releaseNote";
import * as versionUtil from "../../src/utils/versionUtil";
import { ExtensionContext } from "vscode";

chai.use(spies);
const spy = chai.spy;
function gloablStateKeys(): readonly string[] {
  return ["PrereleaseState.Version"];
}
function globalStateGet(key: string): string {
  return "0.0.0";
}
function globalStateUpdate(key: string, value: any): any {}
const reporterSpy = spy.interface({
  sendTelemetryEvent(
    eventName: string,
    properties?: { [p: string]: string },
    measurements?: { [p: string]: number }
  ): void {},
});
const ShowWhatIsNewNotification = "show-what-is-new-notification";
describe("Release Note", () => {
  afterEach(() => {
    sinon.restore();
  });

  describe("stable version shows changelog", () => {
    const sandbox = sinon.createSandbox();
    let context: vscode.ExtensionContext;
    let telemetryStub: sinon.SinonStub;
    const mockGlobalState: vscode.Memento = {
      keys: gloablStateKeys,
      get: globalStateGet,
      update: globalStateUpdate,
    };
    beforeEach(() => {
      context = {
        subscriptions: [],
        globalState: mockGlobalState,
      } as unknown as vscode.ExtensionContext;
      sandbox.stub(versionUtil, "getExtensionId").returns("");
      sandbox.stub(vscode.extensions, "getExtension").returns({
        packageJSON: { version: "5.0.0" },
        id: "",
        extensionPath: "",
        isActive: true,
        exports: {},
        extensionKind: vscode.ExtensionKind.UI,
        extensionUri: vscode.Uri.parse("https://www.test.com"),
        activate(): Thenable<void> {
          return Promise.resolve();
        },
      });
      telemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
    });
    afterEach(() => {
      sandbox.restore();
    });
    it("show changelog notification happy path", async () => {
      const contextSpy = sandbox.spy(context.globalState, "update");
      sandbox.stub(context.globalState, "get").returns("4.99.0");
      let title = "";
      sandbox
        .stub(vscode.window, "showInformationMessage")
        .callsFake((_message: string, option: any, ...items: vscode.MessageItem[]) => {
          title = option.title;
          return Promise.resolve(option);
        });
      const instance = new ReleaseNote(context);
      await instance.show();
      chai.assert(title === "Changelog");
      chai.assert(contextSpy.callCount == 2);
      chai.assert(telemetryStub.calledWith("show-what-is-new-notification"));
    });
    it("should not show changelog if button is not clicked", async () => {
      const contextSpy = sandbox.spy(context.globalState, "update");
      sandbox.stub(context.globalState, "get").returns("4.99.0");
      sandbox.stub(vscode.window, "showInformationMessage").resolves(undefined);
      const instance = new ReleaseNote(context);
      await instance.show();
      chai.assert(contextSpy.callCount == 2);
      chai.assert(telemetryStub.calledOnce);
    });
    it("should not show changelog when version is not changed", async () => {
      const contextSpy = sandbox.spy(context.globalState, "update");
      sandbox.stub(context.globalState, "get").returns("5.0.0");
      sandbox.stub(vscode.window, "showInformationMessage").resolves();
      const instance = new ReleaseNote(context);
      await instance.show();
      sinon.assert.calledOnce(contextSpy);
      chai.assert(telemetryStub.notCalled);
    });
    it("should not show changelog when it's a fresh install", async () => {
      const contextSpy = sandbox.spy(context.globalState, "update");
      sandbox.stub(context.globalState, "get").returns(undefined);
      sandbox.stub(vscode.window, "showInformationMessage").resolves();
      const instance = new ReleaseNote(context);
      await instance.show();
      sinon.assert.calledOnce(contextSpy);
      chai.assert(telemetryStub.notCalled);
    });
  });

  describe("prerelease version shows prerelease note", () => {
    const sandbox = sinon.createSandbox();
    let context: ExtensionContext;
    const mockGlobalState: vscode.Memento = {
      keys: gloablStateKeys,
      get: globalStateGet,
      update: globalStateUpdate,
    };
    before(() => {
      chai.util.addProperty(ExtTelemetry, "reporter", () => reporterSpy);
    });
    beforeEach(() => {
      sandbox.stub(vscode.workspace, "openTextDocument").resolves();
      sandbox.stub(vscode.commands, "executeCommand").resolves();
      context = {
        subscriptions: [],
        globalState: mockGlobalState,
      } as unknown as ExtensionContext;
    });
    afterEach(() => {
      sandbox.restore();
    });
    it("success", async () => {
      sandbox.stub(vscode.extensions, "getExtension").returns({
        packageJSON: { version: "5.1.2023072000" },
        id: "",
        extensionPath: "",
        isActive: true,
        exports: {},
        extensionKind: vscode.ExtensionKind.UI,
        extensionUri: vscode.Uri.parse("https://www.test.com"),
        activate(): Thenable<void> {
          return Promise.resolve();
        },
      });
      sandbox.stub(context.globalState, "get").returns("5.0.1");
      const instance = new ReleaseNote(context);
      const spyChecker = sandbox.spy(context.globalState, "update");
      await instance.show();
      chai.assert(spyChecker.callCount == 1);
      chai
        .expect(reporterSpy.sendTelemetryEvent)
        .to.have.been.called.with(ShowWhatIsNewNotification);
      spyChecker.restore();
    });
    it("returns prerelease version undefined", async () => {
      sandbox.stub(vscode.extensions, "getExtension").returns({
        packageJSON: { version: "5.1.2023072000" },
        id: "",
        extensionPath: "",
        isActive: true,
        exports: {},
        extensionKind: vscode.ExtensionKind.UI,
        extensionUri: vscode.Uri.parse("https://www.test.com"),
        activate(): Thenable<void> {
          return Promise.resolve();
        },
      });
      sandbox.stub(context.globalState, "get").returns(undefined);
      const instance = new ReleaseNote(context);
      const spyChecker = sandbox.spy(context.globalState, "update");
      chai
        .expect(reporterSpy.sendTelemetryEvent)
        .to.have.been.called.with(ShowWhatIsNewNotification);
      await instance.show();
      chai.assert(spyChecker.callCount == 1);
      spyChecker.restore();
    });
    it("has same version", async () => {
      sandbox.stub(vscode.extensions, "getExtension").returns({
        packageJSON: { version: "5.1.2023072000" },
        id: "",
        extensionPath: "",
        isActive: true,
        exports: {},
        extensionKind: vscode.ExtensionKind.UI,
        extensionUri: vscode.Uri.parse("https://www.test.com"),
        activate(): Thenable<void> {
          return Promise.resolve();
        },
      });
      sandbox.stub(context.globalState, "get").returns("5.1.2023072000");
      const instance = new ReleaseNote(context);
      const spyChecker = sandbox.spy(context.globalState, "update");
      await instance.show();
      chai.assert(spyChecker.callCount == 0);
      spyChecker.restore();
    });
    it("has undefined version", async () => {
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      sandbox.stub(context.globalState, "get").returns("5.0.0");
      const instance = new ReleaseNote(context);
      const spyChecker = sandbox.spy(context.globalState, "update");
      await instance.show();
      chai.assert(spyChecker.callCount == 0);
      spyChecker.restore();
    });
  });
});
