// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { openOneDriveSharePointUrlHandler } from "../../src/handlers/openOneDriveSharePointUrlHandler";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../../src/telemetry/extTelemetryEvents";

describe("openOneDriveSharePointUrlHandler", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(vsc_ui, "VS_CODE_UI").value(new vsc_ui.VsCodeUI(<vscode.ExtensionContext>{}));
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should not open URL when no active editor", async () => {
    // Stub window.activeTextEditor to return null
    sandbox.stub(vscode.window, "activeTextEditor").value(undefined);
    const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await openOneDriveSharePointUrlHandler(["https://example.com"]);

    // Verify telemetry was sent
    chai.assert.isTrue(
      sendTelemetryStub.calledOnceWith(TelemetryEvent.OpenOneDriveSharePointUrlStart, {
        [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Other,
      })
    );
    // Verify URL was not opened
    chai.assert.isTrue(openUrlStub.notCalled);
  });

  it("should open URL when active editor exists", async () => {
    // Mock active editor
    sandbox.stub(vscode.window, "activeTextEditor").value({} as vscode.TextEditor);
    const testUrl = "https://example.com";
    const openUrlStub = sandbox.stub(vsc_ui.VS_CODE_UI, "openUrl");
    const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");

    await openOneDriveSharePointUrlHandler([testUrl]);

    // Verify telemetry was sent
    chai.assert.isTrue(
      sendTelemetryStub.calledOnceWith(TelemetryEvent.OpenOneDriveSharePointUrlStart, {
        [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Other,
      })
    );
    // Verify URL was opened with correct parameter
    chai.assert.isTrue(openUrlStub.calledOnceWith(testUrl));
  });
});
