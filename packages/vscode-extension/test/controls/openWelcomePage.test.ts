import * as chai from "chai";
import * as vscode from "vscode";
import { vi } from "vitest";
import {
  openWelcomePageAfterExtensionInstallation,
  openWelcomePageDeps,
} from "../../src/controls/openWelcomePage";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";

describe("openWelcomePageAfterExtensionInstallation()", () => {
  it("will not open welcome page if shown before", async () => {
    vi.spyOn(openWelcomePageDeps, "globalStateGet").mockResolvedValue(true);
    const globalStateUpdateStub = vi.spyOn(openWelcomePageDeps, "globalStateUpdate");

    await openWelcomePageAfterExtensionInstallation();

    chai.assert.isTrue(globalStateUpdateStub.notCalled);
  });

  it("opens welcome page if not shown before", async () => {
    vi.spyOn(openWelcomePageDeps, "globalStateGet").mockResolvedValue(false);
    const globalStateUpdateStub = vi.spyOn(openWelcomePageDeps, "globalStateUpdate");
    vi.spyOn(ExtTelemetry, "sendTelemetryEvent");
    const executeCommandStub = vi.spyOn(vscode.commands, "executeCommand").mockResolvedValue();

    await openWelcomePageAfterExtensionInstallation();

    chai.assert.isTrue(globalStateUpdateStub.calledOnce);
    chai.assert.isTrue(executeCommandStub.calledTwice);
  });
});
