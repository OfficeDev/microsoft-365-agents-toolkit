import { TeamsAppManifest } from "@microsoft/teamsfx-api";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/src/common/projectSettingsHelper";
import * as chai from "chai";
import path from "path";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import { manifestListener } from "../../src/manifestListener";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import TreeViewManagerInstance from "../../src/treeview/treeViewManager";

describe("registerManifestListener", () => {
  const sandbox = sinon.createSandbox();
  let clock: sinon.SinonFakeTimers;

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent").returns();
  });

  afterEach(() => {
    sandbox.restore();
    if (clock) {
      clock.restore();
    }
  });
  it("successfully refresh item", async () => {
    clock = sandbox.useFakeTimers();
    let handler = async (event: any) => {};
    sandbox.stub(projectSettingsHelper, "isValidProjectV3").returns(true);
    sandbox.stub(vscode.workspace, "onDidSaveTextDocument").callsFake((listener: any) => {
      handler = listener;
      return new vscode.Disposable(() => {
        return;
      });
    });
    sandbox.stub(globalVariables, "isDeclarativeCopilotApp").value(false);
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("."));
    sandbox
      .stub(globalVariables, "updateIsDeclarativeCopilotApp")
      .onFirstCall()
      .returns(true)
      .onSecondCall()
      .returns(false);
    sandbox.stub(TreeViewManagerInstance, "updateDevelopmentTreeView").returns();

    const fakeDocument = {
      fileName: path.join(vscode.Uri.file(".").fsPath, "appPackage", "manifest.json"),
      getText: () => {
        return JSON.stringify(new TeamsAppManifest());
      },
    };

    manifestListener();
    let job = handler(fakeDocument);

    await clock.tickAsync(5000);
    let res = await job;
    chai.assert.isTrue(res);

    job = handler(fakeDocument);
    await clock.tickAsync(5000);
    res = await job;
    chai.assert.isFalse(res);
  });

  it("abort previous one", async () => {
    clock = sandbox.useFakeTimers();
    let handler = async (event: any) => {};
    sandbox.stub(projectSettingsHelper, "isValidProjectV3").returns(true);
    sandbox.stub(vscode.workspace, "onDidSaveTextDocument").callsFake((listener: any) => {
      handler = listener;
      return new vscode.Disposable(() => {
        return;
      });
    });
    sandbox.stub(globalVariables, "isDeclarativeCopilotApp").value(false);
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("."));
    sandbox
      .stub(globalVariables, "updateIsDeclarativeCopilotApp")
      .onFirstCall()
      .returns(true)
      .onSecondCall()
      .returns(false);
    sandbox.stub(TreeViewManagerInstance, "updateDevelopmentTreeView").returns();

    const fakeDocument = {
      fileName: path.join(vscode.Uri.file(".").fsPath, "appPackage", "manifest.json"),
      getText: () => {
        return JSON.stringify(new TeamsAppManifest());
      },
    };

    manifestListener();
    const job1 = handler(fakeDocument);
    await clock.tickAsync(1000);
    const job2 = handler(fakeDocument);

    await clock.tickAsync(5000);
    const res1 = await job1;
    const res2 = await job2;

    chai.assert.isUndefined(res1);
    chai.assert.isTrue(res2);
  });

  it("not run if invalid project", async () => {
    clock = sandbox.useFakeTimers();
    let handler = async (event: any) => {};
    sandbox.stub(projectSettingsHelper, "isValidProjectV3").returns(false);
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("."));
    sandbox.stub(vscode.workspace, "onDidSaveTextDocument").callsFake((listener: any) => {
      handler = listener;
      return new vscode.Disposable(() => {
        return;
      });
    });

    const fakeDocument = {
      fileName: path.join(vscode.Uri.file(".").fsPath, "appPackage", "manifest.json"),
      getText: () => {
        return JSON.stringify(new TeamsAppManifest());
      },
    };

    manifestListener();
    const res = await handler(fakeDocument);

    chai.assert.isUndefined(res);
  });

  it("not run if empty workspace", async () => {
    clock = sandbox.useFakeTimers();
    let handler = async (event: any) => {};
    sandbox.stub(globalVariables, "workspaceUri").value("");
    sandbox.stub(projectSettingsHelper, "isValidProjectV3").returns(false);
    sandbox.stub(vscode.workspace, "onDidSaveTextDocument").callsFake((listener: any) => {
      handler = listener;
      return new vscode.Disposable(() => {
        return;
      });
    });

    const fakeDocument = {
      fileName: path.join(vscode.Uri.file(".").fsPath, "appPackage", "manifest.json"),
      getText: () => {
        return JSON.stringify(new TeamsAppManifest());
      },
    };

    manifestListener();
    const res = await handler(fakeDocument);

    chai.assert.isUndefined(res);
  });

  it("not run if not default app manifest", async () => {
    clock = sandbox.useFakeTimers();
    let handler = async (event: any) => {};
    sandbox.stub(globalVariables, "workspaceUri").value(".");
    sandbox.stub(projectSettingsHelper, "isValidProjectV3").returns(false);
    sandbox.stub(vscode.workspace, "onDidSaveTextDocument").callsFake((listener: any) => {
      handler = listener;
      return new vscode.Disposable(() => {
        return;
      });
    });

    const fakeDocument = {
      fileName: path.join(vscode.Uri.file(".").fsPath, "appPackage", "unknown.json"),
      getText: () => {
        return JSON.stringify(new TeamsAppManifest());
      },
    };

    manifestListener();
    const res = await handler(fakeDocument);

    chai.assert.isUndefined(res);
  });
});
