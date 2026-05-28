import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import * as runIconHandlers from "../../src/debug/runIconHandler";
import * as appDefinitionUtils from "../../src/utils/appDefinitionUtils";
import { ok, TeamsAppManifest } from "@microsoft/teamsfx-api";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import {
  showLocalDebugMessage,
  ShowScaffoldingWarningSummary,
} from "../../src/utils/autoOpenHelper";
import VscodeLogInstance from "../../src/commonlib/log";
import * as readmeHandlers from "../../src/handlers/readmeHandlers";
import { manifestUtils, pluginManifestUtils } from "@microsoft/teamsfx-core";
import * as apiSpec from "@microsoft/teamsfx-core/build/component/generator/apiSpec/helper";

describe("autoOpenHelper", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("showLocalDebugMessage() - has local env", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("win32");
    sandbox.stub(fs, "pathExists").onFirstCall().resolves(true);
    const runLocalDebug = sandbox.stub(runIconHandlers, "selectAndDebug").resolves(ok(null));

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Debug",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.calledOnce);
    chai.assert.isTrue(runLocalDebug.called);
  });

  it("showLocalDebugMessage() - local env and non windows", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(fs, "pathExists").onFirstCall().resolves(true);
    const runLocalDebug = sandbox.stub(runIconHandlers, "selectAndDebug").resolves(ok(null));

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Not Debug",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.calledOnce);
    chai.assert.isFalse(runLocalDebug.called);
  });

  it("showLocalDebugMessage() - has local env and not click debug", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("win32");
    sandbox.stub(fs, "pathExists").onFirstCall().resolves(true);
    const runLocalDebug = sandbox.stub(runIconHandlers, "selectAndDebug").resolves(ok(null));

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve(undefined);
        }
      );

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.calledOnce);
    chai.assert.isFalse(runLocalDebug.called);
  });

  it("showLocalDebugMessage() - no local env", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("win32");
    sandbox.stub(fs, "pathExists").onFirstCall().resolves(false);

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Provision",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.called);
    chai.assert.isTrue(executeCommandStub.called);
  });

  it("showLocalDebugMessage() - no local env and non windows", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(appDefinitionUtils, "getAppName").resolves("");
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(fs, "pathExists").onFirstCall().resolves(false);

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Not provision",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.called);
    chai.assert.isTrue(executeCommandStub.notCalled);
  });

  it("showLocalDebugMessage() - no local env and not click provision", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("win32");
    sandbox.stub(fs, "pathExists").onFirstCall().resolves(false);

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve(undefined);
        }
      );
    const executeCommandStub = sandbox.stub(vscode.commands, "executeCommand");

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.called);
    chai.assert.isFalse(executeCommandStub.called);
  });

  it("showLocalDebugMessage() - generate an API key manually (TS - windows)", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("win32");
    sandbox
      .stub(fs, "pathExists")
      .onFirstCall()
      .resolves(true)
      .onSecondCall()
      .resolves(true)
      .onThirdCall()
      .resolves(false);
    const openReadMeHandlerStub = sandbox
      .stub(readmeHandlers, "openReadMeHandler")
      .resolves(ok(null));

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Open README",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.called);
    chai.assert.isTrue(openReadMeHandlerStub.called);
  });

  it("showLocalDebugMessage() - generate an API key manually (TS - windows) not clicked", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("win32");
    sandbox
      .stub(fs, "pathExists")
      .onFirstCall()
      .resolves(true)
      .onSecondCall()
      .resolves(true)
      .onThirdCall()
      .resolves(false);
    const openReadMeHandlerStub = sandbox
      .stub(readmeHandlers, "openReadMeHandler")
      .resolves(ok(null));

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Not Open README",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.called);
    chai.assert.isFalse(openReadMeHandlerStub.called);
  });

  it("showLocalDebugMessage() - generate an API key manually (TS - windows - non selection)", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("win32");
    sandbox
      .stub(fs, "pathExists")
      .onFirstCall()
      .resolves(true)
      .onSecondCall()
      .resolves(true)
      .onThirdCall()
      .resolves(false);
    const openReadMeHandlerStub = sandbox
      .stub(readmeHandlers, "openReadMeHandler")
      .resolves(ok(null));

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve(undefined);
        }
      );

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.called);
    chai.assert.isFalse(openReadMeHandlerStub.called);
  });

  it("showLocalDebugMessage() - generate an API key manually (JS - windows)", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("win32");
    sandbox
      .stub(fs, "pathExists")
      .onFirstCall()
      .resolves(true)
      .onSecondCall()
      .resolves(false)
      .onThirdCall()
      .resolves(true);
    const openReadMeHandlerStub = sandbox
      .stub(readmeHandlers, "openReadMeHandler")
      .resolves(ok(null));

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Open README",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.called);
    chai.assert.isTrue(openReadMeHandlerStub.called);
  });

  it("showLocalDebugMessage() - generate an API key manually (JS - non windows)", async () => {
    sandbox.stub(vscode.workspace, "workspaceFolders").value([{ uri: vscode.Uri.file("test") }]);
    sandbox.stub(vscode.workspace, "openTextDocument");
    sandbox.stub(process, "platform").value("linux");
    sandbox
      .stub(fs, "pathExists")
      .onFirstCall()
      .resolves(true)
      .onSecondCall()
      .resolves(false)
      .onThirdCall()
      .resolves(true);
    const openReadMeHandlerStub = sandbox
      .stub(readmeHandlers, "openReadMeHandler")
      .resolves(ok(null));

    sandbox.stub(globalState, "globalStateGet").callsFake(async (key: string) => {
      if (key === "ShowLocalDebugMessage") {
        return true;
      } else {
        return false;
      }
    });
    sandbox.stub(globalState, "globalStateUpdate");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    sandbox.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("test"));
    const showMessageStub = sandbox
      .stub(vscode.window, "showInformationMessage")
      .callsFake(
        (title: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]) => {
          return Promise.resolve({
            title: "Open README",
            run: (options as any).run,
          } as vscode.MessageItem);
        }
      );

    await showLocalDebugMessage();

    chai.assert.isTrue(showMessageStub.called);
    chai.assert.isTrue(openReadMeHandlerStub.called);
  });

  it("ShowScaffoldingWarningSummary() - copilot agents", async () => {
    const workspacePath = "/path/to/workspace";

    const manifest: TeamsAppManifest = {
      manifestVersion: "version",
      id: "mock-app-id",
      name: { short: "short-name" },
      description: { short: "", full: "" },
      version: "version",
      icons: { outline: "outline.png", color: "color.png" },
      accentColor: "#ffffff",
      developer: {
        privacyUrl: "",
        websiteUrl: "",
        termsOfUseUrl: "",
        name: "",
      },
      staticTabs: [
        {
          name: "name0",
          entityId: "index0",
          scopes: ["personal"],
          contentUrl: "localhost/content",
          websiteUrl: "localhost/website",
        },
      ],
      copilotAgents: {
        plugins: [
          {
            id: "plugin-id",
            file: "copilot-plugin-file",
          },
        ],
      },
    };
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
      .resolves(ok(["/path/to/api/spec"]));
    sandbox.stub(apiSpec, "generateScaffoldingSummary").resolves("fake summary");
    sandbox.stub(VscodeLogInstance, "info").callsFake((message: string) => {
      if (message !== "fake summary") {
        throw new Error(`Unexpected message: ${message}`);
      }
    });
    const fakeOutputChannel = {
      show: sandbox.stub().resolves(),
    };
    sandbox.stub(VscodeLogInstance, "outputChannel").value(fakeOutputChannel);
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent").resolves();
    // Call the function
    await ShowScaffoldingWarningSummary(workspacePath, "");
  });

  it("ShowScaffoldingWarningSummary() - copilot extensions", async () => {
    const workspacePath = "/path/to/workspace";

    const manifest: TeamsAppManifest = {
      manifestVersion: "version",
      id: "mock-app-id",
      name: { short: "short-name" },
      description: { short: "", full: "" },
      version: "version",
      icons: { outline: "outline.png", color: "color.png" },
      accentColor: "#ffffff",
      developer: {
        privacyUrl: "",
        websiteUrl: "",
        termsOfUseUrl: "",
        name: "",
      },
      staticTabs: [
        {
          name: "name0",
          entityId: "index0",
          scopes: ["personal"],
          contentUrl: "localhost/content",
          websiteUrl: "localhost/website",
        },
      ],
      copilotExtensions: {
        plugins: [
          {
            id: "plugin-id",
            file: "copilot-plugin-file",
          },
        ],
      },
    };
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(manifest));
    sandbox
      .stub(pluginManifestUtils, "getApiSpecFilePathFromTeamsManifest")
      .resolves(ok(["/path/to/api/spec"]));
    sandbox.stub(apiSpec, "generateScaffoldingSummary").resolves("fake summary");
    sandbox.stub(VscodeLogInstance, "info").callsFake((message: string) => {
      if (message !== "fake summary") {
        throw new Error(`Unexpected message: ${message}`);
      }
    });
    const fakeOutputChannel = {
      show: sandbox.stub().resolves(),
    };
    sandbox.stub(VscodeLogInstance, "outputChannel").value(fakeOutputChannel);
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent").resolves();
    // Call the function
    await ShowScaffoldingWarningSummary(workspacePath, "");
  });
});
