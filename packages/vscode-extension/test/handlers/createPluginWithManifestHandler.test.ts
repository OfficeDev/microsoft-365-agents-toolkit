import * as sinon from "sinon";
import * as chai from "chai";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import * as vsc_ui from "../../src/qm/vsc_ui";
import * as vscode from "vscode";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import * as localizeUtils from "@microsoft/teamsfx-core/build/common/localizeUtils";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { MockCore } from "../mocks/mockCore";
import { createPluginWithManifest } from "../../src/handlers/createPluginWithManifestHandler";
import * as workspaceUtils from "../../src/utils/workspaceUtils";
import { err, UserError } from "@microsoft/teamsfx-api";

describe("createPluginWithManifestHandler", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("happy path: successfullly create plugin project", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        lastCommand: "createPluginWithManifest",
      },
    ]);
    chai.assert.isTrue(res.isOk());
    chai.assert.isTrue(openFolder.calledOnce);
  });

  it("happy path: successfullly create declarative copilot project", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        lastCommand: "createDeclarativeCopilotWithManifest",
      },
    ]);
    chai.assert.isTrue(res.isOk());
    chai.assert.isTrue(openFolder.calledOnce);
  });

  it("happy path: successfullly create plugin project with folder path", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        lastCommand: "createPluginWithManifest",
      },
      "folder",
    ]);
    chai.assert.isTrue(res.isOk());
    chai.assert.isTrue(openFolder.calledOnce);
  });

  it("should throw error if args length is less than 3", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest(["specPath"]);
    chai.assert.isTrue(res.isErr());
    chai.assert.isTrue(openFolder.notCalled);
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }
  });

  it("should throw error if args length is bigger than 4", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        lastCommand: "createPluginWithManifest",
      },
      "folder",
      "extra",
    ]);
    chai.assert.isTrue(res.isErr());
    chai.assert.isTrue(openFolder.notCalled);
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }
  });

  it("should throw error if command name missing", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        test: "test",
      },
    ]);
    chai.assert.isTrue(res.isErr());
    chai.assert.isTrue(openFolder.notCalled);
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }
  });

  it("should throw error if command name invalid", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        lastCommand: "test",
      },
    ]);
    chai.assert.isTrue(res.isErr());
    chai.assert.isTrue(openFolder.notCalled);
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }
  });

  it("should throw error if args is empty", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest([]);
    chai.assert.isTrue(res.isErr());
    chai.assert.isTrue(openFolder.notCalled);
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }
  });

  it("should throw error if core return error", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    sandbox
      .stub(globalVariables.core, "createProject")
      .resolves(err(new UserError("core", "fakeError", "fakeErrorMessage")));
    const openFolder = sandbox.stub(workspaceUtils, "openFolder").resolves();
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        lastCommand: "createPluginWithManifest",
      },
    ]);
    chai.assert.isTrue(res.isErr());
    chai.assert.isTrue(openFolder.notCalled);
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "fakeError");
    }
  });

  it("happy path: add plugin", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        lastCommand: "addPlugin",
        manifestPath: "manifestPath",
      },
    ]);
    chai.assert.isTrue(res.isOk());
  });

  it("should throw error if add plugin core return error", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    sandbox
      .stub(globalVariables.core, "addPlugin")
      .resolves(err(new UserError("core", "fakeError", "fakeErrorMessage")));
    const res = await createPluginWithManifest([
      "specPath",
      "pluginManifestPath",
      {
        lastCommand: "addPlugin",
        manifestPath: "manifestPath",
      },
    ]);
    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "fakeError");
    }
  });
});
