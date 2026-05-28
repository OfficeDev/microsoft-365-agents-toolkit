import * as sinon from "sinon";
import * as chai from "chai";
import fs from "fs-extra";
import * as globalVariables from "../../src/globalVariables";
import * as vsc_ui from "../../src/qm/vsc_ui";
import * as vscode from "vscode";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import * as localizeUtils from "@microsoft/teamsfx-core/build/common/localizeUtils";
import { kiotaRegenerate } from "../../src/handlers/kiotaRegenerateHandler";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { MockCore } from "../mocks/mockCore";
import * as workspaceUtils from "../../src/utils/workspaceUtils";
import { err, UserError } from "@microsoft/teamsfx-api";

describe("kiotaRegenerateHandler", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("happy path: successfully regenerate", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const res = await kiotaRegenerate(["specPath", "pluginManifestPath"]);
    chai.assert.isTrue(res.isOk());
  });

  it("should throw error if args length not equals 2", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    const res = await kiotaRegenerate(["specPath"]);
    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }

    await kiotaRegenerate(["specPath", "pluginManifestPath", "test"]);
    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }

    await kiotaRegenerate([]);
    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "invalidParameter");
    }
  });

  it("should throw error if core return error", async () => {
    const core = new MockCore();
    sandbox.stub(globalVariables, "core").value(core);
    sandbox
      .stub(globalVariables.core, "kiotaRegenerate")
      .resolves(err(new UserError("core", "fakeError", "fakeErrorMessage")));
    const res = await kiotaRegenerate(["specPath", "pluginManifestPath"]);
    chai.assert.isTrue(res.isErr());
    if (res.isErr()) {
      chai.assert.equal(res.error.name, "fakeError");
    }
  });
});
