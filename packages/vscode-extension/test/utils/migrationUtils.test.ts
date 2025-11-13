import * as chai from "chai";
import * as sinon from "sinon";
import { ExtensionContext } from "vscode";
import * as migrationUtils from "../../src/utils/migrationUtils";
import * as environmentUtils from "../../src/utils/systemEnvUtils";
import * as globalVariables from "../../src/globalVariables";
import { Inputs, UserError, err, ok } from "@microsoft/teamsfx-api";
import { MockCore } from "../mocks/mockCore";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { VsCodeUI } from "../../src/qm/vsc_ui";

describe("migrationUtils", () => {
  const sandbox = sinon.createSandbox();

  describe("triggerV3Migration", () => {
    beforeEach(() => {
      sandbox.stub(environmentUtils, "getSystemInputs").returns({
        locale: "en-us",
        platform: "vsc",
        projectPath: undefined,
        vscodeEnv: "local",
      } as Inputs);
      sandbox.stub(globalVariables, "core").value(new MockCore());
    });

    afterEach(async () => {
      sandbox.restore();
    });

    it("Stop debugging if phantomMigrationV3() returns error", async () => {
      const error = new UserError(
        "test source",
        "test name",
        "test message",
        "test displayMessage"
      );
      const phantomMigrationV3Stub = sandbox
        .stub(globalVariables.core, "phantomMigrationV3")
        .resolves(err(error));
      migrationUtils.triggerV3Migration().catch((e) => {
        chai.assert.equal(e, error);
      });
      chai.assert.isTrue(
        phantomMigrationV3Stub.calledOnceWith({
          locale: "en-us",
          platform: "vsc",
          projectPath: undefined,
          vscodeEnv: "local",
          stage: "debug",
        } as Inputs)
      );
    });

    it("Reload window if phantomMigrationV3() returns ok", async () => {
      const phantomMigrationV3Stub = sandbox
        .stub(globalVariables.core, "phantomMigrationV3")
        .resolves(ok(undefined));
      sandbox.stub(vsc_ui, "VS_CODE_UI").value(new VsCodeUI({} as ExtensionContext));
      const vscUIReloadStub = sandbox.stub(vsc_ui.VS_CODE_UI, "reload").resolves();
      await migrationUtils.triggerV3Migration();
      chai.assert.isTrue(
        phantomMigrationV3Stub.calledOnceWith({
          locale: "en-us",
          platform: "vsc",
          projectPath: undefined,
          vscodeEnv: "local",
          stage: "debug",
        } as Inputs)
      );
      chai.assert.isTrue(vscUIReloadStub.calledOnce);
    });
  });
});
