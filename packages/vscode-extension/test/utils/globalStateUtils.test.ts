import * as globalState from "@microsoft/teamsfx-core/build/src/common/globalState";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/src/common/projectSettingsHelper";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import { updateAutoOpenGlobalKey } from "../../src/utils/globalStateUtils";
import * as telemetryUtils from "../../src/utils/telemetryUtils";

describe("GlobalStateUtils", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("updateAutoOpenGlobalKey", async () => {
    sandbox.stub(telemetryUtils, "isTriggerFromWalkThrough").returns(true);
    sandbox.stub(globalVariables, "checkIsSPFx").returns(true);
    sandbox.stub(projectSettingsHelper, "isValidOfficeAddInProject").returns(false);
    const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate");

    await updateAutoOpenGlobalKey(false, vscode.Uri.file("test"), [
      { type: "type", content: "content" },
    ]);

    chai.assert.isTrue(globalStateUpdateStub.callCount === 4);
  });
});
