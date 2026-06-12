import * as sinon from "sinon";
import * as chai from "chai";
import * as vscode from "vscode";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import { GlobalKey } from "../../src/constants";
import { globalStateUtilsDeps, updateAutoOpenGlobalKey } from "../../src/utils/globalStateUtils";

describe("GlobalStateUtils", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("updateAutoOpenGlobalKey", async () => {
    sandbox.stub(globalStateUtilsDeps, "isTriggerFromWalkThrough").returns(true);
    sandbox.stub(globalStateUtilsDeps, "checkIsSPFx").returns(true);
    sandbox.stub(projectSettingsHelper, "isValidOfficeAddInProject").returns(false);
    const globalStateUpdateStub = sandbox.stub(globalStateUtilsDeps, "globalStateUpdate");

    await updateAutoOpenGlobalKey(false, vscode.Uri.file("test"), [
      { type: "type", content: "content" },
    ]);

    chai.assert.isTrue(globalStateUpdateStub.calledWith(GlobalKey.OpenWalkThrough, true));
    chai.assert.isTrue(globalStateUpdateStub.calledWith(GlobalKey.OpenReadMe, ""));
    chai.assert.isTrue(
      globalStateUpdateStub.calledWith(GlobalKey.CreateWarnings, sinon.match.string)
    );
    chai.assert.isTrue(globalStateUpdateStub.calledWith(GlobalKey.AutoInstallDependency, true));
  });
});
