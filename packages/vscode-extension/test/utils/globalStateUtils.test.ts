import * as chai from "chai";
import * as vscode from "vscode";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import { GlobalKey } from "../../src/constants";
import { globalStateUtilsDeps, updateAutoOpenGlobalKey } from "../../src/utils/globalStateUtils";
import { vi } from "vitest";

describe("GlobalStateUtils", () => {
  it("updateAutoOpenGlobalKey", async () => {
    vi.spyOn(globalStateUtilsDeps, "isTriggerFromWalkThrough").mockReturnValue(true);
    vi.spyOn(globalStateUtilsDeps, "checkIsSPFx").mockReturnValue(true);
    vi.spyOn(projectSettingsHelper, "isValidOfficeAddInProject").mockReturnValue(false);
    const globalStateUpdateStub = vi.spyOn(globalStateUtilsDeps, "globalStateUpdate");

    await updateAutoOpenGlobalKey(false, vscode.Uri.file("test"), [
      { type: "type", content: "content" },
    ]);

    chai.assert.isTrue(globalStateUpdateStub.calledWith(GlobalKey.OpenWalkThrough, true));
    chai.assert.isTrue(globalStateUpdateStub.calledWith(GlobalKey.OpenReadMe, ""));
    chai.assert.isTrue(
      globalStateUpdateStub.calledWith(GlobalKey.CreateWarnings, expect.any(String))
    );
    chai.assert.isTrue(globalStateUpdateStub.calledWith(GlobalKey.AutoInstallDependency, true));
  });
});
