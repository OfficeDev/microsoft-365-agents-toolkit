import * as globalState from "@microsoft/teamsfx-core/build/src/common/globalState";
import * as chai from "chai";
import * as sinon from "sinon";
import { Uri, commands } from "vscode";
import { GlobalKey } from "../../src/constants";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { openOfficeDevFolder } from "../../src/utils/workspaceUtils";

describe("WorkspaceUtils", () => {
  describe("openOfficeDevFolder", () => {
    const sandbox = sinon.createSandbox();

    beforeEach(() => {
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(commands, "executeCommand");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("triggered from walkthrough with local debug message and warnings", async () => {
      const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate");
      const warnings = [{ type: "type", content: "content" }];
      await openOfficeDevFolder(Uri.parse("fakePath"), true, warnings, ["WalkThrough"]);
      chai.expect(globalStateUpdateStub.callCount).equals(5);
      chai
        .expect(globalStateUpdateStub.getCall(0).args)
        .deep.equals([GlobalKey.OpenWalkThrough, false]);
      chai
        .expect(globalStateUpdateStub.getCall(1).args)
        .deep.equals([GlobalKey.AutoInstallDependency, true]);
      chai.expect(globalStateUpdateStub.getCall(2).args).deep.equals([GlobalKey.OpenReadMe, ""]);
      chai
        .expect(globalStateUpdateStub.getCall(3).args)
        .deep.equals([GlobalKey.ShowLocalDebugMessage, true]);
      chai
        .expect(globalStateUpdateStub.getCall(4).args)
        .deep.equals([GlobalKey.CreateWarnings, JSON.stringify(warnings)]);
    });

    it("not triggered from walkthrough with no local debug message and warnings", async () => {
      const globalStateUpdateStub = sandbox.stub(globalState, "globalStateUpdate");
      await openOfficeDevFolder(Uri.parse("fakePath"), false, undefined);
      chai.expect(globalStateUpdateStub.callCount).equals(3);
      chai
        .expect(globalStateUpdateStub.getCall(0).args)
        .deep.equals([GlobalKey.OpenWalkThrough, false]);
      chai
        .expect(globalStateUpdateStub.getCall(1).args)
        .deep.equals([GlobalKey.AutoInstallDependency, true]);
      chai
        .expect(globalStateUpdateStub.getCall(2).args)
        .deep.equals([GlobalKey.OpenReadMe, "fakePath"]);
    });
  });
});
