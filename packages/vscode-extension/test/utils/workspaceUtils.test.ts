import * as chai from "chai";
import * as sinon from "sinon";
import * as globalState from "@microsoft/teamsfx-core";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { Uri, commands } from "vscode";
import { openOfficeDevFolder } from "../../src/utils/workspaceUtils";
import { GlobalKey } from "../../src/constants";

describe("WorkspaceUtils", () => {
  describe("openOfficeDevFolder", () => {
    const sandbox = sinon.createSandbox();
    let inMemoryGlobalState: Map<string, any>;

    beforeEach(async () => {
      inMemoryGlobalState = new Map<string, any>();
      sandbox
        .stub(globalState, "globalStateGet")
        .callsFake(async (key: string, defaultValue?: any) => {
          return inMemoryGlobalState.has(key) ? inMemoryGlobalState.get(key) : defaultValue;
        });
      sandbox.stub(globalState, "globalStateUpdate").callsFake(async (key: string, value: any) => {
        inMemoryGlobalState.set(key, value);
      });

      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(commands, "executeCommand");
      await globalState.globalStateUpdate(GlobalKey.OpenWalkThrough, false);
      await globalState.globalStateUpdate(GlobalKey.AutoInstallDependency, false);
      await globalState.globalStateUpdate(GlobalKey.OpenReadMe, "");
      await globalState.globalStateUpdate(GlobalKey.ShowLocalDebugMessage, false);
      await globalState.globalStateUpdate(GlobalKey.CreateWarnings, "");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("triggered from walkthrough with local debug message and warnings", async () => {
      const warnings = [{ type: "type", content: "content" }];
      await openOfficeDevFolder(Uri.parse("fakePath"), true, warnings, ["WalkThrough"]);
      chai.expect(await globalState.globalStateGet(GlobalKey.OpenWalkThrough, true)).equals(false);
      chai
        .expect(await globalState.globalStateGet(GlobalKey.AutoInstallDependency, false))
        .equals(true);
      chai.expect(await globalState.globalStateGet(GlobalKey.OpenReadMe, "")).equals("");
      chai
        .expect(await globalState.globalStateGet(GlobalKey.ShowLocalDebugMessage, false))
        .equals(true);
      chai
        .expect(await globalState.globalStateGet(GlobalKey.CreateWarnings, ""))
        .equals(JSON.stringify(warnings));
    });

    it("not triggered from walkthrough with no local debug message and warnings", async () => {
      await openOfficeDevFolder(Uri.parse("fakePath"), false, undefined);
      chai.expect(await globalState.globalStateGet(GlobalKey.OpenWalkThrough, true)).equals(false);
      chai
        .expect(await globalState.globalStateGet(GlobalKey.AutoInstallDependency, false))
        .equals(true);
      chai.expect(await globalState.globalStateGet(GlobalKey.OpenReadMe, "")).equals("fakePath");
    });
  });
});
