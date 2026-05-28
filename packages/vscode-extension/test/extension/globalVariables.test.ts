import * as chai from "chai";
import fs from "fs-extra";
import * as sinon from "sinon";
import { ExtensionContext, Uri } from "vscode";

import * as globalVariables from "../../src/globalVariables";
import * as projectSettingHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import { err, ok, SystemError, TeamsAppManifest } from "@microsoft/teamsfx-api";
import { manifestUtils } from "@microsoft/teamsfx-core";

describe("Global Variables", () => {
  describe("isSPFxProject", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("return false for non-spfx project", async () => {
      sandbox.stub(fs, "existsSync").callsFake((path: fs.PathLike) => {
        return false;
      });
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(projectSettingHelper, "isValidProject").returns(true);
      sandbox.stub(globalVariables, "workspaceUri").returns({ fsPath: "/test" });
      sandbox.stub(fs, "readdirSync").returns(["package.json"] as any);

      globalVariables.initializeGlobalVariables({
        globalState: {
          get: () => undefined,
        },
        logUri: Uri.file("test"),
      } as unknown as ExtensionContext);

      chai.expect(globalVariables.isSPFxProject).equals(false);
    });

    it("return true for spfx project", () => {
      sandbox.stub(fs, "existsSync").callsFake((path: fs.PathLike) => {
        return false;
      });
      sandbox.stub(fs, "pathExistsSync").resolves(true);
      sandbox.stub(projectSettingHelper, "isValidProject").returns(true);
      sandbox.stub(projectSettingHelper, "isValidOfficeAddInProject").returns(false);
      sandbox.stub(globalVariables, "workspaceUri").value({ fsPath: "/test" });
      sandbox.stub(fs, "readdirSync").returns([".yo-rc.json"] as any);
      sandbox
        .stub(fs, "readJsonSync")
        .returns({ "@microsoft/generator-sharepoint": { version: " 1.16.0" } });

      globalVariables.initializeGlobalVariables({
        globalState: {
          get: () => undefined,
        },
        logUri: {
          fsPath: "",
        },
      } as unknown as ExtensionContext);

      chai.expect(globalVariables.isSPFxProject).equals(true);
    });

    it("set log folder", () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(fs, "mkdirSync").callsFake(() => {});
      globalVariables.initializeGlobalVariables({
        globalState: {
          get: () => undefined,
        },
        logUri: {
          fsPath: "fakePath",
        },
      } as unknown as ExtensionContext);
      chai.expect(globalVariables.defaultExtensionLogPath).equals("fakePath");
    });

    it("set commandIsRunning", async () => {
      globalVariables.setCommandIsRunning(true);

      chai.expect(globalVariables.commandIsRunning).equals(true);
    });

    it("unsetIsTeamsFxProject()", async () => {
      globalVariables.unsetIsTeamsFxProject();

      chai.expect(globalVariables.isTeamsFxProject).equals(false);
    });
  });

  describe("isDeclarativeCopilotApp", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Declarative copilot project", () => {
      const teamsManifest = new TeamsAppManifest();
      teamsManifest.copilotExtensions = {
        declarativeCopilots: [{ id: "1", file: "testFile" }],
      };
      sandbox.stub(manifestUtils, "readAppManifestSync").returns(ok(teamsManifest));

      const res = globalVariables.checkIsDeclarativeCopilotApp("projectPath");
      chai.expect(res).to.be.true;
    });

    it("Not declarative copilot project", () => {
      const teamsManifest = new TeamsAppManifest();
      sandbox.stub(manifestUtils, "readAppManifestSync").returns(ok(teamsManifest));

      const res = globalVariables.checkIsDeclarativeCopilotApp("projectPath");
      chai.expect(res).to.be.false;
    });

    it("Error: return false", () => {
      sandbox
        .stub(manifestUtils, "readAppManifestSync")
        .returns(err(new SystemError("error", "error", "error", "error")));

      const res = globalVariables.checkIsDeclarativeCopilotApp("projectPath");
      chai.expect(res).to.be.false;
    });
  });

  it("updateIsDeclarativeCopilotApp", () => {
    const manifest = new TeamsAppManifest();
    let res = globalVariables.updateIsDeclarativeCopilotApp(manifest);
    chai.assert.isFalse(res);

    res = globalVariables.updateIsDeclarativeCopilotApp({
      ...manifest,
      copilotExtensions: {
        declarativeCopilots: [
          {
            id: "1",
            file: "test",
          },
        ],
      },
    });
    chai.assert.isTrue(res);
  });
});
