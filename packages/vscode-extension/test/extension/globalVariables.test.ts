import * as chai from "chai";
import fs from "fs-extra";
import { ExtensionContext, Uri } from "vscode";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";

import * as globalVariables from "../../src/globalVariables";
import { err, ok, SystemError, TeamsAppManifest } from "@microsoft/teamsfx-api";
import { manifestUtils, copilotGptManifestUtils } from "@microsoft/teamsfx-core";

describe("Global Variables", () => {
  describe("isSPFxProject", () => {
    it("return false for non-spfx project", async () => {
      vi.spyOn(fs, "existsSync").mockImplementation((path: fs.PathLike) => {
        return false;
      });
      vi.spyOn(fs, "pathExistsSync").mockReturnValue(true);
      vi.spyOn(globalVariables.globalVariablesDeps, "isValidProject").mockReturnValue(true);
      vi.spyOn(globalVariables.globalVariablesDeps, "isValidOfficeAddInProject").mockReturnValue(
        false
      );
      mockValue(globalVariables, "workspaceUri", { fsPath: "/test" } as any);
      vi.spyOn(fs, "readdirSync").mockReturnValue(["package.json"] as any);

      globalVariables.initializeGlobalVariables({
        globalState: {
          get: () => undefined,
        },
        logUri: Uri.file("test"),
      } as unknown as ExtensionContext);

      chai.expect(globalVariables.isSPFxProject).equals(false);
    });

    it("return true for spfx project", () => {
      vi.spyOn(fs, "existsSync").mockImplementation((path: fs.PathLike) => {
        return true;
      });
      vi.spyOn(fs, "pathExistsSync").mockReturnValue(true);
      vi.spyOn(globalVariables.globalVariablesDeps, "isValidProject").mockReturnValue(false);
      vi.spyOn(globalVariables.globalVariablesDeps, "isValidOfficeAddInProject").mockReturnValue(
        false
      );

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
      vi.spyOn(fs, "pathExists").mockResolvedValue(false);
      vi.spyOn(fs, "mkdirSync").mockImplementation(() => {});
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
    it("Declarative copilot project", () => {
      const teamsManifest = new TeamsAppManifest();
      teamsManifest.copilotExtensions = {
        declarativeCopilots: [{ id: "1", file: "testFile" }],
      };
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok(teamsManifest));

      const res = globalVariables.checkIsDeclarativeCopilotApp("projectPath");
      chai.expect(res).to.be.true;
    });

    it("Not declarative copilot project", () => {
      const teamsManifest = new TeamsAppManifest();
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok(teamsManifest));

      const res = globalVariables.checkIsDeclarativeCopilotApp("projectPath");
      chai.expect(res).to.be.false;
    });

    it("Error: return false", () => {
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(
        err(new SystemError("error", "error", "error", "error"))
      );

      const res = globalVariables.checkIsDeclarativeCopilotApp("projectPath");
      chai.expect(res).to.be.false;
    });
  });

  describe("isMetaOSAddinProject", () => {
    it("MetaOS Addin Project: no directory", () => {
      const res = globalVariables.checkIsMetaOSAddinProject("");
      chai.expect(res).equals(false);
    });

    it("MetaOS Addin Project: manifest not ok", () => {
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(
        err(new SystemError("error", "error", "error", "error"))
      );
      const res = globalVariables.checkIsMetaOSAddinProject("abc");
      chai.expect(res).equals(false);
    });

    it("MetaOS Addin Project: manifest is undefined", () => {
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok(undefined as any));
      const res = globalVariables.checkIsMetaOSAddinProject("abc");
      chai.expect(res).equals(false);
    });

    it("MetaOS Addin Project: manifest is not metaOS", () => {
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok({} as any));
      const res = globalVariables.checkIsMetaOSAddinProject("abc");
      chai.expect(res).equals(false);
    });

    it("MetaOS Addin Project: manifest is ok", () => {
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok({ extensions: {} } as any));
      const res = globalVariables.checkIsMetaOSAddinProject("abc");
      chai.expect(res).equals(true);
    });
  });

  describe("checkIsSensitivityLabelSet", () => {
    const fakeDirectory = "fakeDir";

    it("returns true when sensitivity label is set", () => {
      const teamsManifest = new TeamsAppManifest();
      teamsManifest.copilotAgents = {
        declarativeAgents: [{ id: "test-id", file: "test.txt" }],
      };
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok(teamsManifest));
      vi.spyOn(copilotGptManifestUtils, "readDeclarativeAgentManifestFileSync").mockReturnValue(
        ok({
          name: "test-agent",
          description: "test description",
          sensitivity_label: { id: "test-label" },
        } as any)
      );

      const result = globalVariables.checkIsSensitivityLabelSet(fakeDirectory);
      chai.expect(result).to.be.true;
    });

    it("returns false when manifest read fails", () => {
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(
        err(new SystemError("test", "test", "test", "test"))
      );

      const result = globalVariables.checkIsSensitivityLabelSet(fakeDirectory);
      chai.expect(result).to.be.false;
    });

    it("returns false when manifest has no declarative agent path", () => {
      const teamsManifest = new TeamsAppManifest();
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok(teamsManifest));

      const result = globalVariables.checkIsSensitivityLabelSet(fakeDirectory);
      chai.expect(result).to.be.false;
    });

    it("returns false when declarative agent manifest read fails", () => {
      const teamsManifest = new TeamsAppManifest();
      teamsManifest.copilotAgents = {
        declarativeAgents: [{ id: "test-id", file: "test.txt" }],
      };
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok(teamsManifest));
      vi.spyOn(copilotGptManifestUtils, "readDeclarativeAgentManifestFileSync").mockReturnValue(
        err(new SystemError("test", "test", "test", "test"))
      );

      const result = globalVariables.checkIsSensitivityLabelSet(fakeDirectory);
      chai.expect(result).to.be.false;
    });

    it("returns false when declarative agent manifest has no sensitivity label", () => {
      const teamsManifest = new TeamsAppManifest();
      teamsManifest.copilotAgents = {
        declarativeAgents: [{ id: "test-id", file: "test.txt" }],
      };
      vi.spyOn(manifestUtils, "readAppManifestSync").mockReturnValue(ok(teamsManifest));
      vi.spyOn(copilotGptManifestUtils, "readDeclarativeAgentManifestFileSync").mockReturnValue(
        ok({
          name: "test-agent",
          description: "test description",
        } as any)
      );

      const result = globalVariables.checkIsSensitivityLabelSet(fakeDirectory);
      chai.expect(result).to.be.false;
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
