import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import * as status from "../../../../src/chat/commands/nextstep/status";
import { MachineStatus, WholeStatus } from "../../../../src/chat/commands/nextstep/types";
import * as projectStatusUtils from "../../../../src/utils/projectStatusUtils";
import { vi } from "vitest";

chai.use(chaiPromised);

describe("chat nextstep status", () => {
  afterEach(() => {
    // Restore the default sandbox here
    vi.restoreAllMocks();
  });

  describe("func: getWholeStatus", () => {
    it("folder === undefined", async () => {
      vi.spyOn(status.statusDeps, "checkCredential").mockResolvedValue({
        m365LoggedIn: true,
        azureLoggedIn: true,
      });
      vi.spyOn(status.statusDeps, "globalStateGet").mockResolvedValue(true);
      vi.spyOn(status.statusDeps, "globalStateUpdate");
      await chai.expect(status.getWholeStatus()).to.eventually.deep.equal({
        machineStatus: {
          azureLoggedIn: true,
          firstInstalled: true,
          m365LoggedIn: true,
        },
      } as WholeStatus);
    });

    it("folder !== undefined", async () => {
      vi.spyOn(status.statusDeps, "getProjectMetadata").mockReturnValue({ projectId: "test-id" });
      vi.spyOn(status.statusDeps, "getProjectStatus").mockResolvedValue(
        projectStatusUtils.emptyProjectStatus()
      );
      vi.spyOn(status.statusDeps, "getFileModifiedTime").mockResolvedValue(new Date(0));
      vi.spyOn(status.statusDeps, "getREADME").mockResolvedValue(undefined);
      vi.spyOn(status.statusDeps, "getLaunchJSON").mockResolvedValue(undefined);
      vi.spyOn(status.statusDeps, "checkCredential").mockResolvedValue({
        m365LoggedIn: true,
        azureLoggedIn: true,
      });
      vi.spyOn(status.statusDeps, "globalStateGet").mockResolvedValue(true);
      vi.spyOn(status.statusDeps, "globalStateUpdate");
      await chai.expect(status.getWholeStatus("test-folder")).to.eventually.deep.equal({
        machineStatus: {
          azureLoggedIn: true,
          firstInstalled: true,
          m365LoggedIn: true,
        },
        projectOpened: {
          path: "test-folder",
          projectId: "test-id",
          codeModifiedTime: {
            source: new Date(0),
            infra: new Date(0),
          },
          actionStatus: projectStatusUtils.emptyProjectStatus(),
          readmeContent: undefined,
          launchJSONContent: undefined,
        },
      } as WholeStatus);
    });

    it("folder !== undefined (no project id)", async () => {
      vi.spyOn(status.statusDeps, "getProjectMetadata").mockReturnValue(undefined);
      vi.spyOn(status.statusDeps, "getProjectStatus").mockResolvedValue(
        projectStatusUtils.emptyProjectStatus()
      );
      vi.spyOn(status.statusDeps, "getFileModifiedTime").mockResolvedValue(new Date(0));
      vi.spyOn(status.statusDeps, "getREADME").mockResolvedValue(undefined);
      vi.spyOn(status.statusDeps, "getLaunchJSON").mockResolvedValue(undefined);
      vi.spyOn(status.statusDeps, "checkCredential").mockResolvedValue({
        m365LoggedIn: true,
        azureLoggedIn: true,
      });
      vi.spyOn(status.statusDeps, "globalStateGet").mockResolvedValue(true);
      vi.spyOn(status.statusDeps, "globalStateUpdate");
      await chai.expect(status.getWholeStatus("test-folder")).to.eventually.deep.equal({
        machineStatus: {
          azureLoggedIn: true,
          firstInstalled: true,
          m365LoggedIn: true,
        },
        projectOpened: {
          path: "test-folder",
          projectId: undefined,
          codeModifiedTime: {
            source: new Date(0),
            infra: new Date(0),
          },
          actionStatus: projectStatusUtils.emptyProjectStatus(),
          readmeContent: undefined,
          launchJSONContent: undefined,
        },
      } as WholeStatus);
    });
  });

  describe("func: getMachineStatus", () => {
    it("func: getMachineStatus", async () => {
      vi.spyOn(status.statusDeps, "checkCredential").mockResolvedValue({
        m365LoggedIn: true,
        azureLoggedIn: true,
      });
      vi.spyOn(status.statusDeps, "globalStateGet").mockResolvedValue(true);
      vi.spyOn(status.statusDeps, "globalStateUpdate");
      await chai.expect(status.getMachineStatus()).to.eventually.deep.equal({
        azureLoggedIn: true,
        firstInstalled: true,
        m365LoggedIn: true,
      } as MachineStatus);
    });
  });
});
