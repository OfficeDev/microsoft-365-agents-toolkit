import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import fs from "fs-extra";
import * as projectStatusUtils from "../../src/utils/projectStatusUtils";
import { projectStatusUtilsDeps } from "../../src/utils/projectStatusUtils";
import { err, ok } from "@microsoft/teamsfx-api";
import * as helper from "../../src/chat/commands/nextstep/helper";
import * as glob from "glob";
import * as os from "os";
import * as path from "path";
import { UserCancelError } from "@microsoft/teamsfx-core";
import { vi } from "vitest";
import { createMock, mockValue } from "../mocks/vitestMockUtils";

chai.use(chaiPromised);

describe("project status utils", () => {
  describe("func: getProjectStatus", () => {
    it("project state file deos not exist", async () => {
      vi.spyOn(Date, "now").mockReturnValue(1711987200000);
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockResolvedValue(false);
      await chai
        .expect(projectStatusUtils.getProjectStatus("test-id"))
        .to.eventually.deep.equal(projectStatusUtils.emptyProjectStatus());
    });

    it("project state file exists - not a json file", async () => {
      vi.spyOn(Date, "now").mockReturnValue(1711987200000);
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockResolvedValue(false);
      vi.spyOn(projectStatusUtilsDeps, "readFile").mockResolvedValue(
        Buffer.from("not a json file") as any
      );
      await chai
        .expect(projectStatusUtils.getProjectStatus("test-id"))
        .to.eventually.deep.equal(projectStatusUtils.emptyProjectStatus());
    });

    it("project state file exists - a json file", async () => {
      vi.spyOn(Date, "now").mockReturnValue(1711987200000);
      const status = projectStatusUtils.emptyProjectStatus();
      status["fx-extension.provision"] = {
        result: "success",
        time: new Date(1711987200000 + 3600000),
      };
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockResolvedValue(true);
      vi.spyOn(projectStatusUtilsDeps, "readFile").mockResolvedValue(
        Buffer.from(JSON.stringify({ "test-id": status })) as any
      );
      await chai
        .expect(projectStatusUtils.getProjectStatus("test-id"))
        .to.eventually.deep.equal(status);
    });
  });

  describe("func: updateProjectStatus", () => {
    it("command name is not in RecordedActions", async () => {
      vi.spyOn(projectStatusUtilsDeps, "getProjectMetadata").mockReturnValue(undefined as any);
      await projectStatusUtils.updateProjectStatus("test-path", "test-command", ok(undefined));
    });

    it("command name is in RecordedActions - project state file not exist", async () => {
      vi.spyOn(projectStatusUtilsDeps, "getProjectMetadata").mockReturnValue({
        projectId: "test-id",
      } as any);
      vi.spyOn(Date, "now").mockReturnValue(1711987200000);
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockResolvedValue(false);
      const writeFileStub = vi.spyOn(projectStatusUtilsDeps, "writeFile").mockResolvedValue();
      await projectStatusUtils.updateProjectStatus(
        "test-path",
        projectStatusUtils.RecordedActions[0],
        ok(undefined)
      );
      chai.assert.equal(
        writeFileStub.getCall(0).args[1],
        JSON.stringify(
          {
            "test-id": {
              ...projectStatusUtils.emptyProjectStatus(),
              [projectStatusUtils.RecordedActions[0]]: {
                result: "success",
                time: new Date(1711987200000),
              },
            },
          },
          null,
          2
        )
      );
    });

    it("command name is not in RecordedActions but forced - not json", async () => {
      vi.spyOn(projectStatusUtilsDeps, "getProjectMetadata").mockReturnValue({
        projectId: "test-id",
      } as any);
      vi.spyOn(Date, "now").mockReturnValue(1711987200000);
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockImplementation(async (path: string) => {
        return path === projectStatusUtils.projectStatusFilePath;
      });
      vi.spyOn(projectStatusUtilsDeps, "readFile").mockResolvedValue(
        Buffer.from("not a json file") as any
      );
      const writeFileStub = vi.spyOn(projectStatusUtilsDeps, "writeFile").mockResolvedValue();
      await projectStatusUtils.updateProjectStatus(
        "test-path",
        "test-command",
        err(new UserCancelError()),
        true
      );
      chai.assert.equal(
        writeFileStub.getCall(0).args[1],
        JSON.stringify(
          {
            "test-id": {
              ...projectStatusUtils.emptyProjectStatus(),
              "test-command": {
                result: "fail",
                time: new Date(1711987200000),
              },
            },
          },
          null,
          2
        )
      );
    });

    it("command name is not in RecordedActions but forced - json", async () => {
      vi.spyOn(projectStatusUtilsDeps, "getProjectMetadata").mockReturnValue({
        projectId: "test-id",
      } as any);
      vi.spyOn(Date, "now").mockReturnValue(1711987200000);
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockImplementation(async (path: string) => {
        return path === projectStatusUtils.projectStatusFilePath;
      });
      vi.spyOn(projectStatusUtilsDeps, "readFile").mockResolvedValue(Buffer.from("{}") as any);
      const writeFileStub = vi.spyOn(projectStatusUtilsDeps, "writeFile").mockResolvedValue();
      await projectStatusUtils.updateProjectStatus(
        "test-path",
        "test-command",
        ok(undefined),
        true
      );
      chai.assert.equal(
        writeFileStub.getCall(0).args[1],
        JSON.stringify(
          {
            "test-id": {
              ...projectStatusUtils.emptyProjectStatus(),
              "test-command": {
                result: "success",
                time: new Date(1711987200000),
              },
            },
          },
          null,
          2
        )
      );
    });
  });

  it("func: getFileModifiedTime", async () => {
    const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "project-status-test-"));
    const file1 = path.join(tempDir, "test-file1.txt");
    const file2 = path.join(tempDir, "test-file2.txt");
    const latestTime = new Date(1711987200000);
    const oldTime = new Date(1711987200000 - 3600000);

    await fs.writeFile(file1, "test1");
    await fs.writeFile(file2, "test2");
    await fs.utimes(file1, latestTime, latestTime);
    await fs.utimes(file2, oldTime, oldTime);
    mockValue(projectStatusUtilsDeps, "glob", vi.fn().mockResolvedValue([file1, file2]) as any);

    const result = await projectStatusUtils.getFileModifiedTime(path.join(tempDir, "*.txt"));
    chai.expect(result.getTime()).equals(latestTime.getTime());

    await fs.remove(tempDir);
  });

  describe("func: getREADME", () => {
    it("file not exist", async () => {
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockResolvedValue(false);
      await chai.expect(projectStatusUtils.getREADME("test-folder")).to.eventually.equal(undefined);
    });

    it("file exists", async () => {
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockResolvedValue(true);
      vi.spyOn(projectStatusUtilsDeps, "readFile").mockResolvedValue(Buffer.from("123") as any);
      await chai
        .expect(projectStatusUtils.getREADME("test-folder"))
        .to.eventually.deep.equal(Buffer.from("123"));
    });
  });

  describe("func: getLaunchJSON", () => {
    it("file not exist", async () => {
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockResolvedValue(false);
      await chai
        .expect(projectStatusUtils.getLaunchJSON("test-folder"))
        .to.eventually.equal(undefined);
    });

    it("file exists", async () => {
      vi.spyOn(projectStatusUtilsDeps, "pathExists").mockResolvedValue(true);
      vi.spyOn(projectStatusUtilsDeps, "readFile").mockResolvedValue(Buffer.from("123") as any);
      await chai
        .expect(projectStatusUtils.getLaunchJSON("test-folder"))
        .to.eventually.deep.equal(Buffer.from("123"));
    });
  });
});
