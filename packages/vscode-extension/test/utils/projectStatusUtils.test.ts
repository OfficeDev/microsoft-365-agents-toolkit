import * as chai from "chai";
import chaiPromised from "chai-as-promised";
import fs from "fs-extra";
import * as sinon from "sinon";
import * as projectStatusUtils from "../../src/utils/projectStatusUtils";
import { projectStatusUtilsDeps } from "../../src/utils/projectStatusUtils";
import { err, ok } from "@microsoft/teamsfx-api";
import * as helper from "../../src/chat/commands/nextstep/helper";
import * as glob from "glob";
import { UserCancelError } from "@microsoft/teamsfx-core";

chai.use(chaiPromised);

describe("project status utils", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("func: getProjectStatus", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("project state file deos not exist", async () => {
      sandbox.stub(Date, "now").returns(1711987200000);
      sandbox.stub(projectStatusUtilsDeps, "pathExists").resolves(false);
      await chai
        .expect(projectStatusUtils.getProjectStatus("test-id"))
        .to.eventually.deep.equal(projectStatusUtils.emptyProjectStatus());
    });

    it("project state file exists - not a json file", async () => {
      sandbox.stub(Date, "now").returns(1711987200000);
      sandbox.stub(projectStatusUtilsDeps, "pathExists").resolves(false);
      sandbox
        .stub(projectStatusUtilsDeps, "readFile")
        .resolves(Buffer.from("not a json file") as any);
      await chai
        .expect(projectStatusUtils.getProjectStatus("test-id"))
        .to.eventually.deep.equal(projectStatusUtils.emptyProjectStatus());
    });

    it("project state file exists - a json file", async () => {
      sandbox.stub(Date, "now").returns(1711987200000);
      const status = projectStatusUtils.emptyProjectStatus();
      status["fx-extension.provision"] = {
        result: "success",
        time: new Date(1711987200000 + 3600000),
      };
      sandbox.stub(projectStatusUtilsDeps, "pathExists").resolves(true);
      sandbox
        .stub(projectStatusUtilsDeps, "readFile")
        .resolves(Buffer.from(JSON.stringify({ "test-id": status })) as any);
      await chai
        .expect(projectStatusUtils.getProjectStatus("test-id"))
        .to.eventually.deep.equal(status);
    });
  });

  describe("func: updateProjectStatus", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("command name is not in RecordedActions", async () => {
      sandbox.stub(projectStatusUtilsDeps, "getProjectMetadata").returns(undefined as any);
      await projectStatusUtils.updateProjectStatus("test-path", "test-command", ok(undefined));
    });

    it("command name is in RecordedActions - project state file not exist", async () => {
      sandbox
        .stub(projectStatusUtilsDeps, "getProjectMetadata")
        .returns({ projectId: "test-id" } as any);
      sandbox.stub(Date, "now").returns(1711987200000);
      sandbox.stub(projectStatusUtilsDeps, "pathExists").resolves(false);
      const writeFileStub = sandbox.stub(projectStatusUtilsDeps, "writeFile").resolves();
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
      sandbox
        .stub(projectStatusUtilsDeps, "getProjectMetadata")
        .returns({ projectId: "test-id" } as any);
      sandbox.stub(Date, "now").returns(1711987200000);
      sandbox.stub(projectStatusUtilsDeps, "pathExists").callsFake(async (path: string) => {
        return path === projectStatusUtils.projectStatusFilePath;
      });
      sandbox
        .stub(projectStatusUtilsDeps, "readFile")
        .resolves(Buffer.from("not a json file") as any);
      const writeFileStub = sandbox.stub(projectStatusUtilsDeps, "writeFile").resolves();
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
      sandbox
        .stub(projectStatusUtilsDeps, "getProjectMetadata")
        .returns({ projectId: "test-id" } as any);
      sandbox.stub(Date, "now").returns(1711987200000);
      sandbox.stub(projectStatusUtilsDeps, "pathExists").callsFake(async (path: string) => {
        return path === projectStatusUtils.projectStatusFilePath;
      });
      sandbox.stub(projectStatusUtilsDeps, "readFile").resolves(Buffer.from("{}") as any);
      const writeFileStub = sandbox.stub(projectStatusUtilsDeps, "writeFile").resolves();
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
    sandbox.stub(projectStatusUtilsDeps, "glob").resolves(["test-file1", "test-file2"] as any);
    const statInstance1 = sandbox.createStubInstance(fs.Stats);
    statInstance1.mtime = new Date(1711987200000);
    const statInstance2 = sandbox.createStubInstance(fs.Stats);
    statInstance2.mtime = new Date(1711987200000 - 3600000);
    sandbox.stub(projectStatusUtilsDeps, "stat").callsFake(async (path: fs.PathLike) => {
      if (path === "test-file1") {
        return statInstance1;
      } else {
        return statInstance2;
      }
    });
    await chai
      .expect(projectStatusUtils.getFileModifiedTime("test-pattern"))
      .to.eventually.deep.equal(new Date(1711987200000));
  });

  describe("func: getREADME", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("file not exist", async () => {
      sandbox.stub(projectStatusUtilsDeps, "pathExists").resolves(false);
      await chai.expect(projectStatusUtils.getREADME("test-folder")).to.eventually.equal(undefined);
    });

    it("file exists", async () => {
      sandbox.stub(projectStatusUtilsDeps, "pathExists").resolves(true);
      sandbox.stub(projectStatusUtilsDeps, "readFile").resolves(Buffer.from("123") as any);
      await chai
        .expect(projectStatusUtils.getREADME("test-folder"))
        .to.eventually.deep.equal(Buffer.from("123"));
    });
  });

  describe("func: getLaunchJSON", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("file not exist", async () => {
      sandbox.stub(projectStatusUtilsDeps, "pathExists").resolves(false);
      await chai
        .expect(projectStatusUtils.getLaunchJSON("test-folder"))
        .to.eventually.equal(undefined);
    });

    it("file exists", async () => {
      sandbox.stub(projectStatusUtilsDeps, "pathExists").resolves(true);
      sandbox.stub(projectStatusUtilsDeps, "readFile").resolves(Buffer.from("123") as any);
      await chai
        .expect(projectStatusUtils.getLaunchJSON("test-folder"))
        .to.eventually.deep.equal(Buffer.from("123"));
    });
  });
});
